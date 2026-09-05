"""Offline daily-refresh server and history store.

This file is reusable core.  A generated project normally changes only its
PROJECT.json, optional CUSTOM_RULES.py, and dashboard labels/configuration.
"""

import csv
import hashlib
import io
import json
import os
import shutil
import sqlite3
import statistics
import tempfile
import threading
import time
import urllib.parse
import webbrowser
import zlib
from collections import defaultdict
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import calculation_engine as calc

ROOT = Path(__file__).resolve().parent
PROJECT_PATH = Path(os.environ.get("EXCEL_APP_PROJECT", ROOT / "PROJECT.json"))
DATA_DIR = Path(os.environ.get("EXCEL_APP_DATA_DIR", ROOT / "data"))
DB_PATH, STATE_PATH, UPLOAD_DIR = DATA_DIR / "history.db", DATA_DIR / "last_result.json", DATA_DIR / "uploads"
MAX_UPLOAD = 150 * 1024 * 1024
LOCK, ROW_CACHE = threading.RLock(), None


def _runtime_log(message):
    """Keep detailed diagnostics locally without exposing a traceback in the UI."""
    try:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        with (DATA_DIR / "app.log").open("a", encoding="utf-8") as handle:
            handle.write(f"{datetime.now().astimezone().isoformat(timespec='seconds')} | {message}\n")
    except Exception:
        pass


def friendly_error(exc):
    """Explain the common failure modes in business language."""
    text = str(exc).strip() or type(exc).__name__
    lower = text.casefold()
    if isinstance(exc, FileNotFoundError) or "not found" in lower and "sheet" not in lower:
        return f"What happened: the source file could not be found.\nWhy: it may have been moved, renamed, or the upload did not finish.\nHow to fix: choose the current Excel/CSV file from the Upload button and try again.\nDetails: {text}"
    if "required columns" in lower or "header row" in lower:
        return f"What happened: the workbook structure does not match this project.\nWhy: a required column or header row is missing, renamed, or on another sheet.\nHow to fix: upload the correct file, or ask the project AI to update column_aliases, sheet_name, or header_row in PROJECT.json.\nDetails: {text}"
    if "sheet" in lower and ("not found" in lower or "no visible" in lower):
        return f"What happened: the expected worksheet was not found.\nWhy: the sheet name changed, is hidden, or the workbook is empty.\nHow to fix: ask the project AI to use sheet_name: auto or the exact sheet name, then upload again.\nDetails: {text}"
    if "no valid rows" in lower or "empty" in lower:
        return f"What happened: no usable data rows were accepted.\nWhy: rows may be blank, invalid, or all failed the configured critical fields.\nHow to fix: check the rejected-row details and upload a file with at least one valid data row.\nDetails: {text}"
    if "excel fallback" in lower or "excel could not" in lower or "windows only" in lower:
        return f"What happened: the workbook needs desktop Excel to be read.\nWhy: it is protected, legacy, binary, encrypted, or uses a feature the normal reader cannot open.\nHow to fix: use this project on Windows with Microsoft Excel installed, approve the normal company prompt, and keep Excel available while it reads the file.\nDetails: {text}"
    if "permission" in lower or "access is denied" in lower:
        return f"What happened: Windows denied access to the file or local data folder.\nWhy: the workbook may be open exclusively, protected by policy, or the project folder is read-only.\nHow to fix: close the source workbook, extract the project to a writable folder, and try again.\nDetails: {text}"
    if "port" in lower or "socket" in lower:
        return f"What happened: the local dashboard could not open a browser port.\nWhy: another program or a security policy blocked local sockets.\nHow to fix: close an older copy, run START.bat again, or ask IT to allow localhost for this application.\nDetails: {text}"
    return f"What happened: the project could not finish this operation.\nWhy: the source format, configuration, or calculation may not match the agreed logic.\nHow to fix: check the project log and the source workbook, then ask the project AI to make a versioned configuration change.\nDetails: {text}"


def config(): return calc.load_config(PROJECT_PATH)


def init_db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_PATH) as connection:
        connection.execute("CREATE TABLE IF NOT EXISTS current_rows(row_key TEXT PRIMARY KEY,row_json TEXT NOT NULL)")
        existing = {row[1] for row in connection.execute("PRAGMA table_info(runs)")}
        required = {"id", "created_at", "source_name", "file_hash", "mode", "input_count", "accepted_count", "rejected_count", "current_count", "result_json", "history_json", "period_date", "snapshot_blob"}
        if existing and not required.issubset(existing):
            legacy = f"legacy_runs_{datetime.now():%Y%m%d%H%M%S}"
            connection.execute(f"ALTER TABLE runs RENAME TO {legacy}")
            if connection.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='rejected_rows'").fetchone():
                connection.execute(f"ALTER TABLE rejected_rows RENAME TO legacy_rejected_rows_{datetime.now():%Y%m%d%H%M%S}")
        connection.executescript("""
        CREATE TABLE IF NOT EXISTS runs(
          id INTEGER PRIMARY KEY AUTOINCREMENT,created_at TEXT NOT NULL,source_name TEXT NOT NULL,
          file_hash TEXT NOT NULL UNIQUE,mode TEXT NOT NULL,input_count INTEGER NOT NULL,
          accepted_count INTEGER NOT NULL,rejected_count INTEGER NOT NULL,current_count INTEGER NOT NULL,
          result_json TEXT NOT NULL,history_json TEXT NOT NULL,period_date TEXT NOT NULL,snapshot_blob BLOB);
        CREATE TABLE IF NOT EXISTS rejected_rows(run_id INTEGER,row_number INTEGER,reason TEXT,row_json TEXT);
        """)
        connection.commit()


def row_key(row, fields):
    if fields:
        values = [row.get(field) for field in fields]
        if any(value in (None, "") for value in values): return None
        return json.dumps(values, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(json.dumps(row, ensure_ascii=False, sort_keys=True, default=str).encode()).hexdigest()


def public_row(row): return {key: value for key, value in row.items() if not str(key).startswith("_")}


def deduplicate(rows, fields, policy="keep_latest"):
    if policy not in {"keep_latest", "keep_first"}: raise ValueError("duplicate_policy must be keep_latest or keep_first")
    unique, rejected = {}, []
    for index, row in enumerate(rows, start=1):
        source_row = row.get("_source_row", index)
        key = row_key(row, fields)
        if key is None:
            rejected.append({"row_number": source_row, "reason": "Missing business key", "row": public_row(row)}); continue
        if key in unique:
            if policy == "keep_first":
                rejected.append({"row_number": source_row, "reason": "Duplicate business key; first row kept", "row": public_row(row)})
                continue
            previous = unique[key]
            rejected.append({"row_number": previous.get("_source_row", index - 1), "reason": "Duplicate business key; latest row kept", "row": public_row(previous)})
        unique[key] = row
    return [public_row(row) for row in unique.values()], rejected


def current_rows(force=False):
    global ROW_CACHE
    with LOCK:
        if ROW_CACHE is not None and not force: return ROW_CACHE
        init_db()
        with sqlite3.connect(DB_PATH) as connection: ROW_CACHE = [json.loads(item[0]) for item in connection.execute("SELECT row_json FROM current_rows")]
        return ROW_CACHE


def previous_result():
    init_db()
    with sqlite3.connect(DB_PATH) as connection: row = connection.execute("SELECT result_json FROM runs ORDER BY id DESC LIMIT 1").fetchone()
    return json.loads(row[0]) if row else None


def persist(rows, rejected, result, summary, source_name, file_hash, mode, input_count, accepted_count, rejected_count, period_date, cfg):
    global ROW_CACHE
    init_db(); now = datetime.now().isoformat(timespec="seconds"); keys = cfg.get("refresh", {}).get("business_keys", [])
    snapshot = zlib.compress(json.dumps(rows, ensure_ascii=False, default=str).encode("utf-8"), 6)
    with sqlite3.connect(DB_PATH) as connection:
        connection.execute("BEGIN"); connection.execute("DELETE FROM current_rows")
        records = []
        for row in rows:
            key = row_key(row, keys) or hashlib.sha256(json.dumps(row, sort_keys=True, default=str).encode()).hexdigest()
            records.append((key, json.dumps(row, ensure_ascii=False, default=str)))
        connection.executemany("INSERT INTO current_rows(row_key,row_json) VALUES(?,?)", records)
        cursor = connection.execute("INSERT INTO runs(created_at,source_name,file_hash,mode,input_count,accepted_count,rejected_count,current_count,result_json,history_json,period_date,snapshot_blob) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)", (now, source_name, file_hash, mode, input_count, accepted_count, rejected_count, len(rows), "{}", json.dumps(summary, ensure_ascii=False, default=str), period_date, snapshot))
        run_id = cursor.lastrowid
        result.update({"run_id": run_id, "updated_at": now, "source_name": source_name, "refresh_mode": mode, "input_count": input_count, "accepted_count": accepted_count, "rejected_count": rejected_count, "supporting_rejected_count": max(0, len(rejected) - rejected_count)})
        connection.execute("UPDATE runs SET result_json=? WHERE id=?", (json.dumps(result, ensure_ascii=False, default=str), run_id))
        connection.executemany("INSERT INTO rejected_rows(run_id,row_number,reason,row_json) VALUES(?,?,?,?)", [(run_id, item.get("row_number"), item.get("reason"), json.dumps(item.get("row", {}), ensure_ascii=False, default=str)) for item in rejected])
        keep = max(1, int(cfg.get("refresh", {}).get("keep_rollback_snapshots", 3)))
        connection.execute("UPDATE runs SET snapshot_blob=NULL WHERE id NOT IN(SELECT id FROM runs ORDER BY id DESC LIMIT ?)", (keep,))
        history = cfg.get("history", {})
        if not history.get("keep_forever", True):
            keep_uploads = max(1, int(history.get("keep_uploads", 365)))
            old_ids = [row[0] for row in connection.execute("SELECT id FROM runs ORDER BY id DESC LIMIT -1 OFFSET ?", (keep_uploads,))]
            if old_ids:
                placeholders = ",".join("?" for _ in old_ids)
                connection.execute(f"DELETE FROM rejected_rows WHERE run_id IN ({placeholders})", old_ids)
                connection.execute(f"DELETE FROM runs WHERE id IN ({placeholders})", old_ids)
        connection.commit()
    temp = STATE_PATH.with_suffix(".tmp"); temp.write_text(json.dumps(result, ensure_ascii=False, default=str), encoding="utf-8"); temp.replace(STATE_PATH)
    ROW_CACHE = rows
    return result


def archive(path, source_name, cfg):
    keep = int(cfg.get("refresh", {}).get("keep_source_files", 3))
    if keep <= 0: return
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, UPLOAD_DIR / f"{datetime.now():%Y%m%d-%H%M%S-%f}-{Path(source_name).name}")
    files = sorted((item for item in UPLOAD_DIR.iterdir() if item.is_file()), key=lambda item: item.stat().st_mtime, reverse=True)
    for old in files[keep:]: old.unlink(missing_ok=True)


def file_digest(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""): digest.update(chunk)
    return digest.hexdigest()


def reporting_date(rows, cfg):
    history = cfg.get("history", {})
    if history.get("period_source", "data") == "data":
        field = history.get("date_field", "date")
        values = sorted(str(row.get(field))[:10] for row in rows if row.get(field))
        if values: return values[-1]
    return datetime.now().date().isoformat()


def process_file(path, source_name=None):
    source_name, cfg = source_name or Path(path).name, config()
    source_hash = file_digest(path)
    file_hash = hashlib.sha256((source_hash + hashlib.sha256(PROJECT_PATH.read_bytes()).hexdigest()).encode()).hexdigest()
    init_db()
    with LOCK, sqlite3.connect(DB_PATH) as connection:
        if connection.execute("SELECT 1 FROM runs WHERE file_hash=?", (file_hash,)).fetchone():
            state = state_now(); state.update({"duplicate_upload": True, "notice": f"{source_name} was already processed. No data was duplicated."}); return state
    outputs, reader, file_probe = calc.read_sources(path, cfg)
    source_rows, rejected, warnings, mappings, missing_errors = {}, [], [], {}, []
    primary_id = cfg["sources"][0].get("id", "main")
    primary_input_count = 0
    for source_id, output in outputs.items():
        accepted, bad, missing, mapping, notes = output
        source_rows[source_id] = accepted; rejected.extend({**item, "source": source_id} for item in bad); mappings[source_id] = mapping; warnings.extend(notes)
        if source_id == primary_id: primary_input_count = len(accepted) + len(bad)
        if missing: missing_errors.append(f"{source_id}: {', '.join(missing)}")
    if missing_errors:
        raise ValueError("Required columns are missing (" + "; ".join(missing_errors) + "). The previous good result was kept.")
    joined = calc.join_sources(source_rows, cfg)
    warnings.extend(cfg.pop("_runtime_warnings", []))
    calculated, bad, notes = calc.apply_logic(joined, cfg); rejected.extend({**item, "source": primary_id} for item in bad); warnings.extend(notes)
    keys = cfg.get("refresh", {}).get("business_keys", [])
    duplicate_policy = cfg.get("refresh", {}).get("duplicate_policy", "keep_latest")
    calculated, bad = deduplicate(calculated, keys, duplicate_policy); rejected.extend({**item, "source": primary_id} for item in bad)
    if not calculated and not cfg.get("refresh", {}).get("allow_empty_replace", False):
        raise ValueError("No valid rows remained after validation. The previous good result was kept.")
    mode = cfg.get("refresh", {}).get("mode", "replace").lower()
    if mode not in {"replace", "merge"}: raise ValueError("refresh.mode must be replace or merge")
    if mode == "merge":
        merged = {row_key(row, keys): row for row in current_rows() if row_key(row, keys)}
        for row in calculated: merged[row_key(row, keys)] = row
        rows = list(merged.values())
    else: rows = calculated
    try: archive(path, source_name, cfg)
    except Exception as exc: warnings.append(f"Source-file archive was skipped: {exc}")
    result = calc.dashboard(rows, cfg, previous_result())
    summary = calc.history_summary(rows, cfg) if cfg.get("history", {}).get("enabled", True) else {"overall": {}, "dimensions": {}}
    primary_rejected_count = sum(1 for item in rejected if item.get("source") == primary_id)
    reconciliation = {"source_rows": primary_input_count, "accepted_rows": len(calculated), "rejected_rows": primary_rejected_count}
    reconciliation["balanced"] = reconciliation["source_rows"] == reconciliation["accepted_rows"] + reconciliation["rejected_rows"]
    if not reconciliation["balanced"]: raise RuntimeError("Row reconciliation failed. The previous good result was kept.")
    result.update({"reader": reader, "file_probe": file_probe, "mapping": mappings, "reconciliation": reconciliation, "warnings": sorted(set(warnings)), "rejected_preview": rejected[:20]})
    persist(rows, rejected, result, summary, source_name, file_hash, mode, primary_input_count, len(calculated), primary_rejected_count, reporting_date(calculated, cfg), cfg)
    _runtime_log(f"UPLOAD OK | file={source_name} | run_id={result.get('run_id')} | rows={len(rows)} | rejected={len(rejected)} | reader={reader}")
    return result


def state_now():
    if STATE_PATH.exists():
        try: return json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except Exception: pass
    rows, cfg = current_rows(), config(); result = calc.dashboard(rows, cfg)
    result.update({"warnings": [], "rejected_count": 0, "accepted_count": len(rows), "notice": "Upload today's Excel or CSV file to begin."})
    return result


def query(filters):
    old, result = state_now(), calc.dashboard(current_rows(), config(), filters=filters)
    for field in ("run_id", "updated_at", "source_name", "reader", "file_probe", "reconciliation", "refresh_mode", "input_count", "accepted_count", "rejected_count", "supporting_rejected_count", "warnings"):
        if field in old: result[field] = old[field]
    return result


def rollback(run_id):
    cfg = config()
    with LOCK, sqlite3.connect(DB_PATH) as connection: row = connection.execute("SELECT source_name,snapshot_blob FROM runs WHERE id=?", (run_id,)).fetchone()
    if not row: raise ValueError("History item not found")
    if not row[1]: raise ValueError("This older detailed snapshot is no longer retained")
    rows = json.loads(zlib.decompress(row[1]).decode("utf-8")); result = calc.dashboard(rows, cfg, previous_result()); summary = calc.history_summary(rows, cfg)
    result.update({"reader": "rollback", "mapping": {}, "warnings": [f"Restored from run {run_id}"], "rejected_preview": []})
    return persist(rows, [], result, summary, f"Rollback of {row[0]}", f"rollback:{run_id}:{time.time_ns()}", "rollback", len(rows), len(rows), 0, reporting_date(rows, cfg), cfg)


def runs():
    init_db()
    with sqlite3.connect(DB_PATH) as connection: records = connection.execute("SELECT id,created_at,source_name,mode,input_count,accepted_count,rejected_count,current_count,snapshot_blob IS NOT NULL FROM runs ORDER BY id DESC LIMIT 50").fetchall()
    return [{"id": row[0], "created_at": row[1], "source_name": row[2], "mode": row[3], "input_count": row[4], "accepted_count": row[5], "rejected_count": row[6], "current_count": row[7], "can_rollback": bool(row[8])} for row in records]


def period_key(timestamp, period):
    value = datetime.fromisoformat(timestamp)
    if period == "monthly": return value.strftime("%Y-%m")
    if period == "weekly":
        year, week, _ = value.isocalendar(); return f"{year}-W{week:02d}"
    return value.strftime("%Y-%m-%d")


def history_analysis(request):
    cfg = config(); history_cfg = cfg.get("history", {})
    if not history_cfg.get("enabled", True): return {"metric": None, "period": None, "points": [], "statistics": {}, "insight": "History is disabled for this project.", "metrics": [], "dimensions": {}, "periods": []}
    periods = [item for item in history_cfg.get("periods", ["daily", "weekly", "monthly"]) if item in {"daily", "weekly", "monthly"}] or ["daily"]
    metric = request.get("metric") or cfg.get("kpis", [{}])[0].get("id"); period = request.get("period") if request.get("period") in periods else periods[0]; dimension = request.get("dimension"); category = request.get("category")
    metric_spec = next((item for item in cfg.get("kpis", []) if item.get("id") == metric), {})
    aggregation = metric_spec.get("period_aggregation", history_cfg.get("period_aggregation", "latest"))
    with sqlite3.connect(DB_PATH) as connection: records = connection.execute("SELECT id,period_date,source_name,history_json FROM runs WHERE mode!='rollback' ORDER BY id").fetchall()
    buckets, dimensions = defaultdict(list), defaultdict(set)
    for run_id, period_date, source_name, text in records:
        summary = json.loads(text)
        for field, values in summary.get("dimensions", {}).items(): dimensions[field].update(values)
        value = summary.get("overall", {}).get(metric)
        if dimension and category: value = summary.get("dimensions", {}).get(dimension, {}).get(category, {}).get(metric)
        if value is None: continue
        label = period_key(period_date, period)
        buckets[label].append({"run_id": run_id, "label": label, "value": float(value), "source_name": source_name})
    points = []
    for label, items in buckets.items():
        bucket_values = [item["value"] for item in items]
        if aggregation == "sum": value = sum(bucket_values)
        elif aggregation == "average": value = statistics.mean(bucket_values)
        elif aggregation == "minimum": value = min(bucket_values)
        elif aggregation == "maximum": value = max(bucket_values)
        else: value = bucket_values[-1]
        points.append({**items[-1], "value": round(value, 4), "upload_count": len(items)})
    values = [item["value"] for item in points]
    moving = max(1, int(cfg.get("history", {}).get("moving_average_periods", 3)))
    for index, item in enumerate(points): item["moving_average"] = round(statistics.mean(values[max(0, index - moving + 1):index + 1]), 4)
    if values:
        average, median = statistics.mean(values), statistics.median(values); deviation = statistics.pstdev(values) if len(values) > 1 else 0
        slope = sum((i - (len(values)-1)/2) * (value - average) for i, value in enumerate(values)) / max(sum((i - (len(values)-1)/2) ** 2 for i in range(len(values))), 1)
        latest, previous = values[-1], values[-2] if len(values) > 1 else values[-1]
        stats = {"latest": latest, "previous": previous, "change": latest - previous, "change_percent": calc.safe_div(latest - previous, abs(previous)) * 100 if previous else None, "average": average, "median": median, "minimum": min(values), "maximum": max(values), "std_deviation": deviation, "trend": "up" if slope > 0 else "down" if slope < 0 else "stable", "slope": slope, "latest_outlier": abs(latest - average) > 2 * deviation if deviation else False}
        insight = f"Latest value is {latest:,.2f}; trend is {stats['trend']}. It changed {stats['change_percent']:,.1f}% from the previous period." if stats["change_percent"] is not None else f"Latest value is {latest:,.2f}."
        if stats["latest_outlier"]: insight += " The latest value is statistically unusual versus the available history."
    else: stats, insight = {}, "No historical points match the selection."
    return {"metric": metric, "period": period, "aggregation": aggregation, "dimension": dimension, "category": category, "points": points, "statistics": stats, "insight": insight, "metrics": [{"id": item["id"], "label": item["label"], "format": item.get("format")} for item in cfg.get("kpis", [])], "dimensions": {field: sorted(values) for field, values in dimensions.items()}, "periods": periods}


def csv_bytes(rows, columns):
    output = io.StringIO(newline=""); writer = csv.DictWriter(output, fieldnames=columns, extrasaction="ignore"); writer.writeheader(); writer.writerows(rows); return output.getvalue().encode("utf-8-sig")


class Handler(BaseHTTPRequestHandler):
    def send_json(self, payload, status=200):
        body = json.dumps(payload, ensure_ascii=False, default=str).encode(); self.send_response(status); self.send_header("Content-Type", "application/json; charset=utf-8"); self.send_header("Content-Length", str(len(body))); self.end_headers(); self.wfile.write(body)
    def read_json(self):
        length = int(self.headers.get("Content-Length", "0")); return json.loads(self.rfile.read(length).decode()) if length else {}
    def send_csv(self, body, name):
        self.send_response(200); self.send_header("Content-Type", "text/csv; charset=utf-8"); self.send_header("Content-Disposition", f"attachment; filename={name}"); self.send_header("Content-Length", str(len(body))); self.end_headers(); self.wfile.write(body)
    def do_GET(self):
        route = urllib.parse.urlparse(self.path).path
        if route == "/api/state": return self.send_json(state_now())
        if route == "/api/runs": return self.send_json(runs())
        if route == "/api/rejected.csv":
            run_id = int(urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query).get("run_id", [0])[0])
            with sqlite3.connect(DB_PATH) as connection: records = connection.execute("SELECT row_number,reason,row_json FROM rejected_rows WHERE run_id=?", (run_id,)).fetchall()
            rows = [{"row_number": row[0], "reason": row[1], **json.loads(row[2])} for row in records]; return self.send_csv(csv_bytes(rows, list(rows[0]) if rows else ["row_number", "reason"]), "rejected_rows.csv")
        if route in ("/", "/dashboard.html"):
            body = (ROOT / "dashboard.html").read_bytes(); self.send_response(200); self.send_header("Content-Type", "text/html; charset=utf-8"); self.send_header("Content-Length", str(len(body))); self.end_headers(); return self.wfile.write(body)
        self.send_error(404)
    def do_POST(self):
        route = urllib.parse.urlparse(self.path).path
        try:
            if route == "/api/query": return self.send_json(query(self.read_json().get("filters", {})))
            if route == "/api/history": return self.send_json(history_analysis(self.read_json()))
            if route == "/api/rollback": return self.send_json(rollback(int(self.read_json().get("run_id"))))
            if route == "/api/export":
                filters = self.read_json().get("filters", {}); rows = calc.filter_rows(current_rows(), filters); columns = config().get("table", {}).get("columns", []) or (list(rows[0]) if rows else []); return self.send_csv(csv_bytes(rows, columns), "filtered_report.csv")
            if route != "/api/upload": return self.send_error(404)
            length = int(self.headers.get("Content-Length", "0"))
            if not 0 < length <= MAX_UPLOAD: return self.send_json({"error": "File is empty or exceeds 150 MB"}, 400)
            filename = Path(urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query).get("filename", ["upload.xlsx"])[0]).name; suffix = Path(filename).suffix.lower()
            if suffix not in (".xlsx", ".xlsm", ".xls", ".xlsb", ".csv"): return self.send_json({"error": "Use XLSX, XLSM, XLS, XLSB or CSV"}, 400)
            safe_name = filename or f"upload{suffix}"
            temp_dir = tempfile.TemporaryDirectory(prefix="excel_app_")
            temp_path = Path(temp_dir.name) / safe_name
            temp = None
            try:
                temp = temp_path.open("wb")
                remaining = length
                while remaining:
                    chunk = self.rfile.read(min(1024 * 1024, remaining))
                    if not chunk: break
                    temp.write(chunk); remaining -= len(chunk)
                temp.close()
                if remaining: return self.send_json({"error": "The upload ended before the complete file was received"}, 400)
                return self.send_json(process_file(temp_path, filename))
            finally:
                try:
                    if temp: temp.close()
                except Exception: pass
                try: temp_dir.cleanup()
                except Exception: pass
        except Exception as exc:
            _runtime_log(f"ERROR | route={route} | type={type(exc).__name__} | message={exc}")
            return self.send_json({"error": friendly_error(exc)}, 400)
    def log_message(self, fmt, *args): return


def _configured_port():
    value = os.environ.get("EXCEL_APP_PORT", "").strip()
    if not value: return 0
    try:
        port = int(value)
        if 0 <= port <= 65535: return port
    except ValueError: pass
    print(f"Ignoring invalid EXCEL_APP_PORT={value!r}; selecting a free port.")
    return 0


def _create_server(port=None):
    requested = _configured_port() if port is None else int(port)
    try:
        return ThreadingHTTPServer(("127.0.0.1", requested), Handler)
    except OSError as first_error:
        if requested == 0:
            raise RuntimeError(f"Could not open a local port: {first_error}") from first_error
        print(f"Port {requested} is unavailable; selecting a free local port.")
        try:
            return ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        except OSError as fallback_error:
            raise RuntimeError(f"Could not open a local port: {fallback_error}") from fallback_error


def run_server(port=None, open_browser=True):
    init_db(); server = _create_server(port); actual_port = server.server_address[1]; url = f"http://127.0.0.1:{actual_port}"; _runtime_log(f"START | url={url}"); print(f"Excel app is running: {url}"); print("Press Ctrl+C to stop.")
    if open_browser: threading.Thread(target=lambda: (time.sleep(.8), webbrowser.open(url)), daemon=True).start()
    server.serve_forever()


if __name__ == "__main__":
    try:
        run_server()
    except RuntimeError as exc:
        _runtime_log(f"START ERROR | type={type(exc).__name__} | message={exc}")
        print(f"Excel app could not start: {exc}")
        raise SystemExit(1)
