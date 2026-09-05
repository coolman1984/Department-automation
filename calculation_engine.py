"""Small, configuration-driven Excel engine shared by every generated project.

The business assistant normally changes PROJECT.json (and, only when needed,
CUSTOM_RULES.py); the reusable reader, validator, history and dashboard
contract stays stable.
"""

import ast
import csv
import importlib.util
import json
import math
import os
import re
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
for dependency_path in (ROOT / "vendor", ROOT / "vendor.zip"):
    if dependency_path.exists(): sys.path.insert(0, str(dependency_path))


def deep_merge(base, override):
    result = dict(base or {})
    for key, value in (override or {}).items():
        result[key] = deep_merge(result.get(key, {}), value) if isinstance(value, dict) and isinstance(result.get(key), dict) else value
    return result


def load_config(project_path):
    raw = json.loads(Path(project_path).read_text(encoding="utf-8"))
    packs = json.loads((ROOT / "department_packs.json").read_text(encoding="utf-8"))
    pack_name = raw.get("project", {}).get("department_pack", "generic")
    if pack_name not in packs:
        raise ValueError(f"Unknown department pack: {pack_name}")
    pack = packs[pack_name]
    config = {
        "project": raw.get("project", {}), "refresh": raw.get("refresh", {}), "validation": raw.get("validation", {}), "history": raw.get("history", {}), "excel": raw.get("excel", {}), "joins": raw.get("joins", []),
        "derived_fields": pack.get("derived_fields", []), "filters": pack.get("filters", []),
        "kpis": pack.get("kpis", []), "charts": pack.get("charts", []),
        "table": pack.get("table", {}), "rules": pack.get("rules", []), "pack_name": pack_name
    }
    defaults = pack.get("source_defaults", {})
    config["sources"] = [deep_merge(defaults, item) for item in raw.get("sources", [{"id": "main", "sheet_name": "Data"}])]
    config["auxiliary_sources"] = [deep_merge(defaults, item) for item in raw.get("auxiliary_sources", [])]
    config["auxiliary_joins"] = raw.get("auxiliary_joins", [])
    config["auxiliary_date_range_joins"] = raw.get("auxiliary_date_range_joins", [])
    config = deep_merge(config, raw.get("overrides", {}))
    validate_config(config)
    return config


def validate_config(config):
    sources = config.get("sources", [])
    if not sources: raise ValueError("At least one data source is required")
    source_ids = [source.get("id") for source in sources]
    if any(not item for item in source_ids) or len(source_ids) != len(set(source_ids)): raise ValueError("Every source needs a unique id")
    allowed_modes = {"auto", "excel", "excel_open", "excel_attach"}
    for source in sources:
        if source.get("read_mode", "auto") not in allowed_modes: raise ValueError(f"Unsupported read_mode for {source['id']}")
        if source.get("date_order", "DMY").upper() not in {"DMY", "MDY", "YMD"}: raise ValueError(f"date_order must be DMY, MDY or YMD for {source['id']}")
        header = source.get("header_row", 1)
        if str(header).lower() != "auto":
            try:
                if int(header) < 1: raise ValueError
            except (TypeError, ValueError):
                raise ValueError(f"header_row must be a positive number or auto for {source['id']}")
    if config.get("refresh", {}).get("mode", "replace").lower() == "merge" and not config.get("refresh", {}).get("business_keys"):
        raise ValueError("Merge mode requires at least one stable business key")
    kpi_ids = [item.get("id") for item in config.get("kpis", [])]
    if any(not item for item in kpi_ids) or len(kpi_ids) != len(set(kpi_ids)): raise ValueError("Every KPI needs a unique id")
    validation = config.get("validation", {})
    if validation.get("mode", "flexible") not in {"flexible", "strict"}: raise ValueError("validation.mode must be flexible or strict")
    if validation.get("duplicate_source_columns", "first_non_blank") not in {"first_non_blank", "block"}: raise ValueError("duplicate_source_columns must be first_non_blank or block")
    if validation.get("duplicate_join_keys", "keep_latest") not in {"keep_latest", "keep_first", "block"}: raise ValueError("duplicate_join_keys must be keep_latest, keep_first or block")
    if config.get("history", {}).get("period_aggregation", "latest") not in {"latest", "average", "sum", "minimum", "maximum"}: raise ValueError("history.period_aggregation is not supported")
    validate_auxiliary_sources(config)


def validate_auxiliary_sources(config):
    """Optional multi-file HR sources (INT-01): employee, roster, leave. Empty by default, so V1.0 attendance-only projects are unaffected."""
    sources = config.get("auxiliary_sources", [])
    ids = [source.get("id") for source in sources]
    if any(not item for item in ids) or len(ids) != len(set(ids)): raise ValueError("Every auxiliary source needs a unique id")
    known_ids = {source.get("id", "main") for source in config.get("sources", [])} | set(ids)
    for source in sources:
        header = source.get("header_row", 1)
        if str(header).lower() != "auto":
            try:
                if int(header) < 1: raise ValueError
            except (TypeError, ValueError):
                raise ValueError(f"header_row must be a positive number or auto for auxiliary source {source['id']}")
    for join in config.get("auxiliary_joins", []):
        if join.get("right_source") not in known_ids: raise ValueError(f"auxiliary_joins.right_source '{join.get('right_source')}' is not a known source id")
        right_key, left_key = join.get("right_key"), join.get("left_key")
        if not right_key or not left_key: raise ValueError("auxiliary_joins entries need right_key and left_key")
        right_parts = right_key if isinstance(right_key, list) else [right_key]
        left_parts = left_key if isinstance(left_key, list) else [left_key]
        if len(right_parts) != len(left_parts): raise ValueError(f"auxiliary_joins for '{join.get('right_source')}': right_key and left_key must have the same number of fields")
    for join in config.get("auxiliary_date_range_joins", []):
        if join.get("right_source") not in known_ids: raise ValueError(f"auxiliary_date_range_joins.right_source '{join.get('right_source')}' is not a known source id")
        for required in ("employee_field", "date_field", "start_field", "end_field"):
            if not join.get(required): raise ValueError(f"auxiliary_date_range_joins for '{join.get('right_source')}' needs {required}")


def normalize(value):
    text = unicodedata.normalize("NFKC", str(value or "")).strip().casefold()
    return " ".join("".join(character if character.isalnum() else " " for character in text).split())


def number(value):
    if value in (None, ""): return 0.0
    if isinstance(value, bool): return float(value)
    if isinstance(value, (int, float)):
        result = float(value)
        if not math.isfinite(result): raise ValueError(f"invalid number: {value}")
        return result
    text = unicodedata.normalize("NFKC", str(value)).strip().replace("\u00a0", "").replace("٬", ",").replace("٫", ".")
    negative = text.startswith("(") and text.endswith(")")
    if negative: text = text[1:-1].strip()
    text = re.sub(r"[$€£¥₹]|\b(?:USD|EUR|GBP|EGP|SAR|AED)\b", "", text, flags=re.I).strip().removesuffix("%").strip()
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."): text = text.replace(".", "").replace(",", ".")
        else: text = text.replace(",", "")
    elif "," in text:
        pieces = text.split(",")
        text = "".join(pieces) if all(len(piece) == 3 for piece in pieces[1:]) else text.replace(",", ".")
    result = float(text)
    if negative: result = -result
    if not math.isfinite(result): raise ValueError(f"invalid number: {value}")
    return result


def parse_date(value, date_order="DMY"):
    if value in (None, ""): return None
    if isinstance(value, datetime): return value.date().isoformat()
    if isinstance(value, date): return value.isoformat()
    if isinstance(value, (int, float)) and 1000 < value < 100000:
        try:
            from openpyxl.utils.datetime import from_excel
            return from_excel(value).date().isoformat()
        except Exception: pass
    text = str(value).strip()
    order = date_order.upper()
    local_formats = {"DMY": ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"), "MDY": ("%m/%d/%Y", "%m-%d-%Y", "%m.%d.%Y"), "YMD": ("%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d")}[order]
    # ISO dates are unambiguous even when the workbook's local order is DMY.
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", *local_formats):
        try: return datetime.strptime(text[:10], fmt).date().isoformat()
        except ValueError: pass
    raise ValueError(f"invalid date: {text}")


def coerce(value, kind, date_order="DMY"):
    if value in (None, ""): return None
    if kind == "number": return number(value)
    if kind == "date": return parse_date(value, date_order)
    if kind == "boolean":
        if isinstance(value, bool): return value
        text = normalize(value)
        if text in {"true", "yes", "y", "1", "exit", "left"}: return True
        if text in {"false", "no", "n", "0", "active"}: return False
        raise ValueError(f"invalid boolean: {value}")
    return str(value).strip()


def safe_div(a, b):
    return number(a) / number(b) if number(b) else 0.0


FUNCTIONS = {
    "safe_div": safe_div, "coalesce": lambda *v: next((x for x in v if x not in (None, "")), None),
    "abs": abs, "round": round, "min": min, "max": max, "int": int, "float": float, "str": str,
    "upper": lambda v: str(v or "").upper(), "lower": lambda v: str(v or "").lower(),
    "contains": lambda v, p: str(p).lower() in str(v or "").lower(),
    "year": lambda v: int(str(v)[:4]) if v else None, "month": lambda v: str(v)[:7] if v else None,
    "days_between": lambda a, b: (date.fromisoformat(str(b)[:10]) - date.fromisoformat(str(a)[:10])).days
}


class Evaluator(ast.NodeVisitor):
    def __init__(self, values): self.values = values
    def visit_Expression(self, node): return self.visit(node.body)
    def visit_Constant(self, node): return node.value
    def visit_Name(self, node):
        constants = {"True": True, "False": False, "None": None}
        if node.id in constants: return constants[node.id]
        if node.id not in self.values: raise ValueError(f"Missing field or KPI: {node.id}")
        return self.values[node.id]
    def visit_BinOp(self, node):
        left, right = self.visit(node.left), self.visit(node.right)
        ops = {ast.Add: lambda: left + right, ast.Sub: lambda: left - right, ast.Mult: lambda: left * right, ast.Div: lambda: safe_div(left, right), ast.Mod: lambda: left % right, ast.Pow: lambda: left ** right}
        if type(node.op) not in ops: raise ValueError("Unsupported formula operation")
        return ops[type(node.op)]()
    def visit_UnaryOp(self, node):
        value = self.visit(node.operand)
        if isinstance(node.op, ast.USub): return -value
        if isinstance(node.op, ast.UAdd): return +value
        if isinstance(node.op, ast.Not): return not value
        raise ValueError("Unsupported formula operation")
    def visit_BoolOp(self, node):
        if isinstance(node.op, ast.And):
            for item in node.values:
                if not self.visit(item): return False
            return True
        for item in node.values:
            if self.visit(item): return True
        return False
    def visit_Compare(self, node):
        left = self.visit(node.left)
        for op, item in zip(node.ops, node.comparators):
            right = self.visit(item)
            if isinstance(op, ast.Eq): passed = left == right
            elif isinstance(op, ast.NotEq): passed = left != right
            elif isinstance(op, ast.Lt): passed = left < right
            elif isinstance(op, ast.LtE): passed = left <= right
            elif isinstance(op, ast.Gt): passed = left > right
            elif isinstance(op, ast.GtE): passed = left >= right
            elif isinstance(op, ast.In): passed = left in right
            elif isinstance(op, ast.NotIn): passed = left not in right
            else: raise ValueError("Unsupported comparison")
            if not passed: return False
            left = right
        return True
    def visit_IfExp(self, node): return self.visit(node.body) if self.visit(node.test) else self.visit(node.orelse)
    def visit_Call(self, node):
        if not isinstance(node.func, ast.Name) or node.func.id not in FUNCTIONS: raise ValueError("Unsupported formula function")
        return FUNCTIONS[node.func.id](*[self.visit(item) for item in node.args])
    def generic_visit(self, node): raise ValueError(f"Unsupported formula element: {type(node).__name__}")


def evaluate(formula, values):
    return Evaluator(values).visit(ast.parse(formula, mode="eval"))


def used_fields(source, config):
    fields = set(source.get("field_types", {})) | set(source.get("required_fields", [])) | set(source.get("critical_fields", [])) | set(source.get("column_aliases", {}))
    if source is config["sources"][0]:
        fields |= {x.get("field") for x in config.get("filters", [])}
        fields |= set(config.get("table", {}).get("columns", []))
        for item in config.get("kpis", []): fields |= {item.get("field"), item.get("numerator"), item.get("denominator"), item.get("weight")}
        for item in config.get("charts", []):
            fields |= {item.get("group_by"), item.get("series_by"), item.get("value")}
            fields |= {measure.get("field") for measure in item.get("measures", [])}
        formulas = [item.get("formula", "") for item in config.get("derived_fields", [])]
        formulas += [item.get("condition", "") for item in config.get("rules", [])]
        for formula in formulas:
            try:
                fields |= {node.id for node in ast.walk(ast.parse(formula, mode="eval")) if isinstance(node, ast.Name) and node.id not in FUNCTIONS}
            except SyntaxError:
                pass
    return {field for field in fields if field and not str(field).startswith("_")}


def map_and_clean(headers, raw_rows, source, config):
    if not headers or not any(value not in (None, "") for value in headers):
        raise ValueError("No usable header row was found in the selected source")
    validation = config.get("validation", {})
    strict = validation.get("mode", "flexible") == "strict"
    aliases, lookup = source.get("column_aliases", {}), {}
    for canonical, options in aliases.items():
        options = [options] if isinstance(options, str) else list(options or [])
        for option in [canonical, *options]:
            normalized = normalize(option)
            if not normalized: raise ValueError(f"Blank column alias for {canonical}")
            if normalized in lookup and lookup[normalized] != canonical: raise ValueError(f"Column alias '{option}' maps to both {lookup[normalized]} and {canonical}")
            lookup[normalized] = canonical
    wanted, keep_unmapped = used_fields(source, config), bool(source.get("keep_unmapped_columns", False))
    for field in wanted:
        normalized = normalize(field)
        if normalized in lookup and lookup[normalized] != field: raise ValueError(f"Field name '{field}' conflicts with alias for {lookup[normalized]}")
        lookup.setdefault(normalized, field)
    mapped, mapping = [], {}
    for header in headers:
        canonical = lookup.get(normalize(header))
        if not canonical and keep_unmapped: canonical = normalize(header).replace(" ", "_")
        if canonical not in wanted and not keep_unmapped: canonical = None
        mapped.append(canonical)
        if canonical: mapping[str(header)] = canonical
    collisions = sorted({canonical for canonical in mapped if canonical and mapped.count(canonical) > 1})
    collision_policy = validation.get("duplicate_source_columns", "first_non_blank")
    if collisions and (strict or collision_policy == "block"): raise ValueError("Multiple source columns map to the same field: " + ", ".join(collisions))
    present = {item for item in mapped if item}
    missing = [field for field in source.get("required_fields", []) if field not in present]
    types = source.get("field_types", {})
    critical = set(source.get("critical_fields", []))
    if strict: critical |= set(source.get("required_fields", []))
    date_order = source.get("date_order", "DMY")
    accepted, rejected, invalid = [], [], defaultdict(int)
    header_value = source.get("header_row", 1)
    row_start = 2 if str(header_value).lower() == "auto" else int(header_value) + 1
    for row_number, values in enumerate(raw_rows, start=row_start):
        if not any(value not in (None, "") for value in values): continue
        row, reasons = {"_source_row": row_number}, []
        for canonical in dict.fromkeys(item for item in mapped if item):
            positions = [index for index, item in enumerate(mapped) if item == canonical]
            candidates = [values[index] if index < len(values) else None for index in positions]
            raw = next((value for value in candidates if value not in (None, "")), candidates[0] if candidates else None)
            try: row[canonical] = coerce(raw, types.get(canonical, "text"), date_order)
            except Exception:
                invalid[canonical] += 1; row[canonical] = None
                if canonical in critical: reasons.append(f"Invalid {canonical}")
        for field in critical:
            if row.get(field) in (None, ""): reasons.append(f"Missing {field}")
        if reasons: rejected.append({"row_number": row_number, "reason": "; ".join(sorted(set(reasons))), "row": row})
        else: accepted.append(row)
    warnings = [f"{count} invalid value(s) converted to blank in {field}" for field, count in invalid.items() if field not in critical]
    warnings += [f"Multiple source columns matched {field}; the first non-blank value was used" for field in collisions]
    return accepted, rejected, missing, mapping, warnings


def _header_position(value):
    if str(value or "").lower() == "auto": return "auto"
    try: return max(1, int(value or 1))
    except (TypeError, ValueError): return 1


def read_csv_source(path, source, config):
    error = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with Path(path).open("r", encoding=encoding, newline="") as handle: rows = list(csv.reader(handle))
            if not rows: raise ValueError("The CSV file is empty; add a header row and at least one data row")
            position = _header_position(source.get("header_row", 1))
            if position == "auto":
                index = next((i for i, row in enumerate(rows) if sum(value not in (None, "") for value in row) >= 2), 0)
            else: index = position - 1
            if index >= len(rows): raise ValueError(f"Header row {index + 1} was not found in the CSV file")
            headers = rows[index]
            if not any(value not in (None, "") for value in headers): raise ValueError("The CSV header row is blank")
            mapped_source = {**source, "header_row": index + 1}
            return map_and_clean(headers, rows[index + 1:], mapped_source, config)
        except (IndexError, csv.Error) as exc: error = exc
        except UnicodeDecodeError as exc: error = exc
    raise error or ValueError("Could not read CSV")


def _workbook_sheet(workbook, requested):
    """Return the requested sheet, or the first visible sheet with data."""
    requested_text = str(requested or "auto").strip()
    if requested_text.lower() not in {"", "auto", "first"}:
        if requested_text not in workbook.sheetnames:
            visible = ", ".join(workbook.sheetnames) or "none"
            raise ValueError(f"Sheet '{requested_text}' was not found. Available sheets: {visible}")
        return workbook[requested_text]
    for sheet in workbook.worksheets:
        if getattr(sheet, "sheet_state", "visible") != "visible": continue
        for row in sheet.iter_rows(min_row=1, max_row=30, values_only=True):
            if any(value not in (None, "") for value in row): return sheet
    raise ValueError("No visible worksheet contains data")


def _worksheet_header(sheet, source):
    iterator = sheet.iter_rows(values_only=True)
    position = _header_position(source.get("header_row", 1))
    if position == "auto":
        for row_number, row in enumerate(iterator, start=1):
            if sum(value not in (None, "") for value in row) >= 2:
                return list(row), iterator, row_number
        raise ValueError(f"Worksheet '{sheet.title}' has no usable header row")
    for _ in range(position - 1):
        if next(iterator, None) is None: raise ValueError(f"Header row {position} was not found in worksheet '{sheet.title}'")
    headers = list(next(iterator, []) or [])
    if not any(value not in (None, "") for value in headers): raise ValueError(f"Header row {position} in worksheet '{sheet.title}' is blank")
    return headers, iterator, position


def read_workbook(path, config):
    from openpyxl import load_workbook
    data_only = bool(config.get("excel", {}).get("data_only", True))
    workbook = load_workbook(path, read_only=True, data_only=data_only, keep_links=False)
    output = {}
    try:
        for source in config["sources"]:
            sheet = _workbook_sheet(workbook, source.get("sheet_name", "auto"))
            headers, iterator, header_number = _worksheet_header(sheet, source)
            mapped_source = {**source, "header_row": header_number}
            output[source.get("id", sheet.title)] = map_and_clean(headers, iterator, mapped_source, config)
    finally: workbook.close()
    return output, "openpyxl"


def probe_excel_file(path):
    """Small read-only signature check used only to choose the safest reader."""
    path = Path(path)
    result = {"format": path.suffix.lower().lstrip(".") or "unknown", "protection": "none", "reader": "openpyxl", "com_mode": None}
    with path.open("rb") as handle: prefix = handle.read(4 * 1024 * 1024)
    if prefix.startswith(b"<## NASCA DRM FILE - VER1.00 ##>"):
        return {**result, "format": "protected_excel", "protection": "nasca", "reader": "excel", "com_mode": "attach-launch"}
    if prefix.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
        encrypted = "EncryptedPackage".encode("utf-16le") in prefix or "EncryptionInfo".encode("utf-16le") in prefix
        return {**result, "format": "encrypted_excel" if encrypted else result["format"], "protection": "office_encrypted" if encrypted else "legacy_or_binary", "reader": "excel", "com_mode": "attach-launch"}
    if prefix.startswith((b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")):
        try:
            with zipfile.ZipFile(path) as package: names = {name.lower() for name in package.namelist()}
            if "xl/workbook.bin" in names: return {**result, "format": "xlsb", "reader": "excel", "com_mode": "open"}
            if "xl/vbaproject.bin" in names: return {**result, "format": "xlsm", "reader": "excel", "com_mode": "open"}
            if "xl/workbook.xml" in names: return {**result, "format": "xlsx"}
        except (OSError, zipfile.BadZipFile):
            pass
    if path.suffix.lower() in {".xls", ".xlsb"}: return {**result, "reader": "excel", "com_mode": "open"}
    return result


def read_com(path, source, config, mode="open", close_auto_opened=None):
    if os.name != "nt": raise RuntimeError("Excel fallback is available on Windows only")
    with tempfile.TemporaryDirectory(prefix="excel_app_") as folder:
        output = Path(folder) / "converted.csv"
        settings = config.get("excel", {})
        wait_seconds = max(15, min(int(settings.get("authorization_wait_seconds", 120)), 600))
        should_close = settings.get("close_auto_opened_file", True) if close_auto_opened is None else close_auto_opened
        command = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ROOT / "excel_com.ps1"), "-InputFile", str(path), "-OutputCsv", str(output), "-SheetName", str(source.get("sheet_name", "")), "-Mode", mode, "-WaitSeconds", str(wait_seconds), "-Recalculate", str(bool(settings.get("recalculate_before_read", True))).lower(), "-CloseAutoOpened", str(bool(should_close)).lower()]
        result = subprocess.run(command, capture_output=True, text=True, timeout=wait_seconds + 120)
        if result.returncode or not output.exists():
            message = (result.stderr or result.stdout).strip()
            raise RuntimeError(message or "Excel could not read the workbook")
        return read_csv_source(output, source, config)


def read_all_com(path, config, mode):
    sources, output = config["sources"], {}
    close_when_done = bool(config.get("excel", {}).get("close_auto_opened_file", True))
    for index, source in enumerate(sources):
        source_mode = mode if index == 0 else ("attach" if mode == "attach-launch" else mode)
        close_after = close_when_done and index == len(sources) - 1
        output[source.get("id", "main")] = read_com(path, source, config, source_mode, close_after)
    return output


def read_sources(path, config):
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Excel source file was not found: {path.name}. Put the file in the upload box and try again.")
    if not path.is_file(): raise ValueError(f"The selected source is not a file: {path.name}")
    if path.stat().st_size == 0: raise ValueError(f"The selected file is empty: {path.name}")
    if path.suffix.lower() == ".csv":
        source = config["sources"][0]
        return {source.get("id", "main"): read_csv_source(path, source, config)}, "csv", {"format": "csv", "protection": "none", "reader": "csv", "com_mode": None}
    probe = probe_excel_file(path)
    requested = {source.get("read_mode", "auto") for source in config["sources"]}
    force_excel = bool(requested & {"excel", "excel_open", "excel_attach"})
    if "excel_attach" in requested: com_mode = "attach-launch"
    elif "excel_open" in requested: com_mode = "open"
    else: com_mode = probe.get("com_mode") or "open"
    if probe.get("protection") != "none" and not config.get("excel", {}).get("automatic_protected_files", True): com_mode = "attach"
    if force_excel or probe.get("reader") == "excel":
        try:
            output = read_all_com(path, config, com_mode)
            return output, "excel-protected-auto" if com_mode == "attach-launch" else "excel-com", probe
        except Exception as first:
            can_retry_visible = not force_excel and com_mode == "open" and config.get("excel", {}).get("automatic_protected_files", True)
            if not can_retry_visible: raise
            output = read_all_com(path, config, "attach-launch")
            retry_probe = {**probe, "protection": "excel_authorized", "reader": "excel", "com_mode": "attach-launch"}
            return output, "excel-protected-auto", retry_probe
    try:
        output, reader = read_workbook(path, config)
        return output, reader, probe
    except Exception as first:
        try:
            fallback_mode = "attach-launch" if config.get("excel", {}).get("automatic_protected_files", True) else "attach"
            output = read_all_com(path, config, fallback_mode)
            fallback_probe = {**probe, "protection": "excel_authorized", "reader": "excel", "com_mode": fallback_mode}
            return output, "excel-com-fallback", fallback_probe
        except Exception as second: raise RuntimeError(f"Normal read failed: {first}. Excel fallback failed: {second}") from second


def _join_key(row, key_spec):
    """key_spec is a single field name, or a list of field names for a composite key (e.g. employee_id + work_date)."""
    parts = key_spec if isinstance(key_spec, list) else [key_spec]
    values = [row.get(part) for part in parts]
    if any(value in (None, "") for value in values): return None
    return str(values[0]) if len(values) == 1 else "␟".join(str(value) for value in values)


def _apply_join_specs(rows, source_rows, join_specs, config):
    for join in join_specs:
        right_key, left_key = join.get("right_key"), join.get("left_key")
        index, duplicate_keys = {}, set()
        duplicate_policy = config.get("validation", {}).get("duplicate_join_keys", "keep_latest")
        for right_row in source_rows.get(join.get("right_source"), []):
            key = _join_key(right_row, right_key)
            if key is None: continue
            if key in index:
                duplicate_keys.add(key)
                if duplicate_policy == "keep_first": continue
            index[key] = right_row
        if duplicate_keys and duplicate_policy == "block": raise ValueError(f"Join source {join.get('right_source')} has duplicate {right_key} values ({len(duplicate_keys)} key(s))")
        if duplicate_keys:
            winner = "latest" if duplicate_policy == "keep_latest" else "first"
            config.setdefault("_runtime_warnings", []).append(f"Join source {join.get('right_source')} had {len(duplicate_keys)} duplicate {right_key} key(s); {winner} value kept")
        for row in rows:
            match = index.get(_join_key(row, left_key))
            for field in join.get("fields", []): row[join.get("prefix", "") + field] = match.get(field) if match else None
    return rows


def join_sources(source_rows, config):
    rows = [dict(row) for row in source_rows.get(config["sources"][0].get("id", "main"), [])]
    return _apply_join_specs(rows, source_rows, config.get("joins", []), config)


def read_single_workbook_source(path, source, config):
    """Read exactly one auxiliary source from its own, independent workbook file (INT-01 multi-file upload)."""
    from openpyxl import load_workbook
    data_only = bool(config.get("excel", {}).get("data_only", True))
    workbook = load_workbook(path, read_only=True, data_only=data_only, keep_links=False)
    try:
        sheet = _workbook_sheet(workbook, source.get("sheet_name", "auto"))
        headers, iterator, header_number = _worksheet_header(sheet, source)
        mapped_source = {**source, "header_row": header_number}
        return map_and_clean(headers, iterator, mapped_source, config)
    finally:
        workbook.close()


def read_one_auxiliary_source(path, source, config):
    path = Path(path)
    if not path.exists(): raise FileNotFoundError(f"Excel source file was not found: {path.name}.")
    if not path.is_file(): raise ValueError(f"The selected source is not a file: {path.name}")
    if path.stat().st_size == 0: raise ValueError(f"The selected file is empty: {path.name}")
    if path.suffix.lower() == ".csv": return read_csv_source(path, source, config)
    return read_single_workbook_source(path, source, config)


def read_auxiliary_sources(paths_by_id, config):
    """Read whichever auxiliary sources (employee, roster, leave, ...) have a file this upload. A source with no path is simply absent this run, not an error."""
    outputs, missing_ids = {}, []
    for source in config.get("auxiliary_sources", []):
        source_id = source.get("id")
        path = paths_by_id.get(source_id)
        if not path:
            missing_ids.append(source_id)
            continue
        outputs[source_id] = read_one_auxiliary_source(path, source, config)
    return outputs, missing_ids


def enrich_with_auxiliary_sources(rows, auxiliary_outputs, config):
    """Left-join accepted auxiliary rows onto the primary (attendance) rows using config['auxiliary_joins']. Never drops or rejects a primary row for a missing or unmatched auxiliary source."""
    source_rows = {source_id: output[0] for source_id, output in auxiliary_outputs.items()}
    rows = _apply_join_specs([dict(row) for row in rows], source_rows, config.get("auxiliary_joins", []), config)
    return rows


def match_date_range_source(rows, right_rows, join_spec):
    """Match each primary row to at most one right-source row whose [start_field, end_field] date range covers the primary row's date_field, for the same employee (e.g. a leave request covering an attendance work_date). Never duplicates a primary row: if more than one right row matches, the one with the largest tie_break_field (e.g. the latest leave_request_id) wins, deterministically. A primary row with no match keeps the linked fields blank; it is never rejected."""
    employee_field, date_field = join_spec["employee_field"], join_spec["date_field"]
    start_field, end_field = join_spec["start_field"], join_spec["end_field"]
    status_field, approved_value = join_spec.get("status_field"), join_spec.get("approved_value")
    exclude_field, tie_break_field = join_spec.get("exclude_field"), join_spec.get("tie_break_field")
    fields, prefix = join_spec.get("fields", []), join_spec.get("prefix", "")

    by_employee = defaultdict(list)
    for right_row in right_rows:
        if status_field and approved_value is not None and right_row.get(status_field) != approved_value: continue
        if exclude_field and right_row.get(exclude_field): continue
        if right_row.get(start_field) in (None, "") or right_row.get(end_field) in (None, ""): continue
        employee_id = right_row.get(employee_field)
        if employee_id in (None, ""): continue
        by_employee[employee_id].append(right_row)

    matched_count, result_rows = 0, []
    for row in rows:
        row = dict(row)
        work_value = row.get(date_field)
        candidates = []
        if work_value not in (None, ""):
            candidates = [candidate for candidate in by_employee.get(row.get(employee_field), [])
                          if candidate[start_field] <= work_value <= candidate[end_field]]
        match = None
        if candidates:
            match = sorted(candidates, key=lambda item: str(item.get(tie_break_field, "")))[-1] if tie_break_field else candidates[0]
        for field in fields: row[prefix + field] = match.get(field) if match else None
        if match: matched_count += 1
        result_rows.append(row)
    return result_rows, matched_count


def process_auxiliary_sources(rows, paths_by_id, config):
    """INT-01: read whichever of the employee/roster/leave files arrived with this upload, join the ones with a configured link onto the attendance rows, and report what was missing, structurally broken, or referenced an unknown employee. A missing or structurally broken auxiliary file never fails the attendance upload; it is skipped and reported instead."""
    outputs, missing_ids = read_auxiliary_sources(paths_by_id or {}, config)
    warnings, sources_report, usable_outputs = [], {}, {}
    for source in config.get("auxiliary_sources", []):
        source_id = source.get("id")
        if source_id in missing_ids:
            sources_report[source_id] = {"delivered": False, "accepted_count": 0, "rejected_count": 0, "structure_ok": None}
            warnings.append(f"{source_id} source was not delivered in this upload; its data was not refreshed.")
            continue
        accepted, rejected, missing_fields, mapping, notes = outputs[source_id]
        if missing_fields:
            sources_report[source_id] = {"delivered": True, "accepted_count": 0, "rejected_count": len(rejected), "structure_ok": False, "structure_error": ", ".join(missing_fields)}
            warnings.append(f"{source_id} source was delivered but is missing required column(s) ({', '.join(missing_fields)}); it was skipped, not linked this time.")
            continue
        sources_report[source_id] = {"delivered": True, "accepted_count": len(accepted), "rejected_count": len(rejected), "structure_ok": True}
        warnings.extend(notes)
        usable_outputs[source_id] = (accepted, rejected, missing_fields, mapping, notes)
    enriched = enrich_with_auxiliary_sources(rows, usable_outputs, config)
    employee_output = usable_outputs.get("employee")
    if employee_output is not None:
        known_ids = {row.get("employee_id") for row in employee_output[0] if row.get("employee_id")}
        unknown_count = sum(1 for row in enriched if row.get("employee_id") and row.get("employee_id") not in known_ids)
        sources_report["employee"]["unknown_employee_references"] = unknown_count
        if unknown_count:
            warnings.append(f"{unknown_count} attendance row(s) reference an employee_id that was not found in the employee source.")
    for join in config.get("auxiliary_date_range_joins", []):
        source_id = join.get("right_source")
        output = usable_outputs.get(source_id)
        if output is None: continue
        enriched, matched_count = match_date_range_source(enriched, output[0], join)
        if source_id in sources_report: sources_report[source_id]["date_range_matches"] = matched_count
    return enriched, {"sources": sources_report, "missing_sources": missing_ids, "warnings": warnings}


def apply_logic(rows, config):
    accepted, rejected, warnings = [], [], []
    for index, source_row in enumerate(rows, start=1):
        row, reasons = dict(source_row), []
        try:
            for item in config.get("derived_fields", []):
                value = evaluate(item["formula"], row)
                if "round" in item and isinstance(value, (int, float)): value = round(value, int(item["round"]))
                row[item["name"]] = value
            for rule in config.get("rules", []):
                if evaluate(rule.get("condition", "False"), row):
                    message = rule.get("message", rule.get("id", "Business rule"))
                    if rule.get("action", "warn") == "reject": reasons.append(message)
                    else: warnings.append(message)
        except Exception as exc: reasons.append(f"Calculation failed: {exc}")
        if reasons: rejected.append({"row_number": row.get("_source_row", index), "reason": "; ".join(reasons), "row": {key: value for key, value in row.items() if not key.startswith("_")}})
        else: accepted.append(row)
    custom = ROOT / "CUSTOM_RULES.py"
    if custom.exists():
        spec = importlib.util.spec_from_file_location("custom_rules", custom); module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module)
        accepted = module.transform_rows(accepted, config)
    return accepted, rejected, sorted(set(warnings))


def _where(rows, condition):
    if not condition: return rows
    selected = []
    for row in rows:
        try:
            if evaluate(condition, row): selected.append(row)
        except Exception as exc:
            raise ValueError(f"Invalid KPI/chart filter '{condition}': {exc}") from exc
    return selected


def aggregate(rows, kind, field=None, numerator=None, denominator=None, weight=None, multiply=1, where=None):
    rows = _where(rows, where)
    if kind == "count": return len(rows)
    if kind == "count_distinct": return len({str(row.get(field)) for row in rows if row.get(field) not in (None, "")})
    values = [number(row.get(field)) for row in rows if row.get(field) not in (None, "")]
    if kind == "sum": return sum(values)
    if kind == "avg": return sum(values) / len(values) if values else 0
    if kind == "min": return min(values) if values else 0
    if kind == "max": return max(values) if values else 0
    if kind == "ratio": return safe_div(sum(number(row.get(numerator)) for row in rows), sum(number(row.get(denominator)) for row in rows)) * multiply
    if kind == "weighted_avg": return safe_div(sum(number(row.get(field)) * number(row.get(weight)) for row in rows), sum(number(row.get(weight)) for row in rows))
    return 0


def filter_rows(rows, filters):
    output = []
    for row in rows:
        keep = True
        for field, selected in (filters or {}).items():
            if selected in (None, "", [], {}): continue
            value = row.get(field)
            if field == "_search": keep = str(selected).lower() in " ".join(str(item) for item in row.values()).lower()
            elif isinstance(selected, dict):
                if selected.get("from") not in (None, "") and (value is None or str(value) < str(selected["from"])): keep = False
                if selected.get("to") not in (None, "") and (value is None or str(value) > str(selected["to"])): keep = False
                if selected.get("min") not in (None, "") and number(value) < number(selected["min"]): keep = False
                if selected.get("max") not in (None, "") and number(value) > number(selected["max"]): keep = False
            elif isinstance(selected, list): keep = str(value) in {str(item) for item in selected}
            else: keep = str(value) == str(selected)
            if not keep: break
        if keep: output.append(row)
    return output


def chart_data(rows, spec):
    group, series, kind = spec.get("group_by"), spec.get("series_by"), spec.get("aggregation", "count")
    limit = int(spec.get("limit", 12))
    rows = _where(rows, spec.get("where"))
    groups = defaultdict(list)
    for row in rows: groups[str(row.get(group) or "Unknown")].append(row)
    all_ranked = sorted(((name, aggregate(items, kind, field=spec.get("value"))) for name, items in groups.items()), key=lambda item: item[0] if spec.get("type") == "line" or re.search(r"date|month|period", str(group), re.I) else -abs(item[1]))
    ranked = all_ranked[-limit:] if spec.get("type") == "line" else all_ranked[:limit]
    labels = [item[0] for item in ranked]
    if spec.get("measures"):
        datasets = [{"name": measure["name"], "values": [aggregate(groups[label], measure.get("aggregation", kind), field=measure.get("field"), numerator=measure.get("numerator"), denominator=measure.get("denominator"), weight=measure.get("weight"), multiply=measure.get("multiply", 1), where=measure.get("where")) for label in labels]} for measure in spec["measures"]]
        return {**spec, "labels": labels, "datasets": datasets}
    if series:
        series_groups = defaultdict(list)
        for row in rows: series_groups[str(row.get(series) or "Unknown")].append(row)
        series_names = [name for name, _ in sorted(series_groups.items(), key=lambda item: -abs(aggregate(item[1], kind, field=spec.get("value"))))[:12]]
        datasets = [{"name": name, "values": [aggregate([row for row in groups[label] if str(row.get(series) or "Unknown") == name], kind, field=spec.get("value")) for label in labels]} for name in series_names]
        return {**spec, "labels": labels, "datasets": datasets}
    values = [round(item[1], 4) for item in ranked]
    result = {**spec, "labels": labels, "values": values}
    if spec.get("type") == "pareto":
        total, running = sum(abs(value) for _, value in all_ranked) or 1, 0
        result["cumulative"] = []
        for value in values:
            running += abs(value); result["cumulative"].append(round(running / total * 100, 2))
    return result


def options(rows, filters):
    result = {}
    for item in filters:
        field, kind = item.get("field"), item.get("type")
        values = [row.get(field) for row in rows if row.get(field) not in (None, "")]
        if kind == "select": result[field] = sorted({str(value) for value in values})[:500]
        elif kind == "date_range": result[field] = {"min": min(map(str, values)) if values else None, "max": max(map(str, values)) if values else None}
        elif kind == "number_range": result[field] = {"min": min(map(number, values)) if values else None, "max": max(map(number, values)) if values else None}
    return result


def dashboard(rows, config, previous=None, filters=None):
    filtered, currency = filter_rows(rows, filters), config.get("project", {}).get("currency", "")
    project = config.get("project", {})
    kpis, values = [], {}
    for spec in config.get("kpis", []):
        value = evaluate(spec.get("formula", "0"), values) if spec.get("type") == "formula" else aggregate(filtered, spec.get("type", "count"), spec.get("field"), spec.get("numerator"), spec.get("denominator"), spec.get("weight"), spec.get("multiply", 1), spec.get("where"))
        value = round(value, int(spec.get("round", 2))) if isinstance(value, float) else value
        values[spec["id"]] = value; item = {**spec, "value": value}
        if item.get("format") == "currency": item["suffix"] = currency
        kpis.append(item)
    before = {item.get("id"): item.get("value") for item in (previous or {}).get("kpis", [])}
    for item in kpis:
        old = before.get(item["id"])
        if isinstance(old, (int, float)) and isinstance(item["value"], (int, float)):
            item["previous"] = old; item["delta"] = round(item["value"] - old, 4); item["delta_percent"] = round(safe_div(item["value"] - old, abs(old)) * 100, 2) if old else None
    charts = [chart_data(filtered, spec) for spec in config.get("charts", [])]
    insights = []
    for item in kpis:
        if item.get("delta_percent") is not None:
            direction = "increased" if item["delta"] > 0 else "decreased" if item["delta"] < 0 else "did not change"
            insights.append(f"{item['label']} {direction} by {abs(item['delta_percent']):,.1f}% versus the previous upload.")
    for chart in charts[:2]:
        if chart.get("labels") and chart.get("values"):
            if chart.get("type") == "line": insights.append(f"Latest {chart['title'].lower()}: {chart['labels'][-1]} ({chart['values'][-1]:,.2f}).")
            else: insights.append(f"Top {chart['title'].lower()}: {chart['labels'][0]} ({chart['values'][0]:,.2f}).")
    columns = config.get("table", {}).get("columns", []) or (list(filtered[0]) if filtered else [])
    preview = [{field: row.get(field) for field in columns} for row in filtered[:int(config.get("table", {}).get("max_rows", 200))]]
    return {"project_name": project.get("name", "Excel Dashboard"), "project_version": project.get("version", ""), "version_name": project.get("version_name", ""), "project_status": project.get("project_status", ""), "department_pack": config.get("pack_name", "generic"), "currency": currency, "filters": config.get("filters", []), "filter_options": options(rows, config.get("filters", [])), "active_filters": filters or {}, "kpis": kpis, "charts": charts, "insights": insights[:8] or ["Upload data to generate insights."], "table_columns": columns, "table_rows": preview, "filtered_count": len(filtered), "current_count": len(rows)}


def kpi_snapshot(rows, config):
    values = {}
    for spec in config.get("kpis", []):
        value = evaluate(spec.get("formula", "0"), values) if spec.get("type") == "formula" else aggregate(rows, spec.get("type", "count"), spec.get("field"), spec.get("numerator"), spec.get("denominator"), spec.get("weight"), spec.get("multiply", 1), spec.get("where"))
        values[spec["id"]] = round(value, int(spec.get("round", 2))) if isinstance(value, float) else value
    return values


def history_summary(rows, config):
    """Small per-upload analytical snapshot: overall KPIs plus KPI values by useful dimensions."""
    summary = {"overall": kpi_snapshot(rows, config), "dimensions": {}}
    history = config.get("history", {})
    configured = history.get("dimensions", "automatic")
    if configured == "automatic":
        dimensions = [item.get("field") for item in config.get("filters", []) if item.get("type") == "select"]
    else:
        dimensions = list(configured or [])
    for field in dimensions[:8]:
        grouped = defaultdict(list)
        for row in rows:
            if row.get(field) not in (None, ""): grouped[str(row.get(field))].append(row)
        ranked = sorted(grouped.items(), key=lambda item: len(item[1]), reverse=True)[:200]
        summary["dimensions"][field] = {value: kpi_snapshot(group, config) for value, group in ranked}
    return summary
