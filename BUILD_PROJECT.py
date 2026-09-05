import json
import hashlib
import re
import zipfile
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
project_data = json.loads((ROOT / "PROJECT.json").read_text(encoding="utf-8"))
lock_path = ROOT / "LOCKED_CORE.json"
if lock_path.is_file():
    locked = json.loads(lock_path.read_text(encoding="utf-8"))
    changed = [name for name, expected in locked.items() if not (ROOT / name).is_file() or hashlib.sha256((ROOT / name).read_bytes()).hexdigest() != expected]
    if changed:
        approval = project_data.get("core_change", {})
        reason = str(approval.get("reason", "")).strip()
        if not approval.get("allow") or not reason:
            raise SystemExit("Core component changed: " + ", ".join(changed) + ". Use PROJECT.json core_change.allow=true and a short core_change.reason, then run SMOKE_TEST.py and build again.")
        locked = {item: hashlib.sha256((ROOT / item).read_bytes()).hexdigest() for item in locked if (ROOT / item).is_file()}
        lock_path.write_text(json.dumps(locked, indent=2) + "\n", encoding="utf-8")
        print("Approved focused core change: " + reason)
project = project_data.get("project", {})


def file_token(value, fallback):
    """Make a readable Windows-safe token without removing non-Latin names."""
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "-", str(value).strip())
    value = re.sub(r"\s+", "-", value).strip(".-")
    return value or fallback


def table_value(value):
    return str(value).replace("|", "/").replace("\r", " ").replace("\n", " ").strip()


name = project.get("name", "Excel Project")
version = str(project.get("version", "V1.0")).strip() or "V1.0"
version_name = str(project.get("version_name", "Unspecified update")).strip() or "Unspecified update"
safe_name = file_token(name, "Excel-Project")
safe_version = file_token(version, "V1.0")
output = ROOT.parent / f"{safe_name}-{safe_version}.zip"
runtime_files = [
    "START.bat", "RUN_PROJECT.ps1", "PROJECT.json", "README.md", "USER_GUIDE.md",
    "PROJECT_GUIDE.md", "START_HERE_AI.md", "project_memory/PROJECT_LOG.md", "project_memory/PROGRAM_MAP.md", "SKILL.md",
    "BUILD_PROJECT.py", "CHECK_ENVIRONMENT.py", "SMOKE_TEST.py", "LOCKED_CORE.json",
    "engine.py", "calculation_engine.py", "dashboard.html",
    "department_packs.json", "excel_com.ps1", "vendor.zip"
]
if (ROOT / "CUSTOM_RULES.py").is_file(): runtime_files.append("CUSTOM_RULES.py")
for sample_file in sorted((ROOT / "sample").glob("*")):
    if sample_file.is_file(): runtime_files.append(str(sample_file.relative_to(ROOT)))
missing = [name for name in runtime_files if not (ROOT / name).is_file()]
if missing: raise SystemExit("Missing runtime file(s): " + ", ".join(missing))


def build_archive():
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for runtime_name in runtime_files:
            archive.write(ROOT / runtime_name, Path(safe_name) / runtime_name)


def record_build_event():
    """Append one timestamped event per version; reruns do not create noise."""
    log_path = ROOT / "project_memory" / "PROJECT_LOG.md"
    if not log_path.is_file():
        raise SystemExit("Missing project memory log: project_memory/PROJECT_LOG.md")
    marker = f"<!-- BUILD:{version} -->"
    text = log_path.read_text(encoding="utf-8")
    if marker in text:
        return False
    if "## Automated build events" not in text:
        text += (
            "\n## Automated build events\n\n"
            "| Version | Version name | Timestamp | Event | Output | Status |\n"
            "| --- | --- | --- | --- | --- | --- |\n"
        )
    placeholder = "| — | — | — | No build recorded yet; `BUILD_PROJECT.py` will append one. | — | Pending |"
    text = text.replace(placeholder + "\n", "")
    timestamp = datetime.now().astimezone().isoformat(timespec="seconds")
    row = (
        f"{marker}\n"
        f"| {table_value(version)} | {table_value(version_name)} | {timestamp} | "
        f"Automated build | {table_value(output.name)} | Passed |\n"
    )
    log_path.write_text(text.rstrip() + "\n" + row, encoding="utf-8")
    return True


# Build once to validate the package, then record the exact successful build
# in the memory file and rebuild so the event travels inside the ZIP.
build_archive()
if output.stat().st_size >= 5 * 1024 * 1024:
    raise SystemExit("Built project exceeds 5 MB")
if record_build_event():
    build_archive()
if output.stat().st_size >= 5 * 1024 * 1024:
    raise SystemExit("Built project exceeds 5 MB")
print(f"Version: {version} — {version_name}")
print(output)
print(f"{output.stat().st_size / 1024 / 1024:.2f} MB")
