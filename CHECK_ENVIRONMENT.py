"""Small, dependency-free startup check for Python and project packages."""

import argparse
import ast
import importlib
import importlib.util
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
for dependency_path in (ROOT / "vendor", ROOT / "vendor.zip"):
    if dependency_path.exists() and str(dependency_path) not in sys.path:
        sys.path.insert(0, str(dependency_path))

PACKAGE_ALIASES = {
    "cv2": "opencv-python",
    "dateutil": "python-dateutil",
    "dotenv": "python-dotenv",
    "fitz": "PyMuPDF",
    "PIL": "Pillow",
    "sklearn": "scikit-learn",
    "win32api": "pywin32",
    "win32com": "pywin32",
    "yaml": "PyYAML",
}
FALLBACK_STDLIB = {
    "__future__", "argparse", "ast", "asyncio", "base64", "collections", "contextlib", "csv",
    "dataclasses", "datetime", "decimal", "email", "enum", "functools", "hashlib", "http",
    "importlib", "io", "itertools", "json", "logging", "math", "mimetypes", "os", "pathlib",
    "re", "secrets", "shutil", "signal", "socket", "sqlite3", "statistics", "string", "subprocess",
    "sys", "tempfile", "threading", "time", "tkinter", "typing", "unicodedata", "urllib", "uuid",
    "warnings", "webbrowser", "xml", "zipfile", "zlib",
}
SKIP_FILES = {"CHECK_ENVIRONMENT.py", "SMOKE_TEST.py", "BUILD_PROJECT.py"}


def clean_module_name(value):
    return str(value or "").split(".", 1)[0].strip()


def local_module_names():
    names = {item.stem for item in ROOT.glob("*.py")}
    names.update(item.name for item in ROOT.iterdir() if item.is_dir())
    return names


def discovered_imports():
    found = set()
    local_names = local_module_names()
    stdlib = set(getattr(sys, "stdlib_module_names", FALLBACK_STDLIB)) | FALLBACK_STDLIB
    for path in sorted(ROOT.glob("*.py")):
        if path.name in SKIP_FILES:
            continue
        try:
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        except Exception:
            continue
        for node in ast.walk(tree):
            module = None
            if isinstance(node, ast.Import):
                module = node.names[0].name if node.names else None
            elif isinstance(node, ast.ImportFrom):
                module = node.module
            top = clean_module_name(module)
            if top and top not in stdlib and top not in local_names:
                found.add(top)
    return found


def load_requirements():
    try:
        config = json.loads((ROOT / "PROJECT.json").read_text(encoding="utf-8"))
    except Exception as exc:
        return [], f"PROJECT.json could not be read: {exc}"
    configured = config.get("runtime", {}).get("packages", [])
    requirements = {}
    for item in configured:
        if isinstance(item, str):
            item = {"import": item, "package": item}
        if not isinstance(item, dict):
            continue
        module = clean_module_name(item.get("import") or item.get("module") or item.get("package"))
        if not module:
            continue
        requirements[module] = {
            "module": module,
            "package": str(item.get("package") or PACKAGE_ALIASES.get(module, module)),
            "purpose": str(item.get("purpose") or "used by this project"),
            "required": bool(item.get("required", True)),
        }
    for module in discovered_imports():
        requirements.setdefault(module, {
            "module": module,
            "package": PACKAGE_ALIASES.get(module, module),
            "purpose": "imported by the project code",
            "required": True,
        })
    return list(requirements.values()), None


def version_tuple(value):
    numbers = re.findall(r"\d+", str(value or ""))
    return tuple(int(item) for item in numbers[:3]) if numbers else (0,)


def requirement_status(requirement):
    module = requirement["module"]
    try:
        spec = importlib.util.find_spec(module)
        if spec is None:
            return {**requirement, "ok": False, "error": "module was not found"}
        imported = importlib.import_module(module)
        location = str(getattr(imported, "__file__", "") or getattr(spec, "origin", ""))
        bundled = any(str(item) in location for item in (ROOT / "vendor", ROOT / "vendor.zip"))
        return {
            **requirement,
            "ok": True,
            "source": "bundled vendor" if bundled else "installed Python",
            "location": location,
            "version": str(getattr(imported, "__version__", "")),
        }
    except Exception as exc:
        return {**requirement, "ok": False, "error": f"{type(exc).__name__}: {exc}"}


def check():
    requirements, config_error = load_requirements()
    minimum = version_tuple(
        json.loads((ROOT / "PROJECT.json").read_text(encoding="utf-8")).get("runtime", {}).get("minimum_python", "3.10")
    ) if not config_error else (3, 10)
    python_ok = sys.version_info[:3] >= minimum
    statuses = [requirement_status(item) for item in requirements]
    missing = [item for item in statuses if item["required"] and not item["ok"]]
    return {
        "ok": not config_error and python_ok and not missing,
        "python": {"executable": sys.executable, "version": ".".join(map(str, sys.version_info[:3])), "ok": python_ok, "minimum": ".".join(map(str, minimum))},
        "requirements": statuses,
        "missing": missing,
        "config_error": config_error,
    }


def print_report(result, brief=False):
    python = result["python"]
    if brief:
        if result["ok"]:
            ready = ", ".join(item["module"] for item in result["requirements"] if item["ok"]) or "standard library only"
            print(f"OK: Python {python['version']}; ready packages: {ready}")
        else:
            names = ", ".join(f"{item['package']} ({item['module']})" for item in result["missing"])
            reason = result.get("config_error") or (f"Python {python['version']} is below {python['minimum']}" if not python["ok"] else "required package check failed")
            print(f"NOT READY: {reason}; missing/unusable: {names or 'see details'}")
        return
    print("Python environment check")
    print(f"Python: {python['version']} — {python['executable']}")
    print(f"Minimum: {python['minimum']} — {'OK' if python['ok'] else 'Too old'}")
    if result.get("config_error"):
        print(f"Configuration: {result['config_error']}")
    for item in result["requirements"]:
        state = "OK" if item["ok"] else "MISSING"
        print(f"{state}: {item['package']} (import {item['module']}) — {item['purpose']}")
        if item["ok"]:
            print(f"  Source: {item['source']}" + (f"; version {item['version']}" if item.get("version") else ""))
        else:
            print(f"  Reason: {item['error']}")
            print(f"  Install: \"{sys.executable}\" -m pip install {item['package']}")
    if result["ok"]:
        print("Result: ready")
    else:
        print("Result: not ready — install or repair the listed required package(s), then run START.bat again.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--brief", action="store_true")
    parser.add_argument("--json", action="store_true")
    args, _ = parser.parse_known_args()
    report = check()
    if args.json:
        print(json.dumps(report, ensure_ascii=False))
    else:
        print_report(report, brief=args.brief)
    raise SystemExit(0 if report["ok"] else 1)
