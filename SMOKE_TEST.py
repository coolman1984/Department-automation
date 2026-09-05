"""Quick checks for the HR Attendance Control project."""

import argparse
import json
import os
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
parser = argparse.ArgumentParser()
parser.add_argument("--file", default=os.environ.get("SMOKE_FILE", ""), help="also process one supplied workbook")
args = parser.parse_args()


with tempfile.TemporaryDirectory(prefix="hr_attendance_smoke_") as folder:
    test_dir = Path(folder)
    os.environ["EXCEL_APP_DATA_DIR"] = str(test_dir / "data")
    os.environ["EXCEL_APP_PROJECT"] = str(ROOT / "PROJECT.json")

    import engine

    assert engine.calc.number("(1,234.50)") == -1234.5
    assert engine.calc.number("١٬٢٣٤٫٥") == 1234.5
    assert engine.calc.parse_date("04/05/2026", "DMY") == "2026-05-04"
    assert engine.calc.evaluate("late_minutes > 5 and late_minutes <= 10", {"late_minutes": 8}) is True

    sample = ROOT / "sample" / "HR_Time_Attendance_Demo.xlsx"
    first = engine.process_file(sample, sample.name)
    assert first["current_count"] == 200
    assert first["reconciliation"]["balanced"] is True
    assert engine.process_file(sample, "same-content-new-name.xlsx").get("duplicate_upload") is True

    if args.file:
        requested = Path(args.file).resolve()
        result = engine.process_file(requested, requested.name)
        print(json.dumps({"ok": True, "rows": result["current_count"], "reader": result["reader"]}, indent=2))
    else:
        print(json.dumps({"ok": True, "attendance_rows": first["current_count"], "duplicate_guard": True, "calculation_helpers": True}, indent=2))
