"""Foundation acceptance test for HR Attendance Control V1.0.

Uses only synthetic data in sample/. It proves the first vertical slice:
read attendance, reject repeat bytes, merge a correction and a new record,
and preserve the last good result when an upload is structurally wrong.
"""

import json
import os
import shutil
import tempfile
from datetime import date
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent


def write_attendance_book(path, rows, sheet_name="TIM_02_DailyAttendance"):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    sheet.append(["Synthetic HR attendance update"])
    sheet.append(["Report Date", "2026-09-05", "Source", "Foundation test"])
    sheet.append([
        "Attendance_ID", "Employee_ID", "Work_Date", "Scheduled_Shift_ID",
        "First_In", "Last_Out", "Worked_Minutes", "Attendance_Status",
        "Late_Minutes", "Early_Leave_Minutes", "Overtime_Eligible_Minutes",
    ])
    for row in rows:
        sheet.append(row)
    workbook.save(path)


with tempfile.TemporaryDirectory(prefix="hr_attendance_foundation_") as folder:
    data_dir = Path(folder) / "data"
    os.environ["EXCEL_APP_DATA_DIR"] = str(data_dir)

    import engine

    source = ROOT / "sample" / "HR_Time_Attendance_Demo.xlsx"
    first = engine.process_file(source, source.name)
    first_kpis = {item["id"]: item["value"] for item in first["kpis"]}
    assert first["current_count"] == 200
    assert first["reconciliation"]["balanced"] is True
    assert first_kpis["present_records"] == 118
    assert first_kpis["absence_records"] == 34
    assert first_kpis["leave_records"] == 48
    assert first_kpis["attendance_rate"] == 59.0
    assert sum(next(chart for chart in first["charts"] if chart["id"] == "attendance_mix")["values"]) == 200

    duplicate = engine.process_file(source, "same_bytes_different_name.xlsx")
    assert duplicate.get("duplicate_upload") is True
    assert duplicate["current_count"] == 200

    delta = Path(folder) / "HR_Attendance_Delta_Demo.xlsx"
    write_attendance_book(delta, [
        ["TIM02-00001", "E000001", date(2026, 6, 17), "S4", None, None, 480, "Present", 0, 0, 30],
        ["TIM02-00201", "E000001", date(2026, 9, 5), "S4", None, None, 450, "Absent", 15, 0, 0],
    ])
    merged = engine.process_file(delta, delta.name)
    merged_kpis = {item["id"]: item["value"] for item in merged["kpis"]}
    assert merged["current_count"] == 201
    assert merged_kpis["present_records"] == 119
    assert merged_kpis["absence_records"] == 35
    assert merged_kpis["leave_records"] == 47
    assert merged_kpis["late_minutes_total"] == 60238.0
    assert len(engine.runs()) == 2

    malformed = Path(folder) / "Wrong_HR_File.xlsx"
    write_attendance_book(malformed, [["wrong", "row"]], sheet_name="WrongSheet")
    try:
        engine.process_file(malformed, malformed.name)
        raise AssertionError("Wrong source structure must not replace the current result")
    except RuntimeError as exc:
        assert "Normal read failed" in str(exc)
    assert engine.state_now()["current_count"] == 201

    shutil.copy2(delta, ROOT / "sample" / "HR_Attendance_Delta_Demo.xlsx")
    print(json.dumps({
        "ok": True,
        "baseline_records": 200,
        "duplicate_guard": True,
        "after_correction_and_new_record": 201,
        "last_good_result_preserved": True,
        "delta_sample_created": True,
    }, ensure_ascii=False, indent=2))
