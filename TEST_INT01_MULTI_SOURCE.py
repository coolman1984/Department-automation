"""INT-01 acceptance tests: separate files uploaded together (attendance, employee,
roster, leave). Uses only synthetic data built in this file. Proves four cases:
a missing source, an unknown employee id, a late correction, and that upload
history is never wiped.

Roster and leave are read and validated independently but are not linked to
attendance rows yet; only the employee source is joined (employee_id is a
simple 1:1 key). Linking roster/leave by employee + work date is a separate,
not-yet-started stage (see project_memory/INT-01-SOURCE-CONTRACTS_AR.md).
"""

import json
import os
import tempfile
from datetime import date
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent


def write_workbook(path, sheet_name, header_row_blanks, headers, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    for _ in range(header_row_blanks):
        sheet.append(["Synthetic HR source export"])
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    workbook.save(path)


def write_attendance(path, rows):
    write_workbook(
        path, "TIM_02_DailyAttendance", 2,
        ["Attendance_ID", "Employee_ID", "Work_Date", "Scheduled_Shift_ID", "First_In", "Last_Out",
         "Worked_Minutes", "Attendance_Status", "Late_Minutes", "Early_Leave_Minutes", "Overtime_Eligible_Minutes"],
        rows,
    )


def write_employee(path, rows):
    write_workbook(
        path, "EMP_01_EmployeeMaster", 1,
        ["Employee_ID", "Employment_Status", "Hire_Date"],
        rows,
    )


def write_roster(path, rows):
    write_workbook(
        path, "SCH_02_EmployeeRosters", 2,
        ["Roster_ID", "Employee_ID", "Work_Date", "Shift_ID"],
        rows,
    )


def write_leave(path, rows):
    write_workbook(
        path, "TIM_07_LeaveRequests", 1,
        ["Leave_Request_ID", "Employee_ID", "Leave_Type", "Start_Date", "End_Date", "Approval_Status"],
        rows,
    )


with tempfile.TemporaryDirectory(prefix="hr_attendance_int01_") as folder:
    folder = Path(folder)
    data_dir = folder / "data"
    os.environ["EXCEL_APP_DATA_DIR"] = str(data_dir)

    import engine

    attendance_1 = folder / "attendance_1.xlsx"
    write_attendance(attendance_1, [
        ["TIM02-00001", "E000001", date(2026, 6, 1), "S1", None, None, 480, "Present", 0, 0, 0],
        ["TIM02-00002", "E000002", date(2026, 6, 1), "S1", None, None, 480, "Present", 0, 0, 0],
        ["TIM02-00003", "E000999", date(2026, 6, 1), "S1", None, None, 480, "Present", 0, 0, 0],
    ])
    employee = folder / "employee.xlsx"
    write_employee(employee, [
        ["E000001", "Active", date(2015, 5, 8)],
        ["E000002", "Active", date(2017, 6, 2)],
    ])
    roster = folder / "roster.xlsx"
    write_roster(roster, [
        ["SCH02-00001", "E000001", date(2026, 6, 1), "S1"],
        ["SCH02-00002", "E000002", date(2026, 6, 1), "S1"],
    ])
    leave = folder / "leave.xlsx"
    write_leave(leave, [
        ["TIM07-00001", "E000002", "Annual", date(2026, 6, 10), date(2026, 6, 12), "Approved"],
    ])

    # Case 1: all four files delivered together. Employee join enriches known
    # employees; the unknown employee (E000999) is visible as an attendance row
    # that still comes through, plus a reported unmatched reference.
    first = engine.process_file(attendance_1, attendance_1.name, auxiliary_paths={
        "employee": employee, "roster": roster, "leave": leave,
    })
    assert first["current_count"] == 3, first["current_count"]
    rows_by_id = {row["attendance_id"]: row for row in engine.current_rows()}
    assert rows_by_id["TIM02-00001"]["employee_employment_status"] == "Active"
    assert rows_by_id["TIM02-00001"]["employee_hire_date"] == "2015-05-08"
    aux = first["auxiliary_sources"]
    assert aux["employee"]["delivered"] is True and aux["employee"]["accepted_count"] == 2
    assert aux["roster"]["delivered"] is True and aux["roster"]["accepted_count"] == 2
    assert aux["leave"]["delivered"] is True and aux["leave"]["accepted_count"] == 1

    # Case 2: unknown employee. E000999 is not in the employee source; its row
    # must still be accepted (not rejected), with no employee enrichment, and
    # the mismatch must be counted and reported as a warning.
    assert rows_by_id["TIM02-00003"].get("employee_employment_status") is None
    assert aux["employee"]["unknown_employee_references"] == 1
    assert any("employee_id that was not found" in warning for warning in first["warnings"])

    # Case 3: a source file missing from this upload (roster and leave absent).
    # The attendance-only upload must still succeed; the missing sources must
    # be visible in the result, not silently skipped.
    attendance_2 = folder / "attendance_2.xlsx"
    write_attendance(attendance_2, [
        ["TIM02-00001", "E000001", date(2026, 6, 1), "S1", None, None, 480, "Present", 0, 0, 0],
        ["TIM02-00002", "E000002", date(2026, 6, 1), "S1", None, None, 480, "Present", 0, 0, 0],
        ["TIM02-00003", "E000999", date(2026, 6, 1), "S1", None, None, 480, "Present", 0, 0, 0],
        ["TIM02-00004", "E000001", date(2026, 6, 2), "S1", None, None, 480, "Present", 0, 0, 0],
    ])
    second = engine.process_file(attendance_2, attendance_2.name, auxiliary_paths={"employee": employee})
    assert second["current_count"] == 4
    aux2 = second["auxiliary_sources"]
    assert aux2["roster"]["delivered"] is False
    assert aux2["leave"]["delivered"] is False
    assert any("roster source was not delivered" in warning for warning in second["warnings"])
    assert any("leave source was not delivered" in warning for warning in second["warnings"])

    # Case 4: a late correction to an older record must replace only that
    # record's own line, must not touch other, later records, and must not
    # wipe upload history (both prior runs must still be listed).
    attendance_3 = folder / "attendance_3.xlsx"
    write_attendance(attendance_3, [
        ["TIM02-00001", "E000001", date(2026, 6, 1), "S1", None, None, 500, "Present", 20, 0, 0],
    ])
    third = engine.process_file(attendance_3, attendance_3.name, auxiliary_paths={"employee": employee})
    assert third["current_count"] == 4, "the correction must merge, not replace, the other three records"
    rows_by_id_after = {row["attendance_id"]: row for row in engine.current_rows()}
    assert rows_by_id_after["TIM02-00001"]["worked_minutes"] == 500
    assert rows_by_id_after["TIM02-00001"]["late_minutes"] == 20
    assert rows_by_id_after["TIM02-00002"]["worked_minutes"] == 480, "an unrelated later record must not change"
    assert rows_by_id_after["TIM02-00004"]["work_date"] == "2026-06-02", "a later record must not be dropped"

    run_history = engine.runs()
    assert len(run_history) == 3, "upload history must keep every prior run, not just the latest"
    source_names = {item["source_name"] for item in run_history}
    assert {"attendance_1.xlsx", "attendance_2.xlsx", "attendance_3.xlsx"} <= source_names

    print(json.dumps({
        "ok": True,
        "all_four_sources_delivered": True,
        "missing_source_visible": True,
        "unknown_employee_visible": True,
        "late_correction_isolated": True,
        "history_preserved_across_uploads": len(run_history),
    }, ensure_ascii=False, indent=2))
