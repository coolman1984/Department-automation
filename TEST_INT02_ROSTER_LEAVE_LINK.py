"""INT-02 acceptance tests: link roster and leave to attendance rows.

Roster is linked by a composite key (employee_id + work_date): each
attendance row's roster_shift_id/roster_status are filled only when a
matching roster row exists for that same employee and date.

Leave is linked by date-range matching (is work_date within an approved,
non-cancelled leave request's [start_date, end_date] for the same
employee), not a direct key join, because one employee can have several
leave requests. This must never duplicate attendance rows.

Uses only synthetic data built in this file.
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


def write_roster(path, rows):
    write_workbook(
        path, "SCH_02_EmployeeRosters", 2,
        ["Roster_ID", "Employee_ID", "Work_Date", "Shift_ID", "Roster_Status"],
        rows,
    )


def write_leave(path, rows):
    write_workbook(
        path, "TIM_07_LeaveRequests", 1,
        ["Leave_Request_ID", "Employee_ID", "Leave_Type", "Start_Date", "End_Date", "Approval_Status", "Cancellation_Flag"],
        rows,
    )


with tempfile.TemporaryDirectory(prefix="hr_attendance_int02_") as folder:
    folder = Path(folder)
    os.environ["EXCEL_APP_DATA_DIR"] = str(folder / "data")

    import engine

    attendance = folder / "attendance.xlsx"
    write_attendance(attendance, [
        # E000001: has a matching roster row for this date -> should get roster_shift_id/roster_status.
        ["TIM02-00001", "E000001", date(2026, 6, 1), "S1", None, None, 480, "Present", 0, 0, 0],
        # E000002: no roster row for this date at all -> roster fields must stay blank, row still accepted.
        ["TIM02-00002", "E000002", date(2026, 6, 1), "S1", None, None, 480, "Present", 0, 0, 0],
        # E000003: within an approved leave request's date range -> should get leave_type/leave_request_id.
        ["TIM02-00003", "E000003", date(2026, 6, 15), "S1", None, None, 0, "Leave", 0, 0, 0],
        # E000004: has two leave requests, only one of which covers this work date. No duplicate rows allowed.
        ["TIM02-00004", "E000004", date(2026, 7, 20), "S1", None, None, 0, "Leave", 0, 0, 0],
        # E000005: has a cancelled leave request covering this date -> must NOT be linked (cancelled).
        ["TIM02-00005", "E000005", date(2026, 8, 1), "S1", None, None, 480, "Present", 0, 0, 0],
        # E000006: has a pending (not approved) leave request covering this date -> must NOT be linked.
        ["TIM02-00006", "E000006", date(2026, 8, 10), "S1", None, None, 480, "Present", 0, 0, 0],
    ])

    roster = folder / "roster.xlsx"
    write_roster(roster, [
        ["SCH02-00001", "E000001", date(2026, 6, 1), "S1", "Published"],
        # A roster row exists for E000002, but for a different date -> must not match TIM02-00002.
        ["SCH02-00002", "E000002", date(2026, 6, 2), "S1", "Published"],
    ])

    leave = folder / "leave.xlsx"
    write_leave(leave, [
        ["TIM07-00001", "E000003", "Annual", date(2026, 6, 10), date(2026, 6, 20), "Approved", False],
        # E000004 has two requests; only the second covers 2026-07-20.
        ["TIM07-00002", "E000004", "Annual", date(2026, 1, 1), date(2026, 1, 5), "Approved", False],
        ["TIM07-00003", "E000004", "Unpaid", date(2026, 7, 18), date(2026, 7, 25), "Approved", False],
        # E000005's request covers 2026-08-01 but is cancelled -> must not link.
        ["TIM07-00004", "E000005", "Annual", date(2026, 7, 30), date(2026, 8, 5), "Approved", True],
        # E000006's request covers 2026-08-10 but is only Pending -> must not link.
        ["TIM07-00005", "E000006", "Annual", date(2026, 8, 5), date(2026, 8, 15), "Pending", False],
    ])

    result = engine.process_file(attendance, attendance.name, auxiliary_paths={"roster": roster, "leave": leave})
    assert result["current_count"] == 6, "no attendance row may be duplicated or dropped by linking"

    rows_by_id = {row["attendance_id"]: row for row in engine.current_rows()}

    # Case: roster match found -> filled correctly.
    assert rows_by_id["TIM02-00001"]["roster_shift_id"] == "S1"
    assert rows_by_id["TIM02-00001"]["roster_status"] == "Published"

    # Case: no roster match for this (employee_id, work_date) -> blank, not rejected.
    assert rows_by_id["TIM02-00002"].get("roster_shift_id") is None
    assert rows_by_id["TIM02-00002"].get("roster_status") is None

    # Case: leave overlapping work_date -> visible on the matching attendance row.
    assert rows_by_id["TIM02-00003"]["leave_type"] == "Annual"
    assert rows_by_id["TIM02-00003"]["leave_request_id"] == "TIM07-00001"

    # Case: employee with two leave requests -> exactly one attendance row, correctly matched
    # to the request that actually covers the work date (not the unrelated January one).
    assert rows_by_id["TIM02-00004"]["leave_request_id"] == "TIM07-00003"
    assert rows_by_id["TIM02-00004"]["leave_type"] == "Unpaid"

    # Case: cancelled leave request must not be linked even though its date range covers the work date.
    assert rows_by_id["TIM02-00005"].get("leave_request_id") is None

    # Case: a Pending (not Approved) leave request must not be linked.
    assert rows_by_id["TIM02-00006"].get("leave_request_id") is None

    aux = result["auxiliary_sources"]
    assert "date_range_matches" not in aux["roster"], "roster is a plain key join, not a date-range join; it must not report date_range_matches"
    assert aux["leave"]["date_range_matches"] == 2, "exactly two attendance rows should match an active, approved leave request"

    print(json.dumps({
        "ok": True,
        "roster_match_filled": True,
        "roster_no_match_left_blank": True,
        "leave_overlap_visible": True,
        "leave_multiple_requests_no_duplicate_rows": True,
        "cancelled_and_pending_leave_excluded": True,
    }, ensure_ascii=False, indent=2))
