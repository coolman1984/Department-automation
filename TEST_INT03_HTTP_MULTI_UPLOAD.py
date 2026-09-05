"""INT-03 backend acceptance test: the real HTTP endpoint /api/upload_multi.

Starts the actual local server (same Handler used by dashboard.html) and
sends real multipart/form-data requests built by hand with the standard
library only (urllib), the same wire format a browser's FormData upload
produces. No third-party HTTP library is used here on purpose: this
project's environment check (CHECK_ENVIRONMENT.py) treats any imported,
non-stdlib package in a project .py file as a required runtime
dependency, so a test-only import of something like "requests" would
wrongly tell a real user's machine that the attendance program itself
needs it.

Proves: all four files together, attendance-only (missing sources
visible), a required attendance file missing (clear error), and an
unsupported file extension (clear error) — all over real HTTP, not by
calling engine.process_file directly.
"""

import json
import os
import tempfile
import threading
import time
import urllib.error
import urllib.request
from datetime import date
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent
BOUNDARY = "HRAttendanceTestBoundary123456"


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
    write_workbook(path, "EMP_01_EmployeeMaster", 1, ["Employee_ID", "Employment_Status", "Hire_Date"], rows)


def write_roster(path, rows):
    write_workbook(path, "SCH_02_EmployeeRosters", 2, ["Roster_ID", "Employee_ID", "Work_Date", "Shift_ID", "Roster_Status"], rows)


def write_leave(path, rows):
    write_workbook(path, "TIM_07_LeaveRequests", 1,
                    ["Leave_Request_ID", "Employee_ID", "Leave_Type", "Start_Date", "End_Date", "Approval_Status", "Cancellation_Flag"], rows)


def build_multipart_body(files):
    """files: {field_name: (filename, bytes_content)}. Builds the exact wire format a browser's FormData sends."""
    parts = []
    for field_name, (filename, content) in files.items():
        parts.append(f"--{BOUNDARY}\r\n".encode())
        parts.append(f'Content-Disposition: form-data; name="{field_name}"; filename="{filename}"\r\n'.encode())
        parts.append(b"Content-Type: application/octet-stream\r\n\r\n")
        parts.append(content)
        parts.append(b"\r\n")
    parts.append(f"--{BOUNDARY}--\r\n".encode())
    return b"".join(parts)


def post_multipart(url, files):
    body = build_multipart_body(files)
    request = urllib.request.Request(url, data=body, method="POST", headers={
        "Content-Type": f"multipart/form-data; boundary={BOUNDARY}",
        "Content-Length": str(len(body)),
    })
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            return response.status, json.loads(response.read().decode())
    except urllib.error.HTTPError as exc:
        return exc.code, json.loads(exc.read().decode())


with tempfile.TemporaryDirectory(prefix="hr_attendance_int03_") as folder:
    folder = Path(folder)
    os.environ["EXCEL_APP_DATA_DIR"] = str(folder / "data")

    import engine

    server = engine._create_server(0)
    port = server.server_address[1]
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    base = f"http://127.0.0.1:{port}"
    time.sleep(0.2)

    try:
        attendance = folder / "attendance.xlsx"
        write_attendance(attendance, [
            ["TIM02-00001", "E000001", date(2026, 6, 1), "S1", None, None, 480, "Present", 0, 0, 0],
            ["TIM02-00002", "E000999", date(2026, 6, 1), "S1", None, None, 480, "Present", 0, 0, 0],
        ])
        employee = folder / "employee.xlsx"
        write_employee(employee, [["E000001", "Active", date(2015, 5, 8)]])
        roster = folder / "roster.xlsx"
        write_roster(roster, [["SCH02-00001", "E000001", date(2026, 6, 1), "S1", "Published"]])
        leave = folder / "leave.xlsx"
        write_leave(leave, [["TIM07-00001", "E000001", "Annual", date(2026, 6, 1), date(2026, 6, 3), "Approved", False]])

        # Case 1: all four files uploaded together over real HTTP multipart, exactly as the browser will send them.
        status, result = post_multipart(f"{base}/api/upload_multi", {
            "attendance": (attendance.name, attendance.read_bytes()),
            "employee": (employee.name, employee.read_bytes()),
            "roster": (roster.name, roster.read_bytes()),
            "leave": (leave.name, leave.read_bytes()),
        })
        assert status == 200, result
        assert result["current_count"] == 2, result
        aux = result["auxiliary_sources"]
        assert aux["employee"]["delivered"] is True
        assert aux["roster"]["delivered"] is True
        assert aux["leave"]["delivered"] is True
        assert aux["employee"]["unknown_employee_references"] == 1

        # Case 2: attendance only (a real "one of the four files missing today" upload) still succeeds over HTTP,
        # and the missing sources are visible in the response, not silently dropped.
        attendance2 = folder / "attendance2.xlsx"
        write_attendance(attendance2, [
            ["TIM02-00003", "E000001", date(2026, 6, 2), "S1", None, None, 480, "Present", 0, 0, 0],
        ])
        status2, result2 = post_multipart(f"{base}/api/upload_multi", {"attendance": (attendance2.name, attendance2.read_bytes())})
        assert status2 == 200, result2
        aux2 = result2["auxiliary_sources"]
        assert aux2["employee"]["delivered"] is False
        assert aux2["roster"]["delivered"] is False
        assert aux2["leave"]["delivered"] is False

        # Case 3: the attendance file (required) is missing -> a clear 400 error, not a server crash or silent success.
        status3, result3 = post_multipart(f"{base}/api/upload_multi", {"employee": (employee.name, employee.read_bytes())})
        assert status3 == 400
        assert "attendance" in result3["error"].lower()

        # Case 4: an unsupported extension is rejected clearly, over the real endpoint.
        status4, result4 = post_multipart(f"{base}/api/upload_multi", {
            "attendance": (attendance.name, attendance.read_bytes()),
            "employee": ("not_excel.txt", b"not an excel file"),
        })
        assert status4 == 400
        assert "employee" in result4["error"].lower()

        print(json.dumps({
            "ok": True,
            "http_endpoint": "/api/upload_multi",
            "all_four_files_over_http": True,
            "attendance_only_over_http_shows_missing_sources": True,
            "missing_required_attendance_rejected_clearly": True,
            "unsupported_extension_rejected_clearly": True,
        }, ensure_ascii=False, indent=2))
    finally:
        server.shutdown()
        thread.join(timeout=5)
