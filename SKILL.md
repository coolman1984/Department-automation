---
name: build-any-excel-project
description: Adapt this small offline Golden Template to any Excel-based business process without breaking the reusable engine.
---

# Build any Excel project

Read `PROJECT_GUIDE.md`, `project_memory/PROJECT_LOG.md`,
`project_memory/PROGRAM_MAP.md`, then `PROJECT.json`. Inspect the supplied
workbook before changing a formula. Prefer `PROJECT.json`; use
`CUSTOM_RULES.py` only for a small unsupported row rule.

Keep the core files, startup discovery, automatic free port, protected-file
fallback, history, and dashboard. If a locked core change is truly necessary,
set `core_change.allow` and a factual reason, make the smallest change, test,
build, refresh the lock, and turn the flag off. Never silently remove a core
feature or invent source values.

For every material decision append a versioned local ISO timestamp, evidence,
interpretation, assumptions/conflicts, files, and test result to
`project_memory/PROJECT_LOG.md`; update `PROGRAM_MAP.md` when structure or
data flow changes. Record a concise decision summary, not private chain of
thought. The user may be non-technical, so keep warnings visible and factual.

Run `CHECK_ENVIRONMENT.py`, `SMOKE_TEST.py`, and `BUILD_PROJECT.py`. Deliver
one clean ZIP below 5 MB. The normal user workflow is `START.bat` → upload the
next Excel/CSV file → review KPIs, filters, charts, insights, and history.
