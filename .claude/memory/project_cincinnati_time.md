---
name: Time Analyzer — Spun Out to New Project
description: The Cincinnati utilization scripts were removed from this repo on 2026-05-03; a new standalone multi-office time analyzer project is being built separately
type: project
originSessionId: 74a638b7-937f-445f-a911-cf6ed13959c7
---
## Status: Removed from this repo (2026-05-03)

The following scripts were deleted from `etl/` and are no longer in this project:
- `ajera_utilization_api.py` (primary — API-based, was the working version)
- `ajera_utilization.py` (legacy XLS-based)
- `ajera_time_probe.py`, `ajera_time_probe2.py`, `ajera_time_probe3.py`
- `ajera_probe_timesheets.py`

## New Project

A standalone multi-office time & utilization analyzer is being built to cover all four offices (CIN, MN, DAL, ORL) in a single Excel workbook. The kickoff document is at `E:\unanet\time_project_kickoff.md` and contains:
- Recommended project structure
- Canonical `EmployeeRecord` / `HourBucket` data model
- Per-extractor technical notes (Ajera, QBO, QBD)
- Excel output spec (6 tabs, multi-office extensions)
- Phasing plan and what to port from this repo

**Why:** Partners want a single utilization report across all offices, not just Cincinnati.
**How to apply:** For any time/utilization questions, the work lives in the new project, not here. Reference `time_project_kickoff.md` for all design decisions already made.
