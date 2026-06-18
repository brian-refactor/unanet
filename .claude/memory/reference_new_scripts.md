---
name: reference-new-scripts
description: New ETL scripts added during 2026-06-15/16 session not yet in CLAUDE.md
metadata: 
  node_type: memory
  type: reference
  originSessionId: 85e53d6d-ff3b-41ca-ab20-c7fe8e7b6e3a
---

## `etl/extract_monograph_roles.py`
Fetches PM and Principal-In-Charge assignments from Monograph `teamList` for all MN projects and writes `pm_emp_code` / `pic_emp_code` to the Supabase `projects` table.
- Uses saved cookies from `etl/monograph_cookies.json`
- PM: rolesSentence contains "Project Manager"; Principal: contains "Principal" or "Studio Director"
- Requires `output/monograph_phases_raw.json` to exist (run `extract_monograph_phases.py` first)
- `--dry-run` flag supported

## `etl/extract_dal_phases.py`
Extracts Dallas project phases from the QB Desktop estimates Excel export.
- Default input: `input/DAL_Estimates.xlsx`
- L2 = each estimate (fee = estimate total), L3 = each line item (no fee)
- Also updates `start_date`/`end_date` on DAL projects from estimate dates
- `--dry-run` and `--file /path/to/file` flags supported

## `extract_monograph_phases.py` (updated)
Now also updates `projects.start_date` / `projects.end_date` from Monograph project-level dates after inserting phases. The `build_project_date_updates()` function handles this separately from phase insertion.

## `extract_cin_phases.py` (updated 2026-06-16)
- Fee fields now only on L2 parent rows; L3 child rows have null fees
- L2 parent row always inserted when sub-phases exist
- `phase_status` now captured from Ajera `Status` field (ACTIVE, CLOSED, MARKETING, HOLD, etc.)
- L3 rows get their own date fields (not inherited from L2)
