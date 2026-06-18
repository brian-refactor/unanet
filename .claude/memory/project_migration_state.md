---
name: project-migration-state
description: Current state of projects & phases pipeline across all four offices as of 2026-06-16
metadata: 
  node_type: memory
  type: project
  originSessionId: 85e53d6d-ff3b-41ca-ab20-c7fe8e7b6e3a
---

As of 2026-06-16, projects and phases are extracted and in Supabase for MN, CIN, and DAL. ORL has projects but no phases.

**Why:** Active workstream before Andrew check-in 2026-06-24/25.
**How to apply:** Use this as the baseline when resuming phases work.

## Minnesota (MN) — complete
- 893 projects, 1,777 phases (Monograph)
- Dates: 585 from Monograph project-level API, 344 derived from `YY-NNN` code year
- `org_path` on phases: `FUS-MSP-A01` (all)
- `pm_emp_code` / `pic_emp_code`: 237 projects populated via `extract_monograph_roles.py` from Monograph `teamList`
- Employee codes: `UOC000XXX` format
- Remaining gap: `pm_emp_code` blank for ~656 projects (no Monograph team assignment)

## Cincinnati (CIN) — complete except dates
- 1,224 projects, 18,798 phases (Ajera v1 SOAP)
- Dates: ~5% populated — Ajera didn't track dates; [[todo-cin-dates]]
- `org_path` on phases: `FUS-CIN-A01` (all)
- `pm_emp_code` / `pic_emp_code`: 439 projects have real `DXE000XXX` codes; 550 cleared (former employees not in Unanet HR file)
- Fee fix applied 2026-06-16: fee on L2 only, L3 sub-phases have null fee
- `phase_status` populated from Ajera: `ACTIVE`, `CLOSED`, `MARKETING`, `HOLD`, `WORKHOLD`, `BILLINGHOLD`
- Old placeholder codes (`CIN-EMP-XXXXX`) fully replaced with `DXE000XXX`

## Dallas (DAL) — phases from estimates
- 1,149 projects (QB Desktop), 10,236 phase rows from `input/DAL_Estimates.xlsx`
- 1,085 projects matched estimates file; 1,067 got start/end dates from estimate dates
- L2 = each estimate (with total fee), L3 = each line item (no fee)
- `org_path`: blank — DAL has multiple studios, no mapping available yet
- `pm_emp_code` / `pic_emp_code`: blank — not in estimates file

## Orlando (ORL) — no phases yet
- 39 projects from `input/Orlando_Projects.CSV`
- No phase source identified yet
