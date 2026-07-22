# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Repository Is

This is an Unanet AE (Advanced Edition) ERP implementation workspace for a client onboarding engagement across four offices: **Minnesota (MN)**, **Cincinnati (CIN)**, **Dallas (DAL)**, and **Orlando (ORL)**. It contains Python ETL scripts that extract data from source systems, transform it, and load it into Supabase for review before final export to Unanet Fusion load templates.

## Current State (as of 2026-07-22)

### Production Import Status
Core data is **live in Unanet production** (demo data was wiped before load):
- Org units, COA (385 accounts), Clients (~1,200), Vendors (~2,300), Expense Codes ✓
- Employees: **BLOCKED** — code mismatch between pay history file and HR file; blocked on client follow-up
- **Projects & Phases: active workstream** — extracted and in Supabase, pending final load to Unanet
- **Timesheets: COMPLETE** — 84,922 rows generated to `output/templates/timesheets_merged.xlsx`, pending `ProjectPath` format confirmation from Andrew
- **Open AP / Open AR: MN, CIN, DAL complete** (as of 2026-07-12); **Orlando not started** — QB Desktop exports still needed

### Projects & Phases in Supabase
| Office | Projects | Phases | Source |
|--------|----------|--------|--------|
| Minnesota | 897 | 1,777 | Projects from QBO; phases from **Monograph** GraphQL API |
| Cincinnati | ~1,224 | 13,872 | Ajera v1 API (`GetProjects`) |
| Dallas | ~500 | extracted | QB Desktop estimates (`input/DAL_Estimates.xlsx`); use `extract_dal_phases.py` — one L2 phase per estimate line item, no L3 |
| Orlando | 39 | 0 | `input/Orlando_Projects.CSV` (QB Desktop export); phases not yet extracted |

### Open AP / Open AR Workstream
New workstream from the 2026-06-18 meeting, alongside Timesheets, Project Metrics, and Trial Balance. GL codes and BillStatus logic differ by office — see "Open AP/AR Extraction" below.

| Office | AR GL | AP GL | Org | Status |
|--------|-------|-------|-----|--------|
| Minnesota | 11002 | 20002 | FUS-MSP | Complete |
| Cincinnati | 11003 | 20003 | FUS-CIN | Complete |
| Dallas | 11001 | 20001 | FUS-DAL | Complete |
| Orlando | 11004 | 20004 | FUS-ORL | Not started |

Project Metrics and Trial Balance are blocked on templates/direction from Andrew.

### Next Check-in
2026-07-14 (Andrew + Karen) covered Open AP/AR and Timesheets status. No further check-in currently scheduled — confirm next date with Andrew.

---

## ETL Pipeline Architecture

Data flows: **Extract → Normalize → Review (Supabase) → Export Template**

```
Source System              Extract Script           Entity CSVs → Supabase Tables
─────────────────          ──────────────           ──────────────────────────────
QBO (MN)              →   qbo_extract.py        →  COA, Clients, Vendors, Employees, ExpenseCodes
QB Desktop (DAL)      →   qbd_parse.py          →  COA, Clients, Vendors, Employees, ExpenseCodes
QB Desktop (ORL)      →   qbd_parse.py          →  COA, Clients, Vendors, Employees, ExpenseCodes
Ajera REST API (CIN)  →   ajera_extract.py      →  COA, Clients, Vendors, Employees, ExpenseCodes
QBO (MN)              →   extract_projects.py   →  projects
QB Desktop (DAL/ORL)  →   extract_projects.py   →  projects
Ajera (CIN)           →   extract_projects.py   →  projects
Monograph GraphQL(MN) →   extract_monograph_    →  project_phases
                           phases.py
Monograph GraphQL(MN) →   extract_monograph_    →  projects.pm_emp_code / pic_emp_code
                           roles.py
Ajera v1 API (CIN)    →   extract_cin_phases.py →  project_phases
QBO (MN)              →   extract_mn_open_ap.py →  output/templates/open_ap_mn.xlsx
QBO (MN)              →   extract_mn_open_ar.py →  output/templates/open_ar_mn.xlsx
Ajera exports (CIN)   →   extract_cin_open_ap.py→  output/templates/open_ap_cin.xlsx
Ajera exports (CIN)   →   extract_cin_open_ar.py→  output/templates/open_ar_cin.xlsx
QB Desktop (DAL)      →   extract_dal_open_ap.py→  output/templates/open_ap_dal.xlsx
QB Desktop (DAL)      →   extract_dal_open_ar.py→  output/templates/open_ar_dal.xlsx
Time analyzer DB      →   export_timesheets.py  →  output/templates/timesheets_merged.xlsx
```

**Office prefixes** (`MN-`, `DAL-`, `ORL-`, `CIN-`) are applied to FirmCode and ECCode values to prevent key collisions. Project codes retain their native format (e.g., MN uses `YY-NNN`, CIN uses numeric Ajera keys).

**Columns prefixed with `_`** in raw CSVs are source-system reference fields stripped by `normalize_outputs.py` before loading.

---

## Directory Layout

| Path | Purpose |
|---|---|
| `etl/` | All Python ETL scripts |
| `input/` | Raw QB Desktop export files (`.xlsx` for Dallas, `.csv` for Orlando); `input/Orlando_Projects.CSV` |
| `output/<office>/` | Normalized CSVs produced by extract scripts |
| `output/templates/` | Final merged Excel files ready for Unanet upload |
| `output/monograph_phases_raw.json` | Cached raw Monograph ganttChart response (617 projects) |
| `output/monograph_probe.json` | Monograph GraphQL probe capture (diagnostic) |
| `Unanet-Migration-Files/` | Unanet Excel upload templates (Data Upload Templates subfolder) |
| `TimeReport/` | Cincinnati utilization CSV/XLSX exports from Ajera |

> **Note:** `Documentation/` is gitignored (large binaries). Templates are in `Unanet-Migration-Files/Data Upload Templates/`.

---

## Running the ETL Scripts

### One-time setup
```bash
pip install -r etl/requirements.txt
pip install pandas thefuzz playwright requests   # extras needed by various scripts
python -m playwright install chromium            # for Monograph extraction
cp etl/.env.example etl/.env                    # fill in QBO + Supabase credentials
# Also add MONOGRAPH_EMAIL and MONOGRAPH_PASSWORD to etl/.env (not in .env.example)
# Ajera credentials go in etl/ajera.env: AJERA_API_URL, AJERA_USERNAME, AJERA_PASSWORD
python etl/qbo_auth.py                          # QBO OAuth — run once, saves etl/qbo_tokens.json
```

### Extract master data (COA / Clients / Vendors / Employees / ExpenseCodes)
```bash
python etl/qbo_extract.py                        # Minnesota (QBO)
python etl/qbd_parse.py --office dallas          # Dallas (reads input/Dallas_*.xlsx)
python etl/qbd_parse.py --office orlando         # Orlando (reads input/Orlando_*.csv)
python etl/ajera_extract.py                      # Cincinnati (Ajera REST API)
```

### Extract projects
```bash
python etl/extract_projects.py --office minnesota    # from QBO
python etl/extract_projects.py --office cincinnati   # from Ajera
python etl/extract_projects.py --office dallas       # from QB Desktop
python etl/extract_projects.py --office orlando      # from input/Orlando_Projects.CSV
```

### Extract project phases
```bash
python etl/extract_monograph_phases.py           # MN phases from Monograph (uses saved cookies)
python etl/extract_monograph_phases.py --from-cache  # reprocess last raw pull, no network call
python etl/extract_cin_phases.py                 # CIN phases from Ajera v1 GetProjects
python etl/extract_dal_phases.py                 # DAL phases from QB Desktop estimates export
python etl/extract_dal_phases.py --dry-run       # preview without writing to Supabase
python etl/extract_dal_phases.py --file /path/to/estimates.xlsx  # custom file path
```

### Extract MN PM / Principal-In-Charge roles from Monograph
```bash
python etl/extract_monograph_roles.py            # writes pm_emp_code / pic_emp_code to projects
python etl/extract_monograph_roles.py --dry-run  # preview matches without writing
```

### Extract Open AP / Open AR
```bash
python etl/extract_mn_open_ap.py                 # MN AP from QBO (Bills, cutoff TxnDate <= 2026-07-07)
python etl/extract_mn_open_ar.py                 # MN AR from QBO (Invoices, same cutoff)
python etl/extract_mn_open_ap.py --dry-run
python etl/extract_cin_open_ap.py                 # CIN AP from Ajera Vendor Invoice Aging export (input/CINN - AP.xls)
python etl/extract_cin_open_ar.py                 # CIN AR from Ajera Client Invoice Aging export (input/MINN - AR.xls; naming is wrong, it's CIN data)
python etl/extract_dal_open_ap.py                 # DAL AP from QB Desktop Unpaid Bills Detail (input/DAL - Open AP.CSV)
python etl/extract_dal_open_ar.py                 # DAL AR from QB Desktop Open Invoices (input/DAL - Open AR.csv)
python etl/build_cin_ap_ar_guide.py                # builds output/templates/CIN_AP_AR_Guide.xlsx — mapping guide for the CIN team since Ajera API can't pull financials
```
Orlando AP/AR extractors do not exist yet — no QB Desktop exports obtained.

### Export timesheets
```bash
python etl/export_timesheets.py                  # reads live from Supabase via etl/.env, writes output/templates/timesheets_merged.xlsx
```
`etl/_generate_timesheets.py` is a one-shot script from the original 84,922-row run (hardcoded project sets, no `.env` needed) — superseded by `export_timesheets.py` for future runs.

### Export missing employees
```bash
python etl/export_missing_employees.py            # active time-DB employees not present in the Unanet employee file → output/Missing_Employees_For_Review.xlsx
```

### Post-extraction pipeline
```bash
python etl/normalize_outputs.py                  # standardize all CSVs to Unanet columns
python etl/supabase_load.py [--office X] [--entity Y]   # upsert to Supabase
streamlit run etl/review_app.py                  # browse/edit/validate data
python etl/write_templates.py [--entity Y]       # write merged Excel upload templates
python etl/export_coa_preview.py                 # stakeholder COA preview workbook
```

Valid `--office` values: `minnesota`, `cincinnati`, `dallas`, `orlando`  
Valid `--entity` values: `coa`, `clients`, `client_contacts`, `vendors`, `vendor_contacts`, `employees`, `expense_codes`, `projects`, `open_projects`

### Write projects + phases template (preferred for 07a)
```bash
python etl/write_projects_template.py                                          # all offices
python etl/write_projects_template.py --office dallas                          # one office
python etl/write_projects_template.py --org-units input/OrgUnits.xlsx --emp-codes input/EmployeeCodes.xlsx
```
`write_projects_template.py` is the dedicated projects template writer — it supports optional lookup files to populate `owning_org` and `pm_emp_code` columns. Without lookup files those columns are left blank. Prefer this over `write_templates.py --entity open_projects` when Andrew/Chandler have provided the org/employee reference files.

### Load employees from Unanet HR file
```bash
python etl/load_employees.py          # reads Unanet-Migration-Files/01. Data - Initial Pass/05a-Employees_Fusion6.11.2026.xlsx
python etl/load_employees.py --dry-run
```

### QB Desktop live extraction (requires QB Web Connector + `.qwc` file)
```bash
python etl/qbd_server.py --office dallas         # SOAP server on port 5150
```

---

## Key Outputs

| Script | Output |
|---|---|
| `qbo_extract.py` | `output/minnesota/*.csv` |
| `qbd_parse.py` | `output/dallas/*.csv`, `output/orlando/*.csv` |
| `ajera_extract.py` | `output/cincinnati/*.csv` |
| `extract_projects.py` | `output/<office>/<office>_Projects.csv` |
| `extract_monograph_phases.py` | Supabase `project_phases` (minnesota); `output/monograph_phases_raw.json` |
| `extract_cin_phases.py` | Supabase `project_phases` (cincinnati) |
| `extract_dal_phases.py` | Supabase `project_phases` (dallas) |
| `extract_mn_open_ap.py` / `extract_mn_open_ar.py` | `output/templates/open_ap_mn.xlsx` / `open_ar_mn.xlsx` |
| `extract_cin_open_ap.py` / `extract_cin_open_ar.py` | `output/templates/open_ap_cin.xlsx` / `open_ar_cin.xlsx` |
| `extract_dal_open_ap.py` / `extract_dal_open_ar.py` | `output/templates/open_ap_dal.xlsx` / `open_ar_dal.xlsx` |
| `build_cin_ap_ar_guide.py` | `output/templates/CIN_AP_AR_Guide.xlsx` |
| `export_timesheets.py` | `output/templates/timesheets_merged.xlsx` |
| `export_missing_employees.py` | `output/Missing_Employees_For_Review.xlsx` |
| `normalize_outputs.py` | Overwrites CSVs in-place, stripped of `_` columns |
| `write_templates.py` | `output/templates/<entity>_merged.xlsx` |
| `write_templates.py --entity open_projects` | `output/templates/open_projects_merged.xlsx` (07a two-tab: Projects + Phases) |
| `write_projects_template.py` | `output/templates/projects_merged.xlsx` (07a two-tab; supports org/emp lookup) |
| `export_coa_preview.py` | `output/COA_Preview.xlsx` |

---

## Supabase Schema & Data Model

**Project:** `lulhzpqtwjntvijpzwtv`  
**Review app:** https://fusionunanet.streamlit.app/

Each entity maps to a Supabase table. All rows carry an `office` column (enum: `minnesota`, `cincinnati`, `dallas`, `orlando`). **Never pass `"corporate"` to any Supabase query** — it is not a valid enum value and will cause an API error. Use `SOURCE_OFFICES` (defined in `review_app.py`) for any office-filtered queries.

### Upsert conflict keys
| Entity | Table | Conflict Key |
|---|---|---|
| COA | `coa` | `office`, `base_code` |
| Clients | `clients` | `firm_code` |
| Vendors | `vendors` | `firm_code` |
| Expense Codes | `expense_codes` | `ec_code` |
| Projects | `projects` | `office`, `project_code` |
| Project Phases | `project_phases` | `office`, `project_code`, `level2_code`, `level3_code` |
| Client/Vendor Contacts, Employees | respective tables | delete-by-office + re-insert |

### Additional tables
- `org_units` — org hierarchy (`org_code`, `org_name`, `parent_org_code`, `org_path` UNIQUE); `org_path` values are used as the lookup key in `write_projects_template.py`

### COA master tables
- `coa_master` — 389 consolidated master accounts (source of truth for merged COA)
- `coa_crosswalk` — 843 source-to-master mappings; FK to `coa_master.master_code`
- Run `python etl/export_coa_upload.py` to regenerate the Unanet upload file from Supabase

### Schema migrations
SQL migration files live in `etl/migrations/` (e.g., `005_projects.sql`, `006_org_units.sql`, `007_employees_rebuild.sql`). Apply them manually via the Supabase SQL editor — there is no migration runner.

### Resolved views
The review app reads from **`_resolved` views** (e.g., `clients_resolved`), which merge base data with `field_overrides`. Never write directly to resolved views — overrides are stored in `field_overrides` and applied at query time.

### `project_phases` columns
Key columns: `office`, `project_code`, `level2_code`, `level2_name`, `level3_code`, `level3_name`, `contract_type`, `phase_status`, `start_date`, `end_date`, `fixed_fee`, `hours_budget`, `org_path`, plus budget/cap columns (`labor_contract_cap`, `odc_contract_cap`, etc.)

- MN phases sourced from Monograph: `phase_status` values are `ACTIVE`, `COMPLETED`, `CANCELED`, `PAUSED`; `contract_type` values are `Fixed Fee`, `Hourly`, `Not-to-Exceed`, `Retainer`
- CIN phases sourced from Ajera: nested `InvoiceGroups[].Phases[].Phases[]` hierarchy

---

## Review App (`etl/review_app.py`)

Seven-tab Streamlit interface deployed at **https://fusionunanet.streamlit.app/**:

1. **Browse & Edit** — inline cell editing; saves changes to `field_overrides` table
2. **Duplicates** — fuzzy-matches firm names across offices
3. **Validation** — flags rows missing required fields; supports waiving issues
4. **COA Mapping** — review/edit COA master and crosswalk
5. **Projects** — browse projects by office, filter active/inactive
6. **Project Phases** — browse phases by office; contract type filter is dynamic (built from actual values in data, not hardcoded); uses `SOURCE_OFFICES` to avoid the `corporate` enum error
7. **Export** — build 07a (Projects + Phases combined workbook), download COA/Clients/Contacts/Vendors/ExpenseCodes templates

Credentials fall back to `.streamlit/secrets.toml` when `etl/.env` is not present (used on Streamlit Cloud).

---

## Template Writer (`etl/write_templates.py`)

Reads from Supabase `_resolved` views, copies the appropriate Unanet Excel template from `Unanet-Migration-Files/Data Upload Templates/`, and writes merged data starting at the template's data row. Output goes to `output/templates/<entity>_merged.xlsx`.

The `*_COLS` lists in `write_templates.py` must exactly match the template's physical column layout left-to-right (positional, not by header name). `None` entries in the list skip that column position.

`--entity open_projects` writes the 07a two-tab workbook (Projects tab + Phases and Tasks tab).

---

## Monograph Phase Extraction (`etl/extract_monograph_phases.py`)

Minnesota uses **Monograph** (app.monograph.com) as their PM tool — phases live there, not in QBO.

- **Auth:** Cookie-based. First run uses Playwright to login and saves cookies to `etl/monograph_cookies.json` (gitignored). Subsequent runs reuse saved cookies; falls back to fresh login if expired.
- **API:** GraphQL at `https://app.monograph.com/graphql?op=ganttChart`. Paginates with `first=50, offset=0,50,100...` until `hasNextPage=false`.
- **Data captured per phase:** `name`, `feeType` → `contract_type`, `status` → `phase_status`, `budget` → `fixed_fee`, `startDate`, `endDate`, `hoursPlanned` → `hours_budget`
- **Project code mapping:** Monograph `number` field (e.g., `23-053`) maps directly to MN `project_code` in Supabase (same format, no prefix transformation needed)
- **32 projects skipped** — Monograph projects with no number field (internal/marketing projects without codes)
- **Refresh:** `python etl/extract_monograph_phases.py` (re-queries API) or `--from-cache` to reprocess `output/monograph_phases_raw.json` without hitting Monograph

---

## Cincinnati Phase Extraction (`etl/extract_cin_phases.py`)

Uses the **Ajera v1 SOAP API** (`GetProjects` with `RequestedProjects: [keys]`).

- Loads project keys from `output/cincinnati/cincinnati_Projects.csv`
- Traverses `InvoiceGroups[].Phases[].Phases[]` hierarchy for L2/L3 phase structure
- Deduplicates on `(office, project_code, level2_code, level3_code)` before insert
- `org_path` defaults to `"Cincinnati"` for rows where project has no owning org
- Result: 13,872 phase rows across 1,224 CIN projects

---

## Dallas Phase Extraction (`etl/extract_dal_phases.py`)

Reads `input/DAL_Estimates.xlsx` (QB Desktop estimates export) and writes L2/L3 phase rows to Supabase `project_phases` for `office=dallas`.

- **Source structure:** Each estimate becomes an L2 phase (`fixed_fee` = estimate total, `start_date`/`end_date` from estimate dates). Each line item within an estimate becomes an L3 phase (`name` = memo field).
- **Project date update:** Also updates `start_date`/`end_date` on the parent project row in Supabase (`MIN`/`MAX` across all estimates for that project).
- **Prerequisite:** Projects must already exist in Supabase `projects` table for `office=dallas` — runs `extract_projects.py --office dallas` first if needed.
- **Active marker:** Uses `Ö` character (QB Desktop active estimate marker) to detect active estimates.

---

## Monograph Role Extraction (`etl/extract_monograph_roles.py`)

Pulls PM and Principal-In-Charge assignments from Monograph GraphQL and writes to `projects.pm_emp_code` / `projects.pic_emp_code` in Supabase for `office=minnesota`.

- **Role matching:** Substring match on `rolesSentence` (case-insensitive): "Project Manager" → PM; "Principal" or "Studio Director" → Principal. First match wins when multiple people share a role.
- **Auth:** Reuses saved cookies from `etl/monograph_cookies.json` (same as `extract_monograph_phases.py`). Falls back to Playwright login if expired.
- **Prerequisite:** `output/monograph_phases_raw.json` cache must exist (run `extract_monograph_phases.py` first).

---

## Open AP/AR Extraction (`etl/extract_<office>_open_ap.py` / `_open_ar.py`)

Writes flat rows to the `08-OutstandingAP_Fusion.xlsx` / `09-OutstandingAR_Fusion.xlsx` templates. All header-level fields (InvoiceDate, InvoiceAmount, GLBaseCode, Org, GLPeriod) repeat on every GL distribution line — templates expect self-contained rows, not a header + detail structure.

- **BillStatus:** MN uses QBO's native `BillableStatus` field directly (Billable→`Ready To Bill`, HasBeenBilled→`Billed`, NotBillable→`Never bill`). CIN/DAL exports lack this field — has a project code → `Billed`, no project (overhead) → `Never bill`. DAL AP is all `Billed` (subconsultant invoices only, no project field in export).
- **GL defaults (no line-item detail in aging exports):** AR revenue → `40001` (Design Income); AP expense → `60101` (Subconsultant Costs - In-Contract). Flag for review if revenue/expense mix is material.
- **Unmatched clients/vendors are dropped**, not flagged with placeholder codes — likely new records added since the first data pass, not worth including with FirmCodes that would fail validation.
- **Fuzzy matching:** exact match only (threshold=100) for client/vendor lookups — fuzzy scoring produced wrong matches (e.g. "Little Caesars Pizza" → "hideaway pizza").
- **CIN exports:** Ajera "Client Invoice Aging" (AR) / "Vendor Invoice Aging" (AP) reports, aging date set to the cutoff date. One export is misnamed `MINN - AR.xls` (it is CIN data, not Minneapolis — check content, not filename). CIN intercompany vendors ("Fusion-MSP", "Fusion-Orlando") need special handling, not standard vendor-bill treatment.
- **DAL exports:** header-level only (no line-item GL detail); DAL AR has 9 partially-paid invoices needing manually-filled check numbers (not present in the export).
- `build_cin_ap_ar_guide.py` produces a standalone mapping workbook for the CIN team, since the Ajera REST API's migration user can't pull financial data — only master data is accessible via API.

---

## Timesheet Export (`etl/export_timesheets.py`)

Reads time entries from the separate **time analyzer** Supabase project (a different project ref than this repo's migration DB — see `etl/.env` for `TIME_URL`/`TIME_KEY`) and writes to the `11_OpenProjectALLTimesheets_Fusion.xlsx` template.

- **Employee matching:** per-office fuzzy match (≥90 score) time-DB name → Unanet EmpCode. Unmatched active employees fall back to office placeholder codes (`FORMER-MN`, `FORMER-CIN`, `FORMER-DAL`, `FORMER-ORL` — must exist as inactive employees in Unanet before running). Dummy/internal accounts are skipped entirely.
- **Project matching:** only includes entries mapping to a project already loaded in the migration DB; overhead/indirect entries (`project_external_id` null) are skipped. Per-office code extraction differs — CIN fuzzy-matches on project name after stripping the leading display code; MN extracts `YY-NNN`; DAL extracts the code before `' - '`; ORL extracts the 4–5 digit prefix.
- **Known gaps:** `ProjectPath` format is unconfirmed with Andrew (currently uses `project_code` only). Rate/amount fields default to placeholder values — real billing rates are not available in the time DB.

---

## Project Extraction (`etl/extract_projects.py`)

| Office | Source | Notes |
|--------|--------|-------|
| Minnesota | QBO customers (sub-customers of clients) | `project_code = YY-NNN` format |
| Cincinnati | Ajera `GetProjects` API | numeric Ajera project keys |
| Dallas | QB Desktop (via `qbd_parse.py` or live server) | 8-digit codes |
| Orlando | `input/Orlando_Projects.CSV` (QB Desktop transaction detail export) | 4–5 digit codes; non-numeric entries get `ORL-MISC001…` placeholders |

---

## COA Pipeline Scripts (in `etl/`)

| Script | Purpose |
|---|---|
| `build_coa_master.py` | Build `coa_master` table from the working Excel file |
| `load_coa_master.py` | Load `coa_master` to Supabase |
| `apply_unmapped_coa.py` | Push manual resolutions for accounts that had no automatic mapping |
| `populate_coa_metrics.py` | Fill MetricType for blank P&L accounts using series-based rules |
| `rebuild_coa_crosswalk.py` | Rebuild `coa_crosswalk` from scratch |
| `export_coa_upload.py` | Pull from Supabase and write final Unanet upload Excel |
| `export_coa_preview.py` | Write stakeholder COA preview workbook |

---

## QB Desktop Parsing Notes

- Dallas input is `.xlsx`; Orlando input is `.csv` — `qbd_parse.py` detects format automatically
- QB Desktop names the vendor export `Vendor` (not `Vendors`)
- Sub-jobs are filtered out (rows where customer name contains `:`)
- `Item Price List` → `ExpenseCodes`: QB items become Unanet expense codes with `ECCode = <OFFICE_PREFIX>-<ItemName>`
- Orlando project codes: `YYNN` (4-digit) or `YYNNN` (5-digit); non-numeric entries get `ORL-MISC001` placeholders

---

## Ajera-Specific Notes

- Credentials in `etl/ajera.env` (separate from `etl/.env`)
- All API calls require `MethodArguments` in the request body (even when empty)
- Activities in Ajera = ExpenseCodes in Unanet
- Time data is NOT accessible via the Ajera REST API — utilization comes from exported CSV (`TimeReport/cincinnati_utilization.csv`)
- Ajera v1 SOAP API (used by `extract_cin_phases.py`) is different from the Ajera REST API used by `ajera_extract.py`

---

## Diagnostic / One-Off Scripts

Not part of the main pipeline — do not modify unless explicitly debugging:

| Script | Purpose |
|---|---|
| `etl/qbo_qc.py` | QBO data quality checks |
| `etl/ajera_probe_pipeline.py` | Ajera REST API probe / debug |
| `etl/ajera_probe_activities.py` | Ajera activities probe |
| `etl/ajera_test.py` | Ajera connection test |
| `etl/qbo_save_tokens.py` | Manual QBO token save helper |
| `etl/probe_monograph.py` | Playwright probe to discover Monograph GraphQL operations |
| `etl/probe_mn_phases.py` | QBO invoice item diagnostic (found 82 unique items, 683 projects) |
| `etl/probe_mn_estimates.py` | QBO estimate diagnostic (found 764 estimates, 312 with phase fees) |
| `etl/extract_mn_phases.py` | QBO-based MN phase extractor (dry-run reference only; Monograph is now authoritative) |
| `etl/probe_mn_categories.py` | Monograph GraphQL probe for project category fields |
| `etl/probe_mn_no_pm.py` | Checks whether MN projects missing a PM have any Monograph team members at all |
| `etl/probe_mn_roles.py` | Samples Monograph `rolesSentence` values for projects missing `pm_emp_code` |
| `etl/probe_one_project.py` | One-off Monograph team/role lookup for a single project slug |
| `etl/probe_time_db.py` | Diagnostic for matching time-DB employees against the Unanet employee file |
| `etl/_generate_timesheets.py` | One-shot hardcoded timesheet generator from the original 84,922-row run; superseded by `export_timesheets.py` — delete after confirming no further use |

---

## Migration Phase Model

1. **Initial Pass** — First load; structural issues identified
2. **Validation Pass** — Cleansed reload; validated against Unanet business rules
3. **Readiness / Mock Go-Live Pass** — Near-final data; simulates go-live
4. **Go-Live Pass** — Final production load

---

## Key Field Constraints

| Field | Valid Values |
|---|---|
| `FinancialType` | `Asset`, `Liability`, `Capital`, `Income`, `Cost` (NOT `Equity` or `Revenue`) |
| `MetricType` | `Billed Revenue`, `Work In Progress`, `Cost Direct`, `Cost Indirect`, `Bad Debt`, `None` |
| `SubledgerType` | `AP`, `AR`, `Bank`, `None` |
| `charge_type` | `billable`, `indirect`, `opportunity`, `projection`, `plan` |
| `contract_type` (project) | `Fixed Fee`, `Hourly`, `Not-to-Exceed`, `Retainer`, `Percent of Construction` |

---

## Deployment

The review app runs on Railway (and Streamlit Cloud). `Procfile` starts it via `streamlit run etl/review_app.py`. The root `requirements.txt` is for Railway; `etl/requirements.txt` is for local development. Credentials on Streamlit Cloud come from `.streamlit/secrets.toml` (gitignored — configure via the Streamlit Cloud UI).

## Git & Output Files

`output/` and `input/` are gitignored by default. Add generated CSVs or templates explicitly when they need to be committed (`git add output/...`). Credential files gitignored: `etl/.env`, `etl/ajera.env`, `etl/qbo_tokens.json`, `etl/monograph_cookies.json`, `etl/*.qwc`.

---

## Related Projects

**Multi-office Time & Utilization Analyzer** — Standalone project for time extraction and utilization reporting. Design doc at `time_project_kickoff.md`. Cincinnati Ajera connector is the reference implementation.
