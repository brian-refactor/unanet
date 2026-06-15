# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Repository Is

This is an Unanet AE (Advanced Edition) ERP implementation workspace for a client onboarding engagement across four offices: **Minnesota (MN)**, **Cincinnati (CIN)**, **Dallas (DAL)**, and **Orlando (ORL)**. It contains Python ETL scripts that extract data from source systems, transform it, and load it into Supabase for review before final export to Unanet Fusion load templates.

## Current State (as of 2026-06-15)

### Production Import Status
Core data is **live in Unanet production** (demo data was wiped before load):
- Org units, COA (385 accounts), Clients (~1,200), Vendors (~2,300), Expense Codes ✓
- Employees: **BLOCKED** — code mismatch between pay history file and HR file; blocked on client follow-up
- **Projects & Phases: active workstream** — extracted and in Supabase, pending final load to Unanet

### Projects & Phases in Supabase
| Office | Projects | Phases | Source |
|--------|----------|--------|--------|
| Minnesota | 897 | 1,777 | Projects from QBO; phases from **Monograph** GraphQL API |
| Cincinnati | ~1,224 | 13,872 | Ajera v1 API (`GetProjects`) |
| Dallas | ~500 | 0 | QBO/QB Desktop; phases not yet extracted |
| Orlando | 39 | 0 | `input/Orlando_Projects.CSV` (QB Desktop export); phases not yet extracted |

### Next Check-in
2026-06-24 or 06-25 at 1 PM Mountain / 2 PM Central with Andrew — projects migration progress review.

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
Ajera v1 API (CIN)    →   extract_cin_phases.py →  project_phases
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
# Ajera credentials go in etl/ajera.env (same key=value format)
# Monograph credentials: MONOGRAPH_EMAIL and MONOGRAPH_PASSWORD in etl/.env
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
| `normalize_outputs.py` | Overwrites CSVs in-place, stripped of `_` columns |
| `write_templates.py` | `output/templates/<entity>_merged.xlsx` |
| `write_templates.py --entity open_projects` | `output/templates/open_projects_merged.xlsx` (07a two-tab: Projects + Phases) |
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

### COA master tables
- `coa_master` — 389 consolidated master accounts (source of truth for merged COA)
- `coa_crosswalk` — 843 source-to-master mappings; FK to `coa_master.master_code`
- Run `python etl/export_coa_upload.py` to regenerate the Unanet upload file from Supabase

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

## Related Projects

**Multi-office Time & Utilization Analyzer** — Standalone project for time extraction and utilization reporting. Design doc at `time_project_kickoff.md`. Cincinnati Ajera connector is the reference implementation.
