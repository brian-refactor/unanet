# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Repository Is

This is an Unanet AE (Advanced Edition) ERP implementation workspace for a client onboarding engagement. It contains data migration documentation, Excel upload templates, and Python ETL scripts that extract data from source systems (QuickBooks Online, QuickBooks Desktop, Ajera) and transform it for loading into Unanet Fusion.

## ETL Pipeline Architecture

Data flows through four stages: **Extract → Normalize → Review → Load → Template**

```
Source Systems                Extract              Normalize         Review/Load         Templates
─────────────────             ───────              ─────────         ───────────         ─────────
QuickBooks Online (MN)   →  qbo_extract.py    →                →                  →
QuickBooks Desktop (DAL) →  qbd_parse.py      →  normalize_    →  supabase_load  →  write_templates.py
QuickBooks Desktop (ORL) →  qbd_parse.py      →  outputs.py    →  .py            →
Ajera REST API (CIN)     →  ajera_extract.py  →                →  review_app.py  →  export_coa_preview.py
```

Each extract script produces 7 CSVs per office: COA, Clients, ClientContacts, Vendors, VendorContacts, Employees, ExpenseCodes — landing in `output/<office>/`.

**Office prefixes** (`MN-`, `DAL-`, `ORL-`, `CIN-`) are applied to all FirmCode and ECCode values to prevent key collisions in the single shared Unanet instance.

**Columns prefixed with `_`** in raw CSVs are source-system reference fields — they are stripped by `normalize_outputs.py` before loading.

## Directory Layout

| Path | Purpose |
|---|---|
| `etl/` | All Python ETL scripts |
| `input/` | Raw QB Desktop export files (`.xlsx` for Dallas, `.csv` for Orlando) |
| `output/<office>/` | Normalized CSVs produced by extract scripts |
| `output/templates/` | Final merged Excel files ready for Unanet upload |
| `TimeReport/` | Cincinnati utilization CSV/XLSX exports from Ajera |
| `Documentation/OneDrive_1_4-24-2026/` | Unanet Excel upload templates and migration checklists |

## Running the ETL Scripts

**One-time setup:**
```
pip install -r etl/requirements.txt
pip install pandas thefuzz        # not in requirements.txt but needed by review_app
cp etl/.env.example etl/.env      # fill in QBO and Supabase credentials
# Ajera credentials go in etl/ajera.env (separate file, same key=value format)
python etl/qbo_auth.py            # QBO OAuth flow — run once, saves etl/qbo_tokens.json
```

**Extract by source:**
```
python etl/qbo_extract.py                          # Minnesota (QBO)
python etl/qbd_parse.py --office dallas            # Dallas (reads input/Dallas_*.xlsx)
python etl/qbd_parse.py --office orlando           # Orlando (reads input/Orlando_*.csv)
python etl/ajera_extract.py                        # Cincinnati master data (Ajera)
```

**QB Desktop live extraction** (requires QB Web Connector + `.qwc` file):
```
python etl/qbd_server.py --office dallas           # SOAP server on port 5150
```

**Post-extraction pipeline:**
```
python etl/normalize_outputs.py                    # standardize all CSVs to Unanet columns
python etl/supabase_load.py [--office X] [--entity Y]   # upsert to Supabase
streamlit run etl/review_app.py                    # browse/edit/validate data
python etl/write_templates.py [--entity Y]         # write merged Excel upload templates
python etl/export_coa_preview.py                   # stakeholder COA preview workbook
```

Valid `--office` values: `minnesota`, `cincinnati`, `dallas`, `orlando`  
Valid `--entity` values: `coa`, `clients`, `client_contacts`, `vendors`, `vendor_contacts`, `employees`, `expense_codes`

## Key Outputs

| Script | Output |
|---|---|
| `qbo_extract.py` | `output/minnesota/*.csv` |
| `qbd_parse.py` | `output/dallas/*.csv`, `output/orlando/*.csv` |
| `ajera_extract.py` | `output/cincinnati/*.csv` |
| `normalize_outputs.py` | overwrites CSVs in-place, stripped of `_` columns |
| `write_templates.py` | `output/templates/<entity>_merged.xlsx` |
| `export_coa_preview.py` | `output/COA_Preview.xlsx` |

## Supabase Schema & Data Model

Each entity maps to a Supabase table. All rows carry an `office` column (the partition key). Upsert conflict keys:

| Entity | Table | Conflict Key |
|---|---|---|
| COA | `coa` | `office`, `base_code` |
| Clients | `clients` | `firm_code` |
| Vendors | `vendors` | `firm_code` |
| Expense Codes | `expense_codes` | `ec_code` |
| Client/Vendor Contacts, Employees | respective tables | delete-by-office + re-insert |

The review app reads from **`_resolved` views** (e.g., `clients_resolved`), which merge base data with `field_overrides`. Never write directly to the resolved views — overrides are stored in a separate `field_overrides` table and applied at query time.

Validation flags for required fields are tracked per entity in Supabase (not in CSVs). Required fields per entity are defined in `review_app.py:REQUIRED`.

Boolean, integer, numeric, and date columns are coerced from CSV strings in `supabase_load.py:coerce()` — add new typed columns there when extending the schema.

## Review App (`etl/review_app.py`)

Three-tab Streamlit interface:
1. **Browse & Edit** — inline cell editing; saves changes to `field_overrides` table (keyed on entity + office + record key + column name)
2. **Duplicates** — fuzzy-matches firm names across offices; records merge decisions
3. **Validation** — flags rows missing required fields; supports waiving issues

Deployed at **https://fusionunanet.streamlit.app/** (Streamlit Community Cloud). Credentials fall back to `.streamlit/secrets.toml` when `etl/.env` is not present.

## Template Writer (`etl/write_templates.py`)

Reads from Supabase `_resolved` views (so overrides are included), copies the appropriate Unanet Excel template from `Documentation/OneDrive_1_4-24-2026/Data Upload Templates/`, and writes merged data starting at the template's data row. Output goes to `output/templates/<entity>_merged.xlsx`. The template files determine column order — the `*_COLS` list in `write_templates.py` must exactly match the template's physical column layout left-to-right (positional, not by header name).

## QB Desktop Parsing Notes

- Dallas input is `.xlsx`; Orlando input is `.csv` — `qbd_parse.py` detects format automatically.
- QB Desktop names the vendor export `Vendor` (not `Vendors`).
- Sub-jobs are filtered out (rows where the customer name contains `:` indicating a QB sub-job).
- `Item Price List` → `ExpenseCodes` mapping: QB items become Unanet expense codes with `ECCode = <OFFICE_PREFIX>-<ItemName>`.

## Ajera-Specific Notes

- Ajera credentials live in `etl/ajera.env` (separate from `etl/.env`).
- Ajera API calls require `MethodArguments` in the request body (even when empty).
- Activities in Ajera correspond to ExpenseCodes in Unanet.
- Time data is not accessible via the Ajera API — utilization data comes from the exported utilization report CSV (`TimeReport/cincinnati_utilization.csv`), not from API calls.
- The `output/cincinnati/_activities_raw.json` and `_timesheets_raw.json` files are debug caches from probe scripts, not used by the main pipeline.

## Diagnostic / One-Off Scripts

These are not part of the main pipeline and should not be modified unless explicitly debugging:

| Script | Purpose |
|---|---|
| `etl/qbo_qc.py` | QBO data quality checks |
| `etl/ajera_probe_pipeline.py` | Ajera API probe / debug |
| `etl/ajera_probe_activities.py` | Ajera activities probe |
| `etl/ajera_test.py` | Ajera connection test |
| `etl/qbo_save_tokens.py` | Manual token save helper |

## Documentation & Templates

All migration artifacts live under `Documentation/OneDrive_1_4-24-2026/`:

- **Data Upload Templates/** — Excel workbooks loaded into Unanet Fusion. Prefixed with load-order sequence numbers (`00-SetupInformation`, `01-OrgUnits`, `02-COA`, `03a-Clients`, …). Order matters — later templates reference data created by earlier ones.
- **01–04 Data Pass folders** — Checklists per migration phase.
- **Data Validation Log** — `Data Validation Log v25.05.xlsx` tracks issues across all passes.

## Migration Phase Model

1. **Initial Pass** — First load; structural issues identified.
2. **Validation Pass** — Cleansed reload; validated against Unanet business rules.
3. **Readiness / Mock Go-Live Pass** — Near-final data; simulates go-live.
4. **Go-Live Pass** — Final production load.

## Related Projects

**Multi-office Time & Utilization Analyzer** — A standalone project for time extraction and utilization reporting across all four offices is being built separately. The design document and porting guide is at `time_project_kickoff.md` in this repo root. The Cincinnati Ajera connector is the proven reference implementation; other offices will follow the same extractor interface pattern defined in `time_project_kickoff.md`.
