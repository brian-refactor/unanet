---
name: ETL Script Inventory
description: What each ETL script does and how to run it (time/utilization scripts removed 2026-05-03)
type: project
originSessionId: 74a638b7-937f-445f-a911-cf6ed13959c7
---
All scripts live in `etl/`. Run from the repo root (`E:\unanet`).

| Script | Purpose | Command |
|---|---|---|
| `qbo_auth.py` | One-time QBO OAuth flow | `python etl/qbo_auth.py` |
| `qbo_extract.py` | Extract Minnesota from QBO → `output/minnesota/` | `python etl/qbo_extract.py` |
| `qbo_qc.py` | QC report comparing Minnesota CSVs vs live QBO | `python etl/qbo_qc.py` |
| `ajera_extract.py` | Extract Cincinnati from Ajera → `output/cincinnati/` | `python etl/ajera_extract.py` |
| `ajera_probe_activities.py` | Probe Ajera ListActivities/GetActivities for GL account links (diagnostic) | `python etl/ajera_probe_activities.py` |
| `ajera_probe_pipeline.py` | Probe Ajera for revenue pipeline data (projects, contracts, invoices, WIP) | `python etl/ajera_probe_pipeline.py` |
| `ajera_test.py` | Test Ajera connection | `python etl/ajera_test.py` |
| `qbd_parse.py` | Parse QB Desktop exports → `output/dallas/` or `output/orlando/` | `python etl/qbd_parse.py --office dallas --input input/` |
| `qbd_server.py` | QBWC SOAP server (not used — manual export chosen instead) | — |
| `normalize_outputs.py` | Normalize all output CSV column names to Unanet template format | `python etl/normalize_outputs.py` |
| `supabase_load.py` | Upsert normalized CSVs to Supabase | `python etl/supabase_load.py [--office X] [--entity Y]` |
| `write_templates.py` | Write Unanet upload templates from extracted CSVs | `python etl/write_templates.py [--entity Y]` |
| `review_app.py` | Streamlit app for browsing/editing/validating data | `streamlit run etl/review_app.py` |
| `export_coa_preview.py` | Multi-tab COA preview workbook for stakeholders | `python etl/export_coa_preview.py` |

**Credentials:**
- QBO (Minnesota): `etl/.env` + `etl/qbo_tokens.json`
- Ajera (Cincinnati): `etl/ajera.env` (AJERA_API_URL, AJERA_USERNAME, AJERA_PASSWORD)
- QB Desktop (Dallas/Orlando): manual exports dropped into `input/`
- Supabase: `etl/.env` (SUPABASE_URL, SUPABASE_KEY)

**Output convention:** Columns prefixed `_` are reference-only — strip before loading into Unanet templates.

**Why:** Keeps script inventory clear for future sessions.
**How to apply:** When asked to re-run or modify an extraction, reference this to find the right script.
