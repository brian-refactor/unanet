# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Repository Is

This is **not a software development project**. It is an Unanet AE (Advanced Edition) ERP implementation workspace containing data migration documentation and templates for a client onboarding engagement.

There are no build commands, tests, linters, or source code files.

## Content Overview

All working files live under `Documentation/OneDrive_1_4-24-2026/`:

- **Data Upload Templates/** — Excel workbooks (`.xlsx`) used to load master data into Unanet Fusion. Each file maps to a specific data domain (org units, chart of accounts, clients, vendors, pay history, expense codes, etc.).
- **01–04 Data Pass folders** — Checklists and artifacts for each migration phase: Initial Pass, Validation Pass, Readiness/Mock Go-Live Pass, and Go-Live Pass.
- **Discovery & Metrics** — `AMPLIFY Discovery Guide` and `Project Metrics Data` workbooks track implementation scope.
- **Data Validation Log** — Central log (`Data Validation Log v25.05.xlsx`) for recording and resolving data issues across passes.
- **Supporting Docs** — Word/PDF guides describing what clients should expect during each validation phase.

## Migration Phase Model

The engagement follows four sequential data passes, each building on the previous:

1. **Initial Pass** — First load of raw client data; major structural issues are identified.
2. **Validation Pass** — Cleansed data reload; validation against Unanet business rules.
3. **Readiness / Mock Go-Live Pass** — Near-final data; simulates go-live conditions.
4. **Go-Live Pass** — Final production load.

## ETL Scripts (`etl/`)

Python scripts that extract data from source systems and write CSVs for review before loading into Unanet templates.

```
etl/
├── requirements.txt
├── .env.example          # copy to .env, fill in credentials — never commit .env
├── qbo_auth.py           # run once to authorize QBO and save tokens
├── qbo_extract.py        # extract Minnesota (QBO) → output/minnesota/*.csv
└── qbo_tokens.json       # auto-generated after auth — never commit
```

**Setup:** `pip install -r etl/requirements.txt`, copy `.env.example` to `.env`, run `python etl/qbo_auth.py` once, then `python etl/qbo_extract.py`.

**Output convention:** CSVs land in `output/<office>/`. Columns prefixed with `_` are QBO reference fields — strip them before loading into Unanet templates. Pay rates are intentionally blank in the employee CSV (not available via standard QBO API).

**Office prefixes:** `MN-` (Minnesota/QBO), `DAL-` (Dallas/QBD), `ORL-` (Orlando/QBD), `CIN-` (Cincinnati/Ajera) — applied to all FirmCode and ECCode values to prevent collisions in the single Unanet instance.

## File Naming Conventions

Upload templates are prefixed with a two-digit sequence number that reflects the recommended load order (e.g., `00-SetupInformation`, `01-OrgUnits`, `02-COA`, `03a-Clients`). This order matters because later templates reference data created by earlier ones.
