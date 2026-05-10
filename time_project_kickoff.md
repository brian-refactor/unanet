# Multi-Office Time & Utilization Analyzer — Project Kickoff

## What We're Building

A standalone Python project that connects to all four office time systems, extracts timesheet data, and produces a unified Excel utilization report covering all offices in a single workbook. The Cincinnati/Ajera connector is already fully proven — the main work is adding Minnesota (QBO), Dallas, and Orlando (both QBD) and merging everything.

## Source Systems

| Office | System | Prefix | Auth Method |
|--------|--------|--------|-------------|
| Cincinnati | Ajera REST API | `CIN-` | Username/password session token |
| Minnesota | QuickBooks Online | `MN-` | OAuth 2.0 (saved tokens) |
| Dallas | QuickBooks Desktop | `DAL-` | SOAP via Web Connector, or CSV exports |
| Orlando | QuickBooks Desktop | `ORL-` | SOAP via Web Connector, or CSV exports |

## Recommended Project Structure

```
time-analyzer/
├── .env.example
├── .env                    # never commit
├── requirements.txt
├── main.py                 # CLI entry point — runs all extractors, merges, writes Excel
├── extractors/
│   ├── __init__.py
│   ├── ajera.py            # Cincinnati — copy from ajera_utilization_api.py, strip Excel
│   ├── qbo.py              # Minnesota — QBO Time Activities endpoint
│   └── qbd.py             # Dallas/Orlando — parse CSV exports or SOAP
├── output.py               # unified Excel writer (6-tab report, multi-office)
├── rates.py                # pay rate + billing rate loaders (shared across offices)
├── models.py               # EmployeeRecord dataclass — canonical time data shape
└── output/
    └── utilization_<MonthYYYY>.xlsx
```

## Canonical Data Model

All extractors should normalize to a single shape before the Excel writer touches anything. Suggested dataclass (in `models.py`):

```python
@dataclass
class EmployeeRecord:
    office: str              # 'CIN', 'MN', 'DAL', 'ORL'
    name: str
    employee_type: str       # e.g. 'Principal', 'Senior Designer'
    status: str              # 'Active' | 'Inactive'
    target_pct: float        # billable % target (default 85.0)
    ytd_hours: HourBucket    # trailing 12 months
    lq_hours: HourBucket     # last full calendar quarter
    current_hours: HourBucket  # current month-to-date
    project_hours: dict[str, float]   # project_desc → billable hours (T12)
    overhead_hours: dict[str, float]  # cat_key → hours (T12)
    pay_rate: float          # hourly, 0 if unknown
    billing_rate: float      # hourly billing rate, 0 if unknown

@dataclass
class HourBucket:
    billable: float = 0.0
    indirect: float = 0.0
    total: float = 0.0
    # overhead categories
    admin: float = 0.0
    marketing: float = 0.0
    vacation: float = 0.0
    meetings: float = 0.0
    holiday: float = 0.0
    sick: float = 0.0
    cont_ed: float = 0.0
    other: float = 0.0
```

## Extractor Notes

### Cincinnati (Ajera) — `extractors/ajera.py`

Pull directly from `etl/ajera_utilization_api.py` in this repo. Keep only:
- `_connect` / `_disconnect` / `_call` (API session helpers)
- `fetch_details` (ListTimesheets → GetTimesheets in batches of 20)
- `parse_details` (aggregate to per-employee HourBuckets)

Strip everything Excel-related — that moves to `output.py`.

**Key technical facts that must carry over:**
- `MethodArguments` must always be present in the request body, even as `{}`.
- `ListTimesheets` without filters returns only the most-recent batch — you must loop per employee key with `FilterByEmployee` + `FilterByEarliestTimesheetDate` to get trailing 12 months.
- Hours per timesheet entry are in day columns `D1`–`D7`; sum them: `sum(entry.get(f'D{i}') or 0 for i in range(1, 8))`.
- Overhead label → category mapping is in `OVERHEAD_CAT` in the existing script — copy it verbatim and log any unmapped labels at runtime.
- `TargetBillablePercent` comes from the timesheet record itself; default to 85 if missing.

### Minnesota (QuickBooks Online) — `extractors/qbo.py`

Use the `python-quickbooks` library (already in the existing `requirements.txt`). Time data is in the `TimeActivity` entity.

```python
from quickbooks.objects.timeactivity import TimeActivity
activities = TimeActivity.all(qb=client)  # or filter by date range
```

Each `TimeActivity` has: `EmployeeRef.name`, `Hours`, `Minutes`, `BillableStatus` (`"Billable"` | `"NotBillable"` | `"HasBeenBilled"`), `TxnDate`, `ItemRef.name` (service/activity type), `CustomerRef.name` (project/client).

Map `BillableStatus == "Billable"` to billable hours; everything else is indirect. QBO doesn't have overhead categories — put all indirect time in `admin` unless the item name matches a known overhead label.

Tokens are saved in `etl/qbo_tokens.json` after `qbo_auth.py` runs — copy that auth flow or reference it.

### Dallas & Orlando (QuickBooks Desktop) — `extractors/qbd.py`

QBD doesn't have a REST API. Two options — implement whichever is more practical:

**Option A (simpler): CSV exports**
Export the "Time by Employee" or "Time by Job" memorized report from QuickBooks Desktop as CSV. Parse with `qbd_parse.py` logic already in this repo — that code handles the Dallas vs Orlando format differences (merged header rows, sub-job filtering).

**Option B: Live SOAP via Web Connector**
`etl/qbd_server.py` shows the SOAP server pattern on port 5150. For time data specifically, the QBXML request is `TimeTrackingQuery` with `TxnDateRangeFilter`. This requires the Web Connector to be running on the machine with QB Desktop open.

Overhead categorization for QBD: map the QB service item name using the same `OVERHEAD_CAT` dict. Items named "Vacation", "Holiday", etc. → category; all others with no client assigned → `admin`.

## Excel Output — Six Tabs

The existing `ajera_utilization_api.py` produces the target report format. `output.py` should replicate those six tabs but add an `Office` column and office-level subtotals everywhere.

| Tab | Key Change for Multi-Office |
|-----|----------------------------|
| Executive Summary | KPI tiles aggregate all offices; add per-office KPI section below firm totals |
| Utilization by Employee | Add `Office` column (col 1 or 2); sort by office then utilization % |
| Salary & Cost Recovery | Same, add `Office` column |
| Profitability | Same, add `Office` column |
| Work Breakdown by Employee | Same, group by office then employee |
| Indirect Cost Breakdown | Add per-office breakdown section below firm total |

Copy the `_xcell` / `_hdr` / `_title` / `_widths` / `_xf` / `_xb` Excel helpers verbatim — they are stable and well-tested.

Color scheme to preserve:
```python
_DARK  = '1F4E79'  # title bar
_BLUE  = '2E75B6'  # header row
_LBLUE = 'DEEAF1'  # totals / KPI tiles
_GREEN = 'E2EFDA'  # at/above target
_AMBER = 'FFF2CC'  # within 10 pts below
_RED   = 'FCE4D6'  # >10 pts below
```

## Environment Variables

```ini
# Cincinnati (Ajera)
AJERA_API_URL=https://ajera.com/V005845/AjeraAPI.ashx?...
AJERA_USERNAME=migration
AJERA_PASSWORD=...

# Minnesota (QBO) — tokens saved after qbo_auth.py runs
QBO_CLIENT_ID=...
QBO_CLIENT_SECRET=...
QBO_ENVIRONMENT=production

# Dallas / Orlando — only needed if using live SOAP (Option B)
QBD_DALLAS_HOST=...
QBD_ORLANDO_HOST=...
```

## Setup & Run

```bash
python -m venv .venv
.venv\Scripts\activate        # Windows
pip install -r requirements.txt

# One-time QBO auth (copies token flow from etl/qbo_auth.py)
python qbo_auth.py

# Run the full report
python main.py
# Output: output/utilization_<MonthYYYY>.xlsx
```

## What to Port from the Existing Project (`etl/`)

| File | What to take |
|------|-------------|
| `ajera_utilization_api.py` | `_connect`, `_disconnect`, `_call`, `fetch_details`, `parse_details`, `OVERHEAD_CAT`, `CAT_KEYS`, `CAT_LABELS`, `_trail_dates`, `_last_quarter_dates`, all Excel helpers |
| `ajera_utilization_api.py` | `load_pay_rates`, `load_billing_rates`, `_norm`, `_fl`, `_pay`, `_bill_rate` → move to `rates.py` |
| `qbo_auth.py` | Copy whole file |
| `qbo_extract.py` | Auth/client setup pattern for QBO connection |
| `qbd_parse.py` | CSV parsing logic for Dallas/Orlando time exports |

## What NOT to Port

- Supabase load (`supabase_load.py`) — this project writes Excel only, no DB dependency.
- Streamlit review app (`review_app.py`) — not needed.
- Template writer (`write_templates.py`) — not needed.
- All the master-data extract logic (COA, clients, vendors, etc.).

## Phasing Suggestion

1. **Phase 1** — Port Cincinnati alone. Confirm the output matches the existing `ajera_utilization_api.py` report exactly.
2. **Phase 2** — Add Minnesota (QBO). Establish the multi-office data model and merged Excel layout.
3. **Phase 3** — Add Dallas and Orlando via CSV exports (lower friction than SOAP).
4. **Phase 4** — (Optional) Replace CSV exports with live SOAP if real-time pulls become necessary.
