"""
Extract project data from all source systems and write to output/<office>/projects.csv.

Sources:
  minnesota  — QBO sub-customers (Job=True / IsProject=True)
  dallas     — QB Desktop child rows from input/Dallas_Customers.xlsx
  cincinnati — Ajera API: ListProjects + GetProjects (batched, ~5 min for all 1,200+)
  orlando    — Not available via QB Desktop (flat customer list = clients, not projects)

Output:
  output/minnesota/minnesota_Projects.csv
  output/dallas/dallas_Projects.csv
  output/cincinnati/cincinnati_Projects.csv

Columns map directly to the 07a-OpenProjects_Fusion.xlsx Projects tab.
owning_org, pm_emp_code, pa_emp_code are left blank — apply after lookup files arrive.
For Cincinnati, pm_emp_code is populated from the Ajera EmployeeKey via the employees CSV.

Usage:
    python etl/extract_projects.py                    # all supported offices
    python etl/extract_projects.py --office dallas
    python etl/extract_projects.py --office minnesota
    python etl/extract_projects.py --office cincinnati
"""

import argparse
import csv
import re
from pathlib import Path

from dotenv import load_dotenv

HERE = Path(__file__).parent
OUTPUT_DIR = HERE.parent / "output"

# QBO sub-customers that are not real projects — excluded from MN extraction.
# These failed the YY-NNN code pattern and were confirmed as junk entries.
MN_PROJECT_EXCLUDE = {"1", "City-of-Andover-Cust", "Omni-Hotel", "Snyder-Attorneys,-PA"}

# Columns written to the CSV — matches Unanet Projects tab field names
PROJECT_COLS = [
    "office",
    "project_code",
    "project_name",
    "client_firm_code",
    "owning_org",
    "charge_type",
    "start_date",
    "end_date",
    "contract_type",
    "project_note",
    "po_number",
    "pm_emp_code",
    "pic_emp_code",
    "pa_emp_code",
    "billing_term_type",
    "net_days",
    "invoice_email",
    "use_client_bill_to",
    "is_active",
    "_source_id",
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def write_csv(path: Path, rows: list[dict]):
    if not rows:
        print(f"  (no rows — skipping {path.name})")
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=PROJECT_COLS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    print(f"  -> {len(rows)} projects written to {path}")


def load_clients_lookup(office: str) -> dict[str, str]:
    """Return {source_id: firm_code} from the normalized clients CSV."""
    path = OUTPUT_DIR / office / f"{office}_Clients.csv"
    if not path.exists():
        print(f"  [WARN] Clients CSV not found at {path} — client_firm_code will be blank")
        return {}
    lookup = {}
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            sid = row.get("_source_id", "").strip()
            if sid:
                lookup[sid] = row["FirmCode"]
    return lookup


def load_clients_by_name(office: str) -> dict[str, str]:
    """Return {lower_firm_name: firm_code} for name-based joining."""
    path = OUTPUT_DIR / office / f"{office}_Clients.csv"
    if not path.exists():
        return {}
    lookup = {}
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            name = row.get("FirmName", "").strip().lower()
            if name:
                lookup[name] = row["FirmCode"]
    return lookup


def parse_term(term_str: str) -> tuple[str, str]:
    """Return (billing_term_type, net_days) from 'Net 30', 'Net 15', etc."""
    if not term_str:
        return "", ""
    m = re.match(r"^(Net)\s+(\d+)$", term_str.strip(), re.IGNORECASE)
    if m:
        return "Net", m.group(2)
    return term_str.strip(), ""


# ---------------------------------------------------------------------------
# Minnesota — QBO
# ---------------------------------------------------------------------------

def _parse_mn_project_code(display_name: str) -> tuple[str, str]:
    """
    Split QBO sub-customer DisplayName into (project_code, project_name).
    Handles: '23-020 Abdo Programming', '20-101.1 LEED Certification', '99-998 BD'
    """
    m = re.match(r"^(\d{2}-\d{3}(?:\.\d+)?)\s+(.*)", display_name.strip())
    if m:
        return m.group(1), m.group(2).strip()
    # Fallback: whole string is the name, code is unknown
    return display_name.strip().replace(" ", "-")[:20], display_name.strip()


def extract_minnesota():
    print("Extracting Minnesota projects (QBO)...")

    import sys
    sys.path.insert(0, str(HERE))
    from qbo_extract import get_client, query_all
    from quickbooks.objects.customer import Customer
    from quickbooks.objects.term import Term

    qb, _, _ = get_client()
    customers = query_all(Customer, qb)
    terms = {t.Id: t for t in query_all(Term, qb)}

    top_by_id = {c.Id: c for c in customers if not getattr(c, "ParentRef", None)}
    sub = [c for c in customers if getattr(c, "ParentRef", None) and getattr(c, "Job", False)]
    id_to_firm = load_clients_lookup("minnesota")

    rows = []
    unmatched = 0

    for c in sub:
        parent_id = c.ParentRef.value if c.ParentRef else None
        firm_code = id_to_firm.get(str(parent_id), "") if parent_id else ""
        if not firm_code:
            unmatched += 1

        project_code, project_name = _parse_mn_project_code(c.DisplayName or "")

        if project_code in MN_PROJECT_EXCLUDE:
            continue

        term_obj = terms.get(getattr(getattr(c, "SalesTermRef", None), "value", None))
        billing_term = getattr(term_obj, "Name", "") if term_obj else ""
        net_days = str(getattr(term_obj, "DueDays", "") or "")

        rows.append({
            "office": "minnesota",
            "project_code": project_code,
            "project_name": project_name,
            "client_firm_code": firm_code,
            "owning_org": "",
            "charge_type": "Billable",
            "start_date": "",
            "end_date": "",
            "contract_type": "",
            "project_note": "",
            "po_number": "",
            "pm_emp_code": "",
            "pic_emp_code": "",
            "pa_emp_code": "",
            "billing_term_type": billing_term,
            "net_days": net_days,
            "invoice_email": getattr(getattr(c, "PrimaryEmailAddr", None), "Address", "") or "",
            "use_client_bill_to": "TRUE",
            "is_active": "TRUE" if c.Active else "FALSE",
            "_source_id": c.Id,
        })

    if unmatched:
        print(f"  [WARN] {unmatched} projects had no matching client firm code")

    write_csv(OUTPUT_DIR / "minnesota" / "minnesota_Projects.csv", rows)


# ---------------------------------------------------------------------------
# Dallas — QB Desktop
# ---------------------------------------------------------------------------

_DAL_CODE_RE = re.compile(r"^([A-Z0-9]{3}\d{5})\s*[-–]\s*(.+)$")
_DAL_CODE_BARE_RE = re.compile(r"^([A-Z0-9]{3}\d{5})$")


def _parse_dal_project(raw: str) -> tuple[str, str]:
    """
    Parse Dallas project field (everything after the ':' in Customer).
    Formats: 'FTA21001 - Marina Bay Retail Quincy MA'  or  'FTA21001'
    """
    raw = raw.strip()
    m = _DAL_CODE_RE.match(raw)
    if m:
        return m.group(1), m.group(2).strip()
    m2 = _DAL_CODE_BARE_RE.match(raw)
    if m2:
        return m2.group(1), ""
    # Non-standard format — use raw as name, derive code from start
    slug = re.sub(r"[^A-Z0-9]", "", raw.upper())[:12]
    return f"DAL-{slug}" if slug else "DAL-UNKNOWN", raw


def extract_dallas():
    print("Extracting Dallas projects (QB Desktop)...")

    import sys
    sys.path.insert(0, str(HERE))
    from qbd_parse import read_file

    customers_path = next(
        (p for p in (HERE.parent / "input").glob("Dallas_Customers*")), None
    )
    if not customers_path:
        print("  [ERROR] input/Dallas_Customers.xlsx not found — skipping")
        return

    _, all_rows = read_file(customers_path)
    parent_rows = [r for r in all_rows if ":" not in r.get("Customer", "") and r.get("Customer", "").strip()]
    child_rows  = [r for r in all_rows if ":" in r.get("Customer", "")]
    active_rows = [r for r in child_rows if r.get("Active Status", "").lower() not in ("not-active", "inactive")]

    # Build name→firm_code from normalized CSV (FirmName = Company field)
    name_to_firm = load_clients_by_name("dallas")
    # Also index by the raw Customer field (display name used in sub-job parent references)
    dal_clients_csv = list(csv.DictReader(open(OUTPUT_DIR / "dallas" / "dallas_Clients.csv")))
    src_id_to_firm = {r.get("_source_id", "").strip(): r["FirmCode"] for r in dal_clients_csv if r.get("_source_id")}
    # Build Customer-name → FirmCode from the raw parent rows matched by FirmName
    customer_name_to_firm: dict[str, str] = {}
    for pr in parent_rows:
        cust = pr.get("Customer", "").strip().lower()
        firm = name_to_firm.get(cust, "")
        if not firm:
            company = pr.get("Company", "").strip().lower()
            firm = name_to_firm.get(company, "")
        if firm:
            customer_name_to_firm[cust] = firm
    unmatched = 0
    rows = []

    for r in active_rows:
        customer_field = r["Customer"]
        client_name, project_raw = customer_field.split(":", 1)
        client_name = client_name.strip()
        project_code, project_name = _parse_dal_project(project_raw)

        firm_code = customer_name_to_firm.get(client_name.lower(), "") or name_to_firm.get(client_name.lower(), "")
        if not firm_code:
            unmatched += 1

        term_str = r.get("Terms", "").strip()
        billing_term, net_days = parse_term(term_str)

        rows.append({
            "office": "dallas",
            "project_code": project_code,
            "project_name": project_name,
            "client_firm_code": firm_code,
            "owning_org": "",
            "charge_type": "Billable",
            "start_date": "",
            "end_date": "",
            "contract_type": r.get("Job Type", "").strip() or "",
            "project_note": r.get("Job Description", "").strip() or "",
            "po_number": "",
            "pm_emp_code": "",
            "pic_emp_code": "",
            "pa_emp_code": "",
            "billing_term_type": billing_term,
            "net_days": net_days,
            "invoice_email": r.get("Main Email", "").strip() or "",
            "use_client_bill_to": "TRUE",
            "is_active": "TRUE",
            "_source_id": "",
        })

    if unmatched:
        print(f"  [WARN] {unmatched} projects had no matching client firm code")

    write_csv(OUTPUT_DIR / "dallas" / "dallas_Projects.csv", rows)


# ---------------------------------------------------------------------------
# Cincinnati — Ajera
# ---------------------------------------------------------------------------

_BILLING_TYPE_MAP = {
    "PercentComplete":  "Fixed Fee",
    "FixedFee":         "Fixed Fee",
    "TimeAndMaterials": "Time and Materials",
    "TimeAndExpense":   "Time and Materials",
    "CostPlus":         "Cost Plus",
    "Hourly":           "Time and Materials",
}


def extract_cincinnati():
    print("Extracting Cincinnati projects (Ajera)...")

    import sys
    import requests as _req
    sys.path.insert(0, str(HERE))
    from ajera_extract import create_session, HEADERS as AJERA_HEADERS, batched

    load_dotenv(HERE / "ajera.env")
    import os
    api_url = os.environ["AJERA_API_URL"]
    token, _ = create_session()

    # Build client key → firm_code lookup from normalized clients CSV
    cin_clients_path = OUTPUT_DIR / "cincinnati" / "cincinnati_Clients.csv"
    client_key_to_firm: dict[str, str] = {}
    if cin_clients_path.exists():
        with open(cin_clients_path, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                sid = row.get("_source_id", "").strip()
                if sid:
                    client_key_to_firm[sid] = row["FirmCode"]
    else:
        print("  [WARN] Cincinnati clients CSV not found — client_firm_code will be blank")

    # Build Ajera EmployeeKey → CIN EmpCode lookup
    emp_path = OUTPUT_DIR / "cincinnati" / "cincinnati_Employees.csv"
    emp_key_to_code: dict[str, str] = {}
    if emp_path.exists():
        with open(emp_path, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                sid = row.get("_source_id", "").strip()
                code = row.get("EmployeeCode", "").strip()
                if sid and code and sid not in emp_key_to_code:
                    emp_key_to_code[sid] = code

    # Step 1: get full project key list
    r = _req.post(api_url, json={
        "Method": "ListProjects",
        "SessionToken": token,
        "MethodArguments": {},
    }, headers=AJERA_HEADERS, timeout=60)
    r.raise_for_status()
    all_projects = r.json().get("Content", {}).get("Projects", [])
    all_keys = [p["ProjectKey"] for p in all_projects if p.get("ProjectKey")]
    print(f"  {len(all_keys)} projects to fetch (batches of 50)...")

    # Step 2: fetch detail in batches of 25 with retry + checkpoint
    import json as _json
    import time as _time

    BATCH_SIZE  = 25
    MAX_RETRIES = 4
    CHECKPOINT  = OUTPUT_DIR / "cincinnati" / "_projects_checkpoint.json"

    # Load checkpoint if resuming
    checkpoint_data: dict[str, dict] = {}
    if CHECKPOINT.exists():
        checkpoint_data = _json.loads(CHECKPOINT.read_text())
        print(f"  Resuming from checkpoint — {len(checkpoint_data)} projects already fetched")

    rows = []
    unmatched_client = 0
    fetched = 0

    for batch in batched(all_keys, BATCH_SIZE):
        # Skip keys already in checkpoint
        needed = [k for k in batch if str(k) not in checkpoint_data]
        if not needed:
            fetched += len(batch)
            continue

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                r2 = _req.post(api_url, json={
                    "Method": "GetProjects",
                    "SessionToken": token,
                    "MethodArguments": {"RequestedProjects": needed},
                }, headers=AJERA_HEADERS, timeout=180)
                r2.raise_for_status()
                for p in r2.json().get("Content", {}).get("Projects", []):
                    checkpoint_data[str(p["ProjectKey"])] = p
                CHECKPOINT.write_text(_json.dumps(checkpoint_data))
                fetched += len(needed)
                break
            except Exception as exc:
                if attempt == MAX_RETRIES:
                    print(f"  [ERROR] batch starting at key {needed[0]} failed after {MAX_RETRIES} attempts: {exc}")
                    # Save checkpoint so we can resume
                    CHECKPOINT.write_text(_json.dumps(checkpoint_data))
                    raise
                wait = 10 * attempt
                print(f"  [RETRY {attempt}/{MAX_RETRIES}] timeout on batch {needed[0]}..{needed[-1]}, waiting {wait}s")
                _time.sleep(wait)
                # Re-create session on retry
                token, _ = create_session()

        if fetched % 250 == 0 or fetched >= len(all_keys):
            print(f"  ...{fetched} / {len(all_keys)} fetched")

    # Build rows from checkpoint
    for key_str, p in checkpoint_data.items():
        pid = p.get("ID", "").strip()
        if not pid:
            continue

        ig = (p.get("InvoiceGroups") or [])
        client_info = ig[0].get("Client", {}) if ig else {}
        client_key = str(client_info.get("ClientKey", "") or "")
        firm_code = client_key_to_firm.get(client_key, "")
        if not firm_code:
            unmatched_client += 1

        pm = p.get("ProjectManager") or {}
        pm_key = str(pm.get("EmployeeKey", "") or "")
        pm_code = emp_key_to_code.get(pm_key, "")

        pic = p.get("PrincipalInCharge") or {}
        pic_key = str(pic.get("EmployeeKey", "") or "")
        pic_code = emp_key_to_code.get(pic_key, "")

        start = (p.get("ActualStartDate") or p.get("EstimatedStartDate") or "")
        end   = (p.get("ActualCompletionDate") or p.get("EstimatedCompletionDate") or "")

        billing_type  = p.get("BillingType", "") or ""
        contract_type = _BILLING_TYPE_MAP.get(billing_type, billing_type)

        status    = p.get("Status", "") or ""
        is_active_val = "TRUE" if status.lower() in ("open", "active", "") else "FALSE"

        dept = p.get("DepartmentDescription", "") or ""

        rows.append({
            "office":             "cincinnati",
            "project_code":       pid,
            "project_name":       p.get("Description", "").strip(),
            "client_firm_code":   firm_code,
            "owning_org":         "",
            "charge_type":        "Billable" if status.lower() != "internal" else "Indirect",
            "start_date":         start[:10].replace("-", "/") if start else "",
            "end_date":           end[:10].replace("-", "/")   if end   else "",
            "contract_type":      contract_type,
            "project_note":       f"{dept} | {p.get('Notes','')[:200]}".strip(" |") if dept or p.get("Notes") else "",
            "po_number":          "",
            "pm_emp_code":        pm_code,
            "pic_emp_code":       pic_code,
            "pa_emp_code":        "",
            "billing_term_type":  "",
            "net_days":           "",
            "invoice_email":      "",
            "use_client_bill_to": "TRUE",
            "is_active":          is_active_val,
            "_source_id":         str(p.get("ProjectKey", "")),
        })

    if unmatched_client:
        print(f"  [WARN] {unmatched_client} projects had no matching client firm code")

    write_csv(OUTPUT_DIR / "cincinnati" / "cincinnati_Projects.csv", rows)
    # Clean up checkpoint on success
    if CHECKPOINT.exists():
        CHECKPOINT.unlink()


# ---------------------------------------------------------------------------
# Orlando — QB Desktop transaction detail export
# ---------------------------------------------------------------------------

_ORL_CODE_RE = re.compile(r"^(\d{4,5}(?:\.\d+)?)\s+(.+)$")


def extract_orlando():
    """
    Parse input/Orlando_Projects.CSV — a QB Desktop transaction detail export.

    Format: project header row (col 0 = "YYNN ProjectName", rest NaN)
            followed by transaction rows (col 0 = NaN, Date/Type/Amount populated).

    Derives start_date/end_date from first/last Invoice date in each block.
    Entries without a numeric YYNN prefix are skipped (client/retainer names).
    """
    import pandas as pd

    INPUT_FILE = HERE.parent / "input" / "Orlando_Projects.CSV"
    if not INPUT_FILE.exists():
        print(f"[SKIP] orlando: {INPUT_FILE} not found")
        return

    print("Extracting Orlando projects (QB Desktop transaction export)...")

    df = pd.read_csv(INPUT_FILE, header=0)
    name_col = df.columns[0]  # 'Unnamed: 0'

    # Find project header row indices (non-blank first column)
    hdr_mask = df[name_col].notna() & (df[name_col].str.strip() != "")
    hdr_indices = df[hdr_mask].index.tolist()

    rows = []
    placeholder_seq = 0

    for i, idx in enumerate(hdr_indices):
        raw_name = str(df.at[idx, name_col]).strip()
        m = _ORL_CODE_RE.match(raw_name)
        if m:
            raw_code  = m.group(1)   # e.g. "1901", "2020.16", "25008"
            proj_name = m.group(2).strip()
            proj_code = f"ORL-{raw_code}"
        else:
            # No numeric prefix — create a placeholder code
            placeholder_seq += 1
            raw_code  = f"MISC{placeholder_seq:03d}"
            proj_name = raw_name
            proj_code = f"ORL-{raw_code}"
            print(f"  [PLACEHOLDER] {proj_code} = {proj_name}")

        # Transaction rows for this project
        next_idx = hdr_indices[i + 1] if i + 1 < len(hdr_indices) else len(df)
        block = df.iloc[idx + 1 : next_idx]
        invoices = block[block["Type"] == "Invoice"]

        start_date, end_date = "", ""
        if not invoices.empty:
            dates = pd.to_datetime(invoices["Date"], errors="coerce").dropna()
            if not dates.empty:
                start_date = dates.min().date().isoformat()
                end_date   = dates.max().date().isoformat()

        # Active if any invoice in last 18 months (or no date info — assume active)
        is_active = "TRUE"
        if end_date:
            from datetime import date, timedelta
            cutoff = (date.today() - timedelta(days=548)).isoformat()
            is_active = "TRUE" if end_date >= cutoff else "FALSE"

        rows.append({
            "office":            "orlando",
            "project_code":      proj_code,
            "project_name":      proj_name,
            "client_firm_code":  "",
            "owning_org":        "",
            "charge_type":       "Billable",
            "start_date":        start_date,
            "end_date":          end_date,
            "contract_type":     "",
            "project_note":      "",
            "po_number":         "",
            "pm_emp_code":       "",
            "pic_emp_code":      "",
            "pa_emp_code":       "",
            "billing_term_type": "",
            "net_days":          "",
            "invoice_email":     "",
            "use_client_bill_to":"TRUE",
            "is_active":         is_active,
            "_source_id":        raw_code,
        })

    print(f"  {len(rows)} projects extracted ({placeholder_seq} with placeholder codes)")

    write_csv(OUTPUT_DIR / "orlando" / "orlando_Projects.csv", rows)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

EXTRACTORS = {
    "minnesota":  extract_minnesota,
    "dallas":     extract_dallas,
    "cincinnati": extract_cincinnati,
    "orlando":    extract_orlando,
}

UNSUPPORTED = {}


def main():
    parser = argparse.ArgumentParser(description="Extract projects from all source systems")
    parser.add_argument("--office", choices=list(EXTRACTORS) + list(UNSUPPORTED), default=None)
    args = parser.parse_args()

    if args.office:
        EXTRACTORS[args.office]()
    else:
        for fn in EXTRACTORS.values():
            fn()



if __name__ == "__main__":
    load_dotenv(HERE / ".env")
    main()
