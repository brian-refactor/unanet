"""
Load employees from the Unanet HR Excel file into Supabase.

Usage:
    python etl/load_employees.py
    python etl/load_employees.py --dry-run
"""
import argparse
import os
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from supabase import create_client

load_dotenv(Path(__file__).parent / ".env")
sb = create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_KEY"])

SOURCE_FILE = (
    Path(__file__).parent.parent
    / "Unanet-Migration-Files/01. Data - Initial Pass"
    / "05a-Employees_Fusion6.11.2026.xlsx"
)

ORG_TO_OFFICE = {
    "MSP": "minnesota",
    "CIN": "cincinnati",
    "DAL": "dallas",
    "ORL": "orlando",
    "CORP": "corporate",
}


def derive_office(org_path: str) -> str:
    """Map org_path (e.g. FUS-MSP-A01) to office slug."""
    if not org_path:
        return "unknown"
    parts = str(org_path).split("-")
    # Format is FUS-<OFFICE>-<UNIT> — second segment is the office key
    if len(parts) >= 2:
        key = parts[1].upper()
        return ORG_TO_OFFICE.get(key, key.lower())
    return org_path.lower()


def clean_date(val):
    """Return ISO date string or None for blank/whitespace/NaT values."""
    if val is None:
        return None
    s = str(val).strip()
    if not s or s in ("NaT", "nan", "\xa0"):
        return None
    try:
        return pd.to_datetime(val).date().isoformat()
    except Exception:
        return None


def clean_bool(val, default=False) -> bool:
    if val is None:
        return default
    s = str(val).strip().lower()
    return s in ("true", "1", "yes")


def build_full_name(first, mid, last) -> str:
    parts = [p for p in [first, mid, last] if p and str(p).strip() not in ("nan", "")]
    return " ".join(str(p).strip() for p in parts)


def load(dry_run: bool = False):
    print(f"Reading {SOURCE_FILE.name}...")
    df = pd.read_excel(SOURCE_FILE, header=2)
    df = df[df["EmployeeCode"].notna()].copy()
    print(f"  {len(df)} employee rows found")

    rows = []
    for _, r in df.iterrows():
        emp_code = str(r["EmployeeCode"]).strip()
        first    = str(r["FirstName"]).strip() if pd.notna(r["FirstName"]) else ""
        last     = str(r["LastName"]).strip()  if pd.notna(r["LastName"])  else ""
        mid      = str(r["MidName"]).strip()   if pd.notna(r["MidName"])   else None
        if mid == "nan":
            mid = None
        org_path = str(r["Org"]).strip() if pd.notna(r["Org"]) else ""
        office   = derive_office(org_path)

        full_name = build_full_name(first, mid, last)

        email = str(r["WorkEmail"]).strip() if pd.notna(r["WorkEmail"]) else None
        if email == "nan":
            email = None

        job_title_code = str(r["JobTitleCode"]).strip() if pd.notna(r["JobTitleCode"]) else None
        if job_title_code == "nan":
            job_title_code = None
        job_title_name = str(r["JobTitleName"]).strip() if pd.notna(r["JobTitleName"]) else None
        if job_title_name == "nan":
            job_title_name = None

        rows.append({
            "office":           office,
            "record_key":       emp_code,
            "employee_code":    emp_code,
            "first_name":       first,
            "last_name":        last,
            "middle_name":      mid,
            "full_name":        full_name,
            "email":            email,
            "org_path":         org_path,
            "hire_date":        clean_date(r.get("HireDate")),
            "termination_date": clean_date(r.get("TerminationDate")),
            "is_active":        clean_bool(r.get("Active"), default=True),
            "timesheet_group":  str(r["TimesheetGroup"]).strip() if pd.notna(r.get("TimesheetGroup")) else None,
            "pay_group":        str(r["PayrollGroup"]).strip()   if pd.notna(r.get("PayrollGroup"))   else None,
            "job_type":         str(r["EmployeeJobType"]).strip() if pd.notna(r.get("EmployeeJobType")) else None,
            "job_title_code":   job_title_code,
            "job_title_name":   job_title_name,
            "is_subcontractor": clean_bool(r.get("Subcontractor"), default=False),
            "pay_rate":         float(r["JCRate"])   if pd.notna(r.get("JCRate"))   else None,
            "billing_rate":     float(r["BillRate"]) if pd.notna(r.get("BillRate")) else None,
            "target_pct":       float(r["TargetPct"]) if pd.notna(r.get("TargetPct")) else None,
        })

    print(f"\nOffice distribution:")
    from collections import Counter
    for office, count in sorted(Counter(r["office"] for r in rows).items()):
        print(f"  {office}: {count}")

    if dry_run:
        print("\n[DRY RUN] No changes written.")
        print("Sample row:")
        import json
        print(json.dumps(rows[0], indent=2, default=str))
        return

    print(f"\nTruncating employees table...")
    sb.table("employees").delete().neq("employee_code", "__never__").execute()

    print(f"Inserting {len(rows)} rows...")
    batch_size = 100
    for i in range(0, len(rows), batch_size):
        batch = rows[i:i + batch_size]
        sb.table("employees").insert(batch).execute()
        print(f"  {min(i + batch_size, len(rows))} / {len(rows)}")

    final = sb.table("employees").select("employee_code", count="exact").execute()
    print(f"\nDone. {final.count} employees in Supabase.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    load(dry_run=args.dry_run)
