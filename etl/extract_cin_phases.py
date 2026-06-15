"""
Extract project phases from Ajera v1 API for Cincinnati.

Reads all project keys from the existing cincinnati_Projects.csv,
batches GetProjects calls (20 per batch), flattens the nested
InvoiceGroup → Phase → sub-Phase hierarchy, and loads to Supabase.

Usage:
    python etl/extract_cin_phases.py
    python etl/extract_cin_phases.py --dry-run
    python etl/extract_cin_phases.py --limit 100
    python etl/extract_cin_phases.py --from-cache   # re-parse saved JSON without API calls
"""
import argparse
import json
import os
import time
from pathlib import Path

import pandas as pd
import requests
from dotenv import load_dotenv
from supabase import create_client

HERE = Path(__file__).parent
OUTPUT_DIR = HERE.parent / "output" / "cincinnati"
CACHE_FILE = OUTPUT_DIR / "_phases_raw.json"

load_dotenv(HERE / "ajera.env")
load_dotenv(HERE / ".env")

AJERA_URL  = os.environ["AJERA_API_URL"]
AJERA_USER = os.environ["AJERA_USERNAME"]
AJERA_PASS = os.environ["AJERA_PASSWORD"]
HEADERS    = {"Content-Type": "application/json"}
OFFICE     = "cincinnati"
BATCH_SIZE = 20
DEFAULT_ORG = "FUS-CIN-A01"

BILLING_TYPE_MAP = {
    "PercentComplete": "Billable",
    "TimeAndExpense":  "Billable",
    "Nonbillable":     "Non-Billable",
    "Marketing":       "Non-Billable",
    "FixedFee":        "Billable",
}


# ---------------------------------------------------------------------------
# Ajera session helpers
# ---------------------------------------------------------------------------

def create_session() -> str:
    resp = requests.post(AJERA_URL, json={
        "Method": "CreateAPISession",
        "Username": AJERA_USER,
        "Password": AJERA_PASS,
        "APIVersion": 1,
        "UseSessionCookie": False,
    }, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    token = data.get("Content", {}).get("SessionToken")
    if not token:
        raise SystemExit(f"Ajera auth failed: {data.get('Errors')}")
    print(f"Connected — {data['Content'].get('CompanyName', '')}  "
          f"(Ajera {data['Content'].get('AjeraVersion', '')})")
    return token


def end_session(token: str):
    try:
        requests.post(AJERA_URL, json={
            "Method": "EndAPISession", "SessionToken": token,
        }, headers=HEADERS, timeout=15)
    except Exception:
        pass


def fetch_batch(token: str, keys: list[int], attempt: int = 0) -> list[dict]:
    """Call GetProjects for a batch of keys; retry up to 3x on timeout."""
    try:
        resp = requests.post(AJERA_URL, json={
            "Method": "GetProjects",
            "SessionToken": token,
            "MethodArguments": {"RequestedProjects": keys},
        }, headers=HEADERS, timeout=180)
        resp.raise_for_status()
        data = resp.json()
        errors = [e for e in data.get("Errors", []) if e.get("ErrorID", 0) != 0]
        if errors:
            print(f"    [WARN] {errors}")
        return data.get("Content", {}).get("Projects", [])
    except (requests.Timeout, requests.ConnectionError) as e:
        if attempt < 3:
            wait = (attempt + 1) * 15
            print(f"    [RETRY {attempt+1}/3] {e}, waiting {wait}s...")
            time.sleep(wait)
            new_token = create_session()
            return fetch_batch(new_token, keys, attempt + 1)
        raise


# ---------------------------------------------------------------------------
# Phase parsing helpers
# ---------------------------------------------------------------------------

def clean_date(val) -> str | None:
    if not val:
        return None
    s = str(val).strip()
    if not s or s == "null":
        return None
    try:
        return pd.to_datetime(s).date().isoformat()
    except Exception:
        return None


def coerce_amount(val) -> float | None:
    try:
        f = float(val)
        return f if f != 0.0 else None
    except (TypeError, ValueError):
        return None


def flatten_phases(project: dict, org_path: str) -> list[dict]:
    """
    Walk InvoiceGroups → Phases → sub-Phases and produce flat rows.

    Unanet supports two levels (L2 + optional L3). For deeper nesting
    (L2 billing group → L3 sub-group → L4 leaf) we promote the sub-group
    children directly to L3 under the top-level billing group, logging a
    warning. The L2 code is globally sequential across all InvoiceGroups.
    """
    project_code = project.get("ID", "")
    rows = []
    l2_counter = 0

    for ig in project.get("InvoiceGroups", []):
        for l2 in ig.get("Phases", []):
            l2_counter += 1
            l2_code = f"{l2_counter:03d}"
            l2_name = l2.get("Description", "")
            billing = BILLING_TYPE_MAP.get(l2.get("BillingType", ""), "Billable")
            l2_start = clean_date(l2.get("ActualStartDate") or l2.get("EstimatedStartDate"))
            l2_end   = clean_date(l2.get("ActualCompletionDate") or l2.get("EstimatedCompletionDate"))

            base = {
                "office":            OFFICE,
                "project_code":      project_code,
                "contract_type":     billing,
                "level2_name":       l2_name,
                "level2_code":       l2_code,
                "start_date":        l2_start,
                "end_date":          l2_end,
                "org_path":          org_path,
                "fixed_fee":         coerce_amount(l2.get("TotalContractAmount")),
                "labor_contract_cap":coerce_amount(l2.get("LaborContractAmount")),
                "odc_contract_cap":  coerce_amount(l2.get("ExpenseContractAmount")),
                "occ_contract_cap":  None,
                "icc_fixed_fee":     coerce_amount(l2.get("ConsultantContractAmount")),
                "labor_budget":      coerce_amount(l2.get("LaborCostBudget")),
                "odc_budget":        coerce_amount(l2.get("ExpenseCostBudget")),
                "occ_budget":        None,
                "icc_budget":        coerce_amount(l2.get("ConsultantCostBudget")),
                "hours_budget":      coerce_amount(l2.get("HoursCostBudget")),
            }

            sub = l2.get("Phases", [])
            if not sub:
                rows.append({**base, "level3_name": None, "level3_code": None})
                continue

            # Collect leaf L3 rows — handling up to one extra nesting level
            l3_counter = 0
            for l3 in sub:
                grandchildren = l3.get("Phases", [])
                if grandchildren:
                    # Promote grandchildren as L3 under this L2
                    for gc in grandchildren:
                        l3_counter += 1
                        rows.append({
                            **base,
                            "level3_name": gc.get("Description", ""),
                            "level3_code": f"{l3_counter:04d}",
                        })
                else:
                    l3_counter += 1
                    rows.append({
                        **base,
                        "level3_name": l3.get("Description", ""),
                        "level3_code": f"{l3_counter:04d}",
                    })

    return rows


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run",    action="store_true",
                        help="Parse and print stats, no DB writes")
    parser.add_argument("--limit",      type=int, default=None,
                        help="Only process first N projects (for testing)")
    parser.add_argument("--from-cache", action="store_true",
                        help="Re-parse _phases_raw.json without calling Ajera")
    args = parser.parse_args()

    # Load project key → (project_code, org_path) mapping from CSV
    csv_path = OUTPUT_DIR / "cincinnati_Projects.csv"
    proj_df  = pd.read_csv(csv_path)
    proj_df["_source_id"] = pd.to_numeric(proj_df["_source_id"], errors="coerce").dropna().astype(int)
    proj_df  = proj_df.dropna(subset=["_source_id"])

    key_to_code = {
        int(r["_source_id"]): r["project_code"]
        for _, r in proj_df.iterrows()
    }
    key_to_org = {
        int(r["_source_id"]): (
            r["owning_org"]
            if pd.notna(r.get("owning_org")) and str(r["owning_org"]).strip()
            else DEFAULT_ORG
        )
        for _, r in proj_df.iterrows()
    }

    all_keys = sorted(key_to_code.keys())
    if args.limit:
        all_keys = all_keys[:args.limit]

    print(f"Projects to process: {len(all_keys)}")

    # --- Fetch from API or load cache ---
    if args.from_cache:
        print(f"Loading from cache: {CACHE_FILE}")
        with open(CACHE_FILE, encoding="utf-8") as f:
            all_projects = json.load(f)
        print(f"  {len(all_projects)} projects in cache")
    else:
        token = create_session()
        all_projects = []
        total_batches = (len(all_keys) + BATCH_SIZE - 1) // BATCH_SIZE

        for batch_num, start in enumerate(range(0, len(all_keys), BATCH_SIZE), 1):
            batch = all_keys[start:start + BATCH_SIZE]
            print(f"  Batch {batch_num}/{total_batches} "
                  f"(keys {batch[0]}..{batch[-1]})...")
            projects = fetch_batch(token, batch)
            all_projects.extend(projects)
            if batch_num % 10 == 0:
                print(f"  ...{len(all_projects)} projects fetched so far")

        end_session(token)

        # Save raw cache
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(all_projects, f)
        print(f"Raw data cached → {CACHE_FILE}")

    # --- Parse phases ---
    print("\nParsing phases...")
    all_rows = []
    no_phases = []

    for proj in all_projects:
        pk   = proj.get("ProjectKey")
        code = proj.get("ID") or key_to_code.get(pk, "")
        org  = key_to_org.get(pk, DEFAULT_ORG)
        rows = flatten_phases(proj, org)
        if rows:
            all_rows.extend(rows)
        else:
            no_phases.append(code)

    print(f"  Phase rows extracted: {len(all_rows)}")
    print(f"  Projects with no phases: {len(no_phases)}")
    if no_phases[:10]:
        print(f"  Sample no-phase projects: {no_phases[:10]}")

    # Deduplicate by the unique constraint key
    seen: set[tuple] = set()
    deduped: list[dict] = []
    dupes: list[tuple] = []
    for row in all_rows:
        key = (row["office"], row["project_code"], row["level2_code"], row["level3_code"])
        if key in seen:
            dupes.append(key)
        else:
            seen.add(key)
            deduped.append(row)
    if dupes:
        print(f"  [WARN] {len(dupes)} duplicate phase rows dropped: {dupes[:5]}")
    all_rows = deduped

    if args.dry_run:
        print("\n[DRY RUN] — no DB writes.")
        print("Sample rows:")
        for r in all_rows[:3]:
            print(json.dumps(r, indent=2, default=str))
        return

    # --- Load to Supabase ---
    sb = create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_KEY"])

    print(f"\nClearing existing CIN phase rows...")
    sb.table("project_phases").delete().eq("office", OFFICE).execute()

    print(f"Inserting {len(all_rows)} rows (batches of 500)...")
    for i in range(0, len(all_rows), 500):
        batch = all_rows[i:i + 500]
        sb.table("project_phases").insert(batch).execute()
        print(f"  {min(i + 500, len(all_rows))} / {len(all_rows)}")

    final = sb.table("project_phases").select("id", count="exact").eq("office", OFFICE).execute()
    print(f"\nDone. {final.count} phase rows in Supabase for Cincinnati.")


if __name__ == "__main__":
    main()
