"""
Extract project phases for Minnesota from QBO invoice line items + estimates.

Phase list comes from invoice history (which phases were ever billed).
Fixed fees come from QBO Estimates (where available) — the parent-level line
amount per phase is used as fixed_fee. When a project has multiple estimates
the best one is selected: Accepted > Converted > most recent by TxnDate.

312 of 897 MN projects have estimate data. Projects with no invoice history
(≈214) get no phase rows.

Usage:
    python etl/extract_mn_phases.py
    python etl/extract_mn_phases.py --dry-run
"""
import argparse
import os
import re
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from supabase import create_client

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

load_dotenv(HERE / ".env")

OFFICE     = "minnesota"
OUTPUT_DIR = HERE.parent / "output" / "minnesota"

# ---------------------------------------------------------------------------
# Phase catalogue — ordered for display; drives level2_code assignment
# ---------------------------------------------------------------------------
# (qbo_parent_name, level2_code, level2_name, contract_type)
PHASE_CATALOGUE = [
    ("A0 General",                    "A0",  "General",                    "Billable"),
    ("A1 Pre-Design",                 "A1",  "Pre-Design",                 "Billable"),
    ("A2 Schematic Design",           "A2",  "Schematic Design",           "Billable"),
    ("A3 Design Development",         "A3",  "Design Development",         "Billable"),
    ("A4 Construction Document",      "A4",  "Construction Documents",     "Billable"),
    ("A5 Bid/Negotiate",              "A5",  "Bid/Negotiate",              "Billable"),
    ("A6 Construction Administration","A6",  "Construction Administration","Billable"),
    ("A7 Supplemental Services",      "A7",  "Supplemental Services",      "Billable"),
    ("A80 Lump Sum Fee",              "A80", "Lump Sum Fee",               "Billable"),
    ("A9 Hourly Service Contracts",   "A9",  "Hourly Service Contracts",   "Billable"),
    ("M0 Marketing General",          "M0",  "Marketing General",          "Non-Billable"),
    ("M1 Marketing Predesign",        "M1",  "Marketing Pre-Design",       "Non-Billable"),
    ("Procurement FF&E",              "FFE", "Procurement FF&E",           "Billable"),
    ("Reimbursable Expenses",         "RE",  "Reimbursable Expenses",      "Billable"),
]

# Map QBO parent name → (code, display name, contract_type)
PHASE_MAP = {row[0]: (row[1], row[2], row[3]) for row in PHASE_CATALOGUE}
# Ordered list of QBO parent names for sorting phases within a project
PHASE_ORDER = [row[0] for row in PHASE_CATALOGUE]

# Items to discard — billing adjustments, retainers, legacy "Completed" items
SKIP_PATTERNS = [
    re.compile(r"^PmntDiscount_"),
    re.compile(r"^Completed "),
    re.compile(r"^Cash Adjustment$"),
    re.compile(r"^Fin Chg$"),
    re.compile(r"^Sales$"),
    re.compile(r"^Client Retainers"),
]


def should_skip(name: str) -> bool:
    return any(p.match(name) for p in SKIP_PATTERNS)


def parent_phase(item_name: str) -> str | None:
    """
    Extract the parent phase name from a QBO item name.
    QBO formats sub-items as 'Parent:Child', e.g. 'A2 Schematic Design:A220 PM'.
    For top-level items (no colon) the item itself is the parent.
    Returns None for items we don't map to a phase.
    """
    parent = item_name.split(":")[0].strip()
    if should_skip(parent):
        return None
    if parent in PHASE_MAP:
        return parent
    # Some invoices use the sub-item alone without a parent prefix —
    # try to find which catalogue entry it belongs to by code prefix.
    # (e.g. a bare "A220 Project Manager" → "A2 Schematic Design")
    for catalogue_parent in PHASE_ORDER:
        code = PHASE_MAP[catalogue_parent][0]  # e.g. "A2"
        if item_name.startswith(code) and len(code) < len(item_name):
            return catalogue_parent
    return None  # unmapped item (Sales, Fin Chg, etc.)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true",
                        help="Print stats, no DB writes")
    args = parser.parse_args()

    # ── Load project lookup: QBO customer Id → project_code ─────────────────
    csv_path = OUTPUT_DIR / "minnesota_Projects.csv"
    proj_df  = pd.read_csv(csv_path)
    source_to_code = {
        str(r["_source_id"]): r["project_code"]
        for _, r in proj_df.iterrows()
        if pd.notna(r.get("_source_id"))
    }
    print(f"MN projects in CSV: {len(source_to_code)}")

    # ── Connect to QBO ───────────────────────────────────────────────────────
    from qbo_extract import get_client, query_all
    from quickbooks.objects.customer import Customer
    from quickbooks.objects.estimate import Estimate
    from quickbooks.objects.invoice import Invoice

    qb, _, _ = get_client()
    print("Connected to QBO")

    customers = query_all(Customer, qb)
    id_to_cust = {c.Id: c for c in customers}
    print(f"Customers fetched: {len(customers)}")

    invoices = query_all(Invoice, qb)
    print(f"Invoices fetched: {len(invoices)}")

    estimates = query_all(Estimate, qb)
    print(f"Estimates fetched: {len(estimates)}")

    # ── Collect phases per project ───────────────────────────────────────────
    # project_code → set of QBO parent phase names
    project_phases: dict[str, set[str]] = defaultdict(set)
    skipped_items: set[str] = set()
    unknown_items: set[str] = set()

    for inv in invoices:
        cust_ref = getattr(inv, "CustomerRef", None)
        if not cust_ref:
            continue
        cust = id_to_cust.get(cust_ref.value)
        if not cust or not getattr(cust, "ParentRef", None):
            continue  # top-level customer = client, not a project

        proj_code = source_to_code.get(str(cust_ref.value))
        if not proj_code:
            continue  # project not in our CSV (shouldn't happen)

        for line in getattr(inv, "Line", []) or []:
            detail = getattr(line, "SalesItemLineDetail", None)
            if not detail:
                continue
            item_ref = getattr(detail, "ItemRef", None)
            if not item_ref:
                continue
            item_name = getattr(item_ref, "name", "") or str(item_ref.value)

            if should_skip(item_name.split(":")[0].strip()):
                skipped_items.add(item_name)
                continue

            phase = parent_phase(item_name)
            if phase:
                project_phases[proj_code].add(phase)
            else:
                unknown_items.add(item_name)

    print(f"\nProjects with at least one phase: {len(project_phases)}")
    print(f"Projects with no invoice history: "
          f"{len(source_to_code) - len(project_phases)}")
    if unknown_items:
        print(f"Unmapped items (ignored): {sorted(unknown_items)}")

    # ── Collect phase fixed fees from estimates ───────────────────────────────
    # For each project, pick the best estimate (Accepted > Converted > most recent)
    # then read parent-level line amounts as fixed_fee per phase.
    STATUS_RANK = {"Accepted": 0, "Converted": 1, "Pending": 2, "Closed": 3}

    # project_code → {parent_phase_name: amount}
    estimate_fees: dict[str, dict[str, float]] = {}

    # Group estimates by project, pick best
    proj_est_candidates: dict[str, list] = defaultdict(list)
    for est in estimates:
        cust_ref = getattr(est, "CustomerRef", None)
        if not cust_ref:
            continue
        cust = id_to_cust.get(cust_ref.value)
        if not cust or not getattr(cust, "ParentRef", None):
            continue
        proj_code = source_to_code.get(str(cust_ref.value))
        if not proj_code:
            continue
        proj_est_candidates[proj_code].append(est)

    for proj_code, ests in proj_est_candidates.items():
        best = sorted(
            ests,
            key=lambda e: (
                STATUS_RANK.get(getattr(e, "TxnStatus", ""), 9),
                -(int(getattr(e, "TxnDate", "0000-00-00").replace("-", "")) if getattr(e, "TxnDate", None) else 0),
            ),
        )[0]

        fees: dict[str, float] = {}
        for line in getattr(best, "Line", []) or []:
            detail   = getattr(line, "SalesItemLineDetail", None)
            if not detail:
                continue
            item_ref = getattr(detail, "ItemRef", None)
            if not item_ref:
                continue
            item_name = getattr(item_ref, "name", "") or str(item_ref.value)
            amount    = getattr(line, "Amount", None)
            if not amount:
                continue
            phase = parent_phase(item_name)
            if phase:
                fees[phase] = fees.get(phase, 0.0) + float(amount)

        if fees:
            estimate_fees[proj_code] = fees

    print(f"Projects with estimate fee data: {len(estimate_fees)}")

    # ── Build rows ───────────────────────────────────────────────────────────
    rows = []
    for proj_code, phase_set in sorted(project_phases.items()):
        # Sort phases by catalogue order
        ordered = [p for p in PHASE_ORDER if p in phase_set]
        # Any phases not in catalogue order go at the end alphabetically
        ordered += sorted(p for p in phase_set if p not in PHASE_ORDER)

        proj_fees = estimate_fees.get(proj_code, {})

        for seq, phase_name in enumerate(ordered, start=1):
            _, display_name, contract_type = PHASE_MAP[phase_name]
            fee = proj_fees.get(phase_name)
            rows.append({
                "office":            OFFICE,
                "project_code":      proj_code,
                "contract_type":     contract_type,
                "level2_name":       display_name,
                "level2_code":       f"{seq:03d}",
                "level3_name":       None,
                "level3_code":       None,
                "start_date":        None,
                "end_date":          None,
                "org_path":          None,
                "fixed_fee":         fee,
                "labor_contract_cap":None,
                "odc_contract_cap":  None,
                "occ_contract_cap":  None,
                "icc_fixed_fee":     None,
                "labor_budget":      None,
                "odc_budget":        None,
                "occ_budget":        None,
                "icc_budget":        None,
                "hours_budget":      None,
            })

    print(f"Phase rows to insert: {len(rows)}")

    if args.dry_run:
        print("\n[DRY RUN] — no DB writes.")
        # Show phase distribution
        from collections import Counter
        phase_counts = Counter(r["level2_code"] for r in rows)
        print("\nPhase usage across projects:")
        for code, count in sorted(phase_counts.items(), key=lambda x: -x[1]):
            name = next(r["level2_name"] for r in rows if r["level2_code"] == code)
            print(f"  {code:4s}  {name:<32s}  {count} projects")
        return

    # ── Load to Supabase ─────────────────────────────────────────────────────
    sb = create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_KEY"])

    print(f"\nClearing existing MN phase rows...")
    sb.table("project_phases").delete().eq("office", OFFICE).execute()

    print(f"Inserting {len(rows)} rows (batches of 500)...")
    for i in range(0, len(rows), 500):
        batch = rows[i:i + 500]
        sb.table("project_phases").insert(batch).execute()
        print(f"  {min(i + 500, len(rows))} / {len(rows)}")

    final = sb.table("project_phases").select("id", count="exact").eq("office", OFFICE).execute()
    print(f"\nDone. {final.count} phase rows in Supabase for Minnesota.")


if __name__ == "__main__":
    main()
