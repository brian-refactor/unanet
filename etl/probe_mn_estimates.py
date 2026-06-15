"""
Probe QBO Estimates for Minnesota — check if they contain phase fee/date data.

Usage:
    python etl/probe_mn_estimates.py
"""
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

from dotenv import load_dotenv
load_dotenv(HERE / ".env")

from qbo_extract import get_client, query_all
from quickbooks.objects.customer import Customer
from quickbooks.objects.estimate import Estimate

qb, _, _ = get_client()
print("Connected to QBO\n")

customers = query_all(Customer, qb)
id_to_cust = {c.Id: c for c in customers}
print(f"Customers loaded: {len(customers)}")

estimates = query_all(Estimate, qb)
print(f"Total estimates: {len(estimates)}\n")

if not estimates:
    print("No estimates found.")
    sys.exit(0)

# ── Basic stats ──────────────────────────────────────────────────────────────
proj_estimates = defaultdict(list)  # project DisplayName → list of estimates
statuses = defaultdict(int)

for est in estimates:
    cust_ref = getattr(est, "CustomerRef", None)
    if not cust_ref:
        continue
    cust = id_to_cust.get(cust_ref.value)
    status = getattr(est, "TxnStatus", "Unknown")
    statuses[status] += 1
    if cust and getattr(cust, "ParentRef", None):  # sub-customer = project
        proj_estimates[cust.DisplayName].append(est)

print("Estimate status breakdown:")
for status, count in sorted(statuses.items(), key=lambda x: -x[1]):
    print(f"  {status}: {count}")

print(f"\nEstimates against projects (sub-customers): {sum(len(v) for v in proj_estimates.values())}")
print(f"Distinct projects with estimates: {len(proj_estimates)}")

# ── Inspect a sample ─────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("SAMPLE ESTIMATES (first 5 projects with estimates)")
print("=" * 60)

for proj_name, ests in list(proj_estimates.items())[:5]:
    print(f"\nProject: {proj_name}  ({len(ests)} estimate(s))")
    for est in ests[:2]:
        txn_date  = getattr(est, "TxnDate", None)
        exp_date  = getattr(est, "ExpirationDate", None)
        status    = getattr(est, "TxnStatus", None)
        total_amt = getattr(est, "TotalAmt", None)
        print(f"  Estimate #{getattr(est, 'DocNumber', '?')}  status={status}  "
              f"date={txn_date}  expires={exp_date}  total=${total_amt}")

        for line in getattr(est, "Line", []) or []:
            detail  = getattr(line, "SalesItemLineDetail", None)
            desc    = getattr(line, "Description", "") or ""
            amount  = getattr(line, "Amount", None)
            qty     = getattr(detail, "Qty", None) if detail else None
            rate    = getattr(detail, "UnitPrice", None) if detail else None
            item    = getattr(getattr(detail, "ItemRef", None), "name", "") if detail else ""
            svc_dt  = getattr(detail, "ServiceDate", None) if detail else None
            if item or amount:
                print(f"    Line: item={item!r:40s} desc={desc[:40]!r}  "
                      f"qty={qty}  rate={rate}  amt={amount}  svc_date={svc_dt}")

# ── Check what fields have data across all estimates ─────────────────────────
print("\n" + "=" * 60)
print("FIELD COVERAGE across all project estimates")
print("=" * 60)

has_expiry = sum(1 for e in estimates if getattr(e, "ExpirationDate", None))
has_ship   = sum(1 for e in estimates if getattr(e, "ShipDate", None))
has_custom = sum(1 for e in estimates if getattr(e, "CustomField", None))

total_proj_ests = sum(len(v) for v in proj_estimates.values())
print(f"Project estimates with ExpirationDate:  {has_expiry}")
print(f"Project estimates with ShipDate:        {has_ship}")
print(f"Project estimates with CustomField:     {has_custom}")

# Line-item field coverage
line_fields = defaultdict(int)
item_amounts: dict[str, list] = defaultdict(list)

for est in estimates:
    cust_ref = getattr(est, "CustomerRef", None)
    if not cust_ref:
        continue
    cust = id_to_cust.get(cust_ref.value)
    if not (cust and getattr(cust, "ParentRef", None)):
        continue
    for line in getattr(est, "Line", []) or []:
        detail = getattr(line, "SalesItemLineDetail", None)
        if not detail:
            continue
        item_name = getattr(getattr(detail, "ItemRef", None), "name", "") or ""
        amount    = getattr(line, "Amount", None)
        qty       = getattr(detail, "Qty", None)
        rate      = getattr(detail, "UnitPrice", None)
        svc_date  = getattr(detail, "ServiceDate", None)
        if item_name: line_fields["item_name"] += 1
        if amount:    line_fields["amount"] += 1
        if qty:       line_fields["qty"] += 1
        if rate:      line_fields["rate"] += 1
        if svc_date:  line_fields["service_date"] += 1
        if amount and item_name:
            parent = item_name.split(":")[0].strip()
            item_amounts[parent].append(float(amount))

print("\nLine item field population:")
for field, count in sorted(line_fields.items(), key=lambda x: -x[1]):
    print(f"  {field}: {count} lines")

if item_amounts:
    print("\nPhase fee totals across all estimates (sum of line amounts by parent item):")
    for phase, amounts in sorted(item_amounts.items(), key=lambda x: -sum(x[1])):
        print(f"  {phase:<40s}  {len(amounts):4d} lines   total=${sum(amounts):>12,.0f}   "
              f"avg=${sum(amounts)/len(amounts):>8,.0f}")
