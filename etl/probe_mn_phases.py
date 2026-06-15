"""
Probe QBO for Minnesota phase data — two angles:
  1. Sub-sub-customers (3rd level in customer hierarchy)
  2. Invoice line items grouped by project — service items used per project

Usage:
    python etl/probe_mn_phases.py
"""
import json
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

from dotenv import load_dotenv
load_dotenv(HERE / ".env")

from qbo_extract import get_client, query_all
from quickbooks.objects.customer import Customer
from quickbooks.objects.invoice import Invoice

qb, _, _ = get_client()
print("Connected to QBO\n")

# ── 1. Sub-sub-customers ─────────────────────────────────────────────────────
print("=" * 60)
print("1. CUSTOMER HIERARCHY DEPTH")
print("=" * 60)

customers = query_all(Customer, qb)
id_to_cust = {c.Id: c for c in customers}

top_level   = [c for c in customers if not getattr(c, "ParentRef", None)]
sub_level   = [c for c in customers if getattr(c, "ParentRef", None)
               and not getattr(id_to_cust.get(c.ParentRef.value), "ParentRef", None)]
sub_sub     = [c for c in customers if getattr(c, "ParentRef", None)
               and getattr(id_to_cust.get(c.ParentRef.value, object()), "ParentRef", None)]

print(f"Top-level customers (clients):  {len(top_level)}")
print(f"Sub-customers (projects):       {len(sub_level)}")
print(f"Sub-sub-customers (phases?):    {len(sub_sub)}")

if sub_sub:
    print("\nSample sub-sub-customers:")
    for c in sub_sub[:10]:
        parent = id_to_cust.get(c.ParentRef.value)
        grandparent = id_to_cust.get(parent.ParentRef.value) if parent and getattr(parent, "ParentRef", None) else None
        gp_name = grandparent.DisplayName if grandparent else "?"
        p_name  = parent.DisplayName if parent else "?"
        print(f"  {gp_name}  →  {p_name}  →  {c.DisplayName}")

# ── 2. Invoice line items by project ─────────────────────────────────────────
print()
print("=" * 60)
print("2. INVOICE LINE ITEMS — service items used per project")
print("=" * 60)

invoices = query_all(Invoice, qb)
print(f"Total invoices: {len(invoices)}")

# Collect unique service items per project (sub-customer)
project_items: dict[str, set[str]] = defaultdict(set)
item_names: set[str] = set()

for inv in invoices:
    cust_ref = getattr(inv, "CustomerRef", None)
    if not cust_ref:
        continue
    cust = id_to_cust.get(cust_ref.value)
    if not cust or not getattr(cust, "ParentRef", None):
        continue  # top-level customer, not a project

    proj_name = cust.DisplayName or cust_ref.value
    for line in getattr(inv, "Line", []) or []:
        detail = getattr(line, "SalesItemLineDetail", None)
        if not detail:
            continue
        item_ref = getattr(detail, "ItemRef", None)
        if not item_ref:
            continue
        item_name = getattr(item_ref, "name", "") or item_ref.value
        project_items[proj_name].add(item_name)
        item_names.add(item_name)

print(f"\nUnique service items across all project invoices: {len(item_names)}")
print("\nAll unique items (potential phases):")
for name in sorted(item_names):
    print(f"  {name}")

print(f"\nProjects with invoiced items: {len(project_items)}")
print("\nSample projects and their items:")
for proj, items in list(project_items.items())[:10]:
    print(f"  {proj}:")
    for item in sorted(items):
        print(f"    - {item}")
