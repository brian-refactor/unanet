---
name: Cincinnati Billing Rates File
description: Where billing rates live and how the lookup works for profitability calculations
type: reference
originSessionId: 8d6e06b6-d740-4528-bb1b-3a11637df7d6
---
Billing rates for Cincinnati / Reztark Design Studio are stored in:

`output/cincinnati/cincinnati_billing_rates.csv`

Columns: `Employee, EmployeeType, Activity, BillingRate, OverrideCostRate`

**Two-level structure:**
- Rows with `Employee` name = employee-specific overrides (29 employees as of Apr 2026)
- Rows with blank `Employee` + `EmployeeType` = position fallback rates (12 position types)

**Position fallback rates (as of Apr 2026):**
Senior Architect $160, Architect $120, Project Director $200, Junior Professional $100, Interior Designer $120, Graphic Designer $110, Senior Graphic Designer $140, Professional $120, Senior Professional $160, Senior Interior Designer $160, Junior Designer $100, Design Professional $110

**Lookup logic in `lookup_billing_rate()`:** employee name (normalized) → first+last fallback → position type.

**Source:** Exported from Ajera billing rate schedule. Update this file when rates change — it is NOT auto-extracted by `ajera_extract.py`.

**Why:** Billing rates are needed to compute gross margin (billing revenue − cost) per employee.
**How to apply:** When profitability numbers look wrong, check this file first. If a new employee is missing a rate, add a row with their name and rate.
