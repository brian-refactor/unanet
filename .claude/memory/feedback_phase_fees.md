---
name: feedback-phase-fees
description: Fee fields belong on L2 parent rows only — never replicated to L3 sub-phases
metadata: 
  node_type: memory
  type: feedback
  originSessionId: 85e53d6d-ff3b-41ca-ab20-c7fe8e7b6e3a
---

Fee fields (`fixed_fee`, `labor_contract_cap`, `odc_contract_cap`, `icc_fixed_fee`, `labor_budget`, `odc_budget`, `icc_budget`, `hours_budget`) go on the L2 phase row only.

**Why:** When the old CIN extractor spread L2 fee to every L3 child row, Unanet would sum them and show 4x the actual contract value. Confirmed by user: "L2 only."

**How to apply:** Any time phases have a parent (L2) + children (L3) structure, always null out fee fields on L3 rows. The L2 parent row carries the total; L3 rows are descriptive breakdown only. This pattern is implemented in both `extract_cin_phases.py` and `extract_dal_phases.py`.
