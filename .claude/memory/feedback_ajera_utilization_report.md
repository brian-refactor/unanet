---
name: Ajera Utilization Report — Key Technical Facts
description: Non-obvious facts about the Ajera Employee Utilization XLS report structure and data interpretation
type: feedback
originSessionId: 8d6e06b6-d740-4528-bb1b-3a11637df7d6
---
Key facts about the Ajera Employee Utilization XLS export used in `etl/ajera_utilization.py`:

1. **YTD amounts are at COST, not billing rates.** The "YTD Amounts" row in the report shows dollars at cost rates (what the firm pays). To get billing revenue, multiply YTD billable hours by the employee's billing rate from the separate billing rate export.

2. **Report structure is 6-row blocks per employee.** Each employee occupies exactly 6 rows: name/type header row, then Current Hours, Current Amounts, YTD Hours, YTD Amounts, Percent Targeted.

3. **Employees with no target set get 85% default.** `DEFAULT_TARGET = 85.0` is applied when `target_pct == 0`. This is the firm standard.

4. **YTD fraction uses the period end date, NOT the YTD range end date.** The YTD range always ends Dec 31 (100% = wrong). Use the period's end date (e.g., Apr 30 = ~33% of year) from `meta['period']`.

5. **Indirect category dollar amounts are estimated.** Ajera's amounts rows only total billable/indirect/total — no per-category dollars. The script estimates category dollars as `hours × hourly_rate` for Admin, Marketing, Vacation, etc.

6. **Billing rates live in a separate CSV.** `output/cincinnati/cincinnati_billing_rates.csv` has employee-specific overrides plus position-type fallbacks. Lookup: employee name first, then `EmployeeType` fallback.

7. **Name matching requires normalization.** Some Ajera names have `*` prefix (asterisk = inactive/flagged), Unicode accents (e.g., Bazán), and periods in middle initials. `_norm_name()` handles all of these; `_first_last()` provides a first+last fallback for middle-name mismatches.

**Why:** These facts were discovered through trial and error and fix YTD fraction calculation (was 1.0/100%), name matching failures (137/137 now match), and cost vs billing confusion.
**How to apply:** When updating the utilization report or adding new metrics, assume amounts = cost and billing revenue must be computed separately from the billing rates CSV.
