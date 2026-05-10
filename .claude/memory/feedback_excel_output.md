---
name: Excel Report Best Practices (openpyxl)
description: Proven approach for building clean, chartable Excel reports with openpyxl in this project
type: feedback
originSessionId: 8d6e06b6-d740-4528-bb1b-3a11637df7d6
---
Rules for building Excel reports with openpyxl learned from building `ajera_utilization.py`:

1. **Store numbers as numbers, not strings.** Use `cell.number_format = '$#,##0'` or `'0.0"%"'` for display — never pre-format as `f"${val:,.0f}"`. Pre-formatted strings break sorting and charts.

2. **For "%" values stored as 57.0 (not 0.57), use `'0.0"%"'` format.** This appends a literal % sign without dividing by 100. `'0%'` would treat 57.0 as 5700%.

3. **Charts require raw numeric cells.** BarChart and PieChart data references must point to cells containing actual numbers, not formatted strings. Put chart source data below the visible table if needed.

4. **`ws.sheet_view.showGridLines = False`** on every sheet for clean appearance.

5. **Use a single `_xcell()` helper** that accepts fill, font, border, number_format, and alignment in one call — avoids repeating 5 property assignments per cell.

6. **Merged cells: only style the top-left cell.** After `ws.merge_cells(...)`, only `ws.cell(row, col)` at the top-left is writable — the others are `MergedCell` objects.

7. **Row heights must be set explicitly** (`ws.row_dimensions[ri].height = 16`) — openpyxl does not auto-fit.

8. **Horizontal bar charts:** `bar.type = 'bar'` (not 'col'). Categories go on the y-axis, values on x-axis.

**Why:** Learned through iteration building the 5-tab Cincinnati profitability report.
**How to apply:** When building any Excel output with openpyxl, follow these patterns from the start.
