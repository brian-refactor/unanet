---
name: QB Desktop Export Parsing — Key Behaviors
description: Non-obvious facts about parsing QB Desktop memorized report exports for Dallas and Orlando
type: feedback
originSessionId: 8d6e06b6-d740-4528-bb1b-3a11637df7d6
---
Key facts about QB Desktop CSV/XLSX exports and how they parse:

1. **Orlando exports as CSV, Dallas exports as XLSX.** The parser (`etl/qbd_parse.py`) handles both. Do not assume format by office.

2. **Orlando FirmCodes use leading project numbers.** Customer names like "0907 JHA" → FirmCode `ORL-0907`. Decimal sub-codes like "2020.02 Suntrust" → `ORL-2020-02`. Entries without numbers get sequential codes starting after the highest found. FirmName strips the number prefix (just "JHA", "Suntrust Plaza - Rollins").

3. **Dallas FirmCodes use slugified names.** `DAL-1040_Colusa_LLC`. No number extraction.

4. **Dallas customers have sub-jobs.** QB Desktop uses `Parent:Job` colon notation (e.g., "1040 Colusa LLC:DDN16001 - ..."). These are projects, not clients — filter them out. 8,068 raw rows → 561 top-level clients after filtering.

5. **Active Status column name varies.** Dallas/Orlando use "Active Status" with values "Active"/"Not-active". The `active_str()` helper handles all variants.

6. **Dallas Vendor file has duplicate Balance columns.** Three balance columns present — use `Balance Total` (last one).

7. **All exports have a leading empty/None column** that must be stripped before parsing.

**Why:** These differences caused multiple parsing failures when first running against Dallas files.
**How to apply:** When adding a new office, check the raw file structure before assuming it matches existing offices.
