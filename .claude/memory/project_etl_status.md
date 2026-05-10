---
name: ETL Extraction Status — All Offices
description: Current completion state of all four office data extractions into output CSVs
type: project
originSessionId: 8d6e06b6-d740-4528-bb1b-3a11637df7d6
---
All four offices are fully extracted into `output/<office>/` CSVs. All seven data types complete.

| File | MN | CIN | DAL | ORL |
|---|---|---|---|---|
| COA | 295 | 222 | 202 | 131 |
| Clients | 283 | 381 | 561 | 178 |
| ClientContacts | 117 | 196 | 141 | 0 |
| Vendors | 264 | 653 | 2758 | 441 |
| VendorContacts | 56 | 18 | 70 | 3 |
| Employees | 65 | 454* | 252 | 11 |
| ExpenseCodes | 104 | 60 | 114 | 8 |

*CIN Employees: 454 rows = pay rate history rows, ~137 unique employees.

**Next step:** Load CSVs into Unanet upload templates (Excel workbooks in `Documentation/`).

**Why:** All source system extractions are done. The remaining work is template loading and validation passes.
**How to apply:** When asked about status or next steps, extraction is complete — focus is on template loading.
