---
name: Cincinnati Expense Codes Export Pending
description: Manual Ajera export needed for Cincinnati expense codes before Unanet load
type: project
originSessionId: 8d6e06b6-d740-4528-bb1b-3a11637df7d6
---
Expense Codes are not available via the Ajera REST API. Someone at Cincinnati (Reztark Design Studio) needs to export them manually.

**Steps:** In Ajera → Setup → Expense Codes → right-click grid → Grid Options → Export to Excel → send file to Brian.

**How to apply:** When the file arrives, build a parser to map it to `06-ExpenseCodes_Fusion.xlsx` format and write to `output/cincinnati/cincinnati_ExpenseCodes.csv`.
