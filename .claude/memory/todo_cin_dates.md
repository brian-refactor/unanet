---
name: todo-cin-dates
description: "Pending task to get CIN project start/end dates from Reztark's Ajera report"
metadata: 
  node_type: memory
  type: project
  originSessionId: 85e53d6d-ff3b-41ca-ab20-c7fe8e7b6e3a
---

95% of CIN projects are missing start/end dates. Unanet requires dates for the 07a load.

**Why:** Ajera v1 and v2 APIs don't expose invoice or timesheet history in a queryable way. `ListTimesheets` date filters are broken (returns same 500 recent records regardless of filter). `GetInvoices` returns error -150.

**How to apply:** Raise with Andrew at the 2026-06-24/25 check-in. Ask Reztark to run a project summary report from Ajera UI that includes first/last billed date per project, OR export a time report with project code + date columns. Once received, derive min/max dates per project and update Supabase.

Fallback already identified: use `LastModifiedDate` from `output/cincinnati/_phases_raw.json` as a rough end_date proxy for all 1,224 projects.
