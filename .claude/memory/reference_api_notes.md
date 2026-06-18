---
name: reference-api-notes
description: "API access notes for Monograph, Ajera v1/v2, and key limitations discovered"
metadata: 
  node_type: memory
  type: reference
  originSessionId: 85e53d6d-ff3b-41ca-ab20-c7fe8e7b6e3a
---

## Monograph (MN)
- GraphQL at `https://app.monograph.com/graphql?op=ganttChart`
- Auth: cookie-based via Playwright; cookies saved to `etl/monograph_cookies.json`
- `ganttChart` query returns project-level `startDate`/`endDate` AND phase-level dates — both available
- `teamList` on `Project` type has `rolesSentence` (e.g., "Project Manager/Director", "Studio Director") and `name`
- `profiles` on `Project` type has `email` — cross-reference with `teamList` by name to get PM email
- Full schema introspection disabled; `__type` introspection works
- PM role match: "Project Manager" in `rolesSentence`; Principal match: "Principal" or "Studio Director"
- New script: `etl/extract_monograph_roles.py` — fetches PM/PIC per project and writes to Supabase

## Ajera v1 (CIN phases)
- URL in `etl/ajera.env` as `AJERA_API_URL`
- Credentials: `AJERA_USERNAME=migration`, `AJERA_PASSWORD=64Lbdas8$123`
- Available methods: `CreateAPISession`, `EndAPISession`, `GetProjects`, `ListProjects`
- Phase status values: `Active`, `Closed`, `Marketing`, `Hold`, `WorkHold`, `BillingHold`
- Dates on projects/phases are sparse — Ajera CIN didn't track them consistently

## Ajera v2 (CIN)
- Same URL, use `"APIVersion": 2` in `CreateAPISession`
- `ListTimesheets` works but date filters are broken — always returns same 500 recent records
- `GetTimesheets` detail has `Project.Detail[].Project Key` and `TimesheetDate` — useful if date filter worked
- Invoice methods (`GetClientInvoices`, `GetInvoices`) return error -150 — not accessible

## Employee code formats
- MN: `UOC000XXX` (Unanet HR file)
- CIN: `DXE000XXX` (Unanet HR file); old Ajera-derived placeholders were `CIN-EMP-XXXXX` — fully replaced
- DAL: `UOC000XXX` format expected (not yet populated on projects)
- All `@fusion-ae.com` email domain across all offices
