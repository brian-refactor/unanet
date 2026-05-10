---
name: Ajera API — Key Technical Facts
description: Discovered behaviors of the Ajera REST API that are non-obvious and must not be re-learned
type: feedback
originSessionId: 8d6e06b6-d740-4528-bb1b-3a11637df7d6
---
Key facts discovered through probing the Ajera API (Cincinnati / Reztark Design Studio):

1. **Activities = Expense Codes.** `ListActivities` returns what Unanet calls Expense Codes (Meals, Travel, Mileage, Hotel, etc.). There is no `ListExpenseCodes` method. Use `ListActivities` → map ActivityKey to ECCode as `CIN-{ActivityKey}`.

2. **MethodArguments is always required.** Every API call must include `'MethodArguments': args or {}` — even when there are no arguments. Omitting it causes RC=-100 / ErrorID=-200 on all List methods.

3. **Time entry data is not accessible via API.** All transactional methods (ListTimeEntries, GetTimeEntries, ListTimesheets, GetProjectActuals, etc.) return error -150 regardless of parameters, API version (v1 or v2), or user permissions. This is an API scope limitation — time data must come from a manual Ajera report export.

4. **RC=0 + ErrorID=-150 = method exists but failed.** RC=-90 = method does not exist. RC=200 = success.

5. **ICR credentials in auth response** (`icrClientId`, `icrApiKey`) are for Deltek's web time-entry UI, not a callable REST API.

**Why:** These were discovered through extensive probing — re-probing wastes API calls and session time.
**How to apply:** Never attempt time entry methods via the API. Always include MethodArguments. Use ListActivities for expense codes.
