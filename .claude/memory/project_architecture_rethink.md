---
name: Architecture Rethink — ETL Pipeline & Railway App
description: User is considering moving ETL pipeline to Railway/Celery; pending meeting transcript to define scope
type: project
originSessionId: f2ad7fc3-f672-409e-9311-6f18f253eb3e
---
The local ETL scripts are getting complex as the data scope grows (more offices, projects, employees, spreadsheets). Brian has a separate production app on Railway at https://web-production-eeb6.up.railway.app/dashboard/ with a Celery worker (Python task queue — backend is Python, not Node despite initial mention).

**Options discussed:**
- A: Keep scripts local + add Railway Celery scheduling (low migration cost)
- B: Centralize ETL inside the existing Railway app (single dashboard, triggers, results)
- C: Deploy a separate Railway worker service just for ETL (clean separation)

**Status:** Deferred — Brian has a meeting in progress and will share the transcript afterward to review the full complexity of projects, employees, and spreadsheets before deciding on architecture.

**Why:** The data is growing beyond what monthly manual script runs handle cleanly.
**How to apply:** When Brian shares the meeting transcript, review it to understand the full data scope before recommending an architecture. Key question: does the Railway app's existing domain overlap with Unanet ETL work, or is it separate?
