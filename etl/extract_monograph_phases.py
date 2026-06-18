"""
Extract MN project phases from Monograph via Playwright + GraphQL interception.

Strategy:
1. Log into Monograph with credentials from etl/.env
2. Navigate to the gantt chart view (/projects) — Monograph fires a `ganttChart`
   GraphQL operation that returns ALL projects with their phases.
3. Intercept every `ganttChart` response; if pagination (hasNextPage), scroll down
   to trigger the next page load.
4. Aggregate all rows and upsert into Supabase `project_phases`.
5. Also updates `projects.start_date` / `projects.end_date` from the project-level
   dates Monograph returns (these are more reliable than QBO-derived dates).

The `ganttChart` response shape (per project):
    {
      "id": "244310",
      "number": "23-053",
      "name": "3570 Lexington",
      "startDate": "2023-01-15",
      "endDate": "2024-06-30",
      "phases": [
        { "type": { "name": "Pre-Design", "feeType": "FIXED", "budget": 4500.0 },
          "startDate": "2023-06-28", "endDate": "2023-12-12",
          "hoursPlanned": 60.0 }
        ...
      ]
    }

FeeType mapping:
    FIXED   -> contract_type = "Fixed Fee"
    HOURLY  -> contract_type = "Hourly"
    NTE     -> contract_type = "Not-to-Exceed"
    RETAINER-> contract_type = "Retainer"

Usage:
    python etl/extract_monograph_phases.py
    python etl/extract_monograph_phases.py --dry-run
    python etl/extract_monograph_phases.py --capture-only   # save raw JSON, no DB
"""

import argparse
import json
import os
import sys
import time
from collections import defaultdict
from pathlib import Path

from dotenv import load_dotenv

HERE = Path(__file__).parent
load_dotenv(HERE / ".env")

EMAIL    = os.environ.get("MONOGRAPH_EMAIL", "")
PASSWORD = os.environ.get("MONOGRAPH_PASSWORD", "")

OFFICE = "minnesota"
OUTPUT_DIR = HERE.parent / "output"
RAW_CACHE  = OUTPUT_DIR / "monograph_phases_raw.json"

GQL_URL     = "https://app.monograph.com/graphql"
GANTT_URL   = "https://app.monograph.com/projects"
LOGIN_URL   = "https://app.monograph.com/login"
COOKIE_FILE = HERE / "monograph_cookies.json"

# Monograph feeType → Unanet contract_type
FEE_TYPE_MAP = {
    "FIXED":    "Fixed Fee",
    "HOURLY":   "Hourly",
    "NTE":      "Not-to-Exceed",
    "RETAINER": "Retainer",
    "PERCENT":  "Percent of Construction",
}


GANTT_QUERY = """
query ganttChart($isTemplate: Boolean, $filters: GanttChartFiltersInput, $first: Int, $offset: Int) {
  ganttChart(filters: $filters isTemplate: $isTemplate first: $first offset: $offset) {
    totalProjectsCount
    totalFilteredCount
    hasNextPage
    rows {
      id number name slug status startDate endDate clientName
      phases {
        id
        type { id name abbr feeType status budget __typename }
        startDate endDate hoursPlanned hoursConsumed planned
        __typename
      }
      __typename
    }
    __typename
  }
}
"""


def _cookies_as_dict(cookie_list: list) -> dict:
    """Convert Playwright cookie list to requests-compatible dict."""
    return {c["name"]: c["value"] for c in cookie_list}


def _refresh_session() -> dict:
    """Login via Playwright and return fresh cookies as a dict."""
    from playwright.sync_api import sync_playwright

    cookies = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()
        print("Logging in to Monograph...")
        page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=20000)
        page.wait_for_selector('input[name="user[email]"]', timeout=20000)
        page.fill('input[name="user[email]"]', EMAIL)
        page.fill('input[name="user[password]"]', PASSWORD)
        page.click('input[type="submit"], button[type="submit"]')
        page.wait_for_url("**/app.monograph.com/**", timeout=30000)
        time.sleep(2)
        raw = ctx.cookies()
        with open(COOKIE_FILE, "w") as f:
            json.dump(raw, f, indent=2)
        cookies = _cookies_as_dict(raw)
        browser.close()
    print("Login successful, cookies saved.")
    return cookies


def collect_rows_via_api() -> list[dict]:
    """
    Directly paginate the Monograph GraphQL API using saved session cookies.
    Falls back to Playwright login if cookies are expired.
    """
    import requests as req

    # Load saved cookies
    if COOKIE_FILE.exists():
        with open(COOKIE_FILE) as f:
            raw = json.load(f)
        cookies = _cookies_as_dict(raw)
    else:
        cookies = _refresh_session()

    headers = {
        "Content-Type": "application/json",
        "Accept": "*/*",
        "Referer": "https://app.monograph.com/projects/overview",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        ),
    }

    PAGE_SIZE = 50
    offset = 0
    all_rows: dict[str, dict] = {}
    total = None

    session = req.Session()
    session.headers.update(headers)
    session.cookies.update(cookies)

    while True:
        payload = {
            "operationName": "ganttChart",
            "variables": {
                "filters": {"sortBy": "alphabetical"},
                "first": PAGE_SIZE,
                "offset": offset,
            },
            "query": GANTT_QUERY,
        }
        resp = session.post(
            f"{GQL_URL}?op=ganttChart",
            json=payload,
            timeout=30,
        )

        if resp.status_code == 401 or (resp.status_code == 200 and "errors" in resp.json()):
            print("Session expired — refreshing login...")
            cookies = _refresh_session()
            session.cookies.update(cookies)
            continue

        resp.raise_for_status()
        data = resp.json().get("data", {}).get("ganttChart", {})
        rows = data.get("rows") or []
        has_next = data.get("hasNextPage", False)
        if total is None:
            total = data.get("totalFilteredCount", "?")

        for row in rows:
            all_rows[row["id"]] = row

        print(f"  offset={offset:4d}  +{len(rows)} rows  "
              f"(cumulative {len(all_rows)}/{total}, hasNext={has_next})")

        if not has_next or not rows:
            break
        offset += PAGE_SIZE

    return list(all_rows.values())


def build_phase_rows(project_rows: list[dict], proj_df=None) -> list[dict]:
    """
    Convert Monograph ganttChart rows into project_phases DB rows.

    If proj_df is provided (from minnesota_Projects.csv), we cross-reference
    the Monograph project number (e.g. '23-053') to our project_code (e.g. 'MN-23053').
    If no match found, we skip the project and warn.
    """
    # Build number -> project_code lookup from CSV if available
    num_to_code: dict[str, str] = {}
    if proj_df is not None:
        import pandas as pd
        for _, r in proj_df.iterrows():
            src_id = str(r.get("_source_id", "")).strip()
            pcode  = str(r.get("project_code", "")).strip()
            name   = str(r.get("project_name", "")).strip()
            if pcode:
                # Also index by project name fragment for fuzzy matching later
                num_to_code[src_id] = pcode

    rows = []
    skipped = []
    for proj in project_rows:
        mon_number = proj.get("number", "").strip()   # e.g. "23-053"
        mon_name   = proj.get("name", "").strip()
        mon_id     = proj.get("id", "")
        phases     = proj.get("phases") or []

        if not phases:
            continue

        # Monograph project number matches MN project_code in QBO CSV directly
        # e.g. Monograph "23-053" -> project_code "23-053"
        if mon_number:
            project_code = mon_number
        else:
            skipped.append(f"  no number: {mon_name} (id={mon_id})")
            continue

        for seq, phase in enumerate(phases, start=1):
            ph_type   = phase.get("type") or {}
            ph_name   = ph_type.get("name", f"Phase {seq}")
            fee_type  = ph_type.get("feeType", "FIXED")
            budget    = ph_type.get("budget")
            status    = ph_type.get("status")   # COMPLETED, ACTIVE, etc.
            start     = phase.get("startDate")
            end       = phase.get("endDate")
            hours     = phase.get("hoursPlanned")

            contract_type = FEE_TYPE_MAP.get(fee_type, "Fixed Fee")

            rows.append({
                "office":             OFFICE,
                "project_code":       project_code,
                "contract_type":      contract_type,
                "level2_name":        ph_name,
                "level2_code":        f"{seq:03d}",
                "level3_name":        None,
                "level3_code":        None,
                "start_date":         start,
                "end_date":           end,
                "phase_status":       status,
                "org_path":           None,
                "fixed_fee":          float(budget) if budget is not None else None,
                "labor_contract_cap": None,
                "odc_contract_cap":   None,
                "occ_contract_cap":   None,
                "icc_fixed_fee":      None,
                "labor_budget":       None,
                "odc_budget":         None,
                "occ_budget":         None,
                "icc_budget":         None,
                "hours_budget":       float(hours) if hours is not None else None,
            })

    if skipped:
        print(f"\nSkipped {len(skipped)} projects (no Monograph number):")
        for s in skipped:
            print(s)

    return rows


def build_project_date_updates(project_rows: list[dict]) -> list[dict]:
    """
    Extract project-level start/end dates from Monograph ganttChart rows.
    Returns one dict per project that has a valid number and at least one date.
    """
    updates = []
    for proj in project_rows:
        mon_number = proj.get("number", "").strip()
        if not mon_number:
            continue
        start = proj.get("startDate")
        end   = proj.get("endDate")
        if not start and not end:
            continue
        updates.append({
            "project_code": mon_number,
            "start_date":   start,
            "end_date":     end,
        })
    return updates


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run",      action="store_true",
                        help="Print stats but don't write to DB")
    parser.add_argument("--capture-only", action="store_true",
                        help="Collect from Monograph and save raw JSON, no DB write")
    parser.add_argument("--from-cache",   action="store_true",
                        help=f"Skip browser; load from {RAW_CACHE}")
    args = parser.parse_args()

    # ── Collect raw project rows ──────────────────────────────────────────────
    if args.from_cache:
        if not RAW_CACHE.exists():
            print(f"No cache file at {RAW_CACHE}. Run without --from-cache first.")
            sys.exit(1)
        with open(RAW_CACHE) as f:
            project_rows = json.load(f)
        print(f"Loaded {len(project_rows)} projects from cache.")
    else:
        project_rows = collect_rows_via_api()
        print(f"\nTotal projects collected: {len(project_rows)}")
        # Save raw for inspection / reuse
        with open(RAW_CACHE, "w") as f:
            json.dump(project_rows, f, indent=2)
        print(f"Raw data saved to {RAW_CACHE}")

    if args.capture_only:
        print("--capture-only: done.")
        return

    # ── Load MN projects CSV for cross-reference (optional) ──────────────────
    proj_csv = HERE.parent / "output" / "minnesota" / "minnesota_Projects.csv"
    proj_df  = None
    if proj_csv.exists():
        import pandas as pd
        proj_df = pd.read_csv(proj_csv)
        print(f"MN projects CSV loaded: {len(proj_df)} rows")

    # ── Build phase rows ──────────────────────────────────────────────────────
    rows = build_phase_rows(project_rows, proj_df)
    print(f"\nPhase rows built: {len(rows)}")

    if args.dry_run:
        # Summary by phase name
        from collections import Counter
        phase_counts = Counter(r["level2_name"] for r in rows)
        print("\nPhase distribution:")
        for name, count in phase_counts.most_common(30):
            print(f"  {count:4d}  {name}")
        projects_with_phases = len({r["project_code"] for r in rows})
        print(f"\nProjects with phases: {projects_with_phases}")

        date_updates = build_project_date_updates(project_rows)
        has_start = sum(1 for d in date_updates if d["start_date"])
        has_end   = sum(1 for d in date_updates if d["end_date"])
        print(f"\nProject date updates available: {len(date_updates)} projects")
        print(f"  with start_date: {has_start}")
        print(f"  with end_date:   {has_end}")
        print("[DRY RUN] — no DB writes.")
        return

    # ── Load to Supabase ──────────────────────────────────────────────────────
    from supabase import create_client
    sb = create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_KEY"])

    print(f"\nClearing existing Monograph MN phase rows (source=monograph)...")
    # We ONLY delete rows that came from Monograph. The QBO-derived rows (from
    # extract_mn_phases.py) live alongside. We tag them by checking contract_type
    # presence — actually just delete all MN phases and re-insert both.
    # (The two sources shouldn't be used simultaneously; pick one.)
    sb.table("project_phases").delete().eq("office", OFFICE).execute()
    print("Cleared.")

    print(f"Inserting {len(rows)} rows in batches of 500...")
    for i in range(0, len(rows), 500):
        batch = rows[i:i + 500]
        sb.table("project_phases").insert(batch).execute()
        print(f"  {min(i + 500, len(rows))} / {len(rows)}")

    final = sb.table("project_phases") \
              .select("id", count="exact") \
              .eq("office", OFFICE).execute()
    print(f"\nDone. {final.count} MN phase rows in Supabase.")

    # ── Update project-level start/end dates from Monograph ───────────────────
    date_updates = build_project_date_updates(project_rows)
    print(f"\nUpdating start/end dates for up to {len(date_updates)} MN projects...")
    updated = 0
    for item in date_updates:
        data = {}
        if item["start_date"]:
            data["start_date"] = item["start_date"]
        if item["end_date"]:
            data["end_date"] = item["end_date"]
        if data:
            sb.table("projects").update(data) \
              .eq("office", OFFICE) \
              .eq("project_code", item["project_code"]) \
              .execute()
            updated += 1
    print(f"Done. {updated} projects updated with Monograph dates.")


if __name__ == "__main__":
    main()
