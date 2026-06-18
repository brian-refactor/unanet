"""
Extract PM and Principal-In-Charge assignments from Monograph and write to
projects.pm_emp_code and projects.pic_emp_code in Supabase.

Role matching rules (substring, case-insensitive):
  PM        — rolesSentence contains "Project Manager"
  Principal — rolesSentence contains "Principal" or "Studio Director"

When multiple people share a role on one project, the first match wins.

Uses saved cookies from etl/monograph_cookies.json (run extract_monograph_phases.py
first if they're missing). Falls back to fresh Playwright login if expired.

Usage:
    python etl/extract_monograph_roles.py
    python etl/extract_monograph_roles.py --dry-run
"""

import argparse
import json
import os
import time
from pathlib import Path

import requests
from dotenv import load_dotenv

HERE = Path(__file__).parent
load_dotenv(HERE / ".env")

OFFICE     = "minnesota"
RAW_CACHE  = HERE.parent / "output" / "monograph_phases_raw.json"
COOKIE_FILE = HERE / "monograph_cookies.json"
GQL_URL    = "https://app.monograph.com/graphql"

PM_KEYWORDS        = ["project manager"]
PRINCIPAL_KEYWORDS = ["principal", "studio director"]

PROJECT_QUERY = """
query($slug: String!) {
  project(slug: $slug) {
    number
    profiles { fname lname email }
    teamList  { name rolesSentence }
  }
}
"""


def _cookies_as_dict(path: Path) -> dict:
    with open(path) as f:
        return {c["name"]: c["value"] for c in json.load(f)}


def _make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "Content-Type": "application/json",
        "Referer": "https://app.monograph.com/projects",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        ),
    })
    s.cookies.update(_cookies_as_dict(COOKIE_FILE))
    return s


def fetch_project_roles(session: requests.Session, slug: str) -> dict:
    """Return {pm_email, pic_email} for a project slug. Values may be None."""
    resp = session.post(GQL_URL, json={
        "query": PROJECT_QUERY,
        "variables": {"slug": slug},
    }, timeout=15)
    resp.raise_for_status()
    proj = resp.json().get("data", {}).get("project") or {}

    profiles  = proj.get("profiles") or []
    team_list = proj.get("teamList")  or []

    # Build name → email from profiles
    name_to_email: dict[str, str] = {}
    for p in profiles:
        full = f"{p.get('fname', '')} {p.get('lname', '')}".strip()
        if full and p.get("email"):
            name_to_email[full] = p["email"].lower()

    pm_email  = None
    pic_email = None

    for member in team_list:
        role  = (member.get("rolesSentence") or "").lower()
        name  = (member.get("name") or "").strip()
        email = name_to_email.get(name)

        if pm_email is None and any(k in role for k in PM_KEYWORDS):
            pm_email = email

        if pic_email is None and any(k in role for k in PRINCIPAL_KEYWORDS):
            pic_email = email

    return {"pm_email": pm_email, "pic_email": pic_email}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true",
                        help="Print what would be updated without writing to DB")
    args = parser.parse_args()

    if not RAW_CACHE.exists():
        print(f"No raw cache at {RAW_CACHE}. Run extract_monograph_phases.py first.")
        return

    with open(RAW_CACHE) as f:
        raw = json.load(f)

    # Only real numbered projects
    projects = [(r["slug"], r["number"]) for r in raw if r.get("number") and r.get("slug")]
    print(f"Projects to process: {len(projects)}")

    # Build email → employee_code from Supabase (all offices)
    from supabase import create_client
    sb = create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_KEY"])

    emp_rows = sb.table("employees").select("employee_code,email").execute().data
    email_to_code: dict[str, str] = {
        e["email"].lower(): e["employee_code"]
        for e in emp_rows if e.get("email") and e.get("employee_code")
    }
    print(f"Employee email lookup: {len(email_to_code)} entries")

    session = _make_session()

    results   = []   # (project_code, pm_code, pic_code)
    no_pm     = []
    no_pic    = []
    no_match  = []   # email found but not in employees table

    for i, (slug, number) in enumerate(projects, 1):
        roles = fetch_project_roles(session, slug)
        pm_email  = roles["pm_email"]
        pic_email = roles["pic_email"]

        pm_code  = email_to_code.get(pm_email)  if pm_email  else None
        pic_code = email_to_code.get(pic_email) if pic_email else None

        if not pm_email:
            no_pm.append(number)
        elif not pm_code:
            no_match.append(f"{number}  PM={pm_email}")

        if not pic_email:
            no_pic.append(number)
        elif not pic_code:
            no_match.append(f"{number}  PIC={pic_email}")

        results.append((number, pm_code, pic_code))

        if i % 50 == 0:
            print(f"  {i}/{len(projects)} projects fetched...")

        time.sleep(0.07)

    # Stats
    has_pm  = sum(1 for _, pm, _ in results if pm)
    has_pic = sum(1 for _, _, pic in results if pic)
    print(f"\nResults:")
    print(f"  pm_emp_code  populated: {has_pm}/{len(results)}")
    print(f"  pic_emp_code populated: {has_pic}/{len(results)}")
    print(f"  No PM role found:       {len(no_pm)}")
    print(f"  No Principal role found:{len(no_pic)}")
    print(f"  Email not in employees: {len(no_match)}")

    if no_match:
        print("\nEmails not matched to employee_code (need to load more employees):")
        for m in no_match[:20]:
            print(f"  {m}")

    if args.dry_run:
        print("\n[DRY RUN] — no DB writes.")
        return

    # Write to Supabase
    print(f"\nUpdating {len(results)} MN projects in Supabase...")
    updated = 0
    for project_code, pm_code, pic_code in results:
        data = {}
        if pm_code:
            data["pm_emp_code"] = pm_code
        if pic_code:
            data["pic_emp_code"] = pic_code
        if data:
            sb.table("projects").update(data) \
              .eq("office", OFFICE) \
              .eq("project_code", project_code) \
              .execute()
            updated += 1

    print(f"Done. {updated} projects updated.")


if __name__ == "__main__":
    main()
