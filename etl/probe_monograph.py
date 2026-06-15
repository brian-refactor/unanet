"""
Probe Monograph GraphQL API.

Step 1: Log in via Playwright, capture session cookies.
Step 2: Intercept all GraphQL responses on a project page to identify
        which operations return phase data.
Step 3: Print operation names + response shapes so we know what to query.

Credentials: MONOGRAPH_EMAIL and MONOGRAPH_PASSWORD in etl/.env

Usage:
    python etl/probe_monograph.py
    python etl/probe_monograph.py --url https://app.monograph.com/projects/3570-lexington
"""
import argparse
import json
import os
import time
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

EMAIL    = os.environ.get("MONOGRAPH_EMAIL", "")
PASSWORD = os.environ.get("MONOGRAPH_PASSWORD", "")

GQL_URL = "https://app.monograph.com/graphql"


def run_probe(project_url: str):
    from playwright.sync_api import sync_playwright

    captured: list[dict] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)   # headless=False so you can watch
        ctx = browser.new_context()
        page = ctx.new_page()

        # ── Intercept GraphQL responses ───────────────────────────────────────
        def handle_response(response):
            if GQL_URL in response.url and response.status == 200:
                try:
                    body = response.json()
                    # Pull operation name from the URL query string
                    op = response.url.split("op=")[-1] if "op=" in response.url else "unknown"
                    captured.append({"op": op, "body": body})
                except Exception:
                    pass

        page.on("response", handle_response)

        # ── Log in ────────────────────────────────────────────────────────────
        print("Navigating to login page...")
        page.goto("https://app.monograph.com/users/sign_in", wait_until="networkidle")

        print("Filling credentials...")
        page.fill('input[name="user[email]"]',    EMAIL)
        page.fill('input[name="user[password]"]', PASSWORD)
        page.click('input[type="submit"], button[type="submit"]')
        page.wait_for_url("**/app.monograph.com/**", timeout=15000)
        print(f"Logged in. Current URL: {page.url}")

        # ── Navigate to project page ──────────────────────────────────────────
        print(f"\nNavigating to project: {project_url}")
        page.goto(project_url, wait_until="networkidle")
        time.sleep(3)   # let lazy-loaded tabs settle

        # Click Phases tab if visible
        for selector in ['[data-tab="phases"]', 'text=Phases', 'a:text("Phases")', 'button:text("Phases")']:
            try:
                el = page.query_selector(selector)
                if el:
                    el.click()
                    time.sleep(2)
                    print(f"Clicked phases tab via: {selector}")
                    break
            except Exception:
                pass

        # ── Save cookies ──────────────────────────────────────────────────────
        cookies = ctx.cookies()
        cookie_file = Path(__file__).parent / "monograph_cookies.json"
        with open(cookie_file, "w") as f:
            json.dump(cookies, f, indent=2)
        print(f"\nCookies saved to {cookie_file}")

        browser.close()

    # ── Analyse captured operations ───────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"Captured {len(captured)} GraphQL operations")
    print(f"{'='*60}\n")

    phase_ops = []
    for item in captured:
        op   = item["op"]
        body = item["body"]
        data = body.get("data") or {}

        # Summarise top-level keys in the response
        keys = list(data.keys()) if isinstance(data, dict) else []
        print(f"  op={op:40s}  keys={keys}")

        # Flag anything that looks phase-related
        text = json.dumps(data).lower()
        if any(w in text for w in ["phase", "Phase", "wbs", "milestone"]):
            phase_ops.append(item)

    if phase_ops:
        print(f"\n{'='*60}")
        print("PHASE-RELATED OPERATIONS:")
        print(f"{'='*60}")
        for item in phase_ops:
            print(f"\n--- op={item['op']} ---")
            print(json.dumps(item["body"], indent=2)[:3000])
    else:
        print("\nNo phase-related operations detected.")
        print("Try navigating to a project's Phases tab manually and re-running.")

    # Save full capture for manual inspection
    out = Path(__file__).parent.parent / "output" / "monograph_probe.json"
    with open(out, "w") as f:
        json.dump(captured, f, indent=2)
    print(f"\nFull capture saved to {out}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default="https://app.monograph.com/projects",
                        help="Monograph project URL to probe")
    args = parser.parse_args()
    run_probe(args.url)


if __name__ == "__main__":
    main()
