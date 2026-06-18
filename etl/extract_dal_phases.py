"""
Extract Dallas project phases from QB Desktop estimates export.

Source file: input/DAL_Estimates.xlsx  (copy from Downloads before running)
  OR pass a path directly: python etl/extract_dal_phases.py --file /path/to/estimates.xlsx

Structure per project in the Excel export:
  Client row       — col B: client name
  Project row      — col C: "CODE - Project Name"
  Estimate header  — col F="Estimate", col G=date, col H=num, col I=None, col Q=total
  Line items       — col F="Estimate", col G=date, col H=num, col I=memo, col Q=neg amount
  Total row        — col C: "Total CODE - ..."

Unanet mapping:
  L2 phase  → each estimate  (fixed_fee = estimate total, start/end = estimate dates)
  L3 phase  → each line item (fixed_fee = None, name = Memo)

Project dates updated in Supabase:
  start_date = MIN(estimate Date) across all estimates for that project
  end_date   = MAX(Due Date or Date) across all estimates

Only projects already in the Supabase `projects` table for office=dallas are processed.

Usage:
    python etl/extract_dal_phases.py
    python etl/extract_dal_phases.py --dry-run
    python etl/extract_dal_phases.py --file /Users/you/Downloads/estimates.xlsx
"""

import argparse
import os
import re
from pathlib import Path
from collections import defaultdict

import openpyxl
from dotenv import load_dotenv

HERE = Path(__file__).parent
load_dotenv(HERE / ".env")

OFFICE      = "dallas"
DEFAULT_FILE = HERE.parent / "input" / "DAL_Estimates.xlsx"
CODE_RE      = re.compile(r'^([A-Z]{2,6}\d{4,6})', re.IGNORECASE)

# Estimate Active marker in QB Desktop export
ACTIVE_MARK = "Ö"


def parse_estimates(path: Path) -> dict:
    """
    Parse the QB Desktop estimates Excel export.

    Returns dict: project_code -> {
        name, client,
        estimates: [{num, date, due_date, total, active, lines: [{memo, amount}]}]
    }
    """
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb["Sheet1"]

    projects = {}
    current_client  = None
    current_project = None

    for row in ws.iter_rows(values_only=True):
        if not any(v is not None for v in row):
            continue

        col_b = row[1]
        col_c = row[2]
        col_f = row[5]   # Type
        col_g = row[6]   # Date
        col_h = row[7]   # Num (estimate number)
        col_i = row[8]   # Memo
        col_k = row[10]  # Due Date
        col_o = row[14]  # Estimate Active
        col_q = row[16]  # Amount

        # Client row
        if col_b and col_f is None and col_c is None:
            current_client = str(col_b).strip()
            continue

        # Project row — col C has "CODE - Name", no Total prefix
        if col_c and col_f is None:
            raw = str(col_c).strip()
            if raw.startswith("Total "):
                continue
            m = CODE_RE.match(raw)
            if not m:
                continue
            code = m.group(1).upper()
            # Name is everything after "CODE - "
            name_part = raw[len(code):].strip().lstrip("-").strip()
            current_project = code
            if code not in projects:
                projects[code] = {
                    "name":      name_part,
                    "client":    current_client,
                    "estimates": [],
                }
            continue

        # Estimate rows
        if col_f == "Estimate" and current_project:
            active = (str(col_o).strip() == ACTIVE_MARK) if col_o else False
            if col_i is None:
                # Estimate header row
                projects[current_project]["estimates"].append({
                    "num":      str(col_h).strip() if col_h else "",
                    "date":     col_g,
                    "due_date": col_k,
                    "total":    float(col_q) if col_q else 0.0,
                    "active":   active,
                    "lines":    [],
                })
            else:
                # Line item row
                if projects[current_project]["estimates"]:
                    memo   = str(col_i).strip()
                    amount = abs(float(col_q)) if col_q else None
                    # Skip Sales Tax and zero-amount junk lines
                    if memo.lower() == "sales tax" or memo == "":
                        continue
                    projects[current_project]["estimates"][-1]["lines"].append({
                        "memo":   memo,
                        "amount": amount,
                    })

    return projects


def build_phase_rows(code: str, data: dict) -> list[dict]:
    """Build flat project_phases rows for one DAL project."""
    rows = []
    l2_counter = 0

    for est in data["estimates"]:
        l2_counter += 1
        l2_code = f"{l2_counter:03d}"

        # Derive dates
        start = est["date"].date().isoformat()   if est.get("date")     else None
        end   = (est["due_date"] or est["date"])
        end   = end.date().isoformat()            if end                 else None

        phase_status = "ACTIVE" if est["active"] else "CLOSED"

        l2_row = {
            "office":             OFFICE,
            "project_code":       code,
            "contract_type":      "Fixed Fee",
            "level2_name":        f"Estimate #{est['num']}" if est["num"] else f"Estimate {l2_counter}",
            "level2_code":        l2_code,
            "level3_name":        None,
            "level3_code":        None,
            "start_date":         start,
            "end_date":           end,
            "phase_status":       phase_status,
            "org_path":           None,
            "fixed_fee":          est["total"] if est["total"] else None,
            "labor_contract_cap": None,
            "odc_contract_cap":   None,
            "occ_contract_cap":   None,
            "icc_fixed_fee":      None,
            "labor_budget":       None,
            "odc_budget":         None,
            "occ_budget":         None,
            "icc_budget":         None,
            "hours_budget":       None,
        }

        lines = est.get("lines", [])
        if not lines:
            rows.append(l2_row)
            continue

        # L2 parent row (with fee)
        rows.append(l2_row)

        # L3 child rows (no fee — L2 carries the total)
        for l3_idx, line in enumerate(lines, start=1):
            rows.append({
                **l2_row,
                "level3_name":        line["memo"],
                "level3_code":        f"{l3_idx:04d}",
                "fixed_fee":          None,
                "start_date":         None,
                "end_date":           None,
            })

    return rows


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true",
                        help="Parse and print stats without writing to DB")
    parser.add_argument("--file", type=Path, default=DEFAULT_FILE,
                        help=f"Path to estimates Excel file (default: {DEFAULT_FILE})")
    args = parser.parse_args()

    if not args.file.exists():
        print(f"Estimates file not found: {args.file}")
        print("Copy it to input/DAL_Estimates.xlsx or pass --file /path/to/file")
        return

    print(f"Parsing {args.file}...")
    all_projects = parse_estimates(args.file)
    print(f"  Projects in file:     {len(all_projects)}")

    # Load DAL project codes from Supabase
    from supabase import create_client
    sb = create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_KEY"])

    dal_rows, offset = [], 0
    while True:
        batch = sb.table("projects").select("project_code") \
            .eq("office", OFFICE).range(offset, offset + 999).execute().data
        dal_rows.extend(batch)
        if len(batch) < 1000:
            break
        offset += 1000
    dal_codes = {r["project_code"] for r in dal_rows}
    print(f"  DAL projects in Supabase: {len(dal_codes)}")

    matched = {k: v for k, v in all_projects.items() if k in dal_codes}
    print(f"  Matched (estimates ∩ Supabase): {len(matched)}")

    # Build phase rows
    all_phase_rows = []
    project_dates  = {}   # code -> (min_start, max_end)

    for code, data in matched.items():
        rows = build_phase_rows(code, data)
        all_phase_rows.extend(rows)

        # Collect project-level dates
        starts = [e["date"] for e in data["estimates"] if e.get("date")]
        ends   = [(e.get("due_date") or e.get("date")) for e in data["estimates"] if e.get("due_date") or e.get("date")]
        if starts:
            project_dates[code] = (
                min(starts).date().isoformat(),
                max(ends).date().isoformat() if ends else min(starts).date().isoformat(),
            )

    l2_rows = [r for r in all_phase_rows if not r["level3_code"]]
    l3_rows = [r for r in all_phase_rows if r["level3_code"]]
    print(f"\nPhase rows built:   {len(all_phase_rows)}")
    print(f"  L2 parent rows:   {len(l2_rows)}")
    print(f"  L3 child rows:    {len(l3_rows)}")
    print(f"Project dates available: {len(project_dates)}")

    if args.dry_run:
        # Show sample
        sample_code = next(iter(matched))
        sample_rows = [r for r in all_phase_rows if r["project_code"] == sample_code]
        print(f"\nSample — {sample_code}:")
        for r in sample_rows[:6]:
            fee = r.get("fixed_fee")
            print(f"  L2={r['level2_code']} L3={r['level3_code'] or '----'}  "
                  f"{(r['level2_name'] if not r['level3_code'] else r['level3_name'])[:50]}  "
                  f"fee={fee}  start={r['start_date']}")
        print("[DRY RUN] — no DB writes.")
        return

    # Write phases to Supabase
    print(f"\nClearing existing DAL phases...")
    sb.table("project_phases").delete().eq("office", OFFICE).execute()
    print("Cleared.")

    print(f"Inserting {len(all_phase_rows)} rows in batches of 500...")
    for i in range(0, len(all_phase_rows), 500):
        batch = all_phase_rows[i:i + 500]
        sb.table("project_phases").insert(batch).execute()
        print(f"  {min(i + 500, len(all_phase_rows))} / {len(all_phase_rows)}")

    final = sb.table("project_phases").select("id", count="exact") \
        .eq("office", OFFICE).execute()
    print(f"Done. {final.count} DAL phase rows in Supabase.")

    # Update project start/end dates
    print(f"\nUpdating start/end dates for {len(project_dates)} DAL projects...")
    updated = 0
    for code, (start, end) in project_dates.items():
        sb.table("projects").update({"start_date": start, "end_date": end}) \
            .eq("office", OFFICE).eq("project_code", code).execute()
        updated += 1
    print(f"Done. {updated} DAL projects updated with dates.")


if __name__ == "__main__":
    main()
