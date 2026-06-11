"""
Rebuild coa_crosswalk so source_base_code values match the current coa table.

The crosswalk was built from an older extract where codes differed (e.g. CIN
used raw numeric codes, now uses AJ-NNN). This script:
  1. Reads every crosswalk entry's source_base_name + master_code (the mapping intent)
  2. Finds the current base_code in coa that has the same base_name
  3. Deletes all orphaned crosswalk entries (codes not in current coa)
  4. Re-inserts them with the correct current base_code

Entries that already match (source_base_code in current coa) are left untouched.
Entries whose name has no match in coa are dropped and reported.

Usage:
    python etl/rebuild_coa_crosswalk.py
    python etl/rebuild_coa_crosswalk.py --office cincinnati
    python etl/rebuild_coa_crosswalk.py --dry-run
"""
import argparse
import os
from pathlib import Path

from dotenv import load_dotenv
from supabase import create_client

load_dotenv(Path(__file__).parent / ".env")
sb = create_client(os.environ["SUPABASE_URL"], os.environ["SUPABASE_KEY"])

OFFICES = ["minnesota", "cincinnati", "dallas", "orlando"]


def fetch_all(table, office=None, select="*"):
    rows, page, size = [], 0, 1000
    while True:
        q = sb.table(table).select(select)
        if office:
            q = q.eq("office", office)
        batch = q.range(page, page + size - 1).execute().data
        rows.extend(batch)
        if len(batch) < size:
            break
        page += size
    return rows


def rebuild_office(office: str, dry_run: bool):
    print(f"\n{'[DRY RUN] ' if dry_run else ''}=== {office.upper()} ===")

    # Current coa: base_code is the canonical key, base_name is the stable label
    coa_rows = fetch_all("coa", office=office, select="base_code,base_name")
    name_to_code = {r["base_name"]: r["base_code"] for r in coa_rows}
    current_codes = {r["base_code"] for r in coa_rows}
    print(f"  COA accounts:       {len(coa_rows)}")

    # Current crosswalk
    cx_rows = fetch_all("coa_crosswalk", office=office,
                        select="id,source_base_code,source_base_name,master_code,mapped_by")
    print(f"  Crosswalk entries:  {len(cx_rows)}")

    # Split into matched (good) and orphaned (stale code)
    matched   = [r for r in cx_rows if r["source_base_code"] in current_codes]
    orphaned  = [r for r in cx_rows if r["source_base_code"] not in current_codes]
    print(f"  Already matched:    {len(matched)}")
    print(f"  Orphaned (to fix):  {len(orphaned)}")

    if not orphaned:
        print("  Nothing to do.")
        return

    # For each orphan, try to find new base_code by name
    to_insert = []
    no_match  = []
    for r in orphaned:
        name = r["source_base_name"] or ""
        new_code = name_to_code.get(name)
        if new_code:
            to_insert.append({
                "office":           office,
                "source_base_code": new_code,
                "source_base_name": name,
                "master_code":      r["master_code"],
                "mapped_by":        r["mapped_by"] or "rebuild-name-match",
            })
        else:
            no_match.append(r)

    print(f"  Name-matched:       {len(to_insert)}")
    print(f"  No name match:      {len(no_match)}")
    if no_match:
        print("  Unresolvable orphans (will be dropped):")
        for r in no_match:
            print(f"    [{r['source_base_code']}] {r['source_base_name']} → master {r['master_code']}")

    if dry_run:
        print("  (dry run — no changes written)")
        return

    # Delete all orphaned entries
    orphan_ids = [r["id"] for r in orphaned]
    for i in range(0, len(orphan_ids), 500):
        batch_ids = orphan_ids[i:i+500]
        sb.table("coa_crosswalk").delete().in_("id", batch_ids).execute()
    print(f"  Deleted {len(orphaned)} orphaned entries.")

    # Insert rebuilt entries
    if to_insert:
        for i in range(0, len(to_insert), 500):
            sb.table("coa_crosswalk").upsert(
                to_insert[i:i+500],
                on_conflict="office,source_base_code"
            ).execute()
        print(f"  Inserted {len(to_insert)} rebuilt entries.")

    # Final count
    final = fetch_all("coa_crosswalk", office=office, select="source_base_code")
    final_codes = {r["source_base_code"] for r in final}
    matched_final = len(final_codes & current_codes)
    print(f"  Final: {len(final)} crosswalk, {matched_final} matched to current COA")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--office", choices=OFFICES, default=None,
                        help="Rebuild one office only (default: all)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Show what would change without writing")
    args = parser.parse_args()

    offices = [args.office] if args.office else OFFICES
    for office in offices:
        rebuild_office(office, dry_run=args.dry_run)

    print("\nDone.")


if __name__ == "__main__":
    main()
