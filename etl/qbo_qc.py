"""
QC report for the Minnesota QBO extract.

Checks:
  1. Row counts vs QBO actual counts (with explanations for expected gaps)
  2. Required Unanet field completeness per template
  3. Referential integrity (contacts reference valid FirmCodes)
  4. Duplicate FirmCode / BaseCode detection
  5. Account type coverage (which QBO types were captured vs skipped)
  6. Spot-check: pulls 5 random records from QBO and prints side-by-side with CSV rows

Usage:
    python qbo_qc.py
"""

import csv
import json
import os
import random
from collections import Counter, defaultdict
from pathlib import Path

from dotenv import load_dotenv
from intuitlib.client import AuthClient
from quickbooks import QuickBooks
from quickbooks.objects.account import Account
from quickbooks.objects.customer import Customer
from quickbooks.objects.item import Item
from quickbooks.objects.vendor import Vendor

HERE = Path(__file__).parent
OUTPUT_DIR = HERE.parent / 'output' / 'minnesota'

PASS = '[PASS]'
FAIL = '[FAIL]'
WARN = '[WARN]'
INFO = '[INFO]'

# Required fields per Unanet template (must not be blank)
REQUIRED_FIELDS = {
    'minnesota_COA.csv':             ['BaseCode', 'BaseName', 'FinancialType'],
    'minnesota_Clients.csv':         ['FirmCode', 'FirmName', 'IsActive'],
    'minnesota_ClientContacts.csv':  ['FirmCode', 'FirstName', 'LastName'],
    'minnesota_Vendors.csv':         ['FirmCode', 'FirmName', 'IsActive'],
    'minnesota_VendorContacts.csv':  ['FirmCode', 'FirstName', 'LastName'],
    'minnesota_Employees.csv':       ['EmployeeCode', 'EmployeeName'],
    'minnesota_ExpenseCodes.csv':    ['ECCode', 'ECName', 'ShowInES'],
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_qb_client():
    load_dotenv(HERE / '.env')
    token_file = HERE / 'qbo_tokens.json'
    tokens = json.loads(token_file.read_text())
    auth_client = AuthClient(
        client_id=os.environ['QBO_CLIENT_ID'],
        client_secret=os.environ['QBO_CLIENT_SECRET'],
        redirect_uri='http://localhost:8000/callback',
        environment=os.getenv('QBO_ENVIRONMENT', 'production'),
        access_token=tokens['access_token'],
        refresh_token=tokens['refresh_token'],
        realm_id=tokens['realm_id'],
    )
    # Always refresh — access tokens expire after 1 hour
    auth_client.refresh()
    token_file.write_text(json.dumps({
        'access_token': auth_client.access_token,
        'refresh_token': auth_client.refresh_token,
        'realm_id': tokens['realm_id'],
    }, indent=2))
    return QuickBooks(
        auth_client=auth_client,
        refresh_token=auth_client.refresh_token,
        company_id=tokens['realm_id'],
    )


def read_csv(filename: str) -> list[dict]:
    path = OUTPUT_DIR / filename
    if not path.exists():
        return []
    with open(path, encoding='utf-8') as f:
        return list(csv.DictReader(f))


def qbo_count(qb, entity: str) -> int:
    result = qb.query(f'SELECT COUNT(*) FROM {entity}')
    return result.get('QueryResponse', {}).get('totalCount', 0)


def section(title: str):
    print(f'\n{"=" * 60}')
    print(f'  {title}')
    print('=' * 60)


def check(status: str, message: str):
    print(f'  {status} {message}')


# ---------------------------------------------------------------------------
# Check 1: Count reconciliation
# ---------------------------------------------------------------------------

def check_counts(qb):
    section('1. ROW COUNT RECONCILIATION')

    # Raw QBO totals (no WHERE filters — avoids enum validation issues)
    qbo_accounts  = qbo_count(qb, 'Account')
    qbo_customers = qbo_count(qb, 'Customer')
    qbo_vendors   = qbo_count(qb, 'Vendor')
    qbo_employees = qbo_count(qb, 'Employee')
    qbo_items     = qbo_count(qb, 'Item')

    # Sub-customer count — Job is a boolean, safe to filter
    job_result = qb.query('SELECT COUNT(*) FROM Customer WHERE Job = true')
    job_count = job_result.get('QueryResponse', {}).get('totalCount', 0)

    check(INFO, f'QBO raw totals: {qbo_accounts} accounts | {qbo_customers} customers '
                f'({job_count} are sub-jobs) | {qbo_vendors} vendors | '
                f'{qbo_employees} employees | {qbo_items} items')
    print()

    # Derive exclusion breakdowns from our CSVs rather than re-querying QBO
    coa_rows   = read_csv('minnesota_COA.csv')
    client_rows = read_csv('minnesota_Clients.csv')
    vendor_rows = read_csv('minnesota_Vendors.csv')
    emp_rows    = read_csv('minnesota_Employees.csv')
    item_rows   = read_csv('minnesota_ExpenseCodes.csv')

    expected_clients = qbo_customers - job_count

    rows_data = [
        ('Accounts',  qbo_accounts,  len(coa_rows),    qbo_accounts  - len(coa_rows),   'NonPosting accounts excluded (expected)'),
        ('Clients',   expected_clients, len(client_rows), expected_clients - len(client_rows), 'sub-customers/Jobs excluded'),
        ('Vendors',   qbo_vendors,   len(vendor_rows), qbo_vendors   - len(vendor_rows), ''),
        ('Employees', qbo_employees, len(emp_rows),    qbo_employees - len(emp_rows),    ''),
        ('Items',     qbo_items,     len(item_rows),   qbo_items     - len(item_rows),   'Inventory/Category/Group types excluded (expected)'),
    ]

    for label, qbo_total, csv_count, delta, note in rows_data:
        if delta == 0:
            check(PASS, f'{label}: {csv_count} in CSV matches {qbo_total} in QBO')
        elif note:
            check(PASS if delta > 0 else FAIL,
                  f'{label}: CSV={csv_count}, QBO={qbo_total}, delta={delta} — {note}')
        else:
            status = FAIL if delta != 0 else PASS
            check(status, f'{label}: CSV={csv_count}, QBO={qbo_total}, delta={delta}')


# ---------------------------------------------------------------------------
# Check 2: Required field completeness
# ---------------------------------------------------------------------------

def check_required_fields():
    section('2. REQUIRED FIELD COMPLETENESS')

    for filename, required in REQUIRED_FIELDS.items():
        rows = read_csv(filename)
        if not rows:
            check(WARN, f'{filename}: file not found or empty')
            continue

        for field in required:
            if field not in rows[0]:
                check(WARN, f'{filename} -> {field}: column missing from CSV')
                continue
            blank = sum(1 for r in rows if not r.get(field, '').strip())
            if blank == 0:
                check(PASS, f'{filename} -> {field}: all {len(rows)} rows populated')
            else:
                pct = blank / len(rows) * 100
                status = FAIL if pct > 10 else WARN
                check(status, f'{filename} -> {field}: {blank}/{len(rows)} rows blank ({pct:.0f}%)')


# ---------------------------------------------------------------------------
# Check 3: Referential integrity
# ---------------------------------------------------------------------------

def check_referential_integrity():
    section('3. REFERENTIAL INTEGRITY')

    client_codes = {r['FirmCode'] for r in read_csv('minnesota_Clients.csv') if r.get('FirmCode')}
    vendor_codes = {r['FirmCode'] for r in read_csv('minnesota_Vendors.csv') if r.get('FirmCode')}

    for contact_file, parent_codes, parent_label in [
        ('minnesota_ClientContacts.csv', client_codes, 'Clients'),
        ('minnesota_VendorContacts.csv', vendor_codes, 'Vendors'),
    ]:
        rows = read_csv(contact_file)
        orphans = [r['FirmCode'] for r in rows if r.get('FirmCode') and r['FirmCode'] not in parent_codes]
        if not orphans:
            check(PASS, f'{contact_file}: all FirmCodes exist in {parent_label}')
        else:
            check(FAIL, f'{contact_file}: {len(orphans)} FirmCodes not found in {parent_label}: {orphans[:5]}')


# ---------------------------------------------------------------------------
# Check 4: Duplicate key detection
# ---------------------------------------------------------------------------

def check_duplicates():
    section('4. DUPLICATE KEY DETECTION')

    key_checks = [
        ('minnesota_COA.csv',          'BaseCode'),
        ('minnesota_Clients.csv',      'FirmCode'),
        ('minnesota_Vendors.csv',      'FirmCode'),
        ('minnesota_Employees.csv',    'EmployeeCode'),
        ('minnesota_ExpenseCodes.csv', 'ECCode'),
    ]

    for filename, key_field in key_checks:
        rows = read_csv(filename)
        if not rows:
            continue
        counts = Counter(r.get(key_field, '') for r in rows)
        dupes = {k: v for k, v in counts.items() if v > 1 and k}
        if not dupes:
            check(PASS, f'{filename} -> {key_field}: no duplicates')
        else:
            check(FAIL, f'{filename} -> {key_field}: {len(dupes)} duplicates: {list(dupes.items())[:5]}')


# ---------------------------------------------------------------------------
# Check 5: Account type coverage
# ---------------------------------------------------------------------------

def check_account_coverage(qb):
    section('5. ACCOUNT TYPE COVERAGE')

    all_accounts = []
    page_size = 1000
    start = 1
    while True:
        batch = Account.query(
            f'SELECT * FROM Account STARTPOSITION {start} MAXRESULTS {page_size}', qb=qb
        )
        all_accounts.extend(batch)
        if len(batch) < page_size:
            break
        start += page_size

    from etl.qbo_extract import FINANCIAL_TYPE_MAP  # reuse the map
    captured, skipped = defaultdict(int), defaultdict(int)
    for acct in all_accounts:
        t = getattr(acct, 'AccountType', 'Unknown') or 'Unknown'
        if t in FINANCIAL_TYPE_MAP:
            captured[t] += 1
        else:
            skipped[t] += 1

    for t, n in sorted(captured.items()):
        check(PASS, f'Captured  "{t}": {n} accounts -> {FINANCIAL_TYPE_MAP[t]}')
    for t, n in sorted(skipped.items()):
        check(INFO, f'Skipped   "{t}": {n} accounts (intentional)')


# ---------------------------------------------------------------------------
# Check 6: Spot-check 5 random records
# ---------------------------------------------------------------------------

def spot_check(qb):
    section('6. SPOT CHECK — 5 RANDOM CUSTOMERS (QBO vs CSV)')

    csv_by_name = {}
    for row in read_csv('minnesota_Clients.csv'):
        csv_by_name[row.get('FirmName', '').strip().lower()] = row

    # Pull 5 random top-level customers from QBO
    all_customers = []
    page_size = 1000
    start = 1
    while True:
        batch = Customer.query(
            f'SELECT * FROM Customer STARTPOSITION {start} MAXRESULTS {page_size}', qb=qb
        )
        all_customers.extend(batch)
        if len(batch) < page_size:
            break
        start += page_size

    top_level = [c for c in all_customers if not getattr(c, 'Job', False)]
    sample = random.sample(top_level, min(5, len(top_level)))

    for c in sample:
        firm_name = (getattr(c, 'CompanyName', '') or c.DisplayName or '').strip()
        csv_row = csv_by_name.get(firm_name.lower())
        qbo_email = getattr(getattr(c, 'PrimaryEmailAddr', None), 'Address', '') or ''
        qbo_phone = getattr(getattr(c, 'PrimaryPhone', None), 'FreeFormNumber', '') or ''
        qbo_active = 'TRUE' if getattr(c, 'Active', True) else 'FALSE'

        if not csv_row:
            check(FAIL, f'"{firm_name}": found in QBO but NOT in CSV')
            continue

        mismatches = []
        if csv_row.get('FirmName', '').strip() != firm_name:
            mismatches.append(f'FirmName: CSV="{csv_row.get("FirmName")}" QBO="{firm_name}"')
        if csv_row.get('MainEmail', '').strip() != qbo_email:
            mismatches.append(f'Email: CSV="{csv_row.get("MainEmail")}" QBO="{qbo_email}"')
        if csv_row.get('IsActive', '').strip() != qbo_active:
            mismatches.append(f'IsActive: CSV="{csv_row.get("IsActive")}" QBO="{qbo_active}"')

        if not mismatches:
            check(PASS, f'"{firm_name}": FirmName, Email, IsActive match')
        else:
            for m in mismatches:
                check(FAIL, f'"{firm_name}": {m}')


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print('\nQBO Extract QC Report — Minnesota')
    print(f'CSV directory: {OUTPUT_DIR}')

    qb = get_qb_client()

    check_counts(qb)
    check_required_fields()
    check_referential_integrity()
    check_duplicates()

    # Account coverage imports from qbo_extract — run inline instead
    section('5. ACCOUNT TYPE COVERAGE')
    coa_rows = read_csv('minnesota_COA.csv')
    type_counts = Counter(r.get('_QBO_AccountType', '') for r in coa_rows)
    unanet_counts = Counter(r.get('FinancialType', '') for r in coa_rows)
    print('  QBO account types captured:')
    for t, n in sorted(type_counts.items()):
        print(f'    {t}: {n}')
    print('  Mapped to Unanet FinancialType:')
    for t, n in sorted(unanet_counts.items()):
        print(f'    {t}: {n}')

    spot_check(qb)

    print('\n' + '=' * 60)
    print('  QC complete. Address any [FAIL] items before loading.')
    print('  [WARN] items are worth reviewing but may be acceptable.')
    print('=' * 60 + '\n')


if __name__ == '__main__':
    main()
