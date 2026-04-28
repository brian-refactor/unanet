"""
Parse QuickBooks Desktop memorized-report CSV exports and write output CSVs
in the same format as the QBO and Ajera extractors.

Usage:
    python etl/qbd_parse.py --office orlando --input input/
    python etl/qbd_parse.py --office dallas  --input input/

Input files expected (case-insensitive match on prefix):
    <OFFICE>_COA.csv
    <OFFICE>_Customers.csv
    <OFFICE>_Vendor.csv          (QB Desktop names it "Vendor", not "Vendors")
    <OFFICE>_Employees.csv
    <OFFICE>_Item Price List.csv

Output:
    output/<office>/<office>_COA.csv
    output/<office>/<office>_Clients.csv
    output/<office>/<office>_ClientContacts.csv
    output/<office>/<office>_Vendors.csv
    output/<office>/<office>_VendorContacts.csv
    output/<office>/<office>_Employees.csv
    output/<office>/<office>_ExpenseCodes.csv
"""

import argparse
import csv
import datetime
import re
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

OFFICE_PREFIXES = {
    'dallas':  'DAL-',
    'orlando': 'ORL-',
}

# QB Desktop account type → Unanet FinancialType
FINANCIAL_TYPE_MAP = {
    'Bank':                    'Asset',
    'Accounts Receivable':     'Asset',
    'Other Current Asset':     'Asset',
    'Fixed Asset':             'Asset',
    'Other Asset':             'Asset',
    'Accounts Payable':        'Liability',
    'Credit Card':             'Liability',
    'Other Current Liability': 'Liability',
    'Long Term Liability':     'Liability',
    'Equity':                  'Equity',
    'Income':                  'Revenue',
    'Cost of Goods Sold':      'Expense',
    'Expense':                 'Expense',
    'Other Income':            'Revenue',
    'Other Expense':           'Expense',
    'Non-Posting':             None,   # excluded
}

EXPENSE_ITEM_TYPES = {'Service', 'Other Charge'}

# Matches "City, ST 12345" or "City  ST  12345" or "City Arizona 85028" (state spelled out)
_CSZ_RE = re.compile(r'^(.*?),?\s+([A-Za-z]{2,})\s+(\d{5}(?:-\d{4})?)\s*$')

_US_STATES = {
    'AL','AK','AZ','AR','CA','CO','CT','DE','FL','GA','HI','ID','IL','IN',
    'IA','KS','KY','LA','ME','MD','MA','MI','MN','MS','MO','MT','NE','NV',
    'NH','NJ','NM','NY','NC','ND','OH','OK','OR','PA','RI','SC','SD','TN',
    'TX','UT','VT','VA','WA','WV','WI','WY','DC',
}
_ZIP_RE   = re.compile(r'^\d{5}(?:-\d{4})?$')
_EMAIL_RE = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')


# ---------------------------------------------------------------------------
# CSV helpers
# ---------------------------------------------------------------------------

def _cell_str(val) -> str:
    """Normalize a cell value to string regardless of xlsx type."""
    if val is None:
        return ''
    if isinstance(val, datetime.datetime):
        return val.strftime('%m/%d/%Y')
    if isinstance(val, (int, float)):
        return str(val)
    return str(val).strip()


def read_file(path: Path) -> tuple[list[str], list[dict]]:
    """Read a QB Desktop export (CSV or XLSX). Returns (headers, rows)."""
    if path.suffix.lower() in ('.xlsx', '.xls'):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        raw = [[_cell_str(c) for c in row] for row in ws.iter_rows(values_only=True)]
    else:
        with open(path, newline='', encoding='utf-8-sig') as f:
            raw = [[c.strip() for c in row] for row in csv.reader(f)]

    if not raw:
        return [], []

    # QB exports have a leading empty column — strip it
    headers = raw[0]
    if headers and headers[0] == '':
        headers = headers[1:]
        raw = [r[1:] for r in raw]

    rows = []
    for row in raw[1:]:
        while len(row) < len(headers):
            row.append('')
        d = {headers[i]: row[i] for i in range(len(headers))}
        if not any(d.values()):
            continue
        first_val = list(d.values())[0]
        if first_val.upper().startswith('TOTAL') or first_val.startswith('Report'):
            continue
        rows.append(d)
    return headers, rows


# Keep read_csv as an alias for backward compatibility
read_csv = read_file


def write_csv(path: Path, rows: list[dict], fieldnames: list[str]):
    if not rows:
        print(f'  [SKIP] {path.name} — no rows')
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
        w.writeheader()
        w.writerows(rows)
    print(f'  Wrote {len(rows):>4} rows -> {path}')


def find_input(input_dir: Path, office: str, pattern: str) -> Path | None:
    """Case-insensitive glob for office-prefixed file (csv or xlsx)."""
    for p in input_dir.iterdir():
        if (p.suffix.lower() in ('.csv', '.xlsx', '.xls')
                and p.name.lower().startswith(office.lower())
                and pattern.lower() in p.name.lower()):
            return p
    return None


def active_str(row: dict) -> str:
    """Return 'TRUE'/'FALSE' from whichever active/status column is present."""
    for col in ('Active Status', 'Status', 'Active'):
        val = row.get(col, '').strip().lower()
        if val:
            return 'FALSE' if val in ('not-active', 'not active', 'inactive', 'no', 'false', '0', 'released', 'terminated') else 'TRUE'
    return 'TRUE'


def slugify(name: str) -> str:
    """Make a clean code from a display name."""
    return re.sub(r'[^A-Za-z0-9_-]', '_', name.strip()).strip('_')


def _parse_qbd_address(r: dict) -> tuple[str, str, str, str, str]:
    """Return (street1, street2, city, state, zip) from QB Desktop 'Bill to 1-5' columns.

    QB Desktop puts city/state/zip as a combined string in whichever 'Bill to N'
    line happens to follow the last street line (varies by whether a suite exists).
    This function scans all lines for the embedded 'City, ST 12345' pattern so the
    split is always correct regardless of which slot the combined value landed in.
    """
    raw = [r.get(f'Bill to {i}', '').strip() for i in range(1, 6)]

    # Fall back to newline-delimited combined 'Bill to' field
    if not any(raw):
        parts = [p.strip() for p in r.get('Bill to', '').split('\n') if p.strip()]
        raw = (parts + [''] * 5)[:5]

    city = state = zip_code = ''
    street_lines: list[str] = []

    for line in raw:
        if not line:
            continue
        m = _CSZ_RE.match(line)
        if m and not city:
            city     = m.group(1).strip().rstrip(',')
            state    = m.group(2).strip()
            zip_code = m.group(3).strip()
        else:
            street_lines.append(line)

    # If no combined CSZ line was found, treat raw[2/3/4] as separate city/state/zip
    if not city:
        street_lines = [l for l in raw[:2] if l]
        city     = raw[2]
        state    = raw[3]
        zip_code = raw[4]

    return (
        street_lines[0] if street_lines else '',
        street_lines[1] if len(street_lines) > 1 else '',
        city, state, zip_code,
    )


# ---------------------------------------------------------------------------
# Parsers
# ---------------------------------------------------------------------------

def parse_coa(path: Path, prefix: str, office: str, output_dir: Path):
    _, rows = read_csv(path)

    out = []
    for r in rows:
        acct_full = r.get('Account', '').strip()
        acct_type = r.get('Type', '').strip()
        if not acct_full or not acct_type:
            continue

        fin_type = FINANCIAL_TYPE_MAP.get(acct_type)
        if fin_type is None:
            continue  # NonPosting excluded

        # Sub-accounts use "Parent:Child" notation — leaf name is the display name
        if ':' in acct_full:
            base_name = acct_full.split(':')[-1].strip()
        else:
            base_name = acct_full

        # BaseCode: full path with colons → underscores, spaces → underscores
        base_code = re.sub(r'[\s:]+', '_', acct_full).strip('_')

        out.append({
            'BaseCode':          base_code,
            'BaseName':          base_name,
            'Description':       '',
            'IsActive':          active_str(r),
            'Is1099':            '',
            'IsSubcontractor':   '',
            'FinancialType':     fin_type,
            'SubledgerType':     '',
            'MetricType':        '',
            'CostType':          '',
            'PMType':            '',
            'LaborRevenueType':  '',
            'ExpenseRevenueType': '',
            '_QBD_FullName':     acct_full,
            '_QBD_AcctType':     acct_type,
        })

    write_csv(
        output_dir / f'{office}_COA.csv',
        out,
        ['BaseCode', 'BaseName', 'Description', 'IsActive', 'Is1099', 'IsSubcontractor',
         'FinancialType', 'SubledgerType', 'MetricType', 'CostType', 'PMType',
         'LaborRevenueType', 'ExpenseRevenueType', '_QBD_FullName', '_QBD_AcctType'],
    )


def _make_contact(firm_code: str, relationship: str, first: str, last: str, r: dict) -> dict:
    return {
        'FirmCode':         firm_code,
        'FirmRelationship': relationship,
        'Prefix':    '',
        'Suffix':    '',
        'Title':     r.get('Job Title', '').strip(),
        'FirstName': first,
        'LastName':  last,
        'WorkPhone': r.get('Main Phone', '').strip(),
        'CellPhone': r.get('Alt. Phone', '').strip(),
        'WorkEmail': r.get('Main Email', '').strip(),
        'HomeEmail': '',
        'WorkAddress1': '', 'WorkAddress2': '', 'WorkAddress3': '', 'WorkAddress4': '',
        'WorkCity': '', 'WorkState': '', 'WorkZip': '', 'WorkCountry': '',
        'HomeAddress1': '', 'HomeAddress2': '', 'HomeAddress3': '', 'HomeAddress4': '',
        'HomeCity': '', 'HomeState': '', 'HomeZip': '', 'HomeCountry': '',
    }


def _split_name(name: str) -> tuple[str, str]:
    """Return (first, last) from 'Last, First' or 'First Last' formats."""
    if ',' in name:
        last, first = [p.strip() for p in name.split(',', 1)]
    else:
        parts = name.split(None, 1)
        first, last = parts[0], (parts[1] if len(parts) > 1 else '')
    return first, last


def parse_customers(path: Path, prefix: str, office: str, output_dir: Path):
    _, rows = read_csv(path)

    parent_rows = [r for r in rows if r.get('Customer', '').strip() and ':' not in r['Customer']]
    child_rows  = [r for r in rows if r.get('Customer', '').strip() and ':' in r['Customer']]

    clients, contacts = [], []
    cli_seq = 0
    name_to_code: dict[str, str] = {}

    # ── Pass 1: parent rows → clients + primary contacts ─────────────────
    for r in parent_rows:
        name = r['Customer'].strip()

        company = r.get('Company', '').strip()
        if not company:
            company = re.sub(r'^\d+[\.\d]*\s*', '', name).strip() or name

        cli_seq += 1
        firm_code = f"{prefix}CLI-{cli_seq:05d}"
        name_to_code[name] = firm_code

        # Bill to 1 = billing name, Bill to 2 = street, Bill to 3 = suite
        # City and State are now separate columns; Zip falls back to Bill to 5
        bill_street1 = r.get('Bill to 1', '').strip()
        bill_street2 = r.get('Bill to 2', '').strip()
        bill_street3 = r.get('Bill to 3', '').strip()
        bill_city    = r.get('City', '').strip()
        bill_state   = r.get('State', '').strip()
        bill_zip     = r.get('Zip', r.get('Bill to 5', '')).strip()

        clients.append({
            'FirmCode':  firm_code,
            'FirmName':  company,
            'IsActive':  active_str(r),
            'Website':   '',
            'ClientType': '',
            'Specialty': '',
            'Note':      '',
            'PayDays':   '',
            'MainEmail': r.get('Main Email', '').strip(),
            'BillToAddress_Phone':   r.get('Main Phone', '').strip(),
            'BillToAddress_Street1': bill_street1,
            'BillToAddress_Street2': bill_street2,
            'BillToAddress_Street3': bill_street3,
            'BillToAddress_Street4': '',
            'BillToAddress_City':    bill_city,
            'BillToAddress_State':   bill_state,
            'BillToAddress_Zip':     bill_zip,
            'BillToAddress_Country': '',
            'MainContact_Prefix':    '',
            'MainContact_Suffix':    '',
            'MainContact_Title':     '',
            'MainContact_FirstName': '',
            'MainContact_LastName':  '',
            'MainContact_WorkPhone': '',
            'MainContact_CellPhone': '',
            'MainContact_WorkEmail': '',
            'MainContact_HomeEmail': '',
        })

        contact_name = r.get('Primary Contact', '').strip()
        if contact_name:
            first, last = _split_name(contact_name)
            contacts.append(_make_contact(firm_code, 'Primary', first, last, r))

    # ── Pass 2: child rows → additional contacts ──────────────────────────
    orphans = 0
    for r in child_rows:
        parent_name = r['Customer'].split(':')[0].strip()

        firm_code = name_to_code.get(parent_name)
        if not firm_code:
            orphans += 1
            continue

        first = r.get('First Name', '').strip()
        last  = r.get('Last Name', '').strip()
        if not first and not last:
            continue

        contacts.append(_make_contact(firm_code, 'Contact', first, last, r))

    if orphans:
        print(f'  [WARN] {orphans} child rows had no matching parent client — skipped')

    # Deduplicate contacts on (FirmCode, FirstName, LastName, WorkEmail) — case-insensitive
    seen_contacts: set[tuple] = set()
    unique_contacts = []
    for c in contacts:
        key = (
            c['FirmCode'],
            c['FirstName'].lower(),
            c['LastName'].lower(),
            c['WorkEmail'].lower(),
        )
        if key not in seen_contacts:
            seen_contacts.add(key)
            unique_contacts.append(c)
    dupes = len(contacts) - len(unique_contacts)
    if dupes:
        print(f'  Deduped {dupes} duplicate contact rows')
    contacts = unique_contacts

    write_csv(
        output_dir / f'{office}_Clients.csv',
        clients,
        ['FirmCode', 'FirmName', 'IsActive', 'Website', 'ClientType', 'Specialty', 'Note', 'PayDays',
         'MainEmail', 'BillToAddress_Phone',
         'BillToAddress_Street1', 'BillToAddress_Street2', 'BillToAddress_Street3', 'BillToAddress_Street4',
         'BillToAddress_City', 'BillToAddress_State', 'BillToAddress_Zip', 'BillToAddress_Country',
         'MainContact_Prefix', 'MainContact_Suffix', 'MainContact_Title',
         'MainContact_FirstName', 'MainContact_LastName',
         'MainContact_WorkPhone', 'MainContact_CellPhone', 'MainContact_WorkEmail', 'MainContact_HomeEmail'],
    )
    write_csv(
        output_dir / f'{office}_ClientContacts.csv',
        contacts,
        ['FirmCode', 'FirmRelationship', 'Prefix', 'Suffix', 'Title',
         'FirstName', 'LastName', 'WorkPhone', 'CellPhone', 'WorkEmail', 'HomeEmail',
         'WorkAddress1', 'WorkAddress2', 'WorkAddress3', 'WorkAddress4',
         'WorkCity', 'WorkState', 'WorkZip', 'WorkCountry',
         'HomeAddress1', 'HomeAddress2', 'HomeAddress3', 'HomeAddress4',
         'HomeCity', 'HomeState', 'HomeZip', 'HomeCountry'],
    )


def parse_vendors(path: Path, prefix: str, office: str, output_dir: Path):
    _, rows = read_csv(path)

    vendors, contacts = [], []
    vnd_seq = 0
    for r in rows:
        name = r.get('Vendor', '').strip()
        if not name:
            continue

        company = r.get('Company', '').strip() or name
        vnd_seq += 1
        firm_code = f"{prefix}VND-{vnd_seq:05d}"

        vendor_type = r.get('Vendor Type', '').strip()
        terms       = r.get('Terms', '').strip()
        net_m       = re.search(r'\d+', terms)
        net_days    = net_m.group() if net_m else ''

        vendors.append({
            'FirmCode':  firm_code,
            'FirmName':  company,
            'IsActive':  active_str(r),
            'IsConsultant': 'TRUE' if 'consultant' in vendor_type.lower() else 'FALSE',
            'ConsultantType': '',
            'Is1099':    'TRUE' if r.get('Eligible for 1099', '').strip().lower() in ('yes', 'true', '1') else 'FALSE',
            'Website':   '',
            'VendorType': vendor_type,
            'Note':      r.get('Note', '').strip(),
            'NetDays':   net_days,
            'EIN':       r.get('Tax ID', '').strip(),
            'PayToAddress_Phone':   r.get('Main Phone', '').strip(),
            'PayToAddress_Street1': r.get('Bill from Street 1', '').strip(),
            'PayToAddress_Street2': r.get('Bill from Street 2', '').strip(),
            'PayToAddress_Street3': '',
            'PayToAddress_Street4': '',
            'PayToAddress_City':    r.get('Bill from City', '').strip(),
            'PayToAddress_State':   r.get('Bill from State', '').strip(),
            'PayToAddress_Zip':     r.get('Bill from Zip', '').strip(),
            'PayToAddress_Country': '',
            'MainContact_Prefix':    '',
            'MainContact_Suffix':    '',
            'MainContact_Title':     '',
            'MainContact_FirstName': '',
            'MainContact_LastName':  '',
            'MainContact_WorkPhone': '',
            'MainContact_CellPhone': '',
            'MainContact_WorkEmail': '',
            'MainContact_HomeEmail': '',
            'EnableEFT':   '',
            'CompanyID':   '',
            'CompanyName': '',
            'ABA/Routing': '',
            'Account#':    '',
            'Savings':     '',
            'EFType(SEC)': '',
        })

        contact_name = r.get('Primary Contact', '').strip()
        if contact_name:
            parts = contact_name.split(None, 1)
            contacts.append({
                'FirmCode':         firm_code,
                'FirmRelationship': 'Primary',
                'Prefix':    '',
                'Suffix':    '',
                'Title':     '',
                'FirstName': parts[0],
                'LastName':  parts[1] if len(parts) > 1 else '',
                'WorkPhone': r.get('Main Phone', '').strip(),
                'CellPhone': '',
                'WorkEmail': r.get('Main Email', '').strip(),
                'HomeEmail': '',
                'WorkAddress1': '', 'WorkAddress2': '', 'WorkAddress3': '', 'WorkAddress4': '',
                'WorkCity': '', 'WorkState': '', 'WorkZip': '', 'WorkCountry': '',
                'HomeAddress1': '', 'HomeAddress2': '', 'HomeAddress3': '', 'HomeAddress4': '',
                'HomeCity': '', 'HomeState': '', 'HomeZip': '', 'HomeCountry': '',
            })

    write_csv(
        output_dir / f'{office}_Vendors.csv',
        vendors,
        ['FirmCode', 'FirmName', 'IsActive', 'IsConsultant', 'ConsultantType', 'Is1099',
         'Website', 'VendorType', 'Note', 'NetDays', 'EIN',
         'PayToAddress_Phone',
         'PayToAddress_Street1', 'PayToAddress_Street2', 'PayToAddress_Street3', 'PayToAddress_Street4',
         'PayToAddress_City', 'PayToAddress_State', 'PayToAddress_Zip', 'PayToAddress_Country',
         'MainContact_Prefix', 'MainContact_Suffix', 'MainContact_Title',
         'MainContact_FirstName', 'MainContact_LastName',
         'MainContact_WorkPhone', 'MainContact_CellPhone', 'MainContact_WorkEmail', 'MainContact_HomeEmail',
         'MainContact_HomeEmail',
         'EnableEFT', 'CompanyID', 'CompanyName',
         'ABA/Routing', 'Account#', 'Savings', 'EFType(SEC)'],
    )
    write_csv(
        output_dir / f'{office}_VendorContacts.csv',
        contacts,
        ['FirmCode', 'FirmRelationship', 'Prefix', 'Suffix', 'Title',
         'FirstName', 'LastName', 'WorkPhone', 'CellPhone', 'WorkEmail', 'HomeEmail',
         'WorkAddress1', 'WorkAddress2', 'WorkAddress3', 'WorkAddress4',
         'WorkCity', 'WorkState', 'WorkZip', 'WorkCountry',
         'HomeAddress1', 'HomeAddress2', 'HomeAddress3', 'HomeAddress4',
         'HomeCity', 'HomeState', 'HomeZip', 'HomeCountry'],
    )


def parse_employees(path: Path, prefix: str, office: str, output_dir: Path):
    _, rows = read_csv(path)

    out = []
    emp_seq = 0
    for r in rows:
        name = r.get('Employee', '').strip()
        if not name:
            continue

        # Split "Last, First" or "First Last" name formats
        if ',' in name:
            last, first = [p.strip() for p in name.split(',', 1)]
        else:
            parts = name.split(None, 1)
            first = parts[0]
            last = parts[1] if len(parts) > 1 else ''

        emp_seq += 1
        emp_code = f"{prefix}EMP-{emp_seq:05d}"

        out.append({
            'EmployeeCode':      emp_code,
            'EmployeeName':      name,
            'PayRate':           '',
            'salaryperpayperiod': '',
            'PayRateStartDate':  '',
            'PayRateEndDate':    '',
            'IsHourly':          '',
            'OTRate':            '',
            'OTMU':              '',
            '_FirstName':        first,
            '_LastName':         last,
            '_IsActive':         active_str(r),
            '_Phone':            r.get('Main Phone', '').strip(),
            '_Email':            r.get('Main Email', r.get('Email', '')).strip(),
        })

    write_csv(
        output_dir / f'{office}_Employees.csv',
        out,
        ['EmployeeCode', 'EmployeeName', 'PayRate', 'salaryperpayperiod',
         'PayRateStartDate', 'PayRateEndDate', 'IsHourly', 'OTRate', 'OTMU',
         '_FirstName', '_LastName', '_IsActive', '_Phone', '_Email'],
    )


def parse_items(path: Path, prefix: str, office: str, output_dir: Path):
    _, rows = read_csv(path)

    out = []
    exp_seq = 0
    for r in rows:
        name = r.get('Item', '').strip()
        item_type = r.get('Type', '').strip()
        if not name:
            continue
        if item_type not in EXPENSE_ITEM_TYPES:
            continue  # skip Inventory, Group, etc.

        exp_seq += 1
        out.append({
            'ECCode':   f"{prefix}EXP-{exp_seq:05d}",
            'ECName':   name,
            'ShowInES': 'TRUE',
            'IsUnit':   'FALSE',
            'UnitTypename': '',
            'ECTypename': '',
            'ExpMarkupTypename': '',
            'Markup':   '',
            'UnitRate': '',
            'BillStatusname': '',
            'DirectBaseCode': '',
            'DirectBasename': '',
            'OHBaseCode': '',
            'OHBasename': '',
            'BilledDirectBaseCode': '',
            'BilledDirectBasename': '',
            'BilledMarkupBaseCode': '',
            'BilledMarkupBasename': '',
            'UnBilledBaseCode': '',
            'UnBilledBasename': '',
            'CurrencyCode': '',
            'PMCmtRequired': '',
            'IntCmtRequired': '',
            'IsNonReim': '',
            '_IsActive': active_str(r),
            '_QBD_Type': item_type,
        })

    write_csv(
        output_dir / f'{office}_ExpenseCodes.csv',
        out,
        ['ECCode', 'ECName', 'ShowInES', 'IsUnit', 'UnitTypename', 'ECTypename',
         'ExpMarkupTypename', 'Markup', 'UnitRate', 'BillStatusname',
         'DirectBaseCode', 'DirectBasename', 'OHBaseCode', 'OHBasename',
         'BilledDirectBaseCode', 'BilledDirectBasename',
         'BilledMarkupBaseCode', 'BilledMarkupBasename',
         'UnBilledBaseCode', 'UnBilledBasename',
         'CurrencyCode', 'PMCmtRequired', 'IntCmtRequired', 'IsNonReim',
         '_IsActive', '_QBD_Type'],
    )


# ---------------------------------------------------------------------------
# QC
# ---------------------------------------------------------------------------

_QC_SPECS = [
    {
        'file':     '{office}_Clients.csv',
        'key':      'FirmCode',
        'required': ['FirmCode', 'FirmName'],
        'state':    'BillToAddress_State',
        'zip':      'BillToAddress_Zip',
        'email':    'MainEmail',
    },
    {
        'file':     '{office}_Vendors.csv',
        'key':      'FirmCode',
        'required': ['FirmCode', 'FirmName'],
        'state':    'PayToAddress_State',
        'zip':      'PayToAddress_Zip',
    },
    {
        'file':     '{office}_Employees.csv',
        'key':      'EmployeeCode',
        'required': ['EmployeeCode', 'EmployeeName'],
    },
    {
        'file':     '{office}_ExpenseCodes.csv',
        'key':      'ECCode',
        'required': ['ECCode', 'ECName'],
    },
]


def _qc_warn(label: str, bad: list[tuple[int, str]], max_sample: int = 5) -> None:
    sample = ', '.join(f'row {r}: "{v}"' for r, v in bad[:max_sample])
    tail = f'  (+{len(bad) - max_sample} more)' if len(bad) > max_sample else ''
    print(f'  WARN  {label}: {sample}{tail}')


def qc_outputs(output_dir: Path, office: str) -> int:
    """Run QC checks on all output CSVs. Returns total warning count."""
    print('\n--- QC Report ---')
    total_warns = 0

    for spec in _QC_SPECS:
        path = output_dir / spec['file'].format(office=office)
        if not path.exists():
            print(f'\n{path.name}: (not written — skipped)')
            continue

        with open(path, newline='', encoding='utf-8') as f:
            rows = list(csv.DictReader(f))

        print(f'\n{path.name}  ({len(rows)} rows)')
        if not rows:
            print('  WARN  File is empty')
            total_warns += 1
            continue

        warns = 0
        cols = set(rows[0].keys())

        # ── Duplicate primary key ─────────────────────────────────────────
        key_col = spec.get('key', '')
        if key_col in cols:
            from collections import Counter
            counts = Counter(r[key_col] for r in rows)
            dupes = [(k, v) for k, v in counts.items() if v > 1]
            if dupes:
                ds = ', '.join(f'{k} (x{v})' for k, v in dupes[:5])
                print(f'  WARN  Duplicate {key_col}: {ds}{"  ..." if len(dupes) > 5 else ""}')
                warns += 1
            else:
                print(f'  OK    {key_col}: {len(counts)} unique')

        # ── Required fields not blank ─────────────────────────────────────
        for col in spec.get('required', []):
            if col not in cols:
                continue
            blanks = [i + 2 for i, r in enumerate(rows) if not r[col].strip()]
            if blanks:
                sample = str(blanks[:5])[1:-1]
                tail = '  ...' if len(blanks) > 5 else ''
                print(f'  WARN  {col} blank: rows {sample}{tail}  ({len(blanks)} total)')
                warns += 1
            else:
                print(f'  OK    {col}: all populated')

        # ── State abbreviation ────────────────────────────────────────────
        state_col = spec.get('state', '')
        if state_col in cols:
            bad = [(i + 2, r[state_col].strip())
                   for i, r in enumerate(rows)
                   if r[state_col].strip() and r[state_col].strip().upper() not in _US_STATES]
            if bad:
                _qc_warn(f'{state_col} invalid (not a 2-letter US state)', bad)
                warns += 1
            else:
                filled = sum(1 for r in rows if r[state_col].strip())
                print(f'  OK    {state_col}: {filled} populated, all valid')

        # ── Zip code format ───────────────────────────────────────────────
        zip_col = spec.get('zip', '')
        if zip_col in cols:
            bad = [(i + 2, r[zip_col].strip())
                   for i, r in enumerate(rows)
                   if r[zip_col].strip() and not _ZIP_RE.match(r[zip_col].strip())]
            if bad:
                _qc_warn(f'{zip_col} invalid (expected 5 or 9 digits)', bad)
                warns += 1
            else:
                filled = sum(1 for r in rows if r[zip_col].strip())
                print(f'  OK    {zip_col}: {filled} populated, all valid')

        # ── Email basic format ────────────────────────────────────────────
        email_col = spec.get('email', '')
        if email_col in cols:
            bad = [(i + 2, r[email_col].strip())
                   for i, r in enumerate(rows)
                   if r[email_col].strip() and not _EMAIL_RE.match(r[email_col].strip())]
            if bad:
                _qc_warn(f'{email_col} malformed', bad)
                warns += 1

        if warns == 0:
            print('  All checks passed.')
        total_warns += warns

    print(f'\nQC complete — {total_warns} warning(s).')
    return total_warns


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

FILE_PATTERNS = {
    'coa':       'coa',
    'customers': 'customers',
    'vendors':   'vendor',        # QB names it "Vendor" not "Vendors"
    'employees': 'employees',
    'items':     'item price list',
}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--office', choices=['dallas', 'orlando'], required=True)
    parser.add_argument('--input', default='input', help='Directory containing QB export CSVs')
    args = parser.parse_args()

    office = args.office
    prefix = OFFICE_PREFIXES[office]
    input_dir = Path(args.input)
    output_dir = Path(__file__).parent.parent / 'output' / office

    print(f'\nQBD Parse — {office.upper()}')
    print(f'  Input:  {input_dir.resolve()}')
    print(f'  Output: {output_dir.resolve()}')
    print(f'  Prefix: {prefix}')
    print()

    # Locate input files
    files = {}
    for key, pattern in FILE_PATTERNS.items():
        match = find_input(input_dir, office, pattern)
        if match:
            files[key] = match
            print(f'  Found {key:12s}: {match.name}')
        else:
            print(f'  [WARN] {key:12s}: no file matching "*{pattern}*" in {input_dir}')

    print()

    if 'coa' in files:
        parse_coa(files['coa'], prefix, office, output_dir)
    if 'customers' in files:
        parse_customers(files['customers'], prefix, office, output_dir)
    if 'vendors' in files:
        parse_vendors(files['vendors'], prefix, office, output_dir)
    if 'employees' in files:
        parse_employees(files['employees'], prefix, office, output_dir)
    if 'items' in files:
        parse_items(files['items'], prefix, office, output_dir)

    qc_outputs(output_dir, office)
    print(f'\nDone. Review output/{office}/ before loading into Unanet templates.')


if __name__ == '__main__':
    main()
