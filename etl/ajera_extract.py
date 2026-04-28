"""
Ajera extractor — Cincinnati office (Reztark Design Studio, LLC)
Extracts all Unanet-relevant data via Ajera REST API and writes CSVs
to output/cincinnati/ for review before loading into templates.

Usage:
    python etl/ajera_extract.py

Output:
    output/cincinnati/cincinnati_COA.csv
    output/cincinnati/cincinnati_Clients.csv
    output/cincinnati/cincinnati_ClientContacts.csv
    output/cincinnati/cincinnati_Vendors.csv
    output/cincinnati/cincinnati_VendorContacts.csv
    output/cincinnati/cincinnati_Employees.csv        (pay rates expanded per row)
    output/cincinnati/cincinnati_ExpenseCodes.csv

Columns prefixed with _ are reference-only — remove before loading into Unanet.
"""

import csv
import json
import os
import re
from pathlib import Path

import requests
from dotenv import load_dotenv

HERE = Path(__file__).parent
OUTPUT_DIR = HERE.parent / 'output' / 'cincinnati'
OFFICE_PREFIX = 'CIN'
BATCH_SIZE = 50
HEADERS = {'Content-Type': 'application/json'}

# Ajera account type → Unanet FinancialType (case-insensitive match on lowercase)
FINANCIAL_TYPE_MAP = {
    'asset': 'Asset',
    'current asset': 'Asset',
    'fixed asset': 'Asset',
    'other asset': 'Asset',
    'liability': 'Liability',
    'current liability': 'Liability',
    'long-term liability': 'Liability',
    'long term liability': 'Liability',
    'other liability': 'Liability',
    'equity': 'Equity',
    'revenue': 'Revenue',
    'income': 'Revenue',
    'other income': 'Revenue',
    'expense': 'Expense',
    'direct expense': 'Expense',
    'indirect expense': 'Expense',
    'other expense': 'Expense',
    'cost of goods sold': 'Expense',
    'billable cost': 'Expense',
    'non-billable cost': 'Expense',
    'retained earnings': 'Equity',
    'non-current asset': 'Asset',
    'non-current liability': 'Liability',
}

SUBLEDGER_TYPE_MAP = {
    'accounts receivable': 'AR',
    'accounts payable': 'AP',
    'bank': 'Bank',
    'checking': 'Bank',
    'savings': 'Bank',
}


# ---------------------------------------------------------------------------
# Session management
# ---------------------------------------------------------------------------

def create_session() -> tuple[str, str]:
    load_dotenv(HERE / 'ajera.env')
    api_url  = os.environ['AJERA_API_URL']
    username = os.environ['AJERA_USERNAME']
    password = os.environ['AJERA_PASSWORD']

    resp = requests.post(api_url, json={
        'Method': 'CreateAPISession',
        'Username': username,
        'Password': password,
        'APIVersion': 1,
        'UseSessionCookie': False,
    }, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    token = data.get('Content', {}).get('SessionToken')
    if not token:
        raise SystemExit(f'Auth failed: {data.get("Errors")}')

    company = data['Content'].get('CompanyName', '')
    version = data['Content'].get('AjeraVersion', '')
    print(f'Connected — {company}  (Ajera {version})')
    return token, api_url


def end_session(api_url: str, token: str):
    try:
        requests.post(api_url, json={
            'Method': 'EndAPISession', 'SessionToken': token,
        }, headers=HEADERS, timeout=15)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# API call helper
# ---------------------------------------------------------------------------

def call(api_url: str, token: str, method: str, args: dict = None) -> tuple[dict, list]:
    payload = {'Method': method, 'SessionToken': token, 'MethodArguments': args or {}}
    resp = requests.post(api_url, json=payload, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    errors = [e for e in data.get('Errors', []) if e.get('ErrorID', 0) != 0]
    return data.get('Content', {}), errors


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def batched(lst: list, size: int):
    for i in range(0, len(lst), size):
        yield lst[i:i + size]


def make_firm_code(name: str, seen: set) -> str:
    clean = re.sub(r'[^A-Z0-9]', '', (name or '').upper())
    base = f'{OFFICE_PREFIX}-{clean[:8]}' if clean else f'{OFFICE_PREFIX}-UNK'
    code, n = base, 2
    while code in seen:
        code = f'{base}{n}'
        n += 1
    seen.add(code)
    return code


def is_active(status) -> str:
    return 'FALSE' if str(status or '').strip().lower() in ('inactive', 'false', '0') else 'TRUE'


def get_addr(obj: dict, src_prefix: str = 'PrimaryAddress', dest_prefix: str = '', line_field: str = 'Street') -> dict:
    """Extract address fields trying multiple Ajera field name variants."""
    def f(*keys):
        for k in keys:
            v = obj.get(k)
            if v:
                return str(v)
        return ''
    p = src_prefix
    return {
        f'{dest_prefix}{line_field}1': f(f'{p}LineOne',   f'{p}Line1',  f'{p}Address1'),
        f'{dest_prefix}{line_field}2': f(f'{p}LineTwo',   f'{p}Line2',  f'{p}Address2'),
        f'{dest_prefix}{line_field}3': f(f'{p}LineThree', f'{p}Line3'),
        f'{dest_prefix}{line_field}4': f(f'{p}LineFour',  f'{p}Line4'),
        f'{dest_prefix}City':    f(f'{p}City'),
        f'{dest_prefix}State':   f(f'{p}State',     f'{p}Province'),
        f'{dest_prefix}Zip':     f(f'{p}PostalCode', f'{p}Zip'),
        f'{dest_prefix}Country': f(f'{p}Country'),
    }


def first_phone(obj: dict) -> str:
    for key in ('PrimaryPhone1', 'PrimaryPhone', 'Phone', 'PhoneNumber', 'Phone1'):
        val = obj.get(key)
        if val:
            return str(val)
    return ''


def write_csv(path: Path, rows: list[dict], fieldnames: list[str] | None = None):
    if not rows:
        print(f'  (no rows for {path.name})')
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames or list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


# ---------------------------------------------------------------------------
# COA
# ---------------------------------------------------------------------------

def extract_coa(api_url: str, token: str):
    print('Extracting GL Accounts (COA)...')

    content, errors = call(api_url, token, 'ListGLAccounts')
    # Try both possible response key names
    all_accts = content.get('GLAccounts') or content.get('Accounts') or []

    if not all_accts:
        print(f'  ! ListGLAccounts returned no records. Errors: {errors}')
        print('    Keys in response:', list(content.keys()))
        return

    # Extract the key field — try both naming conventions
    def acct_key(a):
        return a.get('AccountKey') or a.get('GLAccountKey') or a.get('Key')

    keys = [acct_key(a) for a in all_accts if acct_key(a)]
    print(f'  Found {len(keys)} accounts — fetching details...')

    rows = []
    unknown_types = set()

    for batch in batched(keys, BATCH_SIZE):
        content, _ = call(api_url, token, 'GetGLAccounts', {'RequestedGLAccounts': batch})
        details = content.get('GLAccounts') or content.get('Accounts') or []

        for acct in details:
            raw_type = str(acct.get('AccountType') or acct.get('Type') or '').strip()
            financial_type = FINANCIAL_TYPE_MAP.get(raw_type.lower(), '')
            if raw_type and not financial_type:
                unknown_types.add(raw_type)

            raw_name = str(acct.get('AccountSubType') or acct.get('SubType') or raw_type).strip()
            subledger = SUBLEDGER_TYPE_MAP.get(raw_name.lower(), '')

            acct_num = (
                str(acct.get('AccountID') or '')
                or str(acct.get('AccountNumber') or '')
                or str(acct.get('Code') or '')
                or f"AJ-{acct.get('AccountKey') or acct.get('GLAccountKey') or ''}"
            )

            rows.append({
                'BaseCode':          acct_num,
                'BaseName':          acct.get('Description') or acct.get('Name') or '',
                'Description':       acct.get('Notes') or acct.get('LongDescription') or '',
                'IsActive':          is_active(acct.get('Status')),
                'Is1099':            '',
                'IsSubcontractor':   '',
                'FinancialType':     financial_type,
                'SubledgerType':     subledger,
                'MetricType':        '',
                'CostType':          '',
                'PMType':            '',
                'LaborRevenueType':  '',
                'ExpenseRevenueType':'',
                '_Ajera_Type':       raw_type,
            })

    write_csv(OUTPUT_DIR / 'cincinnati_COA.csv', rows)
    print(f'  -> {len(rows)} accounts')
    if unknown_types:
        print(f'  ! Unmapped account types (FinancialType blank — review required): {unknown_types}')


# ---------------------------------------------------------------------------
# Clients
# ---------------------------------------------------------------------------

def extract_clients(api_url: str, token: str):
    print('Extracting Clients...')

    content, _ = call(api_url, token, 'ListClients')
    all_clients = content.get('Clients', [])
    keys = [c['ClientKey'] for c in all_clients if c.get('ClientKey')]

    client_rows, contact_rows = [], []
    cli_seq = 0

    for batch in batched(keys, BATCH_SIZE):
        content, _ = call(api_url, token, 'GetClients', {'RequestedClients': batch})
        for c in content.get('Clients', []):
            name = c.get('Description') or c.get('Name') or ''
            cli_seq += 1
            firm_code = f'CIN-CLI-{cli_seq:05d}'
            addr = get_addr(c, 'PrimaryAddress', 'BillToAddress_')
            raw_addr = get_addr(c, 'PrimaryAddress')
            pay_days = (
                c.get('DaysToPay')
                or c.get('NumberOfDaysFromInvoiceDate')
                or c.get('NetDays')
                or 30
            )

            client_rows.append({
                'FirmCode':   firm_code,
                'FirmName':   name,
                'IsActive':   is_active(c.get('Status')),
                'Website':    c.get('Website') or '',
                'ClientType': c.get('ClientTypeDescription') or c.get('ClientType') or '',
                'Specialty':  '',
                'Note':       c.get('Notes') or '',
                'PayDays':    pay_days,
                'MainEmail':  c.get('Email') or '',
                'BillToAddress_Phone':   first_phone(c),
                **addr,
                'MainContact_Prefix':    '',
                'MainContact_Suffix':    '',
                'MainContact_Title':     '',
                'MainContact_FirstName': '',
                'MainContact_LastName':  '',
                'MainContact_WorkPhone': first_phone(c),
                'MainContact_CellPhone': '',
                'MainContact_WorkEmail': c.get('Email') or '',
                'MainContact_HomeEmail': '',
            })

            for contact in (c.get('Contacts') or []):
                first = contact.get('FirstName') or ''
                last  = contact.get('LastName')  or ''
                if not first and not last:
                    continue
                if contact.get('PrimaryAddressCity'):
                    ca = get_addr(contact, 'PrimaryAddress')
                else:
                    ca = raw_addr
                contact_rows.append({
                    'FirmCode':         firm_code,
                    'FirmRelationship': contact.get('ContactTypeDescription') or 'Primary',
                    'Prefix':           contact.get('Prefix') or '',
                    'Suffix':           contact.get('Suffix') or '',
                    'Title':            contact.get('Title') or '',
                    'FirstName':        first,
                    'LastName':         last,
                    'WorkPhone':        first_phone(contact),
                    'CellPhone':        contact.get('PrimaryPhone2') or contact.get('Mobile') or '',
                    'WorkEmail':        contact.get('Email') or '',
                    'HomeEmail':        '',
                    'WorkAddress1':     ca['Street1'],
                    'WorkAddress2':     ca['Street2'],
                    'WorkAddress3':     ca['Street3'],
                    'WorkAddress4':     ca['Street4'],
                    'WorkCity':         ca['City'],
                    'WorkState':        ca['State'],
                    'WorkZip':          ca['Zip'],
                    'WorkCountry':      ca['Country'],
                    'HomeAddress1': '', 'HomeAddress2': '', 'HomeAddress3': '', 'HomeAddress4': '',
                    'HomeCity': '', 'HomeState': '', 'HomeZip': '', 'HomeCountry': '',
                })

    write_csv(OUTPUT_DIR / 'cincinnati_Clients.csv', client_rows)
    write_csv(OUTPUT_DIR / 'cincinnati_ClientContacts.csv', contact_rows)
    print(f'  -> {len(client_rows)} clients, {len(contact_rows)} contacts')


# ---------------------------------------------------------------------------
# Vendors
# ---------------------------------------------------------------------------

def extract_vendors(api_url: str, token: str):
    print('Extracting Vendors...')

    content, _ = call(api_url, token, 'ListVendors')
    all_vendors = content.get('Vendors', [])
    keys = [v['VendorKey'] for v in all_vendors if v.get('VendorKey')]

    vendor_rows, contact_rows = [], []
    vnd_seq = 0

    for batch in batched(keys, BATCH_SIZE):
        content, _ = call(api_url, token, 'GetVendors', {'RequestedVendors': batch})
        for v in content.get('Vendors', []):
            name = v.get('Name') or ''
            vnd_seq += 1
            firm_code = f'CIN-VND-{vnd_seq:05d}'
            addr = get_addr(v, 'PrimaryAddress', 'PayToAddress_')
            raw_addr = get_addr(v, 'PrimaryAddress')
            net_days = (
                v.get('NumberOfDaysFromInvoiceDate')
                or v.get('NetDays')
                or v.get('DaysToPay')
                or 30
            )

            vendor_rows.append({
                'FirmCode':       firm_code,
                'FirmName':       name,
                'IsActive':       is_active(v.get('Status')),
                'IsConsultant':   'TRUE' if v.get('VendorTypeIsConsultant') else 'FALSE',
                'ConsultantType': v.get('VendorTypeDescription') or '',
                'Is1099':         'TRUE' if v.get('Receives1099Form') else 'FALSE',
                'Website':        v.get('Website') or '',
                'VendorType':     v.get('VendorTypeDescription') or '',
                'Note':           v.get('Notes') or '',
                'NetDays':        net_days,
                'EIN':            v.get('TaxIdentifier') or v.get('RecipientID1099') or '',
                'PayToAddress_Phone': first_phone(v),
                **addr,
                'MainContact_Prefix':    '',
                'MainContact_Suffix':    '',
                'MainContact_Title':     '',
                'MainContact_FirstName': '',
                'MainContact_LastName':  '',
                'MainContact_WorkPhone': first_phone(v),
                'MainContact_CellPhone': '',
                'MainContact_WorkEmail': v.get('Email') or '',
                'MainContact_HomeEmail': '',
                'EnableEFT':      '',
                'CompanyID':      v.get('CompanyID') or '',
                'CompanyName':    v.get('CompanyName') or '',
                'ABA/Routing':    v.get('ABARouting') or v.get('RoutingNumber') or '',
                'Account#':       v.get('AccountNum') or v.get('AccountNumber') or '',
                'Savings':        '',
                'EFType(SEC)':    '',
            })

            for contact in (v.get('Contacts') or []):
                first = contact.get('FirstName') or ''
                last  = contact.get('LastName')  or ''
                if not first and not last:
                    continue
                contact_rows.append({
                    'FirmCode':         firm_code,
                    'FirmRelationship': contact.get('ContactTypeDescription') or 'Primary',
                    'Prefix':           contact.get('Prefix') or '',
                    'Suffix':           contact.get('Suffix') or '',
                    'Title':            contact.get('Title') or '',
                    'FirstName':        first,
                    'LastName':         last,
                    'WorkPhone':        first_phone(contact),
                    'CellPhone':        contact.get('PrimaryPhone2') or contact.get('Mobile') or '',
                    'WorkEmail':        contact.get('Email') or '',
                    'HomeEmail':        '',
                    'WorkAddress1':     raw_addr['Street1'],
                    'WorkAddress2':     raw_addr['Street2'],
                    'WorkAddress3':     raw_addr['Street3'],
                    'WorkAddress4':     raw_addr['Street4'],
                    'WorkCity':         raw_addr['City'],
                    'WorkState':        raw_addr['State'],
                    'WorkZip':          raw_addr['Zip'],
                    'WorkCountry':      raw_addr['Country'],
                    'HomeAddress1': '', 'HomeAddress2': '', 'HomeAddress3': '', 'HomeAddress4': '',
                    'HomeCity': '', 'HomeState': '', 'HomeZip': '', 'HomeCountry': '',
                })

    vendor_fields = [
        'FirmCode', 'FirmName', 'IsActive', 'IsConsultant', 'ConsultantType', 'Is1099',
        'Website', 'VendorType', 'Note', 'NetDays', 'EIN',
        'PayToAddress_Phone',
        'PayToAddress_Street1', 'PayToAddress_Street2', 'PayToAddress_Street3', 'PayToAddress_Street4',
        'PayToAddress_City', 'PayToAddress_State', 'PayToAddress_Zip', 'PayToAddress_Country',
        'MainContact_Prefix', 'MainContact_Suffix', 'MainContact_Title',
        'MainContact_FirstName', 'MainContact_LastName',
        'MainContact_WorkPhone', 'MainContact_CellPhone', 'MainContact_WorkEmail',
        'MainContact_HomeEmail', 'MainContact_HomeEmail',
        'EnableEFT', 'CompanyID', 'CompanyName',
        'ABA/Routing', 'Account#', 'Savings', 'EFType(SEC)',
    ]
    write_csv(OUTPUT_DIR / 'cincinnati_Vendors.csv', vendor_rows, vendor_fields)
    write_csv(OUTPUT_DIR / 'cincinnati_VendorContacts.csv', contact_rows)
    print(f'  -> {len(vendor_rows)} vendors, {len(contact_rows)} contacts')


# ---------------------------------------------------------------------------
# Employees (pay rates expanded — one row per pay rate period)
# ---------------------------------------------------------------------------

def extract_employees(api_url: str, token: str):
    print('Extracting Employees...')

    content, _ = call(api_url, token, 'ListEmployees')
    all_emps = content.get('Employees', [])
    keys = [e['EmployeeKey'] for e in all_emps if e.get('EmployeeKey')]

    rows = []
    emp_seq = 0

    for batch in batched(keys, BATCH_SIZE):
        content, _ = call(api_url, token, 'GetEmployees', {'RequestedEmployees': batch})
        for e in content.get('Employees', []):
            parts = [e.get('FirstName') or '', e.get('MiddleName') or '', e.get('LastName') or '']
            full_name = ' '.join(p for p in parts if p).strip()
            emp_seq += 1
            emp_code  = f'CIN-EMP-{emp_seq:05d}'

            pay_rates = e.get('PayRates') or []
            # Sort pay rates by start date so history is in order
            pay_rates.sort(key=lambda r: r.get('StartDate') or '')

            if pay_rates:
                for i, pr in enumerate(pay_rates):
                    # End date = start date of the next rate (blank for the current/last one)
                    next_start = pay_rates[i + 1].get('StartDate') if i + 1 < len(pay_rates) else ''
                    rows.append({
                        'EmployeeCode':       emp_code,
                        'EmployeeName':       full_name,
                        'PayRate':            pr.get('PayRate') or '',
                        'salaryperpayperiod': pr.get('Salary') or '',
                        'PayRateStartDate':   pr.get('StartDate') or '',
                        'PayRateEndDate':     next_start,
                        'IsHourly':           'TRUE' if pr.get('IsHourly') else 'FALSE',
                        'OTRate':             '',
                        'OTMU':               '',
                        '_Status':            is_active(e.get('Status')),
                        '_HireDate':          e.get('DateHired') or '',
                        '_TermDate':          e.get('DateTerminated') or '',
                        '_Title':             e.get('Title') or '',
                        '_Email':             e.get('Email') or '',
                    })
            else:
                rows.append({
                    'EmployeeCode':       emp_code,
                    'EmployeeName':       full_name,
                    'PayRate':            '',
                    'salaryperpayperiod': '',
                    'PayRateStartDate':   '',
                    'PayRateEndDate':     '',
                    'IsHourly':           '',
                    'OTRate':             '',
                    'OTMU':               '',
                    '_Status':            is_active(e.get('Status')),
                    '_HireDate':          e.get('DateHired') or '',
                    '_TermDate':          e.get('DateTerminated') or '',
                    '_Title':             e.get('Title') or '',
                    '_Email':             e.get('Email') or '',
                })

    write_csv(OUTPUT_DIR / 'cincinnati_Employees.csv', rows)
    unique = len({r['EmployeeCode'] for r in rows})
    print(f'  -> {unique} employees, {len(rows)} pay rate rows')


# ---------------------------------------------------------------------------
# Expense Codes — probe several possible API method names
# ---------------------------------------------------------------------------

def extract_expense_codes(api_url: str, token: str):
    # Ajera calls these "Activities" — ListActivities returns all expense/time codes
    print('Extracting Expense Codes (via ListActivities)...')

    content, errors = call(api_url, token, 'ListActivities')
    activities = content.get('Activities', [])

    if not activities:
        print(f'  ! ListActivities returned no records. Errors: {errors}')
        return

    rows = []
    exp_seq = 0
    for a in sorted(activities, key=lambda x: x.get('ActivityKey', 0)):
        is_active = 'TRUE' if str(a.get('Status', '')).lower() == 'active' else 'FALSE'
        is_unit   = 'TRUE' if a.get('UnitBased') else 'FALSE'
        exp_seq += 1
        rows.append({
            'ECCode':                   f'CIN-EXP-{exp_seq:05d}',
            'ECName':                   a.get('Description', ''),
            'ShowInES':                 is_active,
            'IsUnit':                   is_unit,
            'UnitTypename':             a.get('UnitDescription', ''),
            'ECTypename':               '',
            'ExpMarkupTypename':        '',
            'Markup':                   '',
            'UnitRate':                 a.get('UnitCostRate', ''),
            'BillStatusname':           '',
            'DirectBaseCode':           '',
            'DirectBasename':           '',
            'OHBaseCode':               '',
            'OHBasename':               '',
            'BilledDirectBaseCode':     '',
            'BilledDirectBasename':     '',
            'BilledMarkupBaseCode':     '',
            'BilledMarkupBasename':     '',
            'UnBilledBaseCode':         '',
            'UnBilledBasename':         '',
            'CurrencyCode':             '',
            'PMCmtRequired':            '',
            'IntCmtRequired':           '',
            'IsNonReim':                '',
            '_IsActive':                is_active,
        })

    write_csv(OUTPUT_DIR / 'cincinnati_ExpenseCodes.csv', rows)
    active = sum(1 for r in rows if r['_IsActive'] == 'TRUE')
    print(f'  -> {len(rows)} expense codes ({active} active, {len(rows)-active} inactive)')


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    token, api_url = create_session()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f'Output: {OUTPUT_DIR}\n')

    try:
        extract_coa(api_url, token)
        extract_clients(api_url, token)
        extract_vendors(api_url, token)
        extract_employees(api_url, token)
        extract_expense_codes(api_url, token)
    finally:
        end_session(api_url, token)
        print('\nSession closed.')

    print(f'\nDone. Review CSVs in {OUTPUT_DIR} before loading into Unanet templates.')
    print('Remove _ prefixed columns before loading.')


if __name__ == '__main__':
    main()
