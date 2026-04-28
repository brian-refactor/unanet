"""
Extracts all Unanet-relevant data from QuickBooks Online (Minnesota office)
and writes CSV files to output/minnesota/ for review before loading into templates.

Usage:
    python qbo_extract.py

Output files:
    output/minnesota/minnesota_COA.csv
    output/minnesota/minnesota_Clients.csv
    output/minnesota/minnesota_ClientContacts.csv
    output/minnesota/minnesota_Vendors.csv
    output/minnesota/minnesota_VendorContacts.csv
    output/minnesota/minnesota_Employees.csv   (pay rates blank — see note below)
    output/minnesota/minnesota_ExpenseCodes.csv

NOTE on pay rates: QBO's standard API does not expose payroll rates.
If Minnesota uses QBO Payroll, we can add a second pass using the Payroll API.
Otherwise, pay rates will need to be added manually to the Employees CSV.

Columns prefixed with _ are for reference only — remove before loading into Unanet.
"""

import csv
import json
import os
import re
from pathlib import Path

from dotenv import load_dotenv
from intuitlib.client import AuthClient
from quickbooks import QuickBooks
from quickbooks.objects.account import Account
from quickbooks.objects.customer import Customer
from quickbooks.objects.employee import Employee
from quickbooks.objects.item import Item
from quickbooks.objects.term import Term
from quickbooks.objects.vendor import Vendor

HERE = Path(__file__).parent
TOKEN_FILE = HERE / 'qbo_tokens.json'
OUTPUT_DIR = HERE.parent / 'output' / 'minnesota'
OFFICE_PREFIX = 'MN'

# QBO AccountType (as returned by the API, with spaces) -> Unanet FinancialType
FINANCIAL_TYPE_MAP = {
    'Bank': 'Asset',
    'Accounts Receivable': 'Asset',
    'Other Current Asset': 'Asset',
    'Fixed Asset': 'Asset',
    'Other Asset': 'Asset',
    'Accounts Payable': 'Liability',
    'Credit Card': 'Liability',
    'Other Current Liability': 'Liability',
    'Long Term Liability': 'Liability',
    'Equity': 'Equity',
    'Income': 'Revenue',
    'Cost of Goods Sold': 'Expense',
    'Expense': 'Expense',
    'Other Expense': 'Expense',
    # NonPosting = excluded (memo accounts, etc.)
}

SUBLEDGER_TYPE_MAP = {
    'Accounts Receivable': 'AR',
    'Accounts Payable': 'AP',
    'Bank': 'Bank',
}


# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------

def load_tokens() -> dict:
    if not TOKEN_FILE.exists():
        raise SystemExit(f'No tokens found at {TOKEN_FILE}. Run qbo_auth.py first.')
    return json.loads(TOKEN_FILE.read_text())


def save_tokens(auth_client: AuthClient, realm_id: str):
    TOKEN_FILE.write_text(json.dumps({
        'access_token': auth_client.access_token,
        'refresh_token': auth_client.refresh_token,
        'realm_id': realm_id,
    }, indent=2))


def get_client() -> tuple[QuickBooks, AuthClient, str]:
    load_dotenv(HERE / '.env')
    tokens = load_tokens()

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
    save_tokens(auth_client, tokens['realm_id'])

    qb = QuickBooks(
        auth_client=auth_client,
        refresh_token=auth_client.refresh_token,
        company_id=tokens['realm_id'],
    )

    return qb, auth_client, tokens['realm_id']


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def make_firm_code(name: str, seen: set) -> str:
    """Derive a short unique FirmCode from a company or display name."""
    clean = re.sub(r'[^A-Z0-9]', '', (name or '').upper())
    base = f'{OFFICE_PREFIX}-{clean[:8]}' if clean else f'{OFFICE_PREFIX}-UNK'
    code, n = base, 2
    while code in seen:
        code = f'{base}{n}'
        n += 1
    seen.add(code)
    return code


def addr_fields(addr, prefix: str = '', line_field: str = 'Street') -> dict:
    """Flatten a QBO address object into prefixed dict keys."""
    if not addr:
        return {f'{prefix}{k}': '' for k in [f'{line_field}1', f'{line_field}2', 'City', 'State', 'Zip', 'Country']}
    return {
        f'{prefix}{line_field}1': getattr(addr, 'Line1', '') or '',
        f'{prefix}{line_field}2': getattr(addr, 'Line2', '') or '',
        f'{prefix}City': getattr(addr, 'City', '') or '',
        f'{prefix}State': getattr(addr, 'CountrySubDivisionCode', '') or '',
        f'{prefix}Zip': getattr(addr, 'PostalCode', '') or '',
        f'{prefix}Country': getattr(addr, 'Country', '') or '',
    }


def phone(obj, attr: str = 'PrimaryPhone') -> str:
    return getattr(getattr(obj, attr, None), 'FreeFormNumber', '') or ''


def email(obj, attr: str = 'PrimaryEmailAddr') -> str:
    return getattr(getattr(obj, attr, None), 'Address', '') or ''


def query_all(entity_class, qb: QuickBooks) -> list:
    """Paginate through all QBO records for an entity (max 1000 per request)."""
    results = []
    page_size = 1000
    start = 1
    while True:
        batch = entity_class.query(
            f"SELECT * FROM {entity_class.qbo_object_name} STARTPOSITION {start} MAXRESULTS {page_size}",
            qb=qb,
        )
        results.extend(batch)
        if len(batch) < page_size:
            break
        start += page_size
    return results


def build_term_lookup(qb: QuickBooks) -> dict:
    """Return {term_id: due_days} for all payment terms."""
    return {t.Id: getattr(t, 'DueDays', 30) or 30 for t in query_all(Term, qb)}


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
# Extractors
# ---------------------------------------------------------------------------

def extract_coa(qb: QuickBooks):
    print('Extracting Chart of Accounts...')
    rows = []

    for acct in query_all(Account, qb):
        acct_type = getattr(acct, 'AccountType', '') or ''
        financial_type = FINANCIAL_TYPE_MAP.get(acct_type)
        if not financial_type:
            continue  # skip NonPosting, etc.

        sub_type = getattr(acct, 'AccountSubType', '') or ''
        acct_num = getattr(acct, 'AcctNum', '') or ''

        rows.append({
            'BaseCode': acct_num if acct_num else f'QBO-{acct.Id}',
            'BaseName': acct.Name or '',
            'Description': getattr(acct, 'Description', '') or '',
            'IsActive': 'TRUE' if getattr(acct, 'Active', True) else 'FALSE',
            'Is1099': '',
            'IsSubcontractor': '',
            'FinancialType': financial_type,
            'SubledgerType': SUBLEDGER_TYPE_MAP.get(sub_type, ''),
            'MetricType': '',
            'CostType': '',
            'PMType': '',
            'LaborRevenueType': '',
            'ExpenseRevenueType': '',
            '_QBO_AccountType': acct_type,
            '_QBO_SubType': sub_type,
        })

    write_csv(OUTPUT_DIR / 'minnesota_COA.csv', rows)
    print(f'  -> {len(rows)} accounts')

    flagged = [r for r in rows if r['BaseCode'].startswith('QBO-')]
    if flagged:
        print(f'  ! {len(flagged)} accounts have no account number in QBO — BaseCode auto-assigned (review required)')


def extract_clients(qb: QuickBooks, term_lookup: dict):
    print('Extracting Customers...')
    client_rows, contact_rows = [], []
    cli_seq = 0

    for c in query_all(Customer, qb):
        if getattr(c, 'Job', False):
            continue  # skip sub-customers / jobs

        firm_name = getattr(c, 'CompanyName', '') or c.DisplayName or ''
        cli_seq += 1
        firm_code = f'MN-CLI-{cli_seq:05d}'

        term_id = getattr(getattr(c, 'SalesTermRef', None), 'value', None)
        bill_addr = getattr(c, 'BillAddr', None)
        bill = addr_fields(bill_addr, 'BillToAddress_')

        client_rows.append({
            'FirmCode': firm_code,
            'FirmName': firm_name,
            'IsActive': 'TRUE' if getattr(c, 'Active', True) else 'FALSE',
            'Website': getattr(getattr(c, 'WebAddr', None), 'URI', '') or '',
            'ClientType': '',
            'Specialty': '',
            'Note': getattr(c, 'Notes', '') or '',
            'PayDays': term_lookup.get(term_id, 30),
            'MainEmail': email(c),
            'BillToAddress_Phone': phone(c),
            **bill,
            'BillToAddress_Street3': '',
            'BillToAddress_Street4': '',
            'MainContact_Prefix': getattr(c, 'Title', '') or '',
            'MainContact_Suffix': getattr(c, 'Suffix', '') or '',
            'MainContact_Title': '',
            'MainContact_FirstName': getattr(c, 'GivenName', '') or '',
            'MainContact_LastName': getattr(c, 'FamilyName', '') or '',
            'MainContact_WorkPhone': phone(c),
            'MainContact_CellPhone': phone(c, 'Mobile'),
            'MainContact_WorkEmail': email(c),
            'MainContact_HomeEmail': '',
        })

        first = getattr(c, 'GivenName', '') or ''
        last = getattr(c, 'FamilyName', '') or ''
        if first or last:
            work = addr_fields(bill_addr, 'Work', 'Address')
            contact_rows.append({
                'FirmCode': firm_code,
                'FirmRelationship': 'Primary',
                'Prefix': getattr(c, 'Title', '') or '',
                'Suffix': getattr(c, 'Suffix', '') or '',
                'Title': '',
                'FirstName': first,
                'LastName': last,
                'WorkPhone': phone(c),
                'CellPhone': phone(c, 'Mobile'),
                'WorkEmail': email(c),
                'HomeEmail': '',
                **work,
                'WorkAddress3': '',
                'WorkAddress4': '',
                'HomeAddress1': '', 'HomeAddress2': '', 'HomeAddress3': '', 'HomeAddress4': '',
                'HomeCity': '', 'HomeState': '', 'HomeZip': '', 'HomeCountry': '',
            })

    write_csv(OUTPUT_DIR / 'minnesota_Clients.csv', client_rows)
    write_csv(OUTPUT_DIR / 'minnesota_ClientContacts.csv', contact_rows)
    print(f'  -> {len(client_rows)} clients, {len(contact_rows)} contacts')


def extract_vendors(qb: QuickBooks, term_lookup: dict):
    print('Extracting Vendors...')
    vendor_rows, contact_rows = [], []
    vnd_seq = 0

    for v in query_all(Vendor, qb):
        firm_name = getattr(v, 'CompanyName', '') or v.DisplayName or ''
        vnd_seq += 1
        firm_code = f'MN-VND-{vnd_seq:05d}'

        term_id = getattr(getattr(v, 'TermRef', None), 'value', None)
        bill_addr = getattr(v, 'BillAddr', None)
        bill = addr_fields(bill_addr, 'PayToAddress_')

        vendor_rows.append({
            'FirmCode': firm_code,
            'FirmName': firm_name,
            'IsActive': 'TRUE' if getattr(v, 'Active', True) else 'FALSE',
            'IsConsultant': '',
            'ConsultantType': '',
            'Is1099': 'TRUE' if getattr(v, 'Vendor1099', False) else 'FALSE',
            'Website': getattr(getattr(v, 'WebAddr', None), 'URI', '') or '',
            'VendorType': '',
            'Note': '',
            'NetDays': term_lookup.get(term_id, 30),
            'EIN': getattr(v, 'TaxIdentifier', '') or '',
            'PayToAddress_Phone': phone(v),
            **bill,
            'PayToAddress_Street3': '',
            'PayToAddress_Street4': '',
            'MainContact_Prefix': getattr(v, 'Title', '') or '',
            'MainContact_Suffix': getattr(v, 'Suffix', '') or '',
            'MainContact_Title': '',
            'MainContact_FirstName': getattr(v, 'GivenName', '') or '',
            'MainContact_LastName': getattr(v, 'FamilyName', '') or '',
            'MainContact_WorkPhone': phone(v),
            'MainContact_CellPhone': phone(v, 'Mobile'),
            'MainContact_WorkEmail': email(v),
            'MainContact_HomeEmail': '',
            'EnableEFT': '',
            'CompanyID': '',
            'CompanyName': '',
            'ABA/Routing': '',
            'Account#': getattr(v, 'AcctNum', '') or '',
            'Savings': '',
            'EFType(SEC)': '',
        })

        first = getattr(v, 'GivenName', '') or ''
        last = getattr(v, 'FamilyName', '') or ''
        if first or last:
            work = addr_fields(bill_addr, 'Work', 'Address')
            contact_rows.append({
                'FirmCode': firm_code,
                'FirmRelationship': 'Primary',
                'Prefix': getattr(v, 'Title', '') or '',
                'Suffix': getattr(v, 'Suffix', '') or '',
                'Title': '',
                'FirstName': first,
                'LastName': last,
                'WorkPhone': phone(v),
                'CellPhone': phone(v, 'Mobile'),
                'WorkEmail': email(v),
                'HomeEmail': '',
                **work,
                'WorkAddress3': '',
                'WorkAddress4': '',
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
    write_csv(OUTPUT_DIR / 'minnesota_Vendors.csv', vendor_rows, vendor_fields)
    write_csv(OUTPUT_DIR / 'minnesota_VendorContacts.csv', contact_rows)
    print(f'  -> {len(vendor_rows)} vendors, {len(contact_rows)} contacts')


def extract_employees(qb: QuickBooks):
    print('Extracting Employees...')
    rows = []
    emp_seq = 0

    for e in query_all(Employee, qb):
        emp_seq += 1
        rows.append({
            'EmployeeCode': f'MN-EMP-{emp_seq:05d}',
            'EmployeeName': e.DisplayName or '',
            'PayRate': '',
            'salaryperpayperiod': '',
            'PayRateStartDate': '',
            'PayRateEndDate': '',
            'IsHourly': '',
            'OTRate': '',
            'OTMU': '',
            '_FirstName': getattr(e, 'GivenName', '') or '',
            '_LastName': getattr(e, 'FamilyName', '') or '',
            '_IsActive': 'TRUE' if getattr(e, 'Active', True) else 'FALSE',
            '_HireDate': getattr(e, 'HiredDate', '') or '',
            '_ReleaseDate': getattr(e, 'ReleasedDate', '') or '',
            '_WorkPhone': phone(e),
            '_Mobile': phone(e, 'Mobile'),
            '_Email': email(e),
        })

    write_csv(OUTPUT_DIR / 'minnesota_Employees.csv', rows)
    print(f'  -> {len(rows)} employees')
    print('  ! Pay rates not available via standard QBO API — fill in manually or request Payroll API access')


def extract_items(qb: QuickBooks):
    print('Extracting Items / Expense Codes...')
    rows = []
    exp_seq = 0

    RELEVANT_TYPES = {'Service', 'NonInventory', 'OtherCharge'}

    for item in query_all(Item, qb):
        item_type = getattr(item, 'Type', '') or ''
        if item_type not in RELEVANT_TYPES:
            continue

        income_acct = getattr(getattr(item, 'IncomeAccountRef', None), 'name', '') or ''
        expense_acct = getattr(getattr(item, 'ExpenseAccountRef', None), 'name', '') or ''

        exp_seq += 1
        rows.append({
            'ECCode': f'MN-EXP-{exp_seq:05d}',
            'ECName': item.Name or '',
            'ShowInES': 'TRUE',
            'IsUnit': 'FALSE',
            'UnitTypename': '',
            'ECTypename': 'Other Direct Charge' if item_type == 'OtherCharge' else '',
            'ExpMarkupTypename': '',
            'Markup': '',
            'UnitRate': getattr(item, 'UnitPrice', '') or '',
            'BillStatusname': '',
            'DirectBaseCode': expense_acct,
            'DirectBasename': '',
            'OHBaseCode': '',
            'OHBasename': '',
            'BilledDirectBaseCode': income_acct,
            'BilledDirectBasename': '',
            'BilledMarkupBaseCode': '',
            'BilledMarkupBasename': '',
            'UnBilledBaseCode': '',
            'UnBilledBasename': '',
            'CurrencyCode': '',
            'PMCmtRequired': '',
            'IntCmtRequired': '',
            'IsNonReim': '',
            '_IsActive': 'TRUE' if getattr(item, 'Active', True) else 'FALSE',
            '_QBO_Type': item_type,
        })

    write_csv(OUTPUT_DIR / 'minnesota_ExpenseCodes.csv', rows)
    print(f'  -> {len(rows)} items')


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    qb, auth_client, realm_id = get_client()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f'\nConnected — Realm ID: {realm_id}')
    print(f'Output: {OUTPUT_DIR}\n')

    term_lookup = build_term_lookup(qb)

    extract_coa(qb)
    extract_clients(qb, term_lookup)
    extract_vendors(qb, term_lookup)
    extract_employees(qb)
    extract_items(qb)

    save_tokens(auth_client, realm_id)  # persist refreshed tokens

    print(f'\nDone. Review CSVs in {OUTPUT_DIR} before loading into Unanet templates.')
    print('Remove any _ prefixed columns before loading.')


if __name__ == '__main__':
    main()
