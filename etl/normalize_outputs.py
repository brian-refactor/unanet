"""
Normalize all output CSVs to exactly match Unanet upload template column names.
Run after all extractions are complete.

Usage:
    python etl/normalize_outputs.py
"""

import csv
from pathlib import Path

OUTPUT = Path(__file__).parent.parent / 'output'


def rewrite(path: Path, renames: dict, template_cols: list, drop: list = None):
    """
    Read a CSV, apply renames, add any missing template cols as blank,
    prefix any leftover non-template cols with _, write back in template column order.
    """
    if not path.exists():
        return
    drop = set(drop or [])
    with open(path, encoding='utf-8') as f:
        rows = list(csv.DictReader(f))
    if not rows:
        return

    new_rows = []
    for r in rows:
        nr = {}
        for k, v in r.items():
            if k in drop:
                continue
            new_key = renames.get(k, k)
            nr[new_key] = v
        # Fill missing template columns with blank
        for col in template_cols:
            if col not in nr:
                nr[col] = ''
        new_rows.append(nr)

    # Build final fieldnames: template cols in order, then any leftover _ cols
    leftover = [k for k in new_rows[0] if k not in template_cols and not k.startswith('_')]
    prefixed = [f'_{k}' for k in leftover]
    # Rename leftover to _ prefix in rows
    for nr in new_rows:
        for k in leftover:
            nr[f'_{k}'] = nr.pop(k)

    fieldnames = template_cols + [f'_{k}' for k in leftover]

    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
        w.writeheader()
        w.writerows(new_rows)

    print(f'  {path.parent.name}/{path.name}')


# ---------------------------------------------------------------------------
# Template column definitions (row 3 of each Unanet template)
# ---------------------------------------------------------------------------

COA_COLS = [
    'BaseCode', 'BaseName', 'Description', 'IsActive', 'Is1099', 'IsSubcontractor',
    'FinancialType', 'SubledgerType', 'MetricType', 'CostType', 'PMType',
    'LaborRevenueType', 'ExpenseRevenueType',
]

CLIENT_COLS = [
    'FirmCode', 'FirmName', 'IsActive', 'Website', 'ClientType', 'Specialty', 'Note',
    'PayDays', 'MainEmail', 'BillToAddress_Phone',
    'BillToAddress_Street1', 'BillToAddress_Street2', 'BillToAddress_Street3', 'BillToAddress_Street4',
    'BillToAddress_City', 'BillToAddress_State', 'BillToAddress_Zip', 'BillToAddress_Country',
    'MainContact_Prefix', 'MainContact_Suffix', 'MainContact_Title',
    'MainContact_FirstName', 'MainContact_LastName',
    'MainContact_WorkPhone', 'MainContact_CellPhone',
    'MainContact_WorkEmail', 'MainContact_HomeEmail',
]

CONTACT_COLS = [
    'FirmCode', 'FirmRelationship', 'Prefix', 'Suffix', 'Title',
    'FirstName', 'LastName', 'WorkPhone', 'CellPhone', 'WorkEmail', 'HomeEmail',
    'WorkAddress1', 'WorkAddress2', 'WorkAddress3', 'WorkAddress4',
    'WorkCity', 'WorkState', 'WorkZip', 'WorkCountry',
    'HomeAddress1', 'HomeAddress2', 'HomeAddress3', 'HomeAddress4',
    'HomeCity', 'HomeState', 'HomeZip', 'HomeCountry',
]

VENDOR_COLS = [
    'FirmCode', 'FirmName', 'IsActive', 'IsConsultant', 'ConsultantType', 'Is1099',
    'Website', 'VendorType', 'Note', 'NetDays', 'EIN',
    'PayToAddress_Phone',
    'PayToAddress_Street1', 'PayToAddress_Street2', 'PayToAddress_Street3', 'PayToAddress_Street4',
    'PayToAddress_City', 'PayToAddress_State', 'PayToAddress_Zip', 'PayToAddress_Country',
    'MainContact_Prefix', 'MainContact_Suffix', 'MainContact_Title',
    'MainContact_FirstName', 'MainContact_LastName',
    'MainContact_WorkPhone', 'MainContact_CellPhone',
    'MainContact_WorkEmail', 'MainContact_HomeEmail',
    'EnableEFT', 'CompanyID', 'CompanyName',
    'ABA/Routing', 'Account#', 'Savings', 'EFType(SEC)',
]

PAYHISTORY_COLS = [
    'EmployeeCode', 'EmployeeName', 'PayRate', 'salaryperpayperiod',
    'PayRateStartDate', 'PayRateEndDate', 'IsHourly', 'OTRate', 'OTMU',
]

EXPENSECODE_COLS = [
    'ECCode', 'ECName', 'ShowInES', 'IsUnit', 'UnitTypename', 'ECTypename',
    'ExpMarkupTypename', 'Markup', 'UnitRate', 'BillStatusname',
    'DirectBaseCode', 'DirectBasename', 'OHBaseCode', 'OHBasename',
    'BilledDirectBaseCode', 'BilledDirectBasename',
    'BilledMarkupBaseCode', 'BilledMarkupBasename',
    'UnBilledBaseCode', 'UnBilledBasename',
    'CurrencyCode', 'PMCmtRequired', 'IntCmtRequired', 'IsNonReim',
]

# ---------------------------------------------------------------------------
# Column renames per file type
# ---------------------------------------------------------------------------

# Clients — embedded main contact fields use MainContact_ prefix in template
CLIENT_RENAMES = {
    'Country':   'BillToAddress_Country',
    'Prefix':    'MainContact_Prefix',
    'Suffix':    'MainContact_Suffix',
    'Title':     'MainContact_Title',
    'FirstName': 'MainContact_FirstName',
    'LastName':  'MainContact_LastName',
    'WorkPhone': 'MainContact_WorkPhone',
    'CellPhone': 'MainContact_CellPhone',
    'WorkEmail': 'MainContact_WorkEmail',
    'HomeEmail': 'MainContact_HomeEmail',
}
CLIENT_DROP = ['Fax']   # not in template

# Client/Vendor contacts — Street → Address
CONTACT_RENAMES = {
    'WorkStreet1': 'WorkAddress1', 'WorkStreet2': 'WorkAddress2',
    'WorkStreet3': 'WorkAddress3', 'WorkStreet4': 'WorkAddress4',
    'HomeStreet1': 'HomeAddress1', 'HomeStreet2': 'HomeAddress2',
    'HomeStreet3': 'HomeAddress3', 'HomeStreet4': 'HomeAddress4',
    # QB Desktop contact files use Phone/Email without Work prefix
    'Phone': 'WorkPhone',
    'Email': 'WorkEmail',
}

VENDOR_RENAMES = {
    'Country':    'PayToAddress_Country',
    'ABARouting': 'ABA/Routing',
    'AccountNum': 'Account#',
    'EFType':     'EFType(SEC)',
    'Prefix':     'MainContact_Prefix',
    'Suffix':     'MainContact_Suffix',
    'Title':      'MainContact_Title',
    'FirstName':  'MainContact_FirstName',
    'LastName':   'MainContact_LastName',
    'WorkPhone':  'MainContact_WorkPhone',
    'CellPhone':  'MainContact_CellPhone',
    'WorkEmail':  'MainContact_WorkEmail',
    'HomeEmail':  'MainContact_HomeEmail',
}
VENDOR_DROP = ['Fax', 'MainEmail', 'PayToAddress_AltPhone']


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    offices = ['minnesota', 'cincinnati', 'dallas', 'orlando']

    print('Normalizing COA...')
    for o in offices:
        rewrite(OUTPUT / o / f'{o}_COA.csv', {}, COA_COLS)

    print('Normalizing Clients...')
    for o in offices:
        rewrite(OUTPUT / o / f'{o}_Clients.csv', CLIENT_RENAMES, CLIENT_COLS, CLIENT_DROP)

    print('Normalizing ClientContacts...')
    for o in offices:
        rewrite(OUTPUT / o / f'{o}_ClientContacts.csv', CONTACT_RENAMES, CONTACT_COLS)

    print('Normalizing Vendors...')
    for o in offices:
        rewrite(OUTPUT / o / f'{o}_Vendors.csv', VENDOR_RENAMES, VENDOR_COLS, VENDOR_DROP)

    print('Normalizing VendorContacts...')
    for o in offices:
        rewrite(OUTPUT / o / f'{o}_VendorContacts.csv', CONTACT_RENAMES, CONTACT_COLS)

    print('Normalizing PayHistory (Employees)...')
    for o in offices:
        rewrite(OUTPUT / o / f'{o}_Employees.csv', {}, PAYHISTORY_COLS)

    print('Normalizing ExpenseCodes...')
    for o in offices:
        rewrite(OUTPUT / o / f'{o}_ExpenseCodes.csv', {}, EXPENSECODE_COLS)

    print('\nDone.')


if __name__ == '__main__':
    main()
