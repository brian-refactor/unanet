"""
QuickBooks Desktop Web Connector (QBWC) SOAP server.

QB Desktop acts as the HTTP client — it calls this server on a schedule,
sends QBXML responses, and receives QBXML queries in return.

Usage:
    python etl/qbd_server.py --office dallas
    python etl/qbd_server.py --office orlando

Then in QB Desktop Web Connector, load the matching .qwc file.
For remote machines, expose this server with:
    ngrok http 5150
and update the AppURL in the .qwc file with the ngrok https URL.

Outputs:
    output/dallas/   dallas_COA.csv, dallas_Clients.csv, etc.
    output/orlando/  orlando_COA.csv, orlando_Clients.csv, etc.
"""

import argparse
import csv
import os
import uuid
from collections import defaultdict
from pathlib import Path
from textwrap import dedent

from flask import Flask, request, Response

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

VALID_PASSWORD = 'FusionMigration2024!'

OFFICE_PREFIXES = {
    'dallas':  'DAL-',
    'orlando': 'ORL-',
}

# QB Desktop account type → Unanet FinancialType
FINANCIAL_TYPE_MAP = {
    'AccountsPayable':         'Liability',
    'AccountsReceivable':      'Asset',
    'Bank':                    'Asset',
    'CostOfGoodsSold':         'Expense',
    'CreditCard':              'Liability',
    'Equity':                  'Equity',
    'Expense':                 'Expense',
    'FixedAsset':              'Asset',
    'Income':                  'Revenue',
    'LongTermLiability':       'Liability',
    'NonPosting':              None,      # excluded
    'OtherAsset':              'Asset',
    'OtherCurrentAsset':       'Asset',
    'OtherCurrentLiability':   'Liability',
    'OtherExpense':            'Expense',
    'OtherIncome':             'Revenue',
}

# Item types to map as expense codes (exclude inventory/assembly/group)
EXPENSE_ITEM_TYPES = {'Service', 'OtherCharge', 'NonInventory'}

# Sequence of QBXML queries to run per session
QUERY_STEPS = [
    'accounts',
    'customers',
    'vendors',
    'employees',
    'items',
]

# Max records per QBXML page
MAX_RETURNED = 500

# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------

# sessions[ticket] = {
#   'office': str,
#   'step_index': int,        # index into QUERY_STEPS
#   'iterator_id': str|None,  # active QB iterator ID
#   'records': {entity: [rows]},
# }
sessions: dict[str, dict] = {}


# ---------------------------------------------------------------------------
# QBXML builders
# ---------------------------------------------------------------------------

def _xml_header():
    return '<?xml version="1.0" encoding="utf-8"?><?qbxml version="13.0"?>'


def account_query(iterator: str, iterator_id: str | None = None) -> str:
    iter_attr = f'iterator="{iterator}"'
    if iterator_id:
        iter_attr += f' iteratorID="{iterator_id}"'
    return dedent(f"""\
        {_xml_header()}
        <QBXML><QBXMLMsgsRq onError="stopOnError">
          <AccountQueryRq {iter_attr}>
            <MaxReturned>{MAX_RETURNED}</MaxReturned>
          </AccountQueryRq>
        </QBXMLMsgsRq></QBXML>""")


def customer_query(iterator: str, iterator_id: str | None = None) -> str:
    iter_attr = f'iterator="{iterator}"'
    if iterator_id:
        iter_attr += f' iteratorID="{iterator_id}"'
    return dedent(f"""\
        {_xml_header()}
        <QBXML><QBXMLMsgsRq onError="stopOnError">
          <CustomerQueryRq {iter_attr}>
            <MaxReturned>{MAX_RETURNED}</MaxReturned>
            <IncludeRetElement>ListID</IncludeRetElement>
            <IncludeRetElement>Name</IncludeRetElement>
            <IncludeRetElement>CompanyName</IncludeRetElement>
            <IncludeRetElement>IsActive</IncludeRetElement>
            <IncludeRetElement>Phone</IncludeRetElement>
            <IncludeRetElement>AltPhone</IncludeRetElement>
            <IncludeRetElement>Fax</IncludeRetElement>
            <IncludeRetElement>Email</IncludeRetElement>
            <IncludeRetElement>BillAddress</IncludeRetElement>
            <IncludeRetElement>Contact</IncludeRetElement>
            <IncludeRetElement>Contacts</IncludeRetElement>
            <IncludeRetElement>JobStatus</IncludeRetElement>
            <IncludeRetElement>ParentRef</IncludeRetElement>
          </CustomerQueryRq>
        </QBXMLMsgsRq></QBXML>""")


def vendor_query(iterator: str, iterator_id: str | None = None) -> str:
    iter_attr = f'iterator="{iterator}"'
    if iterator_id:
        iter_attr += f' iteratorID="{iterator_id}"'
    return dedent(f"""\
        {_xml_header()}
        <QBXML><QBXMLMsgsRq onError="stopOnError">
          <VendorQueryRq {iter_attr}>
            <MaxReturned>{MAX_RETURNED}</MaxReturned>
            <IncludeRetElement>ListID</IncludeRetElement>
            <IncludeRetElement>Name</IncludeRetElement>
            <IncludeRetElement>CompanyName</IncludeRetElement>
            <IncludeRetElement>IsActive</IncludeRetElement>
            <IncludeRetElement>Phone</IncludeRetElement>
            <IncludeRetElement>AltPhone</IncludeRetElement>
            <IncludeRetElement>Fax</IncludeRetElement>
            <IncludeRetElement>Email</IncludeRetElement>
            <IncludeRetElement>VendorAddress</IncludeRetElement>
            <IncludeRetElement>Contact</IncludeRetElement>
            <IncludeRetElement>Contacts</IncludeRetElement>
          </VendorQueryRq>
        </QBXMLMsgsRq></QBXML>""")


def employee_query(iterator: str, iterator_id: str | None = None) -> str:
    iter_attr = f'iterator="{iterator}"'
    if iterator_id:
        iter_attr += f' iteratorID="{iterator_id}"'
    return dedent(f"""\
        {_xml_header()}
        <QBXML><QBXMLMsgsRq onError="stopOnError">
          <EmployeeQueryRq {iter_attr}>
            <MaxReturned>{MAX_RETURNED}</MaxReturned>
            <IncludeRetElement>ListID</IncludeRetElement>
            <IncludeRetElement>Name</IncludeRetElement>
            <IncludeRetElement>IsActive</IncludeRetElement>
            <IncludeRetElement>FirstName</IncludeRetElement>
            <IncludeRetElement>LastName</IncludeRetElement>
            <IncludeRetElement>Phone</IncludeRetElement>
            <IncludeRetElement>Email</IncludeRetElement>
            <IncludeRetElement>EmployeePayrollInfo</IncludeRetElement>
          </EmployeeQueryRq>
        </QBXMLMsgsRq></QBXML>""")


def item_query(iterator: str, iterator_id: str | None = None) -> str:
    iter_attr = f'iterator="{iterator}"'
    if iterator_id:
        iter_attr += f' iteratorID="{iterator_id}"'
    return dedent(f"""\
        {_xml_header()}
        <QBXML><QBXMLMsgsRq onError="stopOnError">
          <ItemQueryRq {iter_attr}>
            <MaxReturned>{MAX_RETURNED}</MaxReturned>
          </ItemQueryRq>
        </QBXMLMsgsRq></QBXML>""")


# ---------------------------------------------------------------------------
# XML parsing helpers (no external XML library required for this simple shape)
# ---------------------------------------------------------------------------

import re


def _find_all(tag: str, xml: str) -> list[str]:
    """Return all inner text blocks between <tag>...</tag>."""
    return re.findall(rf'<{tag}[^>]*>(.*?)</{tag}>', xml, re.DOTALL)


def _find_one(tag: str, xml: str, default: str = '') -> str:
    hits = _find_all(tag, xml)
    return hits[0].strip() if hits else default


def _attr(attr: str, xml: str, default: str = '') -> str:
    m = re.search(rf'{attr}="([^"]*)"', xml)
    return m.group(1) if m else default


def bool_str(val: str) -> str:
    return 'TRUE' if val.strip().lower() in ('true', '1', 'yes') else 'FALSE'


# ---------------------------------------------------------------------------
# Response parsers
# ---------------------------------------------------------------------------

def parse_accounts(xml: str, prefix: str) -> tuple[list[dict], str | None, int]:
    """Returns (rows, iterator_id_or_None, remaining)."""
    rs_block = _find_one('AccountQueryRs', xml)
    iterator_id = _attr('iteratorID', rs_block) or None
    remaining = int(_attr('iteratorRemainingCount', rs_block) or '0')

    rows = []
    for acct in _find_all('AccountRet', xml):
        acct_type = _find_one('AccountType', acct)
        fin_type = FINANCIAL_TYPE_MAP.get(acct_type)
        if fin_type is None:
            continue  # exclude NonPosting etc.
        full_name = _find_one('FullName', acct)
        name = _find_one('Name', acct)
        number = _find_one('AccountNumber', acct) or full_name.replace(' ', '')
        rows.append({
            'BaseCode':      number,
            'BaseName':      name or full_name,
            'FinancialType': fin_type,
            'IsActive':      bool_str(_find_one('IsActive', acct, 'true')),
            '_QBD_FullName': full_name,
            '_QBD_AcctType': acct_type,
        })
    return rows, iterator_id, remaining


def parse_customers(xml: str, prefix: str) -> tuple[list[dict], list[dict], str | None, int]:
    """Returns (client_rows, contact_rows, iterator_id, remaining)."""
    rs_block = _find_one('CustomerQueryRs', xml)
    iterator_id = _attr('iteratorID', rs_block) or None
    remaining = int(_attr('iteratorRemainingCount', rs_block) or '0')

    clients, contacts = [], []
    for cust in _find_all('CustomerRet', xml):
        # Skip sub-jobs (ParentRef present means it's a child)
        if _find_one('ParentRef', cust):
            continue
        list_id = _find_one('ListID', cust)
        company = _find_one('CompanyName', cust) or _find_one('Name', cust)
        firm_code = f"{prefix}{list_id}"
        clients.append({
            'FirmCode':  firm_code,
            'FirmName':  company,
            'IsActive':  bool_str(_find_one('IsActive', cust, 'true')),
            'Phone':     _find_one('Phone', cust),
            'AltPhone':  _find_one('AltPhone', cust),
            'Fax':       _find_one('Fax', cust),
            'MainEmail': _find_one('Email', cust),
            'Address1':  _find_one('Addr1', cust),
            'Address2':  _find_one('Addr2', cust),
            'City':      _find_one('City', cust),
            'State':     _find_one('State', cust),
            'Zip':       _find_one('PostalCode', cust),
        })
        # Primary contact from Contact element
        primary = _find_one('Contact', cust)
        if primary:
            parts = primary.strip().split(None, 1)
            contacts.append({
                'FirmCode':  firm_code,
                'FirstName': parts[0] if parts else '',
                'LastName':  parts[1] if len(parts) > 1 else '',
                'Email':     _find_one('Email', cust),
                'Phone':     _find_one('Phone', cust),
            })
        # Additional contacts from Contacts list
        for c in _find_all('ContactRet', cust):
            first = _find_one('FirstName', c)
            last = _find_one('LastName', c)
            if not first and not last:
                continue
            contacts.append({
                'FirmCode':  firm_code,
                'FirstName': first,
                'LastName':  last,
                'Email':     _find_one('ContactEmailAddr', c) or _find_one('Email', c),
                'Phone':     _find_one('ContactPhone', c) or _find_one('Phone', c),
            })
    return clients, contacts, iterator_id, remaining


def parse_vendors(xml: str, prefix: str) -> tuple[list[dict], list[dict], str | None, int]:
    rs_block = _find_one('VendorQueryRs', xml)
    iterator_id = _attr('iteratorID', rs_block) or None
    remaining = int(_attr('iteratorRemainingCount', rs_block) or '0')

    vendors, contacts = [], []
    for vend in _find_all('VendorRet', xml):
        list_id = _find_one('ListID', vend)
        company = _find_one('CompanyName', vend) or _find_one('Name', vend)
        firm_code = f"{prefix}{list_id}"
        vendors.append({
            'FirmCode':  firm_code,
            'FirmName':  company,
            'IsActive':  bool_str(_find_one('IsActive', vend, 'true')),
            'Phone':     _find_one('Phone', vend),
            'AltPhone':  _find_one('AltPhone', vend),
            'Fax':       _find_one('Fax', vend),
            'MainEmail': _find_one('Email', vend),
            'Address1':  _find_one('Addr1', vend),
            'Address2':  _find_one('Addr2', vend),
            'City':      _find_one('City', vend),
            'State':     _find_one('State', vend),
            'Zip':       _find_one('PostalCode', vend),
        })
        primary = _find_one('Contact', vend)
        if primary:
            parts = primary.strip().split(None, 1)
            contacts.append({
                'FirmCode':  firm_code,
                'FirstName': parts[0] if parts else '',
                'LastName':  parts[1] if len(parts) > 1 else '',
                'Email':     _find_one('Email', vend),
                'Phone':     _find_one('Phone', vend),
            })
        for c in _find_all('ContactRet', vend):
            first = _find_one('FirstName', c)
            last = _find_one('LastName', c)
            if not first and not last:
                continue
            contacts.append({
                'FirmCode':  firm_code,
                'FirstName': first,
                'LastName':  last,
                'Email':     _find_one('ContactEmailAddr', c) or _find_one('Email', c),
                'Phone':     _find_one('ContactPhone', c) or _find_one('Phone', c),
            })
    return vendors, contacts, iterator_id, remaining


def parse_employees(xml: str, prefix: str) -> tuple[list[dict], str | None, int]:
    rs_block = _find_one('EmployeeQueryRs', xml)
    iterator_id = _attr('iteratorID', rs_block) or None
    remaining = int(_attr('iteratorRemainingCount', rs_block) or '0')

    rows = []
    for emp in _find_all('EmployeeRet', xml):
        list_id = _find_one('ListID', emp)
        first = _find_one('FirstName', emp)
        last = _find_one('LastName', emp)
        name = _find_one('Name', emp) or f"{first} {last}".strip()
        emp_code = f"{prefix}{list_id}"
        # Pay rate — grab first WageItem PayRate if present
        pay_rate = ''
        wage_block = _find_one('WageItem', emp)
        if wage_block:
            pay_rate = _find_one('PayRate', wage_block)
        rows.append({
            'EmployeeCode': emp_code,
            'EmployeeName': name,
            'FirstName':    first,
            'LastName':     last,
            'IsActive':     bool_str(_find_one('IsActive', emp, 'true')),
            'Phone':        _find_one('Phone', emp),
            'Email':        _find_one('Email', emp),
            'PayRate':      pay_rate,
        })
    return rows, iterator_id, remaining


def parse_items(xml: str, prefix: str) -> tuple[list[dict], str | None, int]:
    rs_block = _find_one('ItemQueryRs', xml)
    iterator_id = _attr('iteratorID', rs_block) or None
    remaining = int(_attr('iteratorRemainingCount', rs_block) or '0')

    rows = []
    # Items come back as ItemServiceRet, ItemOtherChargeRet, ItemNonInventoryRet, etc.
    for tag in ('ItemServiceRet', 'ItemOtherChargeRet', 'ItemNonInventoryRet'):
        for item in _find_all(tag, xml):
            list_id = _find_one('ListID', item)
            name = _find_one('Name', item)
            full_name = _find_one('FullName', item) or name
            rows.append({
                'ECCode':     f"{prefix}{list_id}",
                'ECName':     name or full_name,
                'ShowInES':   'TRUE',
                'IsActive':   bool_str(_find_one('IsActive', item, 'true')),
                '_QBD_Type':  tag.replace('Ret', '').replace('Item', ''),
            })
    return rows, iterator_id, remaining


# ---------------------------------------------------------------------------
# CSV writers
# ---------------------------------------------------------------------------

def write_csv(path: Path, rows: list[dict], fieldnames: list[str] | None = None):
    if not rows:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = fieldnames or list(rows[0].keys())
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction='ignore')
        w.writeheader()
        w.writerows(rows)
    print(f'  Wrote {len(rows)} rows → {path}')


def flush_session(ticket: str, output_dir: Path):
    sess = sessions[ticket]
    office = sess['office']
    prefix = OFFICE_PREFIXES[office]
    recs = sess['records']

    print(f'\n[{office.upper()}] Flushing session data to {output_dir}')

    # COA
    write_csv(
        output_dir / f'{office}_COA.csv',
        recs.get('accounts', []),
        ['BaseCode', 'BaseName', 'FinancialType', 'IsActive', '_QBD_FullName', '_QBD_AcctType'],
    )
    # Clients
    write_csv(
        output_dir / f'{office}_Clients.csv',
        recs.get('clients', []),
        ['FirmCode', 'FirmName', 'IsActive', 'Phone', 'AltPhone', 'Fax', 'MainEmail',
         'Address1', 'Address2', 'City', 'State', 'Zip'],
    )
    # ClientContacts
    write_csv(
        output_dir / f'{office}_ClientContacts.csv',
        recs.get('client_contacts', []),
        ['FirmCode', 'FirstName', 'LastName', 'Email', 'Phone'],
    )
    # Vendors
    write_csv(
        output_dir / f'{office}_Vendors.csv',
        recs.get('vendors', []),
        ['FirmCode', 'FirmName', 'IsActive', 'Phone', 'AltPhone', 'Fax', 'MainEmail',
         'Address1', 'Address2', 'City', 'State', 'Zip'],
    )
    # VendorContacts
    write_csv(
        output_dir / f'{office}_VendorContacts.csv',
        recs.get('vendor_contacts', []),
        ['FirmCode', 'FirstName', 'LastName', 'Email', 'Phone'],
    )
    # Employees
    write_csv(
        output_dir / f'{office}_Employees.csv',
        recs.get('employees', []),
        ['EmployeeCode', 'EmployeeName', 'FirstName', 'LastName', 'IsActive', 'Phone', 'Email', 'PayRate'],
    )
    # ExpenseCodes
    write_csv(
        output_dir / f'{office}_ExpenseCodes.csv',
        recs.get('items', []),
        ['ECCode', 'ECName', 'ShowInES', 'IsActive', '_QBD_Type'],
    )

    print(f'[{office.upper()}] Done.\n')


# ---------------------------------------------------------------------------
# QBWC SOAP response builders
# ---------------------------------------------------------------------------

SOAP_ENVELOPE = """\
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <soap:Body>
    {body}
  </soap:Body>
</soap:Envelope>"""


def soap_response(method: str, inner: str) -> str:
    body = f'<{method}Response xmlns="http://developer.intuit.com/"><{method}Result>{inner}</{method}Result></{method}Response>'
    return SOAP_ENVELOPE.format(body=body)


def soap_response_multi(method: str, *results: str) -> str:
    tags = ''.join(f'<{method}Result>{r}</{method}Result>' for r in results)
    body = f'<{method}Response xmlns="http://developer.intuit.com/">{tags}</{method}Response>'
    return SOAP_ENVELOPE.format(body=body)


# ---------------------------------------------------------------------------
# Flask app
# ---------------------------------------------------------------------------

app = Flask(__name__)
OFFICE: str = 'dallas'
OUTPUT_DIR: Path = Path('output/dallas')


def get_soap_action() -> str:
    return request.headers.get('SOAPAction', '').strip('"').split('/')[-1]


def get_body() -> str:
    return request.data.decode('utf-8', errors='replace')


@app.post('/qbwc')
def qbwc():
    action = get_soap_action()
    body = get_body()
    handlers = {
        'serverVersion':    handle_server_version,
        'clientVersion':    handle_client_version,
        'authenticate':     handle_authenticate,
        'sendRequestXML':   handle_send_request,
        'receiveResponseXML': handle_receive_response,
        'getLastError':     handle_get_last_error,
        'closeConnection':  handle_close_connection,
    }
    handler = handlers.get(action)
    if handler is None:
        app.logger.warning(f'Unknown SOAP action: {action}')
        return Response('Unknown action', status=400)
    xml = handler(body)
    return Response(xml, content_type='text/xml; charset=utf-8')


def handle_server_version(body: str) -> str:
    return soap_response('serverVersion', '1.0.0')


def handle_client_version(body: str) -> str:
    # Return empty string = accept any client version
    return soap_response('clientVersion', '')


def handle_authenticate(body: str) -> str:
    password = _find_one('strPassword', body)
    if password != VALID_PASSWORD:
        # Return (ticket, 'nvu') — not valid user
        return soap_response_multi('authenticate', str(uuid.uuid4()), 'nvu')

    ticket = str(uuid.uuid4())
    sessions[ticket] = {
        'office':      OFFICE,
        'step_index':  0,
        'iterator_id': None,
        'records': defaultdict(list),
    }
    app.logger.info(f'Auth OK → ticket {ticket[:8]}...')
    # Return (ticket, '') — empty string = proceed with company file
    return soap_response_multi('authenticate', ticket, '')


def handle_send_request(body: str) -> str:
    ticket = _find_one('ticket', body)
    sess = sessions.get(ticket)
    if not sess:
        return soap_response('sendRequestXML', '')

    step_index = sess['step_index']
    if step_index >= len(QUERY_STEPS):
        return soap_response('sendRequestXML', '')  # done

    step = QUERY_STEPS[step_index]
    iterator_id = sess.get('iterator_id')
    iter_mode = 'Continue' if iterator_id else 'Start'

    query_builders = {
        'accounts':  account_query,
        'customers': customer_query,
        'vendors':   vendor_query,
        'employees': employee_query,
        'items':     item_query,
    }
    qbxml = query_builders[step](iter_mode, iterator_id)
    app.logger.info(f'[{ticket[:8]}] Sending {step} query (iter={iter_mode})')
    return soap_response('sendRequestXML', qbxml)


def handle_receive_response(body: str) -> str:
    ticket = _find_one('ticket', body)
    response_xml = _find_one('response', body)
    sess = sessions.get(ticket)
    if not sess:
        return soap_response('receiveResponseXML', '-1')

    step = QUERY_STEPS[sess['step_index']]
    prefix = OFFICE_PREFIXES[sess['office']]
    recs = sess['records']

    if step == 'accounts':
        rows, iterator_id, remaining = parse_accounts(response_xml, prefix)
        recs['accounts'].extend(rows)
    elif step == 'customers':
        clients, contacts, iterator_id, remaining = parse_customers(response_xml, prefix)
        recs['clients'].extend(clients)
        recs['client_contacts'].extend(contacts)
    elif step == 'vendors':
        vendors, contacts, iterator_id, remaining = parse_vendors(response_xml, prefix)
        recs['vendors'].extend(vendors)
        recs['vendor_contacts'].extend(contacts)
    elif step == 'employees':
        rows, iterator_id, remaining = parse_employees(response_xml, prefix)
        recs['employees'].extend(rows)
    elif step == 'items':
        rows, iterator_id, remaining = parse_items(response_xml, prefix)
        recs['items'].extend(rows)
    else:
        iterator_id, remaining = None, 0

    app.logger.info(
        f'[{ticket[:8]}] {step}: remaining={remaining}, '
        f'accounts={len(recs.get("accounts", []))}, '
        f'clients={len(recs.get("clients", []))}, '
        f'vendors={len(recs.get("vendors", []))}, '
        f'employees={len(recs.get("employees", []))}, '
        f'items={len(recs.get("items", []))}'
    )

    if remaining > 0:
        sess['iterator_id'] = iterator_id
        return soap_response('receiveResponseXML', '50')  # not done yet

    # Step complete — advance
    sess['iterator_id'] = None
    sess['step_index'] += 1

    if sess['step_index'] >= len(QUERY_STEPS):
        flush_session(ticket, OUTPUT_DIR)
        return soap_response('receiveResponseXML', '100')  # 100% = all done

    # Progress percentage
    pct = int(sess['step_index'] / len(QUERY_STEPS) * 100)
    return soap_response('receiveResponseXML', str(pct))


def handle_get_last_error(body: str) -> str:
    return soap_response('getLastError', '')


def handle_close_connection(body: str) -> str:
    ticket = _find_one('ticket', body)
    sessions.pop(ticket, None)
    return soap_response('closeConnection', 'OK')


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--office', choices=['dallas', 'orlando'], required=True)
    parser.add_argument('--port', type=int, default=5150)
    args = parser.parse_args()

    OFFICE = args.office
    OUTPUT_DIR = Path(__file__).parent.parent / 'output' / args.office

    print(f'QB Desktop Web Connector server')
    print(f'  Office:     {OFFICE.upper()}')
    print(f'  Output dir: {OUTPUT_DIR}')
    print(f'  Listening:  http://0.0.0.0:{args.port}/qbwc')
    print(f'  Password:   {VALID_PASSWORD}')
    print()

    app.run(host='0.0.0.0', port=args.port, debug=False)
