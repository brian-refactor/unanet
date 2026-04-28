"""
Probe Ajera API for revenue pipeline data availability.
Tests project, contract, invoice, and WIP methods and shows
a sample of whatever comes back.

Usage:
    python etl/ajera_probe_pipeline.py
"""

import json
import os
from pathlib import Path

import requests
from dotenv import load_dotenv

HERE = Path(__file__).parent
HEADERS = {'Content-Type': 'application/json'}


def create_session():
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
    print(f'Connected — {data["Content"].get("CompanyName")}  '
          f'(Ajera {data["Content"].get("AjeraVersion")})\n')
    return token, api_url


def end_session(api_url, token):
    try:
        requests.post(api_url, json={
            'Method': 'EndAPISession', 'SessionToken': token,
        }, headers=HEADERS, timeout=15)
    except Exception:
        pass


def probe(api_url, token, method, args=None, label=None):
    """Call a method and print a structured summary of the result."""
    label = label or method
    payload = {'Method': method, 'SessionToken': token, 'MethodArguments': args or {}}
    resp = requests.post(api_url, json=payload, headers=HEADERS, timeout=60)
    data = resp.json()

    rc       = data.get('RC', '?')
    errors   = data.get('Errors', [])
    content  = data.get('Content', {})
    err_ids  = [e.get('ErrorID') for e in errors if e.get('ErrorID', 0) != 0]

    status = 'OK' if rc == 200 else f'RC={rc}'
    if err_ids:
        status += f'  errors={err_ids}'

    print('-' * 60)
    print(f'  {label}  [{status}]')

    if rc == 200 and content:
        # Show what keys came back and how many items
        if isinstance(content, list):
            print(f'  >> list of {len(content)} items')
            if content:
                print(f'  >> sample keys: {sorted(content[0].keys())}')
                _print_sample(content[0])
        elif isinstance(content, dict):
            keys = sorted(content.keys())
            print(f'  >> dict keys: {keys}')
            for k in keys:
                v = content[k]
                if isinstance(v, list) and v:
                    print(f'     {k}: {len(v)} items, sample keys: {sorted(v[0].keys())}')
                    _print_sample(v[0], indent=8)
                elif not isinstance(v, (list, dict)):
                    print(f'     {k}: {v}')
    elif err_ids:
        for e in errors:
            if e.get('ErrorID', 0) != 0:
                print(f'  >> {e}')
    print()


def _print_sample(obj, indent=4):
    """Print up to 12 key-value pairs from a dict, skipping nulls."""
    pad = ' ' * indent
    shown = 0
    for k, v in obj.items():
        if v in (None, '', [], {}):
            continue
        print(f'{pad}{k}: {str(v)[:80]}')
        shown += 1
        if shown >= 12:
            remaining = len(obj) - shown
            if remaining > 0:
                print(f'{pad}... ({remaining} more keys)')
            break


# ---------------------------------------------------------------------------
# Main probe sequence
# ---------------------------------------------------------------------------

def main():
    token, api_url = create_session()

    try:
        # ── 1. Project list ──────────────────────────────────────────────
        probe(api_url, token, 'ListProjects',
              label='ListProjects — all projects')

        probe(api_url, token, 'ListProjects',
              {'ProjectStatus': 'Open'},
              label='ListProjects — open only')

        # ── 2. Project detail (need a key from above — try key=1 as guess) ─
        probe(api_url, token, 'GetProject',
              {'ProjectKey': 1},
              label='GetProject(key=1) — detail + budget?')

        probe(api_url, token, 'GetProjectDetail',
              {'ProjectKey': 1},
              label='GetProjectDetail(key=1)')

        # ── 3. Contract / phase data ─────────────────────────────────────
        probe(api_url, token, 'ListContracts',
              label='ListContracts')

        probe(api_url, token, 'ListProjectPhases',
              {'ProjectKey': 1},
              label='ListProjectPhases(key=1)')

        probe(api_url, token, 'ListPhases',
              label='ListPhases — all')

        # ── 4. Billing / invoice data ────────────────────────────────────
        probe(api_url, token, 'ListInvoices',
              label='ListInvoices — all')

        probe(api_url, token, 'ListInvoices',
              {'ProjectKey': 1},
              label='ListInvoices(project=1)')

        probe(api_url, token, 'GetInvoices',
              label='GetInvoices')

        probe(api_url, token, 'ListBillingTransactions',
              label='ListBillingTransactions')

        probe(api_url, token, 'ListARInvoices',
              label='ListARInvoices')

        # ── 5. WIP / unbilled ────────────────────────────────────────────
        probe(api_url, token, 'GetProjectWIP',
              {'ProjectKey': 1},
              label='GetProjectWIP(key=1)')

        probe(api_url, token, 'ListWIP',
              label='ListWIP')

        probe(api_url, token, 'GetProjectActuals',
              {'ProjectKey': 1},
              label='GetProjectActuals(key=1) — unbilled hours?')

        # ── 6. Project financials ────────────────────────────────────────
        probe(api_url, token, 'GetProjectFinancials',
              {'ProjectKey': 1},
              label='GetProjectFinancials(key=1)')

        probe(api_url, token, 'ListProjectSummary',
              label='ListProjectSummary')

        probe(api_url, token, 'GetProjectBilling',
              {'ProjectKey': 1},
              label='GetProjectBilling(key=1)')

        # ── 7. Accounts receivable ───────────────────────────────────────
        probe(api_url, token, 'ListARTransactions',
              label='ListARTransactions')

        probe(api_url, token, 'GetARBalance',
              label='GetARBalance')

    finally:
        end_session(api_url, token)
        print('Session closed.')


if __name__ == '__main__':
    main()
