"""
Probe Ajera v2 API for timesheet / time entry data.
v2 uses the same URL as v1 but creates a session without APIVersion:1.

Usage:
    python etl/ajera_probe_timesheets.py
"""

import json
import os
from pathlib import Path

import requests
from dotenv import load_dotenv

HERE = Path(__file__).parent
HEADERS = {'Content-Type': 'application/json'}


def create_session(api_url, username, password):
    resp = requests.post(api_url, json={
        'Method': 'CreateAPISession',
        'Username': username,
        'Password': password,
        'APIVersion': 2,
        'UseSessionCookie': False,
    }, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    token = data.get('Content', {}).get('SessionToken')
    if not token:
        raise SystemExit(f'Auth failed: {data.get("Errors")}')
    print(f'Connected — {data["Content"].get("CompanyName")}  '
          f'(Ajera {data["Content"].get("AjeraVersion")})\n')
    return token


def call(api_url, token, method, args=None):
    payload = {'Method': method, 'SessionToken': token, 'MethodArguments': args or {}}
    resp = requests.post(api_url, json=payload, headers=HEADERS, timeout=60)
    data = resp.json()
    rc = data.get('ResponseCode') or data.get('RC', '?')
    errors = data.get('Errors', [])
    content = data.get('Content', {})
    err_ids = [e.get('ErrorID') for e in errors if e.get('ErrorID', 0) != 0]
    print(f'  ResponseCode={rc}  content type={type(content).__name__}  content keys={list(content.keys()) if isinstance(content, dict) else len(content) if isinstance(content, list) else content}')
    return rc, content, err_ids, errors


def end_session(api_url, token):
    try:
        requests.post(api_url, json={
            'Method': 'EndAPISession', 'SessionToken': token,
        }, headers=HEADERS, timeout=15)
    except Exception:
        pass


def show(label, rc, content, err_ids, errors):
    print(f'\n{"=" * 60}')
    print(f'{label}  RC={rc}  errors={err_ids}')
    print('=' * 60)
    if rc == 200 and content:
        items = content if isinstance(content, list) else content.get('Timesheets', [])
        print(f'Records returned: {len(items)}')
        if items:
            print('Keys on first record:')
            for k, v in items[0].items():
                print(f'  {k}: {repr(str(v))[:80]}')
    else:
        for e in errors:
            if e.get('ErrorID', 0) != 0:
                print(f'  error: {e}')


def main():
    load_dotenv(HERE / 'ajera.env')
    api_url  = os.environ['AJERA_API_URL']
    username = os.environ['AJERA_USERNAME']
    password = os.environ['AJERA_PASSWORD']

    token = create_session(api_url, username, password)

    try:
        # ── 1. List all timesheets (no filter) ───────────────────────────
        rc, content, err_ids, errors = call(api_url, token, 'ListTimesheets')
        show('ListTimesheets — no filter', rc, content, err_ids, errors)

        # ── 2. List with a date filter going back to 2023 ────────────────
        rc, content, err_ids, errors = call(api_url, token, 'ListTimesheets', {
            'FilterByEarliestTimesheetDate': '2023-01-01',
        })
        show('ListTimesheets — since 2023-01-01', rc, content, err_ids, errors)

        timesheets = content.get('Timesheets', []) if rc == 200 else []
        print(f'Metadata: {content.get("Metadata")}')

        # ── 3. GetTimesheets — grab first 3 keys for detail ──────────────
        if timesheets:
            keys = [t.get('Timesheet Key') for t in timesheets[:3] if t.get('Timesheet Key')]
            print(f'\nSample TimesheetKeys: {keys}')
            rc, content, err_ids, errors = call(api_url, token, 'GetTimesheets', {
                'RequestedTimesheets': keys,
            })
            show('GetTimesheets — detail for first 3', rc, content, err_ids, errors)

            # Dump full detail of first record for inspection
            if rc == 200:
                details = content.get('Timesheets', [])
                if details:
                    out = HERE.parent / 'output' / 'cincinnati' / '_timesheets_raw.json'
                    out.write_text(json.dumps(details[:3], indent=2))
                    print(f'\nRaw detail written to: {out}')

    finally:
        end_session(api_url, token)


if __name__ == '__main__':
    main()
