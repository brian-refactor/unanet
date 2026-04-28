"""
Probe Ajera API for time entry methods and print response structure.
Run this once to confirm method names and field shapes before building the extractor.

Usage:
    python etl/ajera_time_probe.py
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
    resp = requests.post(api_url, json={
        'Method': 'CreateAPISession',
        'Username': os.environ['AJERA_USERNAME'],
        'Password': os.environ['AJERA_PASSWORD'],
        'APIVersion': 1,
        'UseSessionCookie': False,
    }, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    token = data['Content']['SessionToken']
    print(f"Connected: {data['Content'].get('CompanyName')} (Ajera {data['Content'].get('AjeraVersion')})\n")
    return token, api_url


def call(api_url, token, method, args=None):
    resp = requests.post(api_url, json={
        'Method': method,
        'SessionToken': token,
        'MethodArguments': args or {},
    }, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    return resp.json()


def probe(api_url, token, method, args=None):
    r = call(api_url, token, method, args)
    rc = r.get('ResponseCode')
    content = r.get('Content', {})
    errors = [e for e in r.get('Errors', []) if e.get('ErrorID', 0) != 0]
    keys = list(content.keys()) if content else []
    print(f'  [{rc}] {method}')
    if keys:
        print(f'         Content keys: {keys}')
        # Show count and first record shape for any list found
        for k in keys:
            val = content[k]
            if isinstance(val, list) and val:
                print(f'         {k}: {len(val)} records')
                print(f'         First record keys: {list(val[0].keys()) if isinstance(val[0], dict) else type(val[0])}')
                print(f'         First record sample: {json.dumps(val[0], default=str)[:400]}')
            elif isinstance(val, list):
                print(f'         {k}: [] (empty)')
            else:
                print(f'         {k}: {val}')
    if errors:
        print(f'         Errors: {errors}')
    print()
    return content, rc


def main():
    token, api_url = create_session()

    try:
        print('=== Time Entry Methods ===\n')
        time_methods = [
            ('ListTimeEntries',     {}),
            ('GetTimeEntries',      {}),
            ('ListTimesheets',      {}),
            ('GetTimesheets',       {}),
            ('ListTimeRecords',     {}),
            ('GetTimeRecords',      {}),
            ('ListActivityEntries', {}),
            ('GetActivityEntries',  {}),
            ('ListTimeAndExpense',  {}),
            ('GetTimeAndExpense',   {}),
            ('ListProjectTime',     {}),
            ('GetProjectTime',      {}),
            ('ListLabor',           {}),
            ('GetLabor',            {}),
            ('ListTime',            {}),
        ]
        for method, args in time_methods:
            probe(api_url, token, method, args)

        print('=== Project Methods (for cross-referencing) ===\n')
        project_methods = [
            ('ListProjects',   {}),
            ('ListPhases',     {}),
            ('ListActivities', {}),
            ('ListTasks',      {}),
        ]
        for method, args in project_methods:
            probe(api_url, token, method, args)

        # If ListTimeEntries worked, try with date range args
        print('=== ListTimeEntries with date range (last 90 days) ===\n')
        from datetime import date, timedelta
        end_date   = date.today().isoformat()
        start_date = (date.today() - timedelta(days=90)).isoformat()
        for args in [
            {'StartDate': start_date, 'EndDate': end_date},
            {'BeginDate': start_date, 'EndDate': end_date},
            {'FromDate':  start_date, 'ToDate':  end_date},
            {'DateFrom':  start_date, 'DateTo':  end_date},
            {'PageResults': True, 'PageNumber': 1, 'PageSize': 5},
        ]:
            probe(api_url, token, 'ListTimeEntries', args)

    finally:
        requests.post(api_url, json={'Method': 'EndAPISession', 'SessionToken': token}, headers=HEADERS, timeout=15)
        print('Session closed.')


if __name__ == '__main__':
    main()
