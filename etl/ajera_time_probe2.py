"""
Second probe — try time methods with required parameters.
"""

import json, os
from datetime import date, timedelta
from pathlib import Path

import requests
from dotenv import load_dotenv

HERE = Path(__file__).parent
HEADERS = {'Content-Type': 'application/json'}


def create_session():
    load_dotenv(HERE / 'ajera.env')
    api_url = os.environ['AJERA_API_URL']
    resp = requests.post(api_url, json={
        'Method': 'CreateAPISession',
        'Username': os.environ['AJERA_USERNAME'],
        'Password': os.environ['AJERA_PASSWORD'],
        'APIVersion': 1, 'UseSessionCookie': False,
    }, headers=HEADERS, timeout=30)
    data = resp.json()
    return data['Content']['SessionToken'], api_url


def call(api_url, token, method, args=None):
    resp = requests.post(api_url, json={
        'Method': method, 'SessionToken': token, 'MethodArguments': args or {},
    }, headers=HEADERS, timeout=60)
    return resp.json()


def show(label, r):
    rc = r.get('ResponseCode')
    content = r.get('Content', {})
    errors = [e for e in r.get('Errors', []) if e.get('ErrorID', 0) != 0]
    print(f'  [{rc}] {label}')
    if content:
        for k, v in content.items():
            if isinstance(v, list):
                print(f'         {k}: {len(v)} records')
                if v and isinstance(v[0], dict):
                    print(f'         Keys: {list(v[0].keys())}')
                    print(f'         Sample: {json.dumps(v[0], default=str)[:500]}')
            else:
                print(f'         {k}: {v}')
    if errors:
        print(f'         Errors: {errors}')
    print()


def main():
    token, api_url = create_session()
    print('Connected.\n')

    end_dt   = date.today().isoformat()
    start_dt = (date.today() - timedelta(days=365)).isoformat()

    # --- Get a few employee keys to use in parameterized calls ---
    r = call(api_url, token, 'ListEmployees')
    emp_keys = [e['EmployeeKey'] for e in r.get('Content', {}).get('Employees', [])[:3]]
    print(f'Sample employee keys: {emp_keys}\n')

    # --- Get a few project keys ---
    r = call(api_url, token, 'ListProjects')
    proj_keys = [p['ProjectKey'] for p in r.get('Content', {}).get('Projects', [])[:5]]
    print(f'Sample project keys: {proj_keys}\n')

    print('=== ListTimeEntries with various param styles ===\n')
    param_variants = [
        {'EmployeeKeys': emp_keys},
        {'EmployeeKey': emp_keys[0]},
        {'ProjectKeys': proj_keys},
        {'StartDate': start_dt, 'EndDate': end_dt, 'EmployeeKeys': emp_keys},
        {'BeginDate': start_dt, 'EndDate': end_dt, 'EmployeeKeys': emp_keys},
        {'FromDate': start_dt, 'ToDate': end_dt, 'EmployeeKeys': emp_keys},
        {'PageResults': True, 'PageNumber': 1, 'PageSize': 10},
        {'PageResults': True, 'PageNumber': 1, 'PageSize': 10, 'StartDate': start_dt, 'EndDate': end_dt},
        {'RequestedEmployees': emp_keys},
        {'Employees': emp_keys},
    ]
    for args in param_variants:
        r = call(api_url, token, 'ListTimeEntries', args)
        show(f'ListTimeEntries {list(args.keys())}', r)

    print('=== GetTimeEntries variants ===\n')
    for args in [
        {'EmployeeKeys': emp_keys},
        {'RequestedEmployees': emp_keys},
        {'StartDate': start_dt, 'EndDate': end_dt},
        {'TimeEntryKeys': [1, 2, 3, 4, 5]},
        {'RequestedTimeEntries': [1, 2, 3]},
    ]:
        r = call(api_url, token, 'GetTimeEntries', args)
        show(f'GetTimeEntries {list(args.keys())}', r)

    print('=== GetProjects — check if time data is embedded ===\n')
    r = call(api_url, token, 'GetProjects', {'RequestedProjects': proj_keys[:2]})
    show('GetProjects (2 projects)', r)

    print('=== ListTimeRecords variants ===\n')
    for args in [
        {'EmployeeKeys': emp_keys, 'StartDate': start_dt, 'EndDate': end_dt},
        {'PageResults': True, 'PageNumber': 1, 'PageSize': 10},
    ]:
        r = call(api_url, token, 'ListTimeRecords', args)
        show(f'ListTimeRecords {list(args.keys())}', r)

    print('=== ListActivityEntries variants ===\n')
    for args in [
        {'EmployeeKeys': emp_keys},
        {'StartDate': start_dt, 'EndDate': end_dt},
        {'ActivityKeys': [1, 2, 3]},
    ]:
        r = call(api_url, token, 'ListActivityEntries', args)
        show(f'ListActivityEntries {list(args.keys())}', r)

    requests.post(api_url, json={'Method': 'EndAPISession', 'SessionToken': token}, headers=HEADERS, timeout=15)
    print('Session closed.')


if __name__ == '__main__':
    main()
