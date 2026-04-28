"""
Probe 3 — check project LaborEntry detail and try report/summary methods.
"""

import json, os
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


def show(label, r, full=False):
    rc = r.get('ResponseCode')
    content = r.get('Content', {})
    errors = [e for e in r.get('Errors', []) if e.get('ErrorID', 0) != 0]
    print(f'  [{rc}] {label}')
    if content:
        if full:
            print(f'         {json.dumps(content, default=str)[:2000]}')
        else:
            for k, v in content.items():
                if isinstance(v, list):
                    print(f'         {k}: {len(v)} records')
                    if v and isinstance(v[0], dict):
                        print(f'         First keys: {list(v[0].keys())}')
                        print(f'         First: {json.dumps(v[0], default=str)[:600]}')
                else:
                    print(f'         {k}: {v}')
    if errors:
        print(f'         Errors: {errors}')
    print()


def main():
    token, api_url = create_session()
    print('Connected.\n')

    # Get some active projects (skip overhead/blank ID projects)
    r = call(api_url, token, 'ListProjects')
    all_proj = r.get('Content', {}).get('Projects', [])
    active_keys = [p['ProjectKey'] for p in all_proj if p.get('ID')][:5]
    print(f'Active project keys (with IDs): {active_keys}\n')

    print('=== GetProjects — LaborEntry / ExpenseConsultantEntry detail ===\n')
    r = call(api_url, token, 'GetProjects', {'RequestedProjects': active_keys[:3]})
    projects = r.get('Content', {}).get('Projects', [])
    for p in projects:
        print(f'  Project: {p.get("ID")} — {p.get("Description")}')
        labor = p.get('LaborEntry', [])
        expense = p.get('ExpenseConsultantEntry', [])
        print(f'    LaborEntry: {len(labor)} records')
        if labor:
            print(f'    LaborEntry keys: {list(labor[0].keys())}')
            print(f'    LaborEntry[0]: {json.dumps(labor[0], default=str)[:600]}')
        print(f'    ExpenseConsultantEntry: {len(expense)} records')
        if expense:
            print(f'    ExpenseEntry keys: {list(expense[0].keys())}')
            print(f'    ExpenseEntry[0]: {json.dumps(expense[0], default=str)[:400]}')
        print()

    print('=== Reporting / summary method candidates ===\n')
    for method in [
        'GetProjectSummary',
        'ListProjectSummary',
        'GetProjectBudget',
        'GetProjectActuals',
        'ListProjectActuals',
        'GetUtilization',
        'ListUtilization',
        'GetEmployeeUtilization',
        'GetLaborSummary',
        'ListLaborSummary',
        'GetTimesheet',
        'ListTimesheet',
        'GetTimesheetEntries',
        'ListTimesheetEntries',
        'GetDashboard',
        'ListReports',
        'GetReport',
        'RunReport',
    ]:
        r = call(api_url, token, method)
        rc = r.get('ResponseCode')
        errors = [e for e in r.get('Errors', []) if e.get('ErrorID', 0) != 0]
        content = r.get('Content', {})
        if rc == 200:
            print(f'  [200] {method} — WORKS! Keys: {list(content.keys())}')
        elif rc == 0 and errors:
            err_id = errors[0].get('ErrorID') if errors else '?'
            print(f'  [0/{err_id}] {method}')
        else:
            print(f'  [{rc}] {method}')

    print()
    print('=== Try GetEmployees with full detail — check for hours fields ===\n')
    r = call(api_url, token, 'GetEmployees', {'RequestedEmployees': [23]})
    emps = r.get('Content', {}).get('Employees', [])
    if emps:
        print(f'  All employee keys: {list(emps[0].keys())}')

    requests.post(api_url, json={'Method': 'EndAPISession', 'SessionToken': token}, headers=HEADERS, timeout=15)
    print('\nSession closed.')


if __name__ == '__main__':
    main()
