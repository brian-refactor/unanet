"""
Quick connection test for Ajera API — Cincinnati office.
Run after filling in etl/ajera.env.

Usage:
    python etl/ajera_test.py
"""

import os
import requests
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path(__file__).parent / 'ajera.env')

API_URL  = os.environ['AJERA_API_URL']
USERNAME = os.environ['AJERA_USERNAME']
PASSWORD = os.environ['AJERA_PASSWORD']


def post(payload: dict) -> dict:
    r = requests.post(API_URL, json=payload, headers={'Content-Type': 'application/json'}, timeout=30)
    r.raise_for_status()
    return r.json()


def main():
    print(f'\nConnecting to {API_URL.split("?")[0]}...')

    # Step 1 — create session
    resp = post({
        'Method': 'CreateAPISession',
        'Username': USERNAME,
        'Password': PASSWORD,
        'APIVersion': 1,
        'UseSessionCookie': False,
    })

    if resp.get('ResponseCode') != 200:
        raise SystemExit(f'Auth failed: {resp.get("Message")} | Errors: {resp.get("Errors")}')

    content = resp['Content']
    token = content['SessionToken']

    print(f'Connected!')
    print(f'  Company:      {content.get("CompanyName")}')
    print(f'  Ajera version:{content.get("AjeraVersion")}')
    print(f'  API user:     {content.get("EmployeeName")}')

    # Step 2 — quick count of each entity we need
    print('\nEntity counts:')
    for method, label in [
        ('ListClients',   'Clients'),
        ('ListVendors',   'Vendors'),
        ('ListEmployees', 'Employees'),
        ('ListContacts',  'Contacts'),
    ]:
        r = post({
            'Method': method,
            'SessionToken': token,
            'MethodArguments': {'PageResults': True, 'PageNumber': 1, 'PageSize': 1},
        })
        if r.get('ResponseCode') == 200:
            total_pages = r['Content'].get('PageCount', '?')
            # PageSize=1 so PageCount == total record count
            print(f'  {label}: {total_pages} records')
        else:
            print(f'  {label}: ERROR — {r.get("Message")}')

    # Step 3 — end session
    post({'Method': 'EndAPISession', 'SessionToken': token})
    print('\nSession closed. Connection test passed.')


if __name__ == '__main__':
    main()
