"""
Probe Ajera API to see if activity (expense code) records carry GL account links.
Tests ListActivities and GetActivities to find DirectBaseCode/DirectBasename data.

Usage:
    python etl/ajera_probe_activities.py
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


def call(api_url, token, method, args=None):
    payload = {'Method': method, 'SessionToken': token, 'MethodArguments': args or {}}
    resp = requests.post(api_url, json=payload, headers=HEADERS, timeout=60)
    data = resp.json()
    rc = data.get('RC', '?')
    errors = data.get('Errors', [])
    content = data.get('Content', {})
    err_ids = [e.get('ErrorID') for e in errors if e.get('ErrorID', 0) != 0]
    return rc, content, err_ids, errors


def end_session(api_url, token):
    try:
        requests.post(api_url, json={
            'Method': 'EndAPISession', 'SessionToken': token,
        }, headers=HEADERS, timeout=15)
    except Exception:
        pass


def main():
    token, api_url = create_session()

    try:
        # ── 1. List activities — grab the first few ActivityKeys ──────────
        print('=' * 60)
        print('ListActivities — full response structure')
        print('=' * 60)
        rc, content, err_ids, errors = call(api_url, token, 'ListActivities')
        print(f'RC={rc}  errors={err_ids}')

        activities = content.get('Activities', [])
        print(f'Activities count: {len(activities)}')

        if activities:
            print(f'\nAll keys on first activity record:')
            first = activities[0]
            for k, v in first.items():
                print(f'  {k}: {repr(v)[:80]}')

            # Grab up to 3 keys to probe GetActivities
            sample_keys = [
                a.get('ActivityKey') for a in activities[:3]
                if a.get('ActivityKey')
            ]
        else:
            print('No activities returned.')
            sample_keys = []

        # ── 2. Try GetActivities with detail keys ─────────────────────────
        if sample_keys:
            print(f'\n{"=" * 60}')
            print(f'GetActivities — detail for keys {sample_keys}')
            print('=' * 60)

            for arg_name in ['RequestedActivities', 'ActivityKeys', 'ActivityKey']:
                arg_val = sample_keys if arg_name != 'ActivityKey' else sample_keys[0]
                rc, content, err_ids, errors = call(
                    api_url, token, 'GetActivities', {arg_name: arg_val}
                )
                print(f'\nGetActivities({arg_name}=...)  RC={rc}  errors={err_ids}')
                if rc == 200:
                    details = content.get('Activities', [])
                    print(f'  Returned {len(details)} records')
                    if details:
                        print(f'  All keys on first detail record:')
                        for k, v in details[0].items():
                            print(f'    {k}: {repr(v)[:80]}')
                    break
                else:
                    for e in errors:
                        if e.get('ErrorID', 0) != 0:
                            print(f'  error: {e}')

        # ── 3. Try GetActivity (singular) ─────────────────────────────────
        if sample_keys:
            print(f'\n{"=" * 60}')
            print(f'GetActivity (singular) — key={sample_keys[0]}')
            print('=' * 60)
            rc, content, err_ids, errors = call(
                api_url, token, 'GetActivity', {'ActivityKey': sample_keys[0]}
            )
            print(f'RC={rc}  errors={err_ids}')
            if rc == 200 and content:
                print('  Keys:', sorted(content.keys()) if isinstance(content, dict) else type(content))
                if isinstance(content, dict):
                    for k, v in content.items():
                        print(f'  {k}: {repr(v)[:80]}')
            else:
                for e in errors:
                    if e.get('ErrorID', 0) != 0:
                        print(f'  error: {e}')

        # ── 4. Dump all ListActivities records to JSON for inspection ─────
        if activities:
            out = HERE.parent / 'output' / 'cincinnati' / '_activities_raw.json'
            out.write_text(json.dumps(activities, indent=2))
            print(f'\nFull ListActivities dump written to: {out}')

    finally:
        end_session(api_url, token)


if __name__ == '__main__':
    main()
