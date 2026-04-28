"""
Use this instead of qbo_auth.py if your Intuit app uses the OAuth Playground
redirect URL (https://developer.intuit.com/v2/OAuth2Playground/RedirectUrl).

Steps:
    1. Go to https://developer.intuit.com/app/developer/playground
    2. Select your app from the dropdown
    3. Under Scopes, check: com.intuit.quickbooks.accounting
    4. Click "Get Authorization Code" and authorize with the Minnesota QBO account
    5. Click "Get Token" to exchange for access + refresh tokens
    6. Copy the three values shown and paste them in when prompted below

Usage:
    python qbo_save_tokens.py
"""

import json
from pathlib import Path

TOKEN_FILE = Path(__file__).parent / 'qbo_tokens.json'

print('\n--- QuickBooks Online Token Setup ---')
print('Paste the values from the Intuit OAuth Playground:\n')

access_token  = input('Access Token:  ').strip()
refresh_token = input('Refresh Token: ').strip()
realm_id      = input('Realm ID:      ').strip()

if not all([access_token, refresh_token, realm_id]):
    raise SystemExit('ERROR: All three values are required.')

tokens = {
    'access_token': access_token,
    'refresh_token': refresh_token,
    'realm_id': realm_id,
}

TOKEN_FILE.write_text(json.dumps(tokens, indent=2))
print(f'\nTokens saved to {TOKEN_FILE}')
print('You can now run:  python qbo_extract.py')
