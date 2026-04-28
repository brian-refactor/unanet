"""
Run this once to authorize with QuickBooks Online and save tokens.
After running, tokens are stored in qbo_tokens.json for use by qbo_extract.py.

Usage:
    python qbo_auth.py

Prerequisites:
    1. Copy .env.example to .env and fill in QBO_CLIENT_ID and QBO_CLIENT_SECRET
    2. In your Intuit Developer app, add http://localhost:8000/callback as a redirect URI
"""

import os
import json
import webbrowser
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, parse_qs

from dotenv import load_dotenv
from intuitlib.client import AuthClient
from intuitlib.enums import Scopes

HERE = Path(__file__).parent
TOKEN_FILE = HERE / 'qbo_tokens.json'
REDIRECT_URI = 'http://localhost:8000/callback'


class _CallbackHandler(BaseHTTPRequestHandler):
    result = {}

    def do_GET(self):
        params = parse_qs(urlparse(self.path).query)
        _CallbackHandler.result['code'] = params.get('code', [None])[0]
        _CallbackHandler.result['realm_id'] = params.get('realmId', [None])[0]

        self.send_response(200)
        self.send_header('Content-Type', 'text/html')
        self.end_headers()
        self.wfile.write(
            b'<h2>Authorization successful!</h2>'
            b'<p>You can close this tab and return to the terminal.</p>'
        )
        threading.Thread(target=self.server.shutdown, daemon=True).start()

    def log_message(self, *args):
        pass  # suppress server noise


def main():
    load_dotenv(HERE / '.env')

    client_id = os.getenv('QBO_CLIENT_ID')
    client_secret = os.getenv('QBO_CLIENT_SECRET')
    environment = os.getenv('QBO_ENVIRONMENT', 'production')

    if not client_id or not client_secret:
        raise SystemExit('ERROR: QBO_CLIENT_ID and QBO_CLIENT_SECRET must be set in .env')

    auth_client = AuthClient(
        client_id=client_id,
        client_secret=client_secret,
        redirect_uri=REDIRECT_URI,
        environment=environment,
    )

    url = auth_client.get_authorization_url([Scopes.ACCOUNTING])

    print(f'\nOpening browser for QuickBooks Online authorization...')
    print(f'If the browser does not open automatically, visit:\n  {url}\n')
    webbrowser.open(url)

    print('Waiting for authorization callback on http://localhost:8000 ...')
    server = HTTPServer(('localhost', 8000), _CallbackHandler)
    server.serve_forever()  # blocks until callback handler shuts it down

    code = _CallbackHandler.result.get('code')
    realm_id = _CallbackHandler.result.get('realm_id')

    if not code or not realm_id:
        raise SystemExit('ERROR: Authorization failed — no code or realm ID received.')

    auth_client.get_bearer_token(code, realm_id=realm_id)

    tokens = {
        'access_token': auth_client.access_token,
        'refresh_token': auth_client.refresh_token,
        'realm_id': realm_id,
    }

    TOKEN_FILE.write_text(json.dumps(tokens, indent=2))

    print(f'\nTokens saved to {TOKEN_FILE}')
    print(f'Realm ID (Company ID): {realm_id}')
    print('\nYou can now run qbo_extract.py')


if __name__ == '__main__':
    main()
