#!/usr/bin/env python3
"""Microsoft Graph OAuth login using device code flow."""

import msal
import json
import sys
from pathlib import Path

client_id = 'd9260854-0354-44f3-b12a-6aab224803ff'
tenant_id = '14c97b03-9bb9-44ba-a4a6-e44e181ab35e'
authority = f'https://login.microsoftonline.com/{tenant_id}'
scopes = ['Calendars.ReadWrite', 'User.Read']

def main():
    app = msal.PublicClientApplication(client_id, authority=authority)

    # Start device code flow
    flow = app.initiate_device_flow(scopes=scopes)

    if 'user_code' not in flow:
        print(f"Error: {flow.get('error_description', flow)}", flush=True)
        sys.exit(1)

    print("=" * 60, flush=True)
    print("To sign in, open this URL in your browser:", flush=True)
    print(flush=True)
    print(f"  {flow['verification_uri']}", flush=True)
    print(flush=True)
    print(f"Enter this code:  {flow['user_code']}", flush=True)
    print("=" * 60, flush=True)
    print(flush=True)
    print("Waiting for you to complete sign-in...", flush=True)

    # Wait for user to complete login
    result = app.acquire_token_by_device_flow(flow)

    if 'access_token' in result:
        # Save token
        token_data = {
            'access_token': result['access_token'],
            'refresh_token': result.get('refresh_token'),
            'id_token_claims': result.get('id_token_claims', {}),
        }

        token_path = Path(__file__).parent / 'ms_token.json'
        with open(token_path, 'w') as f:
            json.dump(token_data, f, indent=2)

        username = result.get('id_token_claims', {}).get('preferred_username', 'Unknown')
        print(flush=True)
        print(f"SUCCESS! Logged in as: {username}", flush=True)
        print("Token saved to ms_token.json", flush=True)
        print("Calendar integration is now active!", flush=True)
    else:
        print(f"Login failed: {result.get('error_description', result)}", flush=True)
        sys.exit(1)

if __name__ == '__main__':
    main()
