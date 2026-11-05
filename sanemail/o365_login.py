#!/usr/bin/env python3
"""
One-time Office 365 sign-in for SaneMail.

Run this on your own computer (not in GitHub):

    pip install -r requirements.txt
    python o365_login.py <account-name-from-config.yaml>

It shows a code to enter at https://microsoft.com/devicelogin. Sign in with your work
account. The sign-in is saved to o365-token-<account>.json. Copy that file's contents into a
GitHub secret (named by `token_secret` in config.yaml). After that, SaneMail keeps the sign-in
fresh by itself. You only need to repeat this if it stops working, for example after a
password change or if IT revokes it.
"""
import json
import sys

import msal
import yaml

from sanemail import CONFIG_PATH, GRAPH_SCOPES


def main():
    with open(CONFIG_PATH) as fh:
        cfg = yaml.safe_load(fh)
    o365 = [a for a in cfg['accounts'] if a['type'] == 'o365']
    if not o365:
        sys.exit('No account with type: o365 in config.yaml')
    name = sys.argv[1] if len(sys.argv) > 1 else o365[0]['name']
    acct = next((a for a in o365 if a['name'] == name), None)
    if not acct:
        sys.exit(f'No o365 account named {name!r} in config.yaml')

    cache = msal.SerializableTokenCache()
    app = msal.PublicClientApplication(
        acct['client_id'], authority=f"https://login.microsoftonline.com/{acct['tenant']}", token_cache=cache)
    flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
    if 'user_code' not in flow:
        sys.exit('Could not start sign-in: ' + json.dumps(flow, indent=1))
    print('\n' + flow['message'] + '\n')
    result = app.acquire_token_by_device_flow(flow)
    if 'access_token' not in result:
        sys.exit('Sign-in failed: ' + result.get('error_description', json.dumps(result)))

    out = f'o365-token-{name}.json'
    with open(out, 'w') as fh:
        fh.write(cache.serialize())
    print(f'Signed in as {result.get("id_token_claims", {}).get("preferred_username", "?")}.')
    print(f'Saved {out}. Next, put its contents in the GitHub secret {acct["token_secret"]}:')
    print(f'  gh secret set {acct["token_secret"]} < {out}')
    print(f'Then delete {out}.')


if __name__ == '__main__':
    main()
