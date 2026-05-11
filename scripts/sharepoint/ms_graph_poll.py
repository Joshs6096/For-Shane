#!/usr/bin/env python3
import requests, json, time, sys

CLIENT_ID = '9e5f94bc-e8a4-4e73-b8be-63364c29d753'
with open('/tmp/.ms_device_code.json') as f:
    dc = json.load(f)

start = time.time()
result = None
while time.time() - start < dc['expires_in']:
    time.sleep(5)
    r = requests.post(
        'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        data={
            'client_id': CLIENT_ID,
            'grant_type': 'urn:ietf:params:oauth:grant-type:device_code',
            'device_code': dc['device_code']
        }
    )
    if r.status_code == 200:
        token = r.json()
        token['obtained_at'] = int(time.time())
        token['expires_at'] = int(time.time()) + token.get('expires_in', 3600)
        with open('/Users/josh/.openclaw/workspace/.ms_graph_token.json', 'w') as f:
            json.dump(token, f, indent=2)
        result = 'SUCCESS'
        print('SUCCESS', flush=True)
        break
    err = r.json().get('error', '')
    if err == 'authorization_pending':
        print('.', end='', flush=True)
        continue
    if err:
        result = f'FAILED: {err}'
        print(f'\nFAILED: {err}', flush=True)
        break

if not result:
    result = 'TIMEOUT'
    print('\nTIMEOUT', flush=True)

with open('/tmp/.ms_poll_result.txt', 'w') as f:
    f.write(result)
