"""
Re-authentification Google Drive avec scope écriture.
"""
from google_auth_oauthlib.flow import InstalledAppFlow
import json
import os

SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]

CREDS_FILE = os.path.expanduser('~/.config/mcp-gdrive/gcp-oauth.keys.json')
TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')

flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
creds = flow.run_local_server(port=3000)

token_data = {
    'token': creds.token,
    'refresh_token': creds.refresh_token,
    'token_uri': creds.token_uri,
    'client_id': creds.client_id,
    'client_secret': creds.client_secret,
    'scopes': list(creds.scopes),
}

with open(TOKEN_FILE, 'w') as f:
    json.dump(token_data, f, indent=2)

print("Authentification OK avec scope ecriture!")
