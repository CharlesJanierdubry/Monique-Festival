"""
Script d'authentification Google Drive.
Lance ce script une seule fois pour obtenir le refresh token.
Il ouvrira ton navigateur pour te connecter.
"""
from google_auth_oauthlib.flow import InstalledAppFlow
import json
import os

SCOPES = [
    'https://www.googleapis.com/auth/drive.readonly',
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

print(f"\nAuthentification réussie!")
print(f"Token sauvegardé dans: {TOKEN_FILE}")
print(f"Refresh token: {creds.refresh_token[:20]}...")
