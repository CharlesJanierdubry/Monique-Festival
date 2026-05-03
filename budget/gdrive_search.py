"""
Recherche dans Google Drive les documents liés aux artistes du festival.
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json
import os

TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')

with open(TOKEN_FILE) as f:
    token_data = json.load(f)

creds = Credentials(
    token=token_data['token'],
    refresh_token=token_data['refresh_token'],
    token_uri=token_data['token_uri'],
    client_id=token_data['client_id'],
    client_secret=token_data['client_secret'],
    scopes=token_data['scopes']
)

service = build('drive', 'v3', credentials=creds)

# Recherches par mots-clés liés aux artistes
queries = [
    "artiste",
    "programmation",
    "cachet",
    "lineup",
    "prestation",
    "intervenant",
    "budget artiste",
    "Monique festival",
]

seen = set()
for q in queries:
    results = service.files().list(
        q=f"fullText contains '{q}' and trashed = false",
        pageSize=20,
        fields="files(id, name, mimeType, modifiedTime, parents)",
        orderBy="modifiedTime desc"
    ).execute()

    for f in results.get('files', []):
        if f['id'] not in seen:
            seen.add(f['id'])
            mime = f['mimeType'].split('.')[-1] if 'google-apps' in f['mimeType'] else f['mimeType'].split('/')[-1]
            print(f"{f['modifiedTime'][:10]}  {mime:<15} {f['name']}")
            print(f"   ID: {f['id']}")

print(f"\n--- Total: {len(seen)} fichiers trouvés ---")
