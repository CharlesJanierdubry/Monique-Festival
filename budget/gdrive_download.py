"""
Télécharge les documents clés du Drive pour les artistes et le programme.
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json
import os
import io

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

docs_to_fetch = {
    # Google Sheets → export CSV
    "Fiche de contacts artistes": {
        "id": "10Gpv4GMBFpNROmTQ5PPUWEX6nadT3QpinSTA875_Cok",
        "type": "spreadsheet",
    },
    "Budget estimatif cout artiste v2": {
        "id": "1SQU9yUHSEqIYdL5ABco7rmjBAXEryCQw",
        "type": "xlsx",
    },
    "Planning recettes et couts": {
        "id": "1eAH90_LNIgaWY9wBFFvyGHreU96LzCYO",
        "type": "xlsx",
    },
    # Google Docs → export text
    "Programme": {
        "id": "1JCgIjaj9i-FrKwfEvQW2uw2rvPaeQuOcjDyf9aPgATU",
        "type": "document",
    },
    "Suivi 2 scenes": {
        "id": "1Iq9yfWHrhSkHXy53KAYyT4QGmJNx76OQXa31vCqIIX4",
        "type": "document",
    },
    "Edition 1 nos actions": {
        "id": "1ITtQkBlP6D8GVlgOA69SmhSsKGh_taRYXwrRnxHFI-g",
        "type": "document",
    },
    "Fiche technique": {
        "id": "1-U3_mAM4JtsS3a8WtnzdUxtE-mUjdeJfSHnhsdWGGXA",
        "type": "document",
    },
    "Compte rendu 30-03": {
        "id": "1Iz_cYQeInA6HmdcKy0ADDLEjvCFETTeTBVAgbQEfpaE",
        "type": "document",
    },
    "Bio Jakub": {
        "id": "1Ja5i22-QAXQbIjQMFavZwEZKXwZyc_hQKOD7am5Id0s",
        "type": "document",
    },
}

OUTPUT_DIR = "budget/gdrive_exports"
os.makedirs(OUTPUT_DIR, exist_ok=True)

for name, info in docs_to_fetch.items():
    print(f"Downloading: {name}...")
    try:
        if info["type"] == "spreadsheet":
            # Google Sheets - get all sheet names first
            sheets_service = build('sheets', 'v4', credentials=creds)
            spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=info["id"]).execute()
            sheets = spreadsheet.get('sheets', [])

            for sheet in sheets:
                sheet_name = sheet['properties']['title']
                safe_name = sheet_name.replace('/', '-').replace(' ', '_')
                result = sheets_service.spreadsheets().values().get(
                    spreadsheetId=info["id"],
                    range=f"'{sheet_name}'"
                ).execute()
                values = result.get('values', [])

                filename = f"{OUTPUT_DIR}/{name.replace(' ', '_')}_{safe_name}.csv"
                with open(filename, 'w', encoding='utf-8') as f:
                    for row in values:
                        f.write('\t'.join(str(c) for c in row) + '\n')
                print(f"  -> {filename} ({len(values)} rows)")

        elif info["type"] == "document":
            # Google Docs → plain text
            content = service.files().export(
                fileId=info["id"],
                mimeType='text/plain'
            ).execute()
            filename = f"{OUTPUT_DIR}/{name.replace(' ', '_')}.txt"
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(content.decode('utf-8'))
            print(f"  -> {filename}")

        elif info["type"] == "xlsx":
            # Download as xlsx
            content = service.files().get_media(fileId=info["id"]).execute()
            filename = f"{OUTPUT_DIR}/{name.replace(' ', '_')}.xlsx"
            with open(filename, 'wb') as f:
                f.write(content)
            print(f"  -> {filename}")

    except Exception as e:
        print(f"  ERREUR: {e}")

print("\nDone!")
