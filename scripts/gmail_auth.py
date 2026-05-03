"""
Re-authentification Google OAuth avec ajout du scope Gmail (send).

À lancer UNE FOIS pour autoriser l'envoi d'emails via l'API Gmail.

Usage :
    python scripts/gmail_auth.py

Résultat :
    Met à jour ~/.config/mcp-gdrive/token.json avec le scope gmail.send ajouté.
    Un navigateur s'ouvre pour la validation OAuth Google (compte associatif).
"""
from google_auth_oauthlib.flow import InstalledAppFlow
import json
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.compose',  # Permet aussi de créer des brouillons
]

CREDS_FILE = os.path.expanduser('~/.config/mcp-gdrive/gcp-oauth.keys.json')
TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')

print("Lancement de l'authentification OAuth Google avec scope Gmail...")
print("Un navigateur va s'ouvrir — connecte-toi avec le compte Google")
print("qui doit envoyer les mails (probablement charles@janier-dubry.fr)")
print()

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

print()
print(f"OK ! Token sauvegarde dans {TOKEN_FILE}")
print(f"Scopes actives : {', '.join(creds.scopes)}")
