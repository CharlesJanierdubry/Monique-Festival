"""
Re-authentification Google OAuth — ajoute le scope gmail.modify
qui permet de LIRE les emails reçus (et marquer comme lu / déplacer).

À lancer UNE FOIS pour autoriser la lecture des emails (recherche de devis,
analyse des PJ, classement automatique).

Le scope `gmail.modify` est plus large que `gmail.readonly` car il permet
aussi de déplacer/marquer les emails — utile pour classer les devis.

Usage :
    python scripts/gmail_reauth_readonly.py

Le navigateur va s'ouvrir, valide les nouvelles permissions Google sur le compte
festivalljd@gmail.com (le même compte que pour Drive et Gmail send).
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
    'https://www.googleapis.com/auth/gmail.compose',
    'https://www.googleapis.com/auth/gmail.modify',  # NOUVEAU : lire + marquer + déplacer
]

CREDS_FILE = os.path.expanduser('~/.config/mcp-gdrive/gcp-oauth.keys.json')
TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')

print("Lancement de l'authentification OAuth Google avec scope Gmail.modify...")
print("Un navigateur va s'ouvrir — connecte-toi avec festivalljd@gmail.com")
print("(le compte qui a déjà l'accès Drive + Gmail send).")
print()
print("Permissions demandées :")
for s in SCOPES:
    print(f"  - {s}")
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
print(f"OK ! Token mis à jour : {TOKEN_FILE}")
print(f"Scopes actifs : {', '.join(creds.scopes)}")
print()
print("Tu peux maintenant lancer la recherche des devis dans Gmail.")
