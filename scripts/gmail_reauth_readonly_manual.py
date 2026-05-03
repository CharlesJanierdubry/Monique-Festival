"""
Re-authentification Google OAuth (gmail.modify) — version SANS serveur local.

À utiliser si gmail_reauth_readonly.py échoue (navigateur ne s'ouvre pas,
port 3000 occupé, ou tout autre souci avec InstalledAppFlow.run_local_server).

Étapes :
1. Le script affiche une URL d'authentification
2. Tu l'ouvres manuellement dans n'importe quel navigateur
3. Tu te connectes avec festivalljd@gmail.com et tu autorises les permissions
4. Google te redirige vers une page (qui peut afficher une erreur localhost — c'est normal)
5. Tu copies l'URL complète de cette page (qui contient un "code=...") et tu la colles dans le terminal
6. Le token est mis à jour

Usage :
    python scripts/gmail_reauth_readonly_manual.py
"""
from google_auth_oauthlib.flow import Flow
import json
import os
import sys
from urllib.parse import urlparse, parse_qs

sys.stdout.reconfigure(encoding='utf-8')

SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.compose',
    'https://www.googleapis.com/auth/gmail.modify',
]

CREDS_FILE = os.path.expanduser('~/.config/mcp-gdrive/gcp-oauth.keys.json')
TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')

# Redirect URI installed-app standard (out of band)
REDIRECT = 'urn:ietf:wg:oauth:2.0:oob'

flow = Flow.from_client_secrets_file(
    CREDS_FILE,
    scopes=SCOPES,
    redirect_uri=REDIRECT,
)

auth_url, _ = flow.authorization_url(
    access_type='offline',
    prompt='consent',
    include_granted_scopes='true',
)

print("=" * 70)
print("AUTHENTIFICATION OAUTH GOOGLE — MODE MANUEL")
print("=" * 70)
print()
print("ÉTAPE 1 : ouvre cette URL dans n'importe quel navigateur :")
print()
print(auth_url)
print()
print("ÉTAPE 2 : connecte-toi avec festivalljd@gmail.com")
print("ÉTAPE 3 : autorise les permissions demandées")
print("ÉTAPE 4 : Google va t'afficher un CODE à copier ici (ou une URL)")
print()

response = input("Colle le code OU l'URL complète reçue : ").strip()

# Extraire le code si URL complète
if response.startswith('http'):
    parsed = urlparse(response)
    code = parse_qs(parsed.query).get('code', [None])[0]
    if not code:
        # Peut-être un fragment
        code = parse_qs(parsed.fragment).get('code', [None])[0]
    if not code:
        sys.exit("ERREUR : aucun code trouvé dans l'URL collée")
else:
    code = response

print(f"\nCode reçu : {code[:20]}...")
print("Échange du code contre un token...")

flow.fetch_token(code=code)
creds = flow.credentials

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
