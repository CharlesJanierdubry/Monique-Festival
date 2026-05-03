"""
Envoie un email TEST avec un reçu fiscal en pièce jointe via l'API Gmail.

Prérequis : avoir lancé scripts/gmail_auth.py une fois pour autoriser
            le scope gmail.send sur le token OAuth Google.

Usage :
    python scripts/send_test_recu_fiscal.py

Destinataire : charles@janier-dubry.fr
Pièce jointe : officiel/dons/recus_emis/TEST_RF_Charles_Janier-Dubry.docx
"""
import sys
import os
import json
import base64
import mimetypes
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')

# ======== Paramètres du mail test ========
DESTINATAIRE = "charles@janier-dubry.fr"
OBJET = "[TEST] Monique Festival — Exemple de reçu fiscal"
PIECE_JOINTE = Path("officiel/dons/recus_emis/TEST_RF_Charles_Janier-Dubry.docx")

CORPS = """Bonjour Charles,

Ceci est un test d'envoi de reçu fiscal pour valider le modèle
que Monique Festival utilisera pour les 14 donateurs du
crowdfunding du 20-23 avril 2026.

Tu trouveras en pièce jointe un exemple de reçu fiscal
(format Cerfa n° 11580*04) pour un don fictif de 50 euros.

Points à vérifier :
- Raison sociale : MONIQUE FESTIVAL (anciennement JD Production)
- Objet social : article 2 des statuts révisés du 19/04/2026
- N° d'ordre : TEST-2026-04-24-0001
- Article 200 CGI (particuliers) avec réduction 66 %
- Mentions obligatoires Cerfa présentes

Ce reçu ne doit PAS être utilisé pour une déclaration fiscale
(mention RECU TEST en bas du document).

Si le modèle te convient, il sera dupliqué pour les 14 donateurs
réels une fois le récépissé préfectoral reçu (début mai 2026) et
les paramètres Hello Asso mis à jour.

Mail envoyé automatiquement depuis l'API Gmail via le script
send_test_recu_fiscal.py — preuve que l'automatisation fonctionne.

Judith Laithier
Présidente de l'association Monique Festival
"""


def load_creds():
    """Charge et rafraîchit si besoin le token OAuth."""
    with open(TOKEN_FILE) as f:
        data = json.load(f)

    if 'https://www.googleapis.com/auth/gmail.send' not in data.get('scopes', []):
        sys.exit(
            "ERREUR : le scope gmail.send n'est pas active sur le token.\n"
            "Lancer d'abord : python scripts/gmail_auth.py"
        )

    creds = Credentials(
        token=data['token'],
        refresh_token=data.get('refresh_token'),
        token_uri=data['token_uri'],
        client_id=data['client_id'],
        client_secret=data['client_secret'],
        scopes=data['scopes'],
    )

    if not creds.valid and creds.refresh_token:
        creds.refresh(Request())
        # Rétablir le token rafraîchi dans le fichier
        data['token'] = creds.token
        with open(TOKEN_FILE, 'w') as f:
            json.dump(data, f, indent=2)

    return creds


def build_email(destinataire, objet, corps, piece_jointe: Path):
    message = MIMEMultipart()
    message['To'] = destinataire
    message['Subject'] = objet

    message.attach(MIMEText(corps, 'plain', 'utf-8'))

    if piece_jointe and piece_jointe.exists():
        mime_type, _ = mimetypes.guess_type(str(piece_jointe))
        if not mime_type:
            mime_type = 'application/octet-stream'
        main_type, sub_type = mime_type.split('/', 1)

        with open(piece_jointe, 'rb') as f:
            attachment = MIMEBase(main_type, sub_type)
            attachment.set_payload(f.read())
        encoders.encode_base64(attachment)
        attachment.add_header(
            'Content-Disposition',
            f'attachment; filename="{piece_jointe.name}"'
        )
        message.attach(attachment)
    else:
        print(f"ATTENTION : piece jointe introuvable : {piece_jointe}")

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    return {'raw': raw}


def main():
    print(f"Chargement du token OAuth ({TOKEN_FILE})...")
    creds = load_creds()
    print(f"  Scopes : {', '.join(creds.scopes)}")

    service = build('gmail', 'v1', credentials=creds)

    print(f"\nConstruction du mail :")
    print(f"  A       : {DESTINATAIRE}")
    print(f"  Objet   : {OBJET}")
    print(f"  PJ      : {PIECE_JOINTE} ({'OK' if PIECE_JOINTE.exists() else 'MANQUANTE'})")

    message = build_email(DESTINATAIRE, OBJET, CORPS, PIECE_JOINTE)

    print(f"\nEnvoi en cours...")
    try:
        sent = service.users().messages().send(userId='me', body=message).execute()
        print(f"OK ! Message ID : {sent['id']}")
        print(f"Verifie ta boite {DESTINATAIRE}")
    except Exception as e:
        print(f"ERREUR : {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
