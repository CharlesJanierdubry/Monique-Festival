"""
Crée 6 BROUILLONS Gmail (1 par artiste étranger) avec la Note de défraiement EN
en pièce jointe, prêts à envoyer depuis info@monique-festival.fr.

Pas d'envoi — vérification dans Gmail Brouillons.
"""
import sys, os, json, base64, unicodedata, re
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

TOKEN = os.path.expanduser('~/.config/mcp-gdrive/token.json')
REPO = Path(__file__).resolve().parent.parent
PJ = REPO / "budget" / "exports" / "Note_defraiement_artiste_etranger_EN.docx"

FROM = 'Monique Festival <info@monique-festival.fr>'
SUBJECT = "Monique Festival 2026 — Travel reimbursement (please read & prepare)"

ARTISTES = [
    {"name": "Marc",      "email": "m.trembovelski@gmail.com",  "country": "UK",       "group": "Europe Cellists"},
    {"name": "Sophie",    "email": "sophie.ehling@gmail.com",   "country": "Netherlands","group": "Europe Cellists"},
    {"name": "Jakub",     "email": "jakub.wycislik1@gmail.com", "country": "Poland",   "group": "Europe Cellists"},
    {"name": "Beatriz",   "email": "beamucaco@gmail.com",       "country": "Portugal", "group": "Europe Cellists"},
    {"name": "Lisa",      "email": "willemslisa1@gmail.com",    "country": "Belgium",  "group": "À travers la fenêtre des heures"},
    {"name": "Pierre",    "email": "pnbcolombat@gmail.com",     "country": "—",        "group": "À travers la fenêtre des heures"},
]

CORPS_TEMPLATE = """Hi {name},

We're really happy to welcome you at Monique Festival on 28-30 August 2026 in Besançon, France, with {group}.

Since you're traveling from {country}, we want to anticipate the travel reimbursement so that you can book your ticket at the best price.

Please find attached our Travel Reimbursement Agreement, which sets out:

- The mode of transport we cover (train, long-distance bus, plane only when no rail alternative) — in line with our environmental charter
- The cap (€200 per person, refund of actual costs)
- The documents you need to send us
- The expected timeline

A few key points:

1. Book your ticket BEFORE 15 JUNE 2026 — Eurostar / flight prices rise sharply afterwards.
2. We do not reimburse personal car (consistent with our ecological commitment).
3. Accommodation and all meals are provided in kind on-site at the festival venue (Grange Huguenet) — you only need to organize your transport.
4. Send us your IBAN + ticket proof as soon as you've booked.

To sign the agreement, you can either:
- Print, sign, scan and email back, OR
- Reply confirming "I agree to the Travel Reimbursement Agreement" and we'll consider it accepted (we'll send a DocuSign version if you prefer formal signature).

For any question, just reply to this email.

Looking forward to having you with us!

Best regards,

Charles JANIER-DUBRY
Administrative Lead — Monique Festival
info@monique-festival.fr
+33 7 87 43 87 85
"""


def ascii_filename(name):
    nfkd = unicodedata.normalize('NFKD', name)
    no_accents = ''.join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r'\s+', '_', no_accents)


def main():
    if not PJ.exists():
        sys.exit(f"PJ introuvable : {PJ}")
    with open(TOKEN) as f:
        d = json.load(f)
    creds = Credentials(token=d['token'], refresh_token=d.get('refresh_token'),
        token_uri=d['token_uri'], client_id=d['client_id'],
        client_secret=d['client_secret'], scopes=d['scopes'])
    if not creds.valid:
        creds.refresh(Request())
    gmail = build('gmail', 'v1', credentials=creds)

    print(f"PJ : {PJ.name} ({PJ.stat().st_size//1024} KB)\n")

    for a in ARTISTES:
        msg = MIMEMultipart()
        msg['From'] = FROM
        msg['To'] = a['email']
        msg['Subject'] = SUBJECT
        body = CORPS_TEMPLATE.format(**a)
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        # PJ
        with open(PJ, 'rb') as f:
            attach = MIMEBase('application',
                'vnd.openxmlformats-officedocument.wordprocessingml.document')
            attach.set_payload(f.read())
        encoders.encode_base64(attach)
        attach.add_header('Content-Disposition',
            f'attachment; filename="{ascii_filename(PJ.name)}"')
        msg.attach(attach)

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')
        draft = gmail.users().drafts().create(
            userId='me', body={'message': {'raw': raw}}).execute()
        print(f"  OK  {a['name']:<10} → {a['email']:<35}  draft={draft['id']}")

    print(f"\n{len(ARTISTES)} brouillons créés. Vérifie dans Gmail (festivalljd@) > Brouillons.")
    print(f"Avant envoi, contrôle que le champ 'De' affiche info@monique-festival.fr.")


if __name__ == "__main__":
    main()
