"""
Genere les 2 recus fiscaux PDF pour Charles JANIER-DUBRY (dons 2026)
+ cree un brouillon Gmail pour les envoyer en piece jointe.

Reuse le moteur PDF de send_14_recus_retroactifs.py.

Usage : python scripts/generate_recus_charles.py
"""
import sys, os, json, base64
from pathlib import Path
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders

sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# Reuse le module existant pour la fonction build_pdf
sys.path.insert(0, str(Path(__file__).resolve().parent))
from send_14_recus_retroactifs import build_pdf, nombre_en_lettres, ASSO

REPO = Path(__file__).resolve().parent.parent
TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')
LOGO_PATH = REPO / "Logo_Monique_Festival.png"
OUTPUT = REPO / "officiel/dons/charles_dons_2026"
OUTPUT.mkdir(parents=True, exist_ok=True)

# 2 dons a generer
DONS = [
    {
        "numero_recu": "25042026-MF-DON-001",
        "donateur_prenom": "Charles",
        "donateur_nom": "JANIER-DUBRY",
        "donateur_adresse": "54 chemin de Valentin",
        "donateur_cp": "25000",
        "donateur_ville": "Besançon",
        "donateur_pays": "France",
        "montant": 1000,
        "montant_lettres": "mille euros",
        "date_don": "20/04/2026",
        "moyen_paiement": "Virement bancaire",
        "date_emission": "25/04/2026",
        "article_cgi": "200",
        "taux_reduction": "66",
    },
    {
        "numero_recu": "25042026-MF-DON-002",
        "donateur_prenom": "Charles",
        "donateur_nom": "JANIER-DUBRY",
        "donateur_adresse": "54 chemin de Valentin",
        "donateur_cp": "25000",
        "donateur_ville": "Besançon",
        "donateur_pays": "France",
        "montant": 500,
        "montant_lettres": "cinq cents euros",
        "date_don": "25/04/2026",
        "moyen_paiement": "Abandon de créance (avance de fonds 10/01/2026 au propriétaire La Grange Huguenet)",
        "date_emission": "25/04/2026",
        "article_cgi": "200",
        "taux_reduction": "66",
    },
]

EMAIL_HTML = """\
<html>
<body style="font-family: Helvetica, Arial, sans-serif; font-size: 11pt; color: #222; line-height: 1.55;">
<div style="text-align:center; margin-bottom: 18px;">
  <img src="cid:logo" alt="Monique Festival" style="width: 90px; height: auto;">
</div>

<p>Bonjour Charles,</p>

<p>Faisant suite à votre engagement personnel pour l'association
<strong>Monique Festival</strong>, vous trouverez en pièces jointes vos
<strong>deux reçus fiscaux</strong> au format PDF (Cerfa n° 11580*04) :</p>

<ul>
  <li><strong>Reçu n° 25042026-MF-DON-001</strong> : don de <strong>1 000 €</strong> versé le 20/04/2026 par virement sur le compte Qonto de l'association.</li>
  <li><strong>Reçu n° 25042026-MF-DON-002</strong> : don de <strong>500 €</strong> au titre de l'abandon de créance signé le 25/04/2026, correspondant à l'avance de fonds versée le 10/01/2026 au propriétaire de La Grange Huguenet pour le compte de l'association (acompte sur la location du festival).</li>
</ul>

<p>Ces deux dons, d'un montant total de <strong>1 500 €</strong>, ouvrent
droit à une réduction d'impôt sur le revenu de <strong>66 %</strong>
(soit <strong>990 €</strong>) à reporter dans votre déclaration des
revenus 2026 (à effectuer en 2027).</p>

<p>Le Bureau de l'association a accepté ces deux dons à l'unanimité par
délibération du 25 avril 2026 (PV ci-joint en annexe interne).</p>

<p>Au nom de l'ensemble du Bureau, je tiens à exprimer notre sincère
gratitude pour votre soutien financier qui contribue de manière concrète
à la réalisation de la première édition du festival.</p>

<p>Très cordialement,</p>

<p>
<strong>Judith Laithier</strong><br>
Présidente de l'association Monique Festival<br>
<a href="mailto:info@monique-festival.fr">info@monique-festival.fr</a>
</p>

<p style="font-size: 9pt; color: #888; margin-top: 24px; border-top: 1px dashed #bbb; padding-top: 8px;">
Monique Festival — Festival émergent et pluridisciplinaire à Besançon<br>
28, 29 et 30 août 2026 — La Grange Huguenet<br>
Instagram : <a href="https://instagram.com/monique.festival">@monique.festival</a>
</p>
</body>
</html>
"""


def load_creds():
    with open(TOKEN_FILE) as f:
        d = json.load(f)
    creds = Credentials(token=d['token'], refresh_token=d.get('refresh_token'),
        token_uri=d['token_uri'], client_id=d['client_id'],
        client_secret=d['client_secret'], scopes=d['scopes'])
    if not creds.valid: creds.refresh(Request())
    return creds


def build_email(destinataire, objet, html_body, pdf_paths, logo_path):
    message = MIMEMultipart('related')
    message['To'] = destinataire
    message['Subject'] = objet
    alt = MIMEMultipart('alternative')
    message.attach(alt)
    alt.attach(MIMEText(html_body, 'html', 'utf-8'))

    if logo_path.exists():
        with open(logo_path, 'rb') as f:
            img = MIMEImage(f.read(), _subtype='png')
            img.add_header('Content-ID', '<logo>')
            img.add_header('Content-Disposition', 'inline', filename=logo_path.name)
            message.attach(img)

    for pdf_path in pdf_paths:
        with open(pdf_path, 'rb') as f:
            attach = MIMEBase('application', 'pdf')
            attach.set_payload(f.read())
        encoders.encode_base64(attach)
        attach.add_header('Content-Disposition', f'attachment; filename="{pdf_path.name}"')
        message.attach(attach)

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    return {'raw': raw}


def main():
    print("=== Generation des 2 recus fiscaux Charles ===\n")

    pdf_paths = []
    for don in DONS:
        slug = f"RF_{don['numero_recu']}_{don['donateur_prenom']}_{don['donateur_nom']}.pdf"
        pdf_path = OUTPUT / slug.replace(" ", "_")
        build_pdf(don, pdf_path)
        print(f"  PDF : {pdf_path.name} ({pdf_path.stat().st_size // 1024} KB)")
        pdf_paths.append(pdf_path)

    # Brouillon Gmail
    print("\n=== Creation brouillon Gmail ===\n")
    creds = load_creds()
    gmail = build('gmail', 'v1', credentials=creds)
    msg = build_email(
        destinataire="charles@janier-dubry.fr",
        objet="Monique Festival — Vos 2 reçus fiscaux pour vos dons 2026 (1 500 €)",
        html_body=EMAIL_HTML,
        pdf_paths=pdf_paths,
        logo_path=LOGO_PATH,
    )
    draft = gmail.users().drafts().create(userId='me', body={'message': msg}).execute()

    print(f"  Brouillon cree, Draft ID : {draft['id']}")
    print(f"  To : charles@janier-dubry.fr")
    print(f"  PJ : {len(pdf_paths)} reçus ({sum(p.stat().st_size for p in pdf_paths) // 1024} KB)")

    print("\n=== Termine ===")
    print(f"\nRecus PDF : {OUTPUT.relative_to(REPO)}")
    print(f"Brouillon Gmail pret a envoyer.")


if __name__ == "__main__":
    main()
