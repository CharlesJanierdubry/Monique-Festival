"""
Cree un BROUILLON dans Gmail (festivalljd@) en reponse au mail d'Aurelien
de HelloAsso, expedie depuis l'alias info@monique-festival.fr (admin du
compte HelloAsso). Pas d'envoi.

Pieces jointes :
 - officiel/prefecture/Recepisse_modification_24-04-2026.pdf
 - gouvernance/docx/PV_AGE_adoption_statuts 19-04-2026 - Signe.pdf
 - gouvernance/docx/Statuts_Monique_Festival_19-04-2026_Signe.pdf
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

TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')
REPO = Path(__file__).resolve().parent.parent

PJS = [
    REPO / "officiel/prefecture/Recepisse_modification_24-04-2026.pdf",
    REPO / "gouvernance/docx/PV_AGE_adoption_statuts 19-04-2026 - Signé.pdf",
    REPO / "gouvernance/docx/Statuts_Monique_Festival_19-04-2026_Signé.pdf",
]

FROM = "Monique Festival <info@monique-festival.fr>"
TO = "contact@helloasso.org"
SUBJECT = "Re: HelloAsso - Demande de modification de raison sociale — Association W251011013 (JD Production → Monique Festival)"

CORPS = """Bonjour Aurélien,

Merci pour votre retour, et désolé pour la confusion — je vous renvoie ma demande depuis l'adresse administratrice enregistrée sur notre compte HelloAsso (info@monique-festival.fr).

Pour répondre à votre question : il s'agit bien de notre compte de production sur HelloAsso.com (et non du compte de test Sandbox).

Je reformule ci-dessous la demande dans son intégralité.

== Coordonnées de l'association ==

- Nouvelle dénomination : MONIQUE FESTIVAL
- Ancienne dénomination : JD Production
- N° RNA : W251011013 (inchangé)
- SIREN : 991 055 450 (inchangé)
- SIRET : 991 055 450 00019 (inchangé)
- Adresse du siège : 54 chemin de Valentin, 25000 Besançon
- Représentante légale : Judith LAITHIER, Présidente
- Trésorier : Wadih CORMIER

== Nouvel objet social ==

(Article 2 des statuts révisés — pour mise à jour des reçus fiscaux générés automatiquement par HelloAsso)

« Organisation annuelle d'un festival pluridisciplinaire à Besançon et en Bourgogne-Franche-Comté, mêlant musique classique, musiques actuelles, théâtre et arts vivants ; soutien à la création artistique émergente ; actions d'éducation artistique et culturelle, de médiation et de transmission, notamment vers les publics éloignés de la pratique ; valorisation du territoire ; promotion de la diffusion de la culture, à titre non lucratif et d'intérêt général. »

== Contexte juridique ==

La modification a été votée à l'unanimité par notre Assemblée Générale Extraordinaire du 19 avril 2026 et enregistrée par la Sous-Préfecture de Pontarlier le 24 avril 2026 (dossier n° A-6-2KO5BUDQM).

== Pièces jointes ==

1. Récépissé préfectoral de modification du 24 avril 2026
2. Procès-verbal de l'AGE du 19 avril 2026 (signé)
3. Statuts révisés signés en date du 19 avril 2026

== Demande ==

Pourriez-vous, dans notre espace HelloAsso :

1. Mettre à jour la raison sociale de l'association : « JD Production » → « MONIQUE FESTIVAL »
2. Mettre à jour l'objet social tel qu'indiqué ci-dessus, afin que les futurs reçus fiscaux générés automatiquement portent la mention conforme à nos statuts révisés.

Je vous remercie par avance et reste à votre disposition pour tout complément ou justificatif additionnel.

Bien cordialement,

Charles JANIER-DUBRY
Chef du Pôle Administratif — Monique Festival
Mandaté par la Présidente Judith LAITHIER
charles@janier-dubry.fr
"""

def main():
    with open(TOKEN_FILE) as f:
        d = json.load(f)
    creds = Credentials(token=d['token'], refresh_token=d.get('refresh_token'),
        token_uri=d['token_uri'], client_id=d['client_id'],
        client_secret=d['client_secret'], scopes=d['scopes'])
    if not creds.valid:
        creds.refresh(Request())
    gmail = build('gmail', 'v1', credentials=creds)

    print("Vérification des PJ :")
    for p in PJS:
        if not p.exists():
            sys.exit(f"  MANQUE : {p}")
        print(f"  OK  {p.name} ({p.stat().st_size // 1024} KB)")

    msg = MIMEMultipart()
    msg['From'] = FROM
    msg['To'] = TO
    msg['Subject'] = SUBJECT
    msg.attach(MIMEText(CORPS, 'plain', 'utf-8'))

    def ascii_filename(name: str) -> str:
        # Retire accents, remplace espaces par _ : compatible 100% clients mail
        nfkd = unicodedata.normalize('NFKD', name)
        no_accents = ''.join(c for c in nfkd if not unicodedata.combining(c))
        return re.sub(r'\s+', '_', no_accents)

    for p in PJS:
        with open(p, 'rb') as f:
            attach = MIMEBase('application', 'pdf')
            attach.set_payload(f.read())
        encoders.encode_base64(attach)
        safe = ascii_filename(p.name)
        attach.add_header('Content-Disposition', f'attachment; filename="{safe}"')
        msg.attach(attach)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')
    draft = gmail.users().drafts().create(
        userId='me', body={'message': {'raw': raw}}
    ).execute()

    print(f"\nBrouillon créé.")
    print(f"  Draft ID  : {draft['id']}")
    print(f"  Message ID: {draft['message']['id']}")
    print(f"  De        : {FROM}")
    print(f"  À         : {TO}")
    print(f"  Objet     : {SUBJECT}")
    print(f"  PJ        : {len(PJS)} fichiers")
    print(f"\nVérifie dans Gmail (festivalljd@) > Brouillons.")
    print(f"Avant envoi, contrôle bien que le champ 'De' affiche info@monique-festival.fr.")

if __name__ == "__main__":
    main()
