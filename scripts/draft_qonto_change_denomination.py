"""
Cree un BROUILLON dans la boite Gmail (pas d'envoi) pour la demande de
changement de denomination Qonto.

Pieces jointes :
 - Recepisse_modification_24-04-2026.pdf
 - PV_AGE_adoption_statuts 19-04-2026 - Signé.pdf
 - Statuts_Monique_Festival_19-04-2026_Signé.pdf

Usage : python scripts/draft_qonto_change_denomination.py

Le brouillon apparait dans Gmail > Brouillons.
"""
import sys, os, json, base64
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

PV_AGE = REPO / "gouvernance/docx/PV_AGE_adoption_statuts 19-04-2026 - Signé.pdf"
STATUTS = REPO / "gouvernance/docx/Statuts_Monique_Festival_19-04-2026_Signé.pdf"
RECEPISSE = REPO / "officiel/prefecture/Recepisse_modification_24-04-2026.pdf"

DESTINATAIRE = "support@qonto.com"
OBJET = "Demande de modification de raison sociale — Compte associatif (JD Production → Monique Festival, SIREN 991 055 450)"

CORPS = """Bonjour,

Je vous contacte pour solliciter la mise à jour de la raison sociale de notre compte associatif Qonto, suite à une modification statutaire validée par notre Préfecture.

== Coordonnées de l'association ==

- Nouvelle dénomination : MONIQUE FESTIVAL
- Ancienne dénomination : JD Production
- N° RNA : W251011013 (inchangé)
- SIREN : 991 055 450 (inchangé)
- SIRET : 991 055 450 00019 (inchangé)
- Adresse du siège : 54 chemin de Valentin, 25000 Besançon
- Représentante légale : Judith LAITHIER, Présidente
- Trésorier : Wadih CORMIER, Membre du Bureau

== Contexte juridique ==

La modification a été votée à l'unanimité par notre Assemblée Générale Extraordinaire du 19 avril 2026 et enregistrée par la Sous-Préfecture de Pontarlier le 24 avril 2026 (dossier n° A-6-2KO5BUDQM).

Le SIREN, le SIRET, le RNA et le RIB/IBAN restent inchangés — seule la dénomination sociale évolue. La personne morale est inchangée.

== Pièces jointes ==

1. Le récépissé préfectoral de modification du 24 avril 2026
2. Le procès-verbal de l'AGE du 19 avril 2026 (signé)
3. Les statuts révisés signés en date du 19 avril 2026

== Demande ==

Pourriez-vous procéder à la mise à jour suivante dans nos paramètres Qonto :

1. Raison sociale du compte : « JD Production » → « MONIQUE FESTIVAL »
2. Nom de l'association affiché sur les futurs RIB, virements et relevés émis depuis le compte
3. Si nécessaire, mise à jour de l'objet social (article 2 des statuts révisés) :

« Organisation annuelle d'un festival pluridisciplinaire à Besançon et en Bourgogne-Franche-Comté, mêlant musique classique, musiques actuelles, théâtre et arts vivants ; soutien à la création artistique émergente ; actions d'éducation artistique et culturelle, de médiation et de transmission, notamment vers les publics éloignés de la pratique ; valorisation du territoire ; promotion de la diffusion de la culture, à titre non lucratif et d'intérêt général. »

Pourriez-vous également me confirmer qu'aucune fermeture/réouverture de compte n'est nécessaire et que les opérations en cours (mandats prélèvement, virements récurrents, dons HelloAsso reversés sur ce compte) ne seront pas affectées par cette mise à jour.

Je reste à votre disposition pour tout justificatif additionnel ou rendez-vous téléphonique.

Bien cordialement,

Charles JANIER-DUBRY
Chef du Pôle Administratif — Monique Festival
Pour le Bureau (Présidente Judith LAITHIER, Trésorier Wadih CORMIER)
charles@janier-dubry.fr
"""


def load_creds():
    with open(TOKEN_FILE) as f:
        d = json.load(f)
    if 'https://www.googleapis.com/auth/gmail.compose' not in d.get('scopes', []):
        sys.exit("ERREUR : scope gmail.compose manquant. Lance d'abord python scripts/gmail_auth.py")
    creds = Credentials(token=d['token'], refresh_token=d.get('refresh_token'),
        token_uri=d['token_uri'], client_id=d['client_id'],
        client_secret=d['client_secret'], scopes=d['scopes'])
    if not creds.valid: creds.refresh(Request())
    return creds


def main():
    pjs = [
        ("Récépissé préfectoral", RECEPISSE),
        ("PV AGE 19-04-2026 signé", PV_AGE),
        ("Statuts révisés signés", STATUTS),
    ]
    print("Pieces jointes :")
    for label, p in pjs:
        if p.exists():
            print(f"  OK  {label} : {p.name} ({p.stat().st_size // 1024} KB)")
        else:
            sys.exit(f"ERREUR : {label} introuvable : {p}")

    msg = MIMEMultipart()
    msg['To'] = DESTINATAIRE
    msg['Subject'] = OBJET
    msg.attach(MIMEText(CORPS, 'plain', 'utf-8'))

    for label, p in pjs:
        with open(p, 'rb') as f:
            attach = MIMEBase('application', 'pdf')
            attach.set_payload(f.read())
        encoders.encode_base64(attach)
        attach.add_header('Content-Disposition', f'attachment; filename="{p.name}"')
        msg.attach(attach)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')

    gmail = build('gmail', 'v1', credentials=load_creds())
    draft = gmail.users().drafts().create(userId='me', body={'message': {'raw': raw}}).execute()

    print(f"\nOK - Brouillon Qonto cree dans Gmail")
    print(f"   Draft ID : {draft['id']}")
    print(f"\nVerifie ta boite Gmail : Brouillons / Drafts")
    print(f"  - Destinataire : {DESTINATAIRE}")
    print(f"  - Objet : {OBJET}")
    print(f"  - 3 PJ ({sum(p.stat().st_size for _, p in pjs) // 1024} KB total)")


if __name__ == "__main__":
    main()
