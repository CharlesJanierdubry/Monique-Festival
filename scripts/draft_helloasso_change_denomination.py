"""
Cree un BROUILLON dans la boite Gmail (pas d'envoi) pour la demande de
changement de denomination HelloAsso.

Telecharge d'abord le recepisse prefectoral depuis Drive si absent en local.

Pieces jointes :
 - Recepisse_modification_24-04-2026.pdf (depuis Drive)
 - PV_AGE_adoption_statuts 19-04-2026 - Signé.pdf (local)
 - Statuts_Monique_Festival_19-04-2026_Signé.pdf (local)

Usage :
    python scripts/draft_helloasso_change_denomination.py

Le brouillon apparait dans Gmail (web/app) > "Brouillons" pour revue et envoi.
"""
import sys, os, json, io, base64
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')
REPO = Path(__file__).resolve().parent.parent

# Documents en local
PV_AGE_LOCAL = REPO / "gouvernance/docx/PV_AGE_adoption_statuts 19-04-2026 - Signé.pdf"
STATUTS_LOCAL = REPO / "gouvernance/docx/Statuts_Monique_Festival_19-04-2026_Signé.pdf"

# Recepisse a telecharger depuis Drive
RECEPISSE_LOCAL = REPO / "officiel/prefecture/Recepisse_modification_24-04-2026.pdf"

# Brouillon
DESTINATAIRE = "contact@helloasso.org"
OBJET = "Demande de modification de raison sociale — Association W251011013 (JD Production → Monique Festival)"

CORPS = """Bonjour,

Je vous écris pour solliciter la mise à jour de la dénomination sociale de notre association sur HelloAsso, suite à une modification statutaire validée par notre Préfecture.

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

1. Le récépissé préfectoral de modification du 24 avril 2026
2. Le procès-verbal de l'AGE du 19 avril 2026 (signé)
3. Les statuts révisés signés en date du 19 avril 2026

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


def load_creds():
    with open(TOKEN_FILE) as f:
        d = json.load(f)
    if 'https://www.googleapis.com/auth/gmail.compose' not in d.get('scopes', []):
        sys.exit(
            "ERREUR : scope gmail.compose manquant.\n"
            "Lance d'abord : python scripts/gmail_auth.py\n"
            "et autorise les 4 scopes (Drive, Spreadsheets, Gmail send, Gmail compose)."
        )
    creds = Credentials(token=d['token'], refresh_token=d.get('refresh_token'),
        token_uri=d['token_uri'], client_id=d['client_id'],
        client_secret=d['client_secret'], scopes=d['scopes'])
    if not creds.valid: creds.refresh(Request())
    return creds


def find_file_by_name(svc, name, parent_id=None):
    name_esc = name.replace("'", "\\'")
    q = f"name='{name_esc}' and trashed=false"
    if parent_id:
        q += f" and '{parent_id}' in parents"
    res = svc.files().list(q=q, fields="files(id, name, mimeType, parents)",
        pageSize=10, supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    return res.get('files', [])


def download_drive(svc, file_id, dest: Path):
    dest.parent.mkdir(parents=True, exist_ok=True)
    req = svc.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.FileIO(str(dest), 'wb')
    downloader = MediaIoBaseDownload(fh, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.close()
    return dest


def main():
    creds = load_creds()
    drive = build('drive', 'v3', credentials=creds)
    gmail = build('gmail', 'v1', credentials=creds)

    # 1. Recuperer le recepisse depuis Drive si absent en local
    if not RECEPISSE_LOCAL.exists():
        print("Telechargement du recepisse depuis Drive...")
        candidates = find_file_by_name(drive, "Recepisse_modification_24-04-2026.pdf")
        if not candidates:
            sys.exit("ERREUR : Recepisse_modification_24-04-2026.pdf introuvable sur Drive")
        download_drive(drive, candidates[0]['id'], RECEPISSE_LOCAL)
        print(f"  OK ({RECEPISSE_LOCAL.stat().st_size // 1024} KB)")
    else:
        print(f"Recepisse en local : {RECEPISSE_LOCAL.relative_to(REPO)}")

    # 2. Verifier les autres pieces
    pjs = [
        ("Recepisse préfectoral", RECEPISSE_LOCAL),
        ("PV AGE 19-04-2026 signé", PV_AGE_LOCAL),
        ("Statuts révisés signés", STATUTS_LOCAL),
    ]
    print("\nPieces jointes :")
    for label, p in pjs:
        if p.exists():
            print(f"  OK  {label} : {p.name} ({p.stat().st_size // 1024} KB)")
        else:
            sys.exit(f"ERREUR : {label} introuvable : {p}")

    # 3. Construire le mail
    print("\nConstruction du brouillon...")
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

    # 4. Creer un brouillon (PAS d'envoi)
    draft = gmail.users().drafts().create(
        userId='me',
        body={'message': {'raw': raw}},
    ).execute()

    print(f"\nOK - Brouillon cree dans Gmail")
    print(f"   Draft ID : {draft['id']}")
    print(f"   Message ID : {draft['message']['id']}")
    print(f"\nVerifie ta boite Gmail : Brouillons / Drafts")
    print(f"  Le brouillon contient :")
    print(f"  - Destinataire : {DESTINATAIRE}")
    print(f"  - Objet : {OBJET}")
    print(f"  - 3 PJ ({sum(p.stat().st_size for _, p in pjs) // 1024} KB total)")
    print(f"\nPour envoyer : ouvre le brouillon dans Gmail et clique 'Envoyer'.")


if __name__ == "__main__":
    main()
