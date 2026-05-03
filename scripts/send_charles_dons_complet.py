"""
Envoie a charles@janier-dubry.fr les 6 documents lies aux dons 2026 :
- 4 documents a signer (DOCX)
- 2 recus fiscaux PDF (deja generes)

+ uploade les memes documents sur Drive dans
  Festival 2026/Administratif/2_Finances/Dons_et_recus_fiscaux/2026/Charles_dons_membres/

Usage : python scripts/send_charles_dons_complet.py
"""
import sys, os, json, base64, mimetypes
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders

sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

REPO = Path(__file__).resolve().parent.parent
TOKEN = os.path.expanduser('~/.config/mcp-gdrive/token.json')
LOGO = REPO / "Logo_Monique_Festival.png"
DOCS_DIR = REPO / "officiel/dons/charles_dons_2026"

# 6 documents a envoyer + uploader
DOCUMENTS = [
    DOCS_DIR / "1_Lettre_don_1000_2026-04-20.docx",
    DOCS_DIR / "2_Note_de_frais_acompte_500.docx",
    DOCS_DIR / "3_Acte_abandon_creance_500.docx",
    DOCS_DIR / "4_PV_Bureau_acceptation_dons_2026-04-25.docx",
    DOCS_DIR / "RF_25042026-MF-DON-001_Charles_JANIER-DUBRY.pdf",
    DOCS_DIR / "RF_25042026-MF-DON-002_Charles_JANIER-DUBRY.pdf",
]

# Drive : trouver Festival 2026 puis sous-dossier cible
ROOT_MONIQUE = "1v_VnHyEHYxW8LeRbi7TQL4NL5JGjTzn-"


EMAIL_HTML = """\
<html>
<body style="font-family: Helvetica, Arial, sans-serif; font-size: 11pt; color: #222; line-height: 1.55;">
<div style="text-align:center; margin-bottom: 18px;">
  <img src="cid:logo" alt="Monique Festival" style="width: 90px; height: auto;">
</div>

<p>Bonjour Charles,</p>

<p>Vous trouverez en pièces jointes les <strong>6 documents</strong> formalisant vos
deux dons 2026 à l'association Monique Festival, pour un montant total de
<strong>1 500 €</strong> (réduction d'impôt de <strong>990 €</strong> sur revenus 2026,
déclaration 2027).</p>

<p><strong>Documents à signer (4)</strong> :</p>
<ol>
  <li><strong>Lettre de don</strong> de 1 000 € (versement du 20/04/2026) — à signer par Charles</li>
  <li><strong>Note de frais</strong> de 500 € (avance de fonds du 10/01/2026 pour acompte location Grange Huguenet) — à signer par Charles</li>
  <li><strong>Acte d'abandon de créance</strong> de 500 € (transformant l'avance en don) — à signer par Charles
    <br><em>Important : recopier la mention manuscrite obligatoire au stylo, sous peine d'invalidité du don au regard du fisc.</em></li>
  <li><strong>PV de Bureau du 25/04/2026</strong> — à signer par Judith LAITHIER (Présidente) et Wadih CORMIER (Trésorier)</li>
</ol>

<p><strong>Reçus fiscaux générés (2)</strong> :</p>
<ol start="5">
  <li><strong>Reçu fiscal n° 25042026-MF-DON-001</strong> — don de 1 000 € du 20/04/2026</li>
  <li><strong>Reçu fiscal n° 25042026-MF-DON-002</strong> — don de 500 € du 25/04/2026 (date de l'abandon de créance)</li>
</ol>

<p>Une fois les 4 documents signés, merci de les scanner et de les déposer
dans <code>Drive &gt; Administratif &gt; 2_Finances &gt; Dons_et_recus_fiscaux &gt;
2026 &gt; Charles_dons_membres</code> (sous-dossier signe/) — la copie non signée
de chaque document est déjà uploadée sur Drive.</p>

<p>Un grand merci pour ce soutien financier qui contribue concrètement à
la première édition du festival.</p>

<p>Très cordialement,</p>

<p>
<strong>Judith Laithier</strong><br>
Présidente de l'association Monique Festival<br>
<a href="mailto:info@monique-festival.fr">info@monique-festival.fr</a>
</p>

<p style="font-size: 9pt; color: #888; margin-top: 24px; border-top: 1px dashed #bbb; padding-top: 8px;">
Monique Festival — Festival émergent et pluridisciplinaire à Besançon<br>
28, 29 et 30 août 2026 — La Grange Huguenet
</p>
</body>
</html>
"""


def load_creds():
    with open(TOKEN) as f:
        d = json.load(f)
    creds = Credentials(token=d['token'], refresh_token=d.get('refresh_token'),
        token_uri=d['token_uri'], client_id=d['client_id'],
        client_secret=d['client_secret'], scopes=d['scopes'])
    if not creds.valid: creds.refresh(Request())
    return creds


def build_msg(destinataire, objet, html, files, logo):
    m = MIMEMultipart('related')
    m['To'] = destinataire
    m['Subject'] = objet
    alt = MIMEMultipart('alternative')
    m.attach(alt)
    alt.attach(MIMEText(html, 'html', 'utf-8'))

    if logo.exists():
        with open(logo, 'rb') as f:
            img = MIMEImage(f.read(), _subtype='png')
            img.add_header('Content-ID', '<logo>')
            img.add_header('Content-Disposition', 'inline', filename=logo.name)
            m.attach(img)

    for fp in files:
        if not fp.exists():
            continue
        mime, _ = mimetypes.guess_type(str(fp))
        if not mime:
            mime = 'application/octet-stream'
        main_t, sub_t = mime.split('/', 1)
        with open(fp, 'rb') as f:
            attach = MIMEBase(main_t, sub_t)
            attach.set_payload(f.read())
        encoders.encode_base64(attach)
        attach.add_header('Content-Disposition', f'attachment; filename="{fp.name}"')
        m.attach(attach)

    return base64.urlsafe_b64encode(m.as_bytes()).decode('utf-8')


def find_child(svc, parent, name):
    name_esc = name.replace("'", "\\'")
    res = svc.files().list(q=f"'{parent}' in parents and name='{name_esc}' and trashed=false",
        fields='files(id, mimeType)', supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    files = res.get('files', [])
    return files[0] if files else None

def ensure_folder(svc, parent, name):
    e = find_child(svc, parent, name)
    if e and e['mimeType'] == 'application/vnd.google-apps.folder':
        return e['id']
    return svc.files().create(
        body={'name': name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent]},
        fields='id', supportsAllDrives=True).execute()['id']

def ensure_path(svc, root, path):
    cur = root
    for part in path.split('/'):
        if part:
            cur = ensure_folder(svc, cur, part)
    return cur

def upload(svc, fp, parent_id):
    mime, _ = mimetypes.guess_type(str(fp))
    if not mime:
        mime = 'application/octet-stream'
    existing = find_child(svc, parent_id, fp.name)
    media = MediaFileUpload(str(fp), mimetype=mime, resumable=False)
    if existing:
        return svc.files().update(fileId=existing['id'], media_body=media, supportsAllDrives=True).execute()
    return svc.files().create(
        body={'name': fp.name, 'parents': [parent_id], 'mimeType': mime},
        media_body=media, fields='id, name', supportsAllDrives=True).execute()


def main():
    creds = load_creds()
    gmail = build('gmail', 'v1', credentials=creds)
    drive = build('drive', 'v3', credentials=creds)

    # Verifier que tous les fichiers existent
    print("Verification des 6 documents :")
    for f in DOCUMENTS:
        if f.exists():
            print(f"  OK  {f.name} ({f.stat().st_size // 1024} KB)")
        else:
            sys.exit(f"ERREUR : fichier manquant {f}")

    # 1. ENVOI EMAIL (pas brouillon, vrai envoi)
    print(f"\nEnvoi email a charles@janier-dubry.fr...")
    raw = build_msg(
        destinataire="charles@janier-dubry.fr",
        objet="Monique Festival — Documents pour vos 2 dons 2026 (1 500 €) — à signer + reçus fiscaux",
        html=EMAIL_HTML,
        files=DOCUMENTS,
        logo=LOGO,
    )
    sent = gmail.users().messages().send(userId='me', body={'raw': raw}).execute()
    print(f"  OK  Message ID : {sent['id']}")

    # 2. UPLOAD DRIVE
    print(f"\nUpload Drive...")
    festival = drive.files().list(
        q=f"'{ROOT_MONIQUE}' in parents and mimeType='application/vnd.google-apps.folder' and name contains 'Monique Festival #1' and trashed=false",
        fields='files(id, name)', supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()['files'][0]

    target = ensure_path(drive, festival['id'],
        "Administratif/2_Finances/Dons_et_recus_fiscaux/2026/Charles_dons_membres")
    print(f"  Cible : Festival 2026/Administratif/2_Finances/Dons_et_recus_fiscaux/2026/Charles_dons_membres/")
    print(f"  Folder ID : {target}")

    for f in DOCUMENTS:
        try:
            r = upload(drive, f, target)
            print(f"    OK  {f.name} (id {r['id']})")
        except Exception as e:
            print(f"    ERR {f.name} : {e}")

    print(f"\n=== Termine ===")
    print(f"Email envoye a charles@janier-dubry.fr (Message ID : {sent['id']})")
    print(f"6 documents uploades sur Drive : https://drive.google.com/drive/folders/{target}")


if __name__ == "__main__":
    main()
