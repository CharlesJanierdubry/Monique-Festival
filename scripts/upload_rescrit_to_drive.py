"""
Upload du dossier de rescrit fiscal sur Google Drive
(Administratif > 1_Admin_legale > Rescrit_fiscal_DDFIP).

Structure créée :
- Rescrit_fiscal_DDFIP/
  - Dossier_rescrit_complet_a_signer.pdf  (PDF unique 68 pages)
  - Sources_Word/                          (sources éditables)
    - 9 fichiers DOCX

Usage : python scripts/upload_rescrit_to_drive.py
"""
import sys, json, os
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

TOKEN = os.path.expanduser('~/.config/mcp-gdrive/token.json')
REPO = Path(__file__).resolve().parent.parent

RESCRIT_DRIVE_ID = '1oelsJUvJJJD4TQNCxY32qXy7RPxrBenM'  # Rescrit_fiscal_DDFIP

LOCAL_PDF = REPO / "officiel/rescrit_fiscal/Dossier_rescrit_complet_a_signer.pdf"
LOCAL_DOCX_DIR = REPO / "officiel/rescrit_fiscal/exports_a_signer"

DOCX_FILES = [
    "00_Sommaire.docx",
    "01_Courrier_demande.docx",
    "02_Memoire_argumentaire.docx",
    "08_Liste_dirigeants.docx",
    "09_Annexe_7_Budget.docx",
    "10_Annexe_8_Artistes.docx",
    "11_Annexe_9_Tarifs.docx",
    "12_Annexe_10_Charte.docx",
    "13_Annexe_11_Presentation.docx",
]


def get_or_create_folder(svc, name, parent_id):
    """Retourne l'id du dossier (le crée s'il n'existe pas)."""
    q = (f"name='{name}' and '{parent_id}' in parents and "
         f"mimeType='application/vnd.google-apps.folder' and trashed=false")
    r = svc.files().list(q=q, fields='files(id,name)', pageSize=10).execute()
    files = r.get('files', [])
    if files:
        return files[0]['id']
    meta = {'name': name, 'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_id]}
    f = svc.files().create(body=meta, fields='id').execute()
    return f['id']


def upload_or_replace(svc, local_path: Path, parent_id, mime_type):
    """Upload (remplace si existe déjà)."""
    name = local_path.name
    name_esc = name.replace("'", "\\'")
    q = (f"name='{name_esc}' and '{parent_id}' in parents and trashed=false")
    r = svc.files().list(q=q, fields='files(id,name)', pageSize=5).execute()
    existing = r.get('files', [])
    media = MediaFileUpload(str(local_path), mimetype=mime_type, resumable=True)
    if existing:
        fid = existing[0]['id']
        svc.files().update(fileId=fid, media_body=media).execute()
        return fid, 'remplacé'
    else:
        meta = {'name': name, 'parents': [parent_id]}
        f = svc.files().create(body=meta, media_body=media, fields='id').execute()
        return f['id'], 'créé'


def main():
    with open(TOKEN) as f:
        d = json.load(f)
    creds = Credentials(token=d['token'], refresh_token=d.get('refresh_token'),
        token_uri=d['token_uri'], client_id=d['client_id'],
        client_secret=d['client_secret'], scopes=d['scopes'])
    if not creds.valid:
        creds.refresh(Request())
    svc = build('drive', 'v3', credentials=creds)

    # 1. PDF unique au niveau racine du dossier Rescrit
    if not LOCAL_PDF.exists():
        sys.exit(f"ERREUR : {LOCAL_PDF} introuvable.")
    print(f"Upload PDF unique :")
    fid, status = upload_or_replace(svc, LOCAL_PDF, RESCRIT_DRIVE_ID,
                                     'application/pdf')
    size_mb = LOCAL_PDF.stat().st_size / (1024 * 1024)
    print(f"  {status}  {LOCAL_PDF.name}  ({size_mb:.2f} Mo)  id={fid}")

    # 2. Sous-dossier Sources_Word + DOCX
    print(f"\nUpload sources Word :")
    sources_id = get_or_create_folder(svc, 'Sources_Word', RESCRIT_DRIVE_ID)
    print(f"  Sous-dossier 'Sources_Word' : id={sources_id}")
    for name in DOCX_FILES:
        p = LOCAL_DOCX_DIR / name
        if not p.exists():
            print(f"  SKIP  {name} (introuvable)")
            continue
        fid, status = upload_or_replace(svc, p, sources_id,
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        kb = p.stat().st_size // 1024
        print(f"  {status}  {name}  ({kb} KB)")

    print(f"\nOK — dossier accessible : "
          f"https://drive.google.com/drive/folders/{RESCRIT_DRIVE_ID}")


if __name__ == "__main__":
    main()
