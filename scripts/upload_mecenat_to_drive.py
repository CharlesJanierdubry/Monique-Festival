"""
Upload du dossier local mecenat/ vers Drive Administratif/2_Finances/Mecenat/
"""
import sys, os, json
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')
ADMINISTRATIF_NEW = "1CDUX2_T6etLUbNeKcxyyXJ_zK4LP1cfU"
LOCAL_MECENAT = Path(__file__).resolve().parent.parent / "mecenat"

def load_creds():
    with open(TOKEN_FILE) as f:
        d = json.load(f)
    creds = Credentials(token=d['token'], refresh_token=d.get('refresh_token'),
        token_uri=d['token_uri'], client_id=d['client_id'],
        client_secret=d['client_secret'], scopes=d['scopes'])
    if not creds.valid: creds.refresh(Request())
    return creds

def find_child(svc, parent, name):
    name_esc = name.replace("'", "\\'")
    res = svc.files().list(
        q=f"'{parent}' in parents and name='{name_esc}' and trashed=false",
        fields="files(id, mimeType)",
        pageSize=5, supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    files = res.get('files', [])
    return files[0] if files else None

def ensure_folder(svc, parent, name):
    e = find_child(svc, parent, name)
    if e and e['mimeType'] == 'application/vnd.google-apps.folder':
        return e['id']
    created = svc.files().create(
        body={'name': name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent]},
        fields='id', supportsAllDrives=True,
    ).execute()
    return created['id']

def upload_file(svc, path: Path, parent_id):
    mime = "text/markdown" if path.suffix == ".md" else "application/octet-stream"
    existing = find_child(svc, parent_id, path.name)
    media = MediaFileUpload(str(path), mimetype=mime, resumable=False)
    if existing:
        svc.files().update(fileId=existing['id'], media_body=media, supportsAllDrives=True).execute()
        return "MAJ"
    else:
        svc.files().create(
            body={'name': path.name, 'parents': [parent_id], 'mimeType': mime},
            media_body=media, fields='id', supportsAllDrives=True,
        ).execute()
        return "NEW"

def main():
    svc = build('drive', 'v3', credentials=load_creds())

    # Structure : Administratif/2_Finances/Mecenat/{,Fiches_prospects/}
    finances = ensure_folder(svc, ADMINISTRATIF_NEW, "2_Finances")
    mecenat = ensure_folder(svc, finances, "Mecenat")
    fiches = ensure_folder(svc, mecenat, "Fiches_prospects")

    print(f"Cible : Administratif/2_Finances/Mecenat/")
    print(f"  -> id Mecenat : {mecenat}")
    print(f"  -> id Fiches_prospects : {fiches}\n")

    # Upload racine mecenat (tout sauf le sous-dossier)
    n_root = 0
    for f in sorted(LOCAL_MECENAT.iterdir()):
        if f.is_file():
            status = upload_file(svc, f, mecenat)
            print(f"  {status}  {f.name}")
            n_root += 1

    # Upload Fiches_prospects/
    n_fiches = 0
    for f in sorted((LOCAL_MECENAT / "Fiches_prospects").iterdir()):
        if f.is_file():
            status = upload_file(svc, f, fiches)
            print(f"  {status}  Fiches_prospects/{f.name}")
            n_fiches += 1

    total = n_root + n_fiches
    print(f"\n=== Termine. {total} fichiers uploades ({n_root} racine + {n_fiches} fiches) ===")
    print(f"Lien : https://drive.google.com/drive/folders/{mecenat}")

if __name__ == "__main__":
    main()
