"""
Upload des documents Word de gouvernance vers le dossier Administratif Google Drive.

Utilise le token OAuth existant (~/.config/mcp-gdrive/token.json).
"""
import os
import sys
import json
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')
DOCX_DIR = Path(__file__).parent / 'docx'
MIME_DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'


def load_creds() -> Credentials:
    with open(TOKEN_FILE) as f:
        data = json.load(f)
    return Credentials(
        token=data['token'],
        refresh_token=data.get('refresh_token'),
        token_uri=data['token_uri'],
        client_id=data['client_id'],
        client_secret=data['client_secret'],
        scopes=data['scopes'],
    )


def find_folder_by_name(service, name: str, parent_id: str | None = None) -> list:
    """Cherche les dossiers portant ce nom (optionnel : dans un parent)."""
    q_parts = [
        "mimeType = 'application/vnd.google-apps.folder'",
        "trashed = false",
        f"name = '{name}'",
    ]
    if parent_id:
        q_parts.append(f"'{parent_id}' in parents")
    res = service.files().list(
        q=" and ".join(q_parts),
        fields="files(id, name, parents)",
        pageSize=50,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    return res.get('files', [])


def find_administratif_folder(service):
    """Trouve le sous-dossier 'Statuts - AG' dans Administratif (cible des uploads)."""
    # Dossier cible : Administratif > Statuts - AG
    TARGET_SUBFOLDER_ID = '1-hJJ53zzJwvMoCLLv10a4tW5WZGN1sxY'
    try:
        meta = service.files().get(fileId=TARGET_SUBFOLDER_ID, fields='id,name',
                                   supportsAllDrives=True).execute()
        return meta
    except Exception as e:
        sys.exit(f"❌ Sous-dossier 'Statuts - AG' introuvable (id {TARGET_SUBFOLDER_ID}) : {e}")


def find_existing_file(service, name: str, folder_id: str):
    """Cherche un fichier existant dans le dossier cible."""
    res = service.files().list(
        q=f"name = '{name}' and '{folder_id}' in parents and trashed = false",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = res.get('files', [])
    return files[0] if files else None


def upload_file(service, path: Path, folder_id: str):
    existing = find_existing_file(service, path.name, folder_id)
    media = MediaFileUpload(str(path), mimetype=MIME_DOCX, resumable=False)
    if existing:
        # Mise à jour du contenu
        updated = service.files().update(
            fileId=existing['id'],
            media_body=media,
            supportsAllDrives=True,
        ).execute()
        return f"🔄 MAJ   {path.name} (id {updated['id']})"
    else:
        metadata = {'name': path.name, 'parents': [folder_id], 'mimeType': MIME_DOCX}
        created = service.files().create(
            body=metadata, media_body=media, fields='id, name',
            supportsAllDrives=True,
        ).execute()
        return f"✅ NEW   {path.name} (id {created['id']})"


def main():
    if not os.path.exists(TOKEN_FILE):
        sys.exit(f"❌ Token Google Drive introuvable : {TOKEN_FILE}")
    if not DOCX_DIR.exists():
        sys.exit(f"❌ Dossier docx introuvable : {DOCX_DIR}")

    creds = load_creds()
    service = build('drive', 'v3', credentials=creds)

    print("🔍 Recherche du dossier 'Administratif' sur Google Drive…")
    folder = find_administratif_folder(service)
    folder_id = folder['id']
    print(f"📁 Dossier cible : {folder.get('name')} (id {folder_id})\n")

    docx_files = sorted(f for f in DOCX_DIR.glob('*.docx') if not f.name.startswith('~$'))
    print(f"📤 Upload de {len(docx_files)} fichier(s)…\n")
    for f in docx_files:
        try:
            print(upload_file(service, f, folder_id))
        except Exception as e:
            print(f"❌ ERREUR sur {f.name} : {e}")

    print(f"\n✅ Terminé. Vérifier : https://drive.google.com/drive/folders/{folder_id}")


if __name__ == '__main__':
    main()
