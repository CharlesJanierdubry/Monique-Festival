"""
Upload des documents officiels (Préfecture, Partenaires, Rescrit fiscal) vers Google Drive.

Cible : sous-dossier "Statuts - AG" dans Administratif.
Utilise le token OAuth existant (~/.config/mcp-gdrive/token.json).

Upload :
- officiel/prefecture/*.pdf, *.docx  (preuve dépôt, mandat signé, courrier)
- officiel/partenaires/*.pdf, *.docx (courriers partenaires — si générés)
- officiel/rescrit_fiscal/*.pdf, *.docx (rescrit — si généré)

Re-run safe : si un fichier porte déjà ce nom dans Drive, il est mis à jour (pas dupliqué).
"""
import os
import sys
import json
import mimetypes
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')
ROOT = Path(__file__).parent
STATUTS_AG_FOLDER_ID = '1-hJJ53zzJwvMoCLLv10a4tW5WZGN1sxY'

MIME_PDF = 'application/pdf'
MIME_DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

EXTENSIONS_MIME = {
    '.pdf': MIME_PDF,
    '.docx': MIME_DOCX,
}

# Sous-dossiers à parcourir
SUBFOLDERS = ['prefecture', 'partenaires', 'rescrit_fiscal']


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


def find_or_create_subfolder(service, name: str, parent_id: str) -> str:
    """Cherche un sous-dossier par nom, le crée sinon. Retourne son id."""
    q = (
        f"mimeType = 'application/vnd.google-apps.folder' "
        f"and trashed = false "
        f"and name = '{name}' "
        f"and '{parent_id}' in parents"
    )
    res = service.files().list(
        q=q,
        fields="files(id, name)",
        pageSize=10,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = res.get('files', [])
    if files:
        print(f"📁 Sous-dossier existant : {name} (id {files[0]['id']})")
        return files[0]['id']
    # Creer
    metadata = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id],
    }
    created = service.files().create(
        body=metadata, fields='id, name', supportsAllDrives=True
    ).execute()
    print(f"🆕 Sous-dossier cree : {name} (id {created['id']})")
    return created['id']


def find_existing_file(service, name: str, folder_id: str):
    # Le nom peut contenir des apostrophes — on les echappe
    name_escaped = name.replace("'", "\\'")
    res = service.files().list(
        q=f"name = '{name_escaped}' and '{folder_id}' in parents and trashed = false",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = res.get('files', [])
    return files[0] if files else None


def upload_file(service, path: Path, folder_id: str) -> str:
    mime = EXTENSIONS_MIME.get(path.suffix.lower())
    if not mime:
        return f"⏭  IGNORE {path.name} (extension non supportee)"

    existing = find_existing_file(service, path.name, folder_id)
    media = MediaFileUpload(str(path), mimetype=mime, resumable=False)

    if existing:
        updated = service.files().update(
            fileId=existing['id'],
            media_body=media,
            supportsAllDrives=True,
        ).execute()
        return f"🔄 MAJ    {path.name} (id {updated['id']})"
    else:
        metadata = {'name': path.name, 'parents': [folder_id], 'mimeType': mime}
        created = service.files().create(
            body=metadata, media_body=media, fields='id, name',
            supportsAllDrives=True,
        ).execute()
        return f"✅ NEW    {path.name} (id {created['id']})"


def main():
    if not os.path.exists(TOKEN_FILE):
        sys.exit(f"❌ Token Google Drive introuvable : {TOKEN_FILE}\n"
                 f"   Relancer : python budget/gdrive_auth_write.py")

    creds = load_creds()
    service = build('drive', 'v3', credentials=creds)

    print(f"🔍 Dossier parent : 'Statuts - AG' (id {STATUTS_AG_FOLDER_ID})\n")

    # On cree un sous-dossier "Declaration Prefecture 04-2026" dans Statuts - AG
    target_folder_id = find_or_create_subfolder(
        service, 'Declaration Prefecture 04-2026', STATUTS_AG_FOLDER_ID
    )
    print()

    total_uploaded = 0
    for sub in SUBFOLDERS:
        local_dir = ROOT / sub
        if not local_dir.exists():
            continue

        # Sous-sous-dossier par categorie
        cat_folder_id = find_or_create_subfolder(
            service, sub.capitalize(), target_folder_id
        )

        files = []
        for ext in EXTENSIONS_MIME.keys():
            files.extend(sorted(local_dir.glob(f'*{ext}')))

        if not files:
            print(f"   (aucun fichier PDF/DOCX dans {sub})\n")
            continue

        print(f"📤 Upload de {len(files)} fichier(s) dans {sub}...")
        for f in files:
            try:
                print(f"   {upload_file(service, f, cat_folder_id)}")
                total_uploaded += 1
            except Exception as e:
                print(f"   ❌ ERREUR sur {f.name} : {e}")
        print()

    print(f"\n✅ Termine. {total_uploaded} fichier(s) uploade(s).")
    print(f"🔗 Verifier : https://drive.google.com/drive/folders/{target_folder_id}")


if __name__ == '__main__':
    main()
