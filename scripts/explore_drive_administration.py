"""
Explore l'arborescence des dossiers Administration / Administratif sur Drive
et liste tous les documents pour proposer un classement unifie.
"""
import sys
import os
import json
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')


def load_creds():
    with open(TOKEN_FILE) as f:
        data = json.load(f)
    creds = Credentials(
        token=data['token'], refresh_token=data.get('refresh_token'),
        token_uri=data['token_uri'], client_id=data['client_id'],
        client_secret=data['client_secret'], scopes=data['scopes'],
    )
    if not creds.valid and creds.refresh_token:
        creds.refresh(Request())
    return creds


def find_folders_by_name(svc, name):
    res = svc.files().list(
        q=f"mimeType='application/vnd.google-apps.folder' and trashed=false and name='{name}'",
        fields="files(id, name, parents)",
        pageSize=20, supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    return res.get('files', [])


def get_file(svc, fid):
    return svc.files().get(fileId=fid, fields="id, name, parents, mimeType",
                            supportsAllDrives=True).execute()


def list_children(svc, parent_id):
    """Liste tous les enfants directs d'un dossier."""
    res = svc.files().list(
        q=f"'{parent_id}' in parents and trashed=false",
        fields="files(id, name, mimeType, modifiedTime, size)",
        pageSize=500, supportsAllDrives=True, includeItemsFromAllDrives=True,
        orderBy="folder, name",
    ).execute()
    return res.get('files', [])


def walk(svc, folder_id, prefix=""):
    """Parcours récursif d'un dossier."""
    children = list_children(svc, folder_id)
    for c in children:
        is_folder = c['mimeType'] == 'application/vnd.google-apps.folder'
        icon = "📁" if is_folder else "📄"
        size = ""
        if not is_folder and c.get('size'):
            size = f" ({int(c['size']) // 1024} KB)"
        print(f"{prefix}{icon} {c['name']}{size}")
        if is_folder:
            walk(svc, c['id'], prefix + "    ")


def get_path(svc, fid):
    """Construit le chemin d'un dossier en remontant les parents."""
    parts = []
    current = fid
    while current:
        try:
            f = get_file(svc, current)
            parts.append(f['name'])
            parents = f.get('parents', [])
            if not parents:
                break
            current = parents[0]
        except Exception:
            break
    return " > ".join(reversed(parts))


def main():
    creds = load_creds()
    svc = build('drive', 'v3', credentials=creds)

    for name in ("Administration", "Administratif"):
        print(f"\n{'='*70}")
        print(f"  RECHERCHE DOSSIER : '{name}'")
        print('='*70)
        folders = find_folders_by_name(svc, name)
        for f in folders:
            path = get_path(svc, f['id'])
            print(f"\n📁 {f['name']}  (id {f['id']})")
            print(f"   Chemin : {path}")
            print(f"   Contenu :")
            walk(svc, f['id'], "      ")


if __name__ == "__main__":
    main()
