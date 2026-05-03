"""
Correctif post-migration : recupere les fichiers restants dans les anciens dossiers
et les place dans la bonne cible (les noms avec accents/apostrophes typo n'ont pas matche).
"""
import sys, os, json
sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')
ROOT_MONIQUE = "1v_VnHyEHYxW8LeRbi7TQL4NL5JGjTzn-"
ADMIN_OUTER = "13DluZWg4v6ZVBZuUJ1YITPa6n0eTT75_"
ADMIN_INNER = "1WX7a5X-XbF59mwnRJDOvgHWGT5J1V-l-"
ADMINISTRATIF_NEW = "1CDUX2_T6etLUbNeKcxyyXJ_zK4LP1cfU"


def load_creds():
    with open(TOKEN_FILE) as f:
        d = json.load(f)
    creds = Credentials(
        token=d['token'], refresh_token=d.get('refresh_token'),
        token_uri=d['token_uri'], client_id=d['client_id'],
        client_secret=d['client_secret'], scopes=d['scopes'],
    )
    if not creds.valid and creds.refresh_token:
        creds.refresh(Request())
    return creds


def list_children(svc, parent_id):
    res = svc.files().list(
        q=f"'{parent_id}' in parents and trashed=false",
        fields="files(id, name, mimeType, parents)",
        pageSize=500, supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    return res.get('files', [])


def find_child(svc, parent_id, name):
    name_esc = name.replace("'", "\\'")
    res = svc.files().list(
        q=f"'{parent_id}' in parents and name='{name_esc}' and trashed=false",
        fields="files(id, name, mimeType)", pageSize=5,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    files = res.get('files', [])
    return files[0] if files else None


def ensure_folder(svc, parent_id, name):
    existing = find_child(svc, parent_id, name)
    if existing and existing['mimeType'] == 'application/vnd.google-apps.folder':
        return existing['id']
    created = svc.files().create(
        body={'name': name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]},
        fields='id', supportsAllDrives=True,
    ).execute()
    return created['id']


def ensure_path(svc, root_id, rel_path):
    current = root_id
    for part in rel_path.split('/'):
        if part:
            current = ensure_folder(svc, current, part)
    return current


def move_file(svc, file_id, new_parent_id, current_parents):
    return svc.files().update(
        fileId=file_id,
        addParents=new_parent_id,
        removeParents=",".join(current_parents),
        fields='id, name', supportsAllDrives=True,
    ).execute()


def trash(svc, fid):
    return svc.files().update(fileId=fid, body={'trashed': True}, supportsAllDrives=True).execute()


def main():
    svc = build('drive', 'v3', credentials=load_creds())

    print("=== Correctif migration : fichiers restants ===\n")

    # 1) Lister ce qui reste dans Administration/Administration
    print("Restant dans ADMIN_INNER :")
    inner = list_children(svc, ADMIN_INNER)
    for f in inner:
        print(f"  - {f['name']} ({f['mimeType']})")

    # 2) Mapping par mots-cles dans le nom
    def classify(name):
        nl = name.lower()
        # Bancaire / Comptes
        if "ouverture compte bancaire" in nl:
            return "2_Finances/Banque_Qonto/PV_ouverture_compte"
        # AG constitutive 2025
        if "assembl" in nl and "constitutive" in nl:
            return "0_Gouvernance/AG_et_Bureau/2025_AG_constitutive"
        # Charte ethique
        if "charte" in nl and ("éth" in nl or "eth" in nl):
            return "0_Gouvernance/Charte_ethique"
        # MD Recapitulatif (recap declaration prefecture)
        if "recapitulatif" in nl or "récapitulatif" in nl:
            return "1_Admin_legale/Prefecture/2026-04_modif_denomination"
        return None

    # 3) Re-classer ce qui reste dans INNER
    print("\nMigration des fichiers restants dans ADMIN_INNER :")
    for f in inner:
        if f['mimeType'] == 'application/vnd.google-apps.folder':
            continue
        dest_path = classify(f['name'])
        if dest_path is None:
            print(f"  WARN  {f['name']} : pas de classification automatique")
            continue
        target = ensure_path(svc, ADMINISTRATIF_NEW, dest_path)
        try:
            move_file(svc, f['id'], target, f.get('parents', []))
            print(f"  MOVE  {f['name']} -> {dest_path}")
        except Exception as e:
            print(f"  ERR   {f['name']} : {e}")

    # 4) Trouver et migrer les fichiers restants dans Adm2026/Administratif
    festival_folders = svc.files().list(
        q=f"'{ROOT_MONIQUE}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id,name)", supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute().get('files', [])
    festival_2026 = next((x for x in festival_folders if x['name'].startswith("Monique Festival #1")), None)
    if festival_2026:
        old_admin_2026 = find_child(svc, festival_2026['id'], "Administratif")
        if old_admin_2026:
            print(f"\nRestant dans Adm2026/Administratif :")
            adm2026_kids = list_children(svc, old_admin_2026['id'])
            for f in adm2026_kids:
                print(f"  - {f['name']} ({f['mimeType']})")
                if f['mimeType'] == 'application/vnd.google-apps.folder':
                    continue
                dest_path = classify(f['name'])
                if dest_path is None:
                    print(f"    WARN : pas de classification")
                    continue
                target = ensure_path(svc, ADMINISTRATIF_NEW, dest_path)
                try:
                    move_file(svc, f['id'], target, f.get('parents', []))
                    print(f"    MOVE -> {dest_path}")
                except Exception as e:
                    print(f"    ERR : {e}")

    # 5) Re-tenter la suppression des dossiers vides
    print("\nNettoyage final des dossiers vides :")
    folders_to_check = []
    if festival_2026:
        old_admin_2026 = find_child(svc, festival_2026['id'], "Administratif")
        if old_admin_2026:
            for sub in list_children(svc, old_admin_2026['id']):
                if sub['mimeType'] == 'application/vnd.google-apps.folder':
                    folders_to_check.append((sub['id'], f"Adm2026/{sub['name']}"))
            folders_to_check.append((old_admin_2026['id'], "Adm2026/Administratif"))

    folders_to_check.append((ADMIN_INNER, "Administration/Administration"))
    folders_to_check.append((ADMIN_OUTER, "Administration"))

    for fid, label in folders_to_check:
        try:
            kids = list_children(svc, fid)
            if not kids:
                trash(svc, fid)
                print(f"  TRASH {label} (vide)")
            else:
                names = [k['name'] for k in kids]
                print(f"  KEEP  {label} : {len(kids)} elem -> {names}")
        except Exception as e:
            print(f"  ERR   {label} : {e}")


if __name__ == "__main__":
    main()
