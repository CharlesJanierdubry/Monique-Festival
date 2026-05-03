"""
Deplace le dossier unifie Administratif/ (racine MONIQUE FESTIVAL) vers
Monique Festival #1 - 28, 29, 30 août 2026/.

Etapes :
 1. Deplacer le Mecenat existant (avec contenu) vers le 2_Finances/Mecenat
    vide du dossier unifie (apres suppression du Mecenat empty)
 2. Renommer l'ancien Administratif vide de Festival 2026 en _old_Administratif_a_nettoyer
    (impossible a supprimer car contient un dossier orphelin)
 3. Deplacer le dossier unifie Administratif/ vers Festival 2026/
"""
import sys, os, json
sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

with open(os.path.expanduser('~/.config/mcp-gdrive/token.json')) as f:
    d = json.load(f)
creds = Credentials(token=d['token'], refresh_token=d.get('refresh_token'),
    token_uri=d['token_uri'], client_id=d['client_id'],
    client_secret=d['client_secret'], scopes=d['scopes'])
if not creds.valid: creds.refresh(Request())
svc = build('drive', 'v3', credentials=creds)

ROOT = '1v_VnHyEHYxW8LeRbi7TQL4NL5JGjTzn-'
ADMIN_UNIFIED = '1CDUX2_T6etLUbNeKcxyyXJ_zK4LP1cfU'  # Administratif unifie a la racine
MECENAT_FULL = '113FzPABrUrljtSmmznSCYHvBSJcPBpuB'   # Mecenat avec ses 13 fichiers (actuellement dans Festival 2026/Administratif)


def find_child(parent, name):
    name_esc = name.replace("'", "\\'")
    res = svc.files().list(q=f"'{parent}' in parents and name='{name_esc}' and trashed=false",
        fields='files(id, mimeType, parents)', supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    files = res.get('files', [])
    return files[0] if files else None

def list_children(parent):
    res = svc.files().list(q=f"'{parent}' in parents and trashed=false",
        fields='files(id, name, mimeType, parents)', supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    return res.get('files', [])

def move(file_id, new_parent, current_parents):
    return svc.files().update(fileId=file_id,
        addParents=new_parent, removeParents=','.join(current_parents),
        fields='id, name, parents', supportsAllDrives=True).execute()

def trash(fid):
    return svc.files().update(fileId=fid, body={'trashed': True}, supportsAllDrives=True).execute()

def rename(fid, new_name):
    return svc.files().update(fileId=fid, body={'name': new_name},
        fields='id, name', supportsAllDrives=True).execute()


# === ETAPE 1 : Trouver Festival 2026 et son Administratif ===
festivals = svc.files().list(
    q=f"'{ROOT}' in parents and mimeType='application/vnd.google-apps.folder' and name contains 'Monique Festival #1' and trashed=false",
    fields='files(id, name)', supportsAllDrives=True, includeItemsFromAllDrives=True).execute()['files']
festival = festivals[0]
print(f"Festival 2026 : {festival['name']} (id {festival['id']})")

old_admin = find_child(festival['id'], 'Administratif')
print(f"Ancien Administratif dans Festival 2026 : id {old_admin['id'] if old_admin else 'NONE'}")

# === ETAPE 2 : Trouver 2_Finances/Mecenat (vide) dans unifie ===
finances = find_child(ADMIN_UNIFIED, '2_Finances')
mecenat_empty = find_child(finances['id'], 'Mecenat')
print(f"Empty Mecenat in unified 2_Finances/ : id {mecenat_empty['id'] if mecenat_empty else 'NONE'}")

# Verifier qu'il est bien vide
if mecenat_empty:
    kids = list_children(mecenat_empty['id'])
    if not kids:
        print("  -> vide, suppression")
        trash(mecenat_empty['id'])
    else:
        print(f"  -> non-vide ({len(kids)} elem) : skip suppression")

# === ETAPE 3 : Deplacer Mecenat (avec 13 fichiers) vers 2_Finances/ ===
mecenat_meta = svc.files().get(fileId=MECENAT_FULL,
    fields='id, name, parents', supportsAllDrives=True).execute()
print(f"\nDeplacement Mecenat ({mecenat_meta['name']}) :")
print(f"  Source parent : {mecenat_meta['parents']}")
print(f"  Destination   : 2_Finances/ (id {finances['id']})")
move(MECENAT_FULL, finances['id'], mecenat_meta['parents'])
print("  OK")

# === ETAPE 4 : Renommer l'ancien Administratif Festival 2026 ===
if old_admin:
    remaining = list_children(old_admin['id'])
    print(f"\nAncien Administratif Festival 2026 contient : {[c['name'] for c in remaining]}")
    if remaining:
        # Renomme pour eviter le doublon avec le nouveau qu'on va deplacer
        rename(old_admin['id'], '_Administratif_orphelin_a_nettoyer')
        print(f"  Renomme en '_Administratif_orphelin_a_nettoyer' (contient dossier orphelin non supprimable)")
    else:
        # Vide, on peut trash
        trash(old_admin['id'])
        print(f"  Vide -> corbeille")

# === ETAPE 5 : Deplacer Administratif unifie vers Festival 2026 ===
unified_meta = svc.files().get(fileId=ADMIN_UNIFIED,
    fields='id, name, parents', supportsAllDrives=True).execute()
print(f"\nDeplacement Administratif unifie ({unified_meta['name']}) :")
print(f"  Source parent : {unified_meta['parents']}")
print(f"  Destination   : Festival 2026 (id {festival['id']})")
move(ADMIN_UNIFIED, festival['id'], unified_meta['parents'])
print("  OK")

print(f"\n=== Termine ===")
print(f"Nouveau lien : https://drive.google.com/drive/folders/{ADMIN_UNIFIED}")
