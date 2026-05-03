"""
Verifie si Administration/Administration est un doublon ou un raccourci/shortcut.
"""
import sys, os, json
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


creds = load_creds()
svc = build('drive', 'v3', credentials=creds)

# IDs des deux dossiers Administration trouvés
OUTER_ID = "13DluZWg4v6ZVBZuUJ1YITPa6n0eTT75_"  # Administration racine
INNER_ID = "1WX7a5X-XbF59mwnRJDOvgHWGT5J1V-l-"  # Administration/Administration

# Récupérer chaque dossier avec metadata complète
for label, fid in [("OUTER", OUTER_ID), ("INNER", INNER_ID)]:
    f = svc.files().get(
        fileId=fid,
        fields="id, name, mimeType, parents, shortcutDetails, createdTime, modifiedTime",
        supportsAllDrives=True,
    ).execute()
    print(f"\n=== {label} ===")
    print(f"  ID         : {f['id']}")
    print(f"  Nom        : {f['name']}")
    print(f"  MimeType   : {f['mimeType']}")
    print(f"  Parents    : {f.get('parents', [])}")
    print(f"  Cree       : {f.get('createdTime')}")
    print(f"  Modifie    : {f.get('modifiedTime')}")
    if f.get('shortcutDetails'):
        print(f"  >>> SHORTCUT vers : {f['shortcutDetails']}")
    else:
        print(f"  Type       : Dossier reel (pas un raccourci)")

# Compter les enfants pour savoir si c'est vraiment un doublon de contenu
def count_and_compare(svc, fid):
    res = svc.files().list(
        q=f"'{fid}' in parents and trashed=false",
        fields="files(id, name, size, modifiedTime)",
        pageSize=500, supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    return res.get('files', [])

print("\n\n=== Comparaison du contenu ===")
outer_children = count_and_compare(svc, OUTER_ID)
inner_children = count_and_compare(svc, INNER_ID)

# Les enfants directs de OUTER incluent INNER lui-même + d'autres
outer_ids_by_name = {c['name']: c['id'] for c in outer_children}
inner_ids_by_name = {c['name']: c['id'] for c in inner_children}

print(f"OUTER ({OUTER_ID}) : {len(outer_children)} enfants directs")
print(f"INNER ({INNER_ID}) : {len(inner_children)} enfants directs")

# Verifier si les fichiers de INNER ont les memes IDs que les fichiers de OUTER
common_names = set(outer_ids_by_name.keys()) & set(inner_ids_by_name.keys())
same_id = sum(1 for n in common_names if outer_ids_by_name[n] == inner_ids_by_name[n])
diff_id = sum(1 for n in common_names if outer_ids_by_name[n] != inner_ids_by_name[n])

print(f"\nFichiers en commun par nom : {len(common_names)}")
print(f"  - meme ID (= meme fichier physique) : {same_id}")
print(f"  - ID different (= duplication) : {diff_id}")

if same_id > 0 and diff_id == 0:
    print("\n>>> CONCLUSION : INNER contient les memes fichiers que OUTER (memes IDs)")
    print(">>> = Affichage parallele du meme contenu (multi-parent ou raccourci dossier)")
    print(">>> = SUPPRIMER INNER ne supprime PAS les fichiers physiques")
elif diff_id > 0:
    print("\n>>> CONCLUSION : INNER contient des copies distinctes (IDs differents)")
    print(">>> = Vraie duplication, supprimer INNER supprime ses fichiers")
