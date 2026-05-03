"""
Migration de l'arborescence Drive Administration -> Administratif unifie.

Plan :
 1. Creer la nouvelle structure 'Administratif/' a la racine MONIQUE FESTIVAL
 2. Deplacer 47 fichiers (changer parent)
 3. Renommer 2 fichiers (recepisse 2025 + 2026)
 4. Supprimer FRM.pdf (corbeille)
 5. Supprimer les anciens dossiers vides
 6. Logger toutes les operations dans un CSV

Usage : python scripts/migrate_drive_admin.py
"""
import sys
import os
import csv
import json
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')
LOG_DIR = Path(__file__).resolve().parent.parent / "officiel" / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / f"migration_drive_admin_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.csv"

# IDs connus (decouverts a l'exploration)
ROOT_MONIQUE = "1v_VnHyEHYxW8LeRbi7TQL4NL5JGjTzn-"     # 🌟 MONIQUE FESTIVAL 🌟
ADMIN_OUTER = "13DluZWg4v6ZVBZuUJ1YITPa6n0eTT75_"     # Administration (racine, conteneur)
ADMIN_INNER = "1WX7a5X-XbF59mwnRJDOvgHWGT5J1V-l-"     # Administration/Administration (24 fichiers)


# ===== Mapping =====
# Source folder path -> { filename: dest_path }
# Les paths sont relatifs au nouveau "Administratif/"
MAPPINGS = {
    # 24 fichiers de Administration/Administration/
    ADMIN_INNER: {
        "ATTESTATION SUR L'HONNEUR DE DECLARATION DE BUDGET PREVISIONNEL.pdf": "2_Finances/Budgets/Previsionnels",
        "ATTESTATION+DE+DECLARATION+DE+CA+PREVISIONNEL JD productions.docx": "2_Finances/Budgets/Previsionnels",
        "ATTESTATION+DE+DECLARATION+DE+CA+PREVISIONNEL.docx": "2_Finances/Budgets/Previsionnels",
        "cerfa-demande-de-licence-dentrepreneur-de-spectacles-vivants-pour-la-duree-des-representations.pdf": "1_Admin_legale/Licences_spectacle_vivant",
        "Cerfa.pdf": "1_Admin_legale/Cerfas_references",
        "F-ASSO-PV.doc": "9_Archives_obsoletes/JD_Production_2025",
        "FRM.pdf": "__DELETE__",  # Corbeille
        "JD.PROD_Procès Verbal - Ouverture compte bancaire.doc": "2_Finances/Banque_Qonto/PV_ouverture_compte",
        "JD.PROD_Procès Verbal - Ouverture compte bancaire.pdf": "2_Finances/Banque_Qonto/PV_ouverture_compte",
        "JD.PROD_Procès Verbal de l'assemblée constitutive.docx": "0_Gouvernance/AG_et_Bureau/2025_AG_constitutive",
        "JD.PROD_Procès Verbal de l'assemblée constitutive.pdf": "0_Gouvernance/AG_et_Bureau/2025_AG_constitutive",
        "JD.PROD_Procès_Verbal_de_l'assemblée_constitutive - Signé.pdf": "0_Gouvernance/AG_et_Bureau/2025_AG_constitutive",
        "Journal officiel": "1_Admin_legale/Journal_Officiel",
        "LDC.pdf": "1_Admin_legale/Cerfas_references",
        "Mandat de décharge de déclaration": "1_Admin_legale/Mandats",
        "Mandat de décharge de déclaration  - signé.pdf": "1_Admin_legale/Mandats",
        "Mandat de décharge de déclaration -signé.docx": "1_Admin_legale/Mandats",
        "Monique-Insee.pdf": "1_Admin_legale/INSEE_SIRENE",
        "QUESTIONNAIRE PAYS  2025.pdf": "1_Admin_legale/Questionnaires_PAYS",
        "QUESTIONNAIRE PAYS  2025_.docx": "1_Admin_legale/Questionnaires_PAYS",
        # Renommer ce fichier !
        "Recepisse_CR (2).pdf": ("1_Admin_legale/Prefecture/2025_creation", "Recepisse_creation_JD_Production_2025.pdf"),
        "STATUTS JD PRODUCTION": "9_Archives_obsoletes/JD_Production_2025",
        "STATUTS JD PRODUCTION.pdf": "9_Archives_obsoletes/JD_Production_2025",
        "STATUTS_JD_PRODUCTION - Signé.pdf": "9_Archives_obsoletes/JD_Production_2025",
    },
}


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


def find_child_by_name(svc, parent_id, name):
    name_esc = name.replace("'", "\\'")
    res = svc.files().list(
        q=f"'{parent_id}' in parents and name='{name_esc}' and trashed=false",
        fields="files(id, name, mimeType, parents)",
        pageSize=10, supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    files = res.get('files', [])
    return files[0] if files else None


def list_children(svc, parent_id):
    res = svc.files().list(
        q=f"'{parent_id}' in parents and trashed=false",
        fields="files(id, name, mimeType, parents)",
        pageSize=500, supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    return res.get('files', [])


def ensure_folder(svc, parent_id, name):
    """Trouve ou cree un dossier par nom dans un parent donne."""
    existing = find_child_by_name(svc, parent_id, name)
    if existing and existing['mimeType'] == 'application/vnd.google-apps.folder':
        return existing['id'], False
    metadata = {'name': name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
    created = svc.files().create(body=metadata, fields='id,name', supportsAllDrives=True).execute()
    return created['id'], True


def ensure_path(svc, root_id, rel_path):
    """Cree (ou trouve) recursivement les dossiers du chemin relatif."""
    current = root_id
    for part in rel_path.split('/'):
        if not part:
            continue
        current, _ = ensure_folder(svc, current, part)
    return current


def move_file(svc, file_id, new_parent_id, file_meta=None):
    """Change le(s) parent(s) d'un fichier vers new_parent_id."""
    if file_meta is None:
        file_meta = svc.files().get(fileId=file_id, fields='id,parents,name', supportsAllDrives=True).execute()
    old_parents = ",".join(file_meta.get('parents', []))
    return svc.files().update(
        fileId=file_id,
        addParents=new_parent_id,
        removeParents=old_parents,
        fields='id, name, parents',
        supportsAllDrives=True,
    ).execute()


def rename_file(svc, file_id, new_name):
    return svc.files().update(
        fileId=file_id, body={'name': new_name},
        fields='id, name', supportsAllDrives=True,
    ).execute()


def trash_file(svc, file_id):
    return svc.files().update(fileId=file_id, body={'trashed': True}, supportsAllDrives=True).execute()


def folder_is_empty(svc, folder_id):
    children = list_children(svc, folder_id)
    return len(children) == 0


# ============================================================================
# MAIN
# ============================================================================

def main():
    creds = load_creds()
    svc = build('drive', 'v3', credentials=creds)

    log_rows = [["Action", "Source", "Destination", "ID", "Statut"]]

    def log(action, source, dest, fid, status):
        log_rows.append([action, source, dest, fid, status])
        print(f"  {action:<8} | {source[:50]:<50} | {dest[:50]:<50} | {status}")

    print(f"=== Migration Drive Administration -> Administratif unifie ===")
    print(f"Log : {LOG_FILE.relative_to(LOG_FILE.parent.parent.parent)}")

    # =========================================================================
    # ETAPE 1 : Creer la nouvelle structure Administratif/ a la racine
    # =========================================================================
    print(f"\n[1/5] Creation de la nouvelle structure 'Administratif/'...")
    admin_id, _ = ensure_folder(svc, ROOT_MONIQUE, "Administratif")
    log("MKDIR", "(racine)", "Administratif", admin_id, "OK")

    # Pre-creer toutes les categories pour avoir une structure visible meme avant deplacement
    structure = [
        "0_Gouvernance",
        "0_Gouvernance/Statuts/2025_JD_Production_originaux",
        "0_Gouvernance/Statuts/2026_Monique_Festival_revises",
        "0_Gouvernance/AG_et_Bureau/2025_AG_constitutive",
        "0_Gouvernance/AG_et_Bureau/2026-04-19_AGE_revision",
        "0_Gouvernance/Reglement_interieur",
        "0_Gouvernance/Charte_ethique",
        "0_Gouvernance/Delegations",
        "0_Gouvernance/Modeles",
        "1_Admin_legale/Prefecture/2025_creation",
        "1_Admin_legale/Prefecture/2026-04_modif_denomination",
        "1_Admin_legale/INSEE_SIRENE",
        "1_Admin_legale/Journal_Officiel",
        "1_Admin_legale/Licences_spectacle_vivant",
        "1_Admin_legale/Cerfas_references",
        "1_Admin_legale/Mandats",
        "1_Admin_legale/Obligations",
        "1_Admin_legale/Questionnaires_PAYS",
        "1_Admin_legale/Rescrit_fiscal_DDFIP",
        "2_Finances/Banque_Qonto/PV_ouverture_compte",
        "2_Finances/Banque_Qonto/Releves_mensuels",
        "2_Finances/Budgets/Previsionnels",
        "2_Finances/Budgets/Realises",
        "2_Finances/Dons_et_recus_fiscaux/2026/Crowdfunding_20-23_avril",
        "2_Finances/Dons_et_recus_fiscaux/2026/Hello_Asso_auto_post-activation",
        "2_Finances/Mecenat",
        "2_Finances/Subventions",
        "2_Finances/Pieces_comptables",
        "9_Archives_obsoletes/JD_Production_2025",
        "9_Archives_obsoletes/2026_brouillons",
    ]
    for path in structure:
        ensure_path(svc, admin_id, path)
    log("STRUCT", "Administratif/", f"{len(structure)} sous-dossiers", admin_id, "OK")

    # =========================================================================
    # ETAPE 2 : Migrer les fichiers de Administration/Administration/ (INNER)
    # =========================================================================
    print(f"\n[2/5] Migration des 24 fichiers de Administration/Administration/...")
    inner_children = list_children(svc, ADMIN_INNER)
    name_to_dest = MAPPINGS[ADMIN_INNER]

    for child in inner_children:
        if child['mimeType'] == 'application/vnd.google-apps.folder':
            log("SKIP", child['name'], "(dossier - voir suite)", child['id'], "skip")
            continue

        dest = name_to_dest.get(child['name'])
        if dest is None:
            log("WARN", child['name'], "(non mappe)", child['id'], "MAPPING MANQUANT")
            continue

        if dest == "__DELETE__":
            try:
                trash_file(svc, child['id'])
                log("TRASH", child['name'], "(corbeille)", child['id'], "OK")
            except Exception as e:
                log("TRASH", child['name'], "(corbeille)", child['id'], f"ERR: {e}")
            continue

        if isinstance(dest, tuple):
            dest_path, new_name = dest
        else:
            dest_path, new_name = dest, None

        try:
            target_id = ensure_path(svc, admin_id, dest_path)
            move_file(svc, child['id'], target_id, child)
            if new_name:
                rename_file(svc, child['id'], new_name)
                log("RENAME", child['name'], f"{dest_path}/{new_name}", child['id'], "OK")
            else:
                log("MOVE", child['name'], dest_path, child['id'], "OK")
        except HttpError as e:
            log("MOVE", child['name'], dest_path, child['id'], f"ERR: {e}")

    # =========================================================================
    # ETAPE 3 : Migrer Administration/Comptabilite/* + Idees de nom
    # =========================================================================
    print(f"\n[3/5] Migration Comptabilite/ et Idees de nom...")
    outer_children = list_children(svc, ADMIN_OUTER)
    for child in outer_children:
        name = child['name']
        if name == "Administration":
            continue  # On traite INNER separement
        if name == "Comptabilite" or name == "Comptabilité":
            # Migrer les fichiers du dossier vers 2_Finances/Pieces_comptables/
            sub_children = list_children(svc, child['id'])
            target_id = ensure_path(svc, admin_id, "2_Finances/Pieces_comptables")
            for sub in sub_children:
                if sub['mimeType'] == 'application/vnd.google-apps.folder':
                    continue
                try:
                    move_file(svc, sub['id'], target_id, sub)
                    log("MOVE", f"Comptabilité/{sub['name']}", "2_Finances/Pieces_comptables", sub['id'], "OK")
                except Exception as e:
                    log("MOVE", f"Comptabilité/{sub['name']}", "...", sub['id'], f"ERR: {e}")
        elif name == "Idées de nom" or name == "Idees de nom":
            target_id = ensure_path(svc, admin_id, "9_Archives_obsoletes/JD_Production_2025")
            try:
                move_file(svc, child['id'], target_id, child)
                log("MOVE", "Idées de nom", "9_Archives_obsoletes/JD_Production_2025", child['id'], "OK")
            except Exception as e:
                log("MOVE", "Idées de nom", "...", child['id'], f"ERR: {e}")

    # =========================================================================
    # ETAPE 4 : Migrer le contenu de "Monique Festival #1 - 2026/Administratif/"
    # =========================================================================
    print(f"\n[4/5] Migration de Monique Festival #1 - 2026/Administratif/...")

    # Trouver "Monique Festival #1 - 28, 29, 30 août 2026"
    festival_folders = svc.files().list(
        q=f"'{ROOT_MONIQUE}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id,name)", supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute().get('files', [])

    festival_2026 = next((f for f in festival_folders if f['name'].startswith("Monique Festival #1")), None)
    if not festival_2026:
        log("WARN", "(introuvable)", "Monique Festival #1...", "", "SKIP")
    else:
        # Trouver Administratif/ dans Monique Festival #1
        old_admin_2026 = find_child_by_name(svc, festival_2026['id'], "Administratif")
        if old_admin_2026:
            # Sous-dossiers a migrer
            sub_dossiers = list_children(svc, old_admin_2026['id'])

            mapping_2026 = {
                # Statuts - AG/* (fichiers directs)
                "Convocation_AGE.docx": "0_Gouvernance/AG_et_Bureau/2026-04-19_AGE_revision",
                "Detail_taches_par_pole.docx": "0_Gouvernance/Modeles",
                "Feuille de Présence 19-04-2026 - Signé.pdf": "0_Gouvernance/AG_et_Bureau/2026-04-19_AGE_revision",
                "Modele_bulletin_adhesion.docx": "0_Gouvernance/Modeles",
                "Note_explicative_changements_statuts.docx": "0_Gouvernance/AG_et_Bureau/2026-04-19_AGE_revision",
                "PV_AGE_adoption_statuts 19-04-2026 - Signé.pdf": "0_Gouvernance/AG_et_Bureau/2026-04-19_AGE_revision",
                "PV_Bureau_post_AGE19-04-2026_ Signé.pdf": "0_Gouvernance/AG_et_Bureau/2026-04-19_AGE_revision",
                "Reglement_Interieur 19-04-2026_ Signé.pdf": "0_Gouvernance/Reglement_interieur",
                "Statuts_Monique_Festival_19-04-2026_Signé.pdf": "0_Gouvernance/Statuts/2026_Monique_Festival_revises",
                # Racine Administratif/*
                "Charte éthique.pdf": "0_Gouvernance/Charte_ethique",
                "Obligations administratives festival.docx": "1_Admin_legale/Obligations",
                # Renommage
                "Recepisse_MD.pdf": ("1_Admin_legale/Prefecture/2026-04_modif_denomination", "Recepisse_modification_24-04-2026.pdf"),
            }

            archive_2026 = {
                "Feuille_presence_AGE_signé.docx": "9_Archives_obsoletes/2026_brouillons",
                "PV_AGE_adoption_statuts-signé.docx": "9_Archives_obsoletes/2026_brouillons",
                "Reglement_Interieur_Signé.docx": "9_Archives_obsoletes/2026_brouillons",
                "Statuts_Monique_Festival-signé.docx": "9_Archives_obsoletes/2026_brouillons",
            }

            licences_dest = "1_Admin_legale/Licences_spectacle_vivant"
            recus_dest = "2_Finances/Dons_et_recus_fiscaux/2026"
            prefecture_dest = "1_Admin_legale/Prefecture/2026-04_modif_denomination"

            # Parcourir : fichiers et sous-dossiers de old_admin_2026
            for child in sub_dossiers:
                cname = child['name']
                cmime = child['mimeType']

                if cmime != 'application/vnd.google-apps.folder':
                    # Fichier direct dans Administratif (Charte éthique, Obligations, Recepisse_MD)
                    dest = mapping_2026.get(cname)
                    if not dest:
                        log("WARN", f"Adm2026/{cname}", "(non mappe)", child['id'], "SKIP")
                        continue
                    if isinstance(dest, tuple):
                        dest_path, new_name = dest
                    else:
                        dest_path, new_name = dest, None
                    try:
                        target = ensure_path(svc, admin_id, dest_path)
                        move_file(svc, child['id'], target, child)
                        if new_name:
                            rename_file(svc, child['id'], new_name)
                            log("RENAME", f"Adm2026/{cname}", f"{dest_path}/{new_name}", child['id'], "OK")
                        else:
                            log("MOVE", f"Adm2026/{cname}", dest_path, child['id'], "OK")
                    except Exception as e:
                        log("MOVE", f"Adm2026/{cname}", dest_path, child['id'], f"ERR: {e}")

                elif cname == "Statuts - AG":
                    # Migrer les fichiers de Statuts - AG/ + sous-dossiers archive et Declaration Prefecture
                    statuts_children = list_children(svc, child['id'])
                    for sc in statuts_children:
                        sname = sc['name']
                        smime = sc['mimeType']

                        if smime != 'application/vnd.google-apps.folder':
                            dest = mapping_2026.get(sname)
                            if not dest:
                                log("WARN", f"Statuts-AG/{sname}", "(non mappe)", sc['id'], "SKIP")
                                continue
                            try:
                                target = ensure_path(svc, admin_id, dest)
                                move_file(svc, sc['id'], target, sc)
                                log("MOVE", f"Statuts-AG/{sname}", dest, sc['id'], "OK")
                            except Exception as e:
                                log("MOVE", f"Statuts-AG/{sname}", dest, sc['id'], f"ERR: {e}")

                        elif sname == "archive":
                            # Migrer chaque fichier de archive vers 9_Archives_obsoletes/2026_brouillons
                            arch_children = list_children(svc, sc['id'])
                            target = ensure_path(svc, admin_id, "9_Archives_obsoletes/2026_brouillons")
                            for ac in arch_children:
                                if ac['mimeType'] == 'application/vnd.google-apps.folder':
                                    continue
                                try:
                                    move_file(svc, ac['id'], target, ac)
                                    log("MOVE", f"archive/{ac['name']}", "9_Archives_obsoletes/2026_brouillons", ac['id'], "OK")
                                except Exception as e:
                                    log("MOVE", f"archive/{ac['name']}", "...", ac['id'], f"ERR: {e}")

                        elif sname == "Declaration Prefecture 04-2026":
                            # Recuperer Prefecture/ contenu vers prefecture_dest
                            decl_children = list_children(svc, sc['id'])
                            for dc in decl_children:
                                if dc['name'] == "Prefecture":
                                    pref_files = list_children(svc, dc['id'])
                                    target = ensure_path(svc, admin_id, prefecture_dest)
                                    for pf in pref_files:
                                        if pf['mimeType'] == 'application/vnd.google-apps.folder':
                                            continue
                                        try:
                                            move_file(svc, pf['id'], target, pf)
                                            log("MOVE", f"Decl-Pref/{pf['name']}", prefecture_dest, pf['id'], "OK")
                                        except Exception as e:
                                            log("MOVE", f"Decl-Pref/{pf['name']}", "...", pf['id'], f"ERR: {e}")
                                # Partenaires/ et Rescrit_fiscal/ sont vides - on les ignore (suppression en etape 5)

                elif cname == "Licences d'entrepreneurs de spectacle":
                    lic_children = list_children(svc, child['id'])
                    target = ensure_path(svc, admin_id, licences_dest)
                    for lc in lic_children:
                        if lc['mimeType'] == 'application/vnd.google-apps.folder':
                            continue
                        try:
                            move_file(svc, lc['id'], target, lc)
                            log("MOVE", f"Licences/{lc['name']}", licences_dest, lc['id'], "OK")
                        except Exception as e:
                            log("MOVE", f"Licences/{lc['name']}", "...", lc['id'], f"ERR: {e}")

                elif cname.startswith("Reçus fiscaux"):
                    # Deplacer le dossier entier vers 2_Finances/Dons_et_recus_fiscaux/2026/
                    target = ensure_path(svc, admin_id, recus_dest)
                    try:
                        move_file(svc, child['id'], target, child)
                        log("MOVE", f"Adm2026/{cname}/", recus_dest, child['id'], "OK (dossier entier)")
                    except Exception as e:
                        log("MOVE", f"Adm2026/{cname}/", recus_dest, child['id'], f"ERR: {e}")

                else:
                    log("WARN", f"Adm2026/{cname}", "(non gere)", child['id'], "SKIP")

    # =========================================================================
    # ETAPE 5 : Nettoyer les anciens dossiers vides
    # =========================================================================
    print(f"\n[5/5] Nettoyage des anciens dossiers vides...")

    # Anciens dossiers a nettoyer (corbeille si vides)
    candidates = []

    # Administration (OUTER + INNER)
    for fid, label in [(ADMIN_INNER, "Administration/Administration"), (ADMIN_OUTER, "Administration")]:
        candidates.append((fid, label))

    # Sous-dossiers vides de Statuts - AG (si encore presents)
    if festival_2026 and old_admin_2026:
        old_statuts = find_child_by_name(svc, old_admin_2026['id'], "Statuts - AG")
        if old_statuts:
            decl = find_child_by_name(svc, old_statuts['id'], "Declaration Prefecture 04-2026")
            if decl:
                # sous-dossiers vides
                for subname in ("Partenaires", "Rescrit_fiscal", "Prefecture"):
                    sub = find_child_by_name(svc, decl['id'], subname)
                    if sub:
                        candidates.append((sub['id'], f"Decl-Pref/{subname}"))
                candidates.append((decl['id'], "Declaration Prefecture 04-2026"))
            for subname in ("archive",):
                sub = find_child_by_name(svc, old_statuts['id'], subname)
                if sub:
                    candidates.append((sub['id'], f"Statuts-AG/{subname}"))
            candidates.append((old_statuts['id'], "Statuts - AG"))

        for subname in ("Licences d'entrepreneurs de spectacle",):
            sub = find_child_by_name(svc, old_admin_2026['id'], subname)
            if sub:
                candidates.append((sub['id'], f"Adm2026/{subname}"))
        candidates.append((old_admin_2026['id'], "Adm2026/Administratif"))

    # Comptabilité
    comp = find_child_by_name(svc, ADMIN_OUTER, "Comptabilité") or find_child_by_name(svc, ADMIN_OUTER, "Comptabilite")
    if comp:
        candidates.append((comp['id'], "Comptabilité"))

    for fid, label in candidates:
        try:
            if folder_is_empty(svc, fid):
                trash_file(svc, fid)
                log("RM_DIR", label, "(corbeille)", fid, "OK (vide)")
            else:
                kids = list_children(svc, fid)
                log("KEEP", label, f"({len(kids)} elem restants)", fid, "non-vide")
        except Exception as e:
            log("RM_DIR", label, "...", fid, f"ERR: {e}")

    # =========================================================================
    # Sauvegarde du log CSV
    # =========================================================================
    with LOG_FILE.open('w', encoding='utf-8', newline='') as f:
        w = csv.writer(f)
        w.writerows(log_rows)

    print(f"\n=== TERMINÉ ===")
    print(f"Operations totales : {len(log_rows) - 1}")
    ok = sum(1 for r in log_rows[1:] if r[-1].startswith("OK"))
    err = sum(1 for r in log_rows[1:] if r[-1].startswith("ERR"))
    warn = sum(1 for r in log_rows[1:] if r[-1] in ("MAPPING MANQUANT", "SKIP", "non-vide", "skip"))
    print(f"  OK    : {ok}")
    print(f"  WARN  : {warn}")
    print(f"  ERR   : {err}")
    print(f"\nLog complet : {LOG_FILE}")
    print(f"Dossier cible : https://drive.google.com/drive/folders/{admin_id}")


if __name__ == "__main__":
    main()
