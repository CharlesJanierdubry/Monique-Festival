"""
Génère les DOCX de l'ensemble du dossier de rescrit fiscal DDFIP, prêts pour
relecture Word puis "Enregistrer en PDF" puis signature DocuSign.

Sortie : officiel/rescrit_fiscal/exports_a_signer/

Usage : python scripts/generate_rescrit_pdf.py
"""
import sys, subprocess
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')

REPO = Path(__file__).resolve().parent.parent
RESCRIT = REPO / "officiel" / "rescrit_fiscal"
OUT = RESCRIT / "exports_a_signer"

DOCS = [
    ("Sommaire_dossier_rescrit.md", "00_Sommaire.docx"),
    ("Courrier_rescrit_DDFIP.md", "01_Courrier_demande.docx"),
    ("Memoire_rescrit_mecenat.md", "02_Memoire_argumentaire.docx"),
    ("Liste_dirigeants.md", "08_Liste_dirigeants.docx"),
    ("annexes/Annexe_7_Budget_previsionnel_2026.md", "09_Annexe_7_Budget.docx"),
    ("annexes/Annexe_8_Liste_artistes_transparence.md", "10_Annexe_8_Artistes.docx"),
    ("annexes/Annexe_9_Comparatif_tarifs.md", "11_Annexe_9_Tarifs.docx"),
    ("annexes/Annexe_10_Charte_ethique.md", "12_Annexe_10_Charte.docx"),
    ("annexes/Annexe_11_Presentation_festival.md", "13_Annexe_11_Presentation.docx"),
]

def main():
    if not RESCRIT.exists():
        sys.exit(f"ERREUR : {RESCRIT} introuvable.")
    OUT.mkdir(exist_ok=True)

    print(f"Génération de {len(DOCS)} DOCX dans {OUT.relative_to(REPO)}\n")
    for src_rel, out_name in DOCS:
        src = RESCRIT / src_rel
        if not src.exists():
            print(f"  SKIP {src_rel} (introuvable)")
            continue
        out = OUT / out_name
        cmd = [
            "pandoc", str(src), "-o", str(out),
            "--from=gfm",
            "--to=docx",
        ]
        try:
            subprocess.run(cmd, check=True, capture_output=True, text=True)
            kb = out.stat().st_size // 1024
            print(f"  OK  {out_name}  ({kb} KB)")
        except subprocess.CalledProcessError as e:
            print(f"  FAIL {src_rel} : {e.stderr}")

    print(f"\nProchaines étapes :")
    print(f"1. Ouvrir chaque DOCX dans Word, vérifier rendu")
    print(f"2. 'Enregistrer en PDF' chaque document")
    print(f"3. Signer DocuSign les 3 documents principaux (Judith) :")
    print(f"   - 01_Courrier_demande.docx")
    print(f"   - 02_Memoire_argumentaire.docx")
    print(f"   - 08_Liste_dirigeants.docx")
    print(f"4. Récupérer aussi les PDF déjà signés :")
    print(f"   - Statuts révisés (gouvernance/docx/Statuts_Monique_Festival_19-04-2026_Signé.pdf)")
    print(f"   - Règlement intérieur (gouvernance/docx/Reglement_Interieur 19-04-2026_ Signé.pdf)")
    print(f"   - PV AGE (gouvernance/docx/PV_AGE_adoption_statuts 19-04-2026 - Signé.pdf)")
    print(f"   - PV Bureau post-AGE (gouvernance/docx/PV_Bureau_post_AGE19-04-2026_ Signé.pdf)")
    print(f"   - Récépissé préfectoral (officiel/prefecture/Recepisse_modification_24-04-2026.pdf)")
    print(f"5. Compiler en un dossier paginé (LRAR) ou un PDF unique (téléservice)")
    print(f"6. Déposer (LRAR ou impots.gouv.fr 'Mes rescrits')")

if __name__ == "__main__":
    main()
