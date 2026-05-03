"""
Fusionne les 14 PDF du dossier de rescrit fiscal en un seul PDF paginé,
prêt pour signature DocuSign unique par la Présidente OU dépôt sur le
téléservice "Mes rescrits" sur impots.gouv.fr.

Sortie : officiel/rescrit_fiscal/Dossier_rescrit_complet_a_signer.pdf

Usage : python scripts/merge_rescrit_pdf.py
"""
import sys
from pathlib import Path
from pypdf import PdfWriter, PdfReader
sys.stdout.reconfigure(encoding='utf-8')

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "officiel" / "rescrit_fiscal" / "exports_a_signer"
OUT = REPO / "officiel" / "rescrit_fiscal" / "Dossier_rescrit_complet_a_signer.pdf"

ORDER = [
    "00_Sommaire.pdf",
    "01_Courrier_demande.pdf",
    "02_Memoire_argumentaire.pdf",
    "03_Statuts_revises.pdf",
    "04_Reglement_interieur.pdf",
    "05_PV_AGE.pdf",
    "06_PV_Bureau_post_AGE.pdf",
    "07_Recepisse_prefectoral.pdf",
    "08_Liste_dirigeants.pdf",
    "09_Annexe_7_Budget.pdf",
    "10_Annexe_8_Artistes.pdf",
    "11_Annexe_9_Tarifs.pdf",
    "12_Annexe_10_Charte.pdf",
    "13_Annexe_11_Presentation.pdf",
]

def main():
    writer = PdfWriter()
    total_pages = 0
    print(f"Fusion de {len(ORDER)} PDF :\n")
    for name in ORDER:
        p = SRC / name
        if not p.exists():
            print(f"  MANQUE  {name} (introuvable)")
            sys.exit(1)
        reader = PdfReader(str(p))
        n_pages = len(reader.pages)
        for page in reader.pages:
            writer.add_page(page)
        # Bookmark / signet pour navigation
        writer.add_outline_item(name.replace('.pdf', ''), total_pages)
        total_pages += n_pages
        print(f"  OK  {name:<35} {n_pages:>3} pages  (total {total_pages})")

    with open(OUT, 'wb') as f:
        writer.write(f)
    size_mb = OUT.stat().st_size / (1024 * 1024)
    print(f"\n=> {OUT.relative_to(REPO)}")
    print(f"   {total_pages} pages  /  {size_mb:.2f} Mo")
    print(f"\nProchaines étapes :")
    print(f"1. Téléverser ce PDF unique sur DocuSign pour signature Judith")
    print(f"   - Placer la zone de signature à la fin du Courrier (~page 4)")
    print(f"   - Optionnel : 2e zone à la fin du Mémoire et de la Liste dirigeants")
    print(f"2. Récupérer le PDF signé")
    print(f"3. Déposer :")
    print(f"   - LRAR à DDFIP du Doubs, 1 Place Jean Cornet, 25000 Besançon")
    print(f"   - OU téléservice 'Mes rescrits' sur impots.gouv.fr (max ~50 Mo : OK ici)")

if __name__ == "__main__":
    main()
