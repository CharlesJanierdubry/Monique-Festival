"""Verifie la presence de 'Wadih CORMIER' (bon) ou 'Wadih JANIER-DUBRY' (mauvais) dans les PDF signes."""
import sys
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')
from pypdf import PdfReader

PDFS = [
    "gouvernance/docx/Statuts_Monique_Festival_19-04-2026_Signé.pdf",
    "gouvernance/docx/PV_AGE_adoption_statuts 19-04-2026 - Signé.pdf",
    "gouvernance/docx/Feuille de Présence 19-04-2026 - Signé.pdf",
    "gouvernance/docx/Reglement_Interieur 19-04-2026_ Signé.pdf",
    "gouvernance/docx/PV_Bureau_post_AGE19-04-2026_ Signé.pdf",
]

ROOT = Path(r"c:\Users\charl\Monique-Festival")

MAUVAIS = ["Wadih JANIER-DUBRY", "Wadih Janier-Dubry", "JANIER-DUBRY, Wadih"]
BONS = ["Wadih CORMIER", "Wadih Cormier", "CORMIER, Wadih"]

for rel in PDFS:
    path = ROOT / rel
    print(f"\n--- {path.name} ---")
    if not path.exists():
        print("  FICHIER ABSENT")
        continue
    try:
        reader = PdfReader(str(path))
        text = "\n".join(p.extract_text() or "" for p in reader.pages)
    except Exception as e:
        print(f"  ERREUR : {e}")
        continue

    mauvaises = [(p, text.count(p)) for p in MAUVAIS if text.count(p) > 0]
    bonnes = [(p, text.count(p)) for p in BONS if text.count(p) > 0]
    if mauvaises:
        print(f"  ⚠️ MAUVAIS : {mauvaises}")
    if bonnes:
        print(f"  ✅ BON : {bonnes}")
    if not mauvaises and not bonnes:
        print("  (aucune mention de Wadih)")
