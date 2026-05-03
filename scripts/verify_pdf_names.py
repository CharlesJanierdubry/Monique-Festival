"""
Extrait le texte des 5 PDF signes et verifie la presence du bon nom.
"""
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        print("ERREUR : installer pypdf  ->  pip install pypdf")
        sys.exit(1)

PDFS = [
    "gouvernance/docx/Statuts_Monique_Festival_19-04-2026_Signé.pdf",
    "gouvernance/docx/PV_AGE_adoption_statuts 19-04-2026 - Signé.pdf",
    "gouvernance/docx/Feuille de Présence 19-04-2026 - Signé.pdf",
    "gouvernance/docx/Reglement_Interieur 19-04-2026_ Signé.pdf",
    "gouvernance/docx/PV_Bureau_post_AGE19-04-2026_ Signé.pdf",
]

ROOT = Path(r"c:\Users\charl\Monique-Festival")

PATTERNS_MAUVAIS = ["Judith JANIER-DUBRY", "Judith Janier-Dubry", "Judith JANIER DUBRY", "Judith Janier Dubry"]
PATTERNS_BONS = ["Judith LAITHIER", "Judith Laithier"]

def verify_pdf(path: Path):
    try:
        reader = PdfReader(str(path))
    except Exception as e:
        return f"  ERREUR LECTURE : {e}"

    all_text = ""
    for page in reader.pages:
        try:
            all_text += page.extract_text() + "\n"
        except Exception as e:
            all_text += f"[page non extractible : {e}]\n"

    mauvaises = []
    bonnes = []
    for p in PATTERNS_MAUVAIS:
        n = all_text.count(p)
        if n > 0:
            mauvaises.append((p, n))
    for p in PATTERNS_BONS:
        n = all_text.count(p)
        if n > 0:
            bonnes.append((p, n))

    verdict = []
    if mauvaises:
        verdict.append(f"  MAUVAIS NOM DETECTE : {mauvaises}")
    if bonnes:
        verdict.append(f"  BON NOM detecte : {bonnes}")
    if not mauvaises and not bonnes:
        verdict.append(f"  AUCUNE MENTION de Judith")
    return "\n".join(verdict)

def main():
    print(f"\n=== Verification des PDF signes ===\n")
    for rel in PDFS:
        path = ROOT / rel
        print(f"--- {path.name} ---")
        if not path.exists():
            print(f"  FICHIER ABSENT")
            continue
        print(verify_pdf(path))
        print()

if __name__ == "__main__":
    main()
