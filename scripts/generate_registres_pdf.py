"""
Génère les DOCX horodatés des 3 registres légaux (special, délibérations,
membres) à partir des fichiers .md, prêts pour signature DocuSign.

Pourquoi DOCX et pas PDF directement ?
- Aucun moteur PDF n'est installé localement (LaTeX/wkhtmltopdf/weasyprint absents)
- DocuSign accepte les DOCX et produit le PDF signé
- Cohérent avec le workflow existant (PV, statuts → DOCX → DocuSign)

Usage : python scripts/generate_registres_pdf.py
"""
import sys, subprocess
from datetime import date
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')

REPO = Path(__file__).resolve().parent.parent
REG_DIR = REPO / "gouvernance" / "registres"
OUT_DIR = REG_DIR / "exports_a_signer"

REGISTRES = [
    "Registre_special.md",
    "Registre_deliberations.md",
    "Registre_membres.md",
]

def main():
    if not REG_DIR.exists():
        sys.exit(f"ERREUR : {REG_DIR} introuvable.")
    OUT_DIR.mkdir(exist_ok=True)
    today = date.today().isoformat()

    for name in REGISTRES:
        src = REG_DIR / name
        if not src.exists():
            print(f"  SKIP {name} (introuvable)")
            continue
        stem = src.stem  # ex: Registre_special
        out = OUT_DIR / f"{stem}_arrete_{today}.docx"
        cmd = [
            "pandoc", str(src), "-o", str(out),
            "--from=gfm",
            "--to=docx",
            f"--metadata=title:{stem.replace('_', ' ')}",
        ]
        try:
            subprocess.run(cmd, check=True, capture_output=True, text=True)
            kb = out.stat().st_size // 1024
            print(f"  OK  {out.relative_to(REPO)}  ({kb} KB)")
        except subprocess.CalledProcessError as e:
            print(f"  FAIL {name} : {e.stderr}")
            sys.exit(1)

    print(f"\nProchaines étapes :")
    print(f"1. Ouvrir chaque DOCX dans Word, vérifier rendu, sauvegarder en PDF si besoin")
    print(f"2. Téléverser le DOCX (ou PDF) dans DocuSign → signature Présidente")
    print(f"3. Récupérer le PDF signé, le déposer dans gouvernance/registres/signes/")
    print(f"4. Mettre à jour la table 'Suivi de signature électronique' du .md correspondant")
    print(f"5. Sauvegarder sur Drive : 1_Gouvernance/Registres/")

if __name__ == "__main__":
    main()
