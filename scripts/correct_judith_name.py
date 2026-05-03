"""
Script de correction : "Judith LAITHIER" -> "Judith LAITHIER"
Parcourt tous les .md, .html, .py, .json, .env.example identifiés.
"""
import sys
from pathlib import Path
import re

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path(r"c:\Users\charl\Monique-Festival")

# Patterns à corriger. L'ordre importe : les patterns plus longs d'abord
# pour éviter qu'un remplacement court ne brise un pattern long.
REPLACEMENTS = [
    ("Judith LAITHIER", "Judith LAITHIER"),
    ("Judith Laithier", "Judith Laithier"),
    ("Judith LAITHIER", "Judith LAITHIER"),
    ("Judith Laithier", "Judith Laithier"),
]

# Extensions à traiter
EXTENSIONS = {".md", ".html", ".py", ".json", ".example", ".txt"}

# Exclusions : on ne touche pas aux binaires et aux PDF signés
EXCLUDE_DIRS = {".git", ".claude", "__pycache__", "node_modules", "archive"}

def should_process(path: Path) -> bool:
    if any(part in EXCLUDE_DIRS for part in path.parts):
        return False
    if path.suffix.lower() not in EXTENSIONS:
        return False
    # .env.example est traite (fichier texte)
    if path.name == ".env.example":
        return True
    return True

def correct_file(path: Path) -> tuple[int, list[str]]:
    try:
        content = path.read_text(encoding='utf-8')
    except UnicodeDecodeError:
        return 0, [f"SKIP (encoding) {path}"]

    original = content
    changes = []
    for old, new in REPLACEMENTS:
        count = content.count(old)
        if count > 0:
            content = content.replace(old, new)
            changes.append(f"  - '{old}' -> '{new}' ({count}x)")

    if content != original:
        path.write_text(content, encoding='utf-8')
        return sum([original.count(o) for o, _ in REPLACEMENTS]), changes
    return 0, []

def main():
    total_files = 0
    total_changes = 0
    modified_files = []

    for path in ROOT.rglob("*"):
        if not path.is_file():
            continue
        if not should_process(path):
            continue

        count, changes = correct_file(path)
        if count > 0:
            total_files += 1
            total_changes += count
            rel = path.relative_to(ROOT)
            modified_files.append((rel, count, changes))

    print(f"\n=== BILAN ===")
    print(f"Fichiers modifies : {total_files}")
    print(f"Occurrences remplacees : {total_changes}\n")
    print("Detail par fichier :\n")
    for rel, count, changes in sorted(modified_files):
        print(f"  {rel} ({count} occurrences)")
        for c in changes:
            print(f"    {c}")

if __name__ == "__main__":
    main()
