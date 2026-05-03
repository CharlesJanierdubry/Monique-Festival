"""
Script de correction : "Wadih CORMIER" -> "Wadih CORMIER"
Parcourt tous les .md, .html, .py, .json, .txt identifies.
"""
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

ROOT = Path(r"c:\Users\charl\Monique-Festival")

REPLACEMENTS = [
    ("Wadih CORMIER", "Wadih CORMIER"),
    ("Wadih Cormier", "Wadih Cormier"),
    ("Wadih CORMIER", "Wadih CORMIER"),
    ("Wadih Cormier", "Wadih Cormier"),
    ("CORMIER, Wadih", "CORMIER, Wadih"),
    ("CORMIER, Wadih", "CORMIER, Wadih"),
]

EXTENSIONS = {".md", ".html", ".py", ".json", ".example", ".txt"}
EXCLUDE_DIRS = {".git", ".claude", "__pycache__", "node_modules", "archive"}


def should_process(path: Path) -> bool:
    if any(part in EXCLUDE_DIRS for part in path.parts):
        return False
    if path.suffix.lower() not in EXTENSIONS:
        return False
    return True


def correct_file(path: Path) -> tuple[int, list[str]]:
    try:
        content = path.read_text(encoding='utf-8')
    except UnicodeDecodeError:
        return 0, [f"SKIP (encoding) {path}"]

    original = content
    changes = []
    total = 0
    for old, new in REPLACEMENTS:
        count = content.count(old)
        if count > 0:
            content = content.replace(old, new)
            changes.append(f"  - '{old}' -> '{new}' ({count}x)")
            total += count

    if content != original:
        path.write_text(content, encoding='utf-8')
        return total, changes
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
