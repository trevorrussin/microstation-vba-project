"""
Print Book 3 sheet -> owned pages -> sign/table candidates for human
review when extending Data/sheet-registry.tsv.

Does NOT write the TSV (signs can be contaminated by legend/cross-refs;
always verify against the PDF page before seeding a row).

Usage:
  python scripts/extract_sheet_signs.py
  python scripts/extract_sheet_signs.py 619-201 619-310
"""
from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path

import fitz

REPO = Path(__file__).resolve().parents[1]
PDF = REPO / "Project Documentation" / "2026_1_stdsht_usc_book_3.pdf"

SIGN_RE = re.compile(r"\b(?:NY)?(?:W|R|G|M|D|S)\d{1,2}-\d{1,2}[A-Za-z]{0,3}\b")
TABLE_RE = re.compile(r"TABLE\s+(\d{3}-\d{2})", re.I)
DGN_RE = re.compile(r"(619-\d{3})(?:-\d+)?[_A-Za-z0-9]*\.dgn", re.I)


def owned_pages(doc: fitz.Document) -> dict[str, list[int]]:
    """Pages whose embedded .dgn filename starts with 619-NNN."""
    owner: dict[str, list[int]] = defaultdict(list)
    for i in range(doc.page_count):
        text = doc.load_page(i).get_text("text")
        m = DGN_RE.search(text)
        if m:
            owner[m.group(1)].append(i + 1)
    return owner


def summarize(doc: fitz.Document, sheet: str, pages: list[int]) -> None:
    text = "\n".join(doc.load_page(p - 1).get_text("text") for p in pages)
    signs = sorted(set(SIGN_RE.findall(text)))
    prefix = sheet.split("-")[1]
    tables = sorted({t for t in TABLE_RE.findall(text) if t.startswith(prefix)})
    print(f"{sheet}  pages={pages}")
    print(f"  signs  {'|'.join(signs) if signs else '(none)'}")
    print(f"  tables {','.join(tables) if tables else '(none sheet-prefixed)'}")


def main(argv: list[str]) -> int:
    if not PDF.exists():
        print(f"PDF not found: {PDF}", file=sys.stderr)
        return 1
    doc = fitz.open(PDF)
    owner = owned_pages(doc)
    wanted = argv[1:] if len(argv) > 1 else sorted(owner.keys())
    for sn in wanted:
        pages = owner.get(sn, [])
        if not pages:
            print(f"{sn}  (no owned .dgn pages found)")
            continue
        summarize(doc, sn, pages)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
