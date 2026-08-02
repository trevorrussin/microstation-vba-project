"""
One-time ingestion of the three NYSDOT/MUTCD reference PDFs into a local
FTS5 full-text index (Data/manual-index.sqlite), so search_reference_manual
(server.py) can ground engineer-facing answers in the actual manuals
instead of the model's own recollection.

Run: python ingest_manuals.py
Re-run any time the source PDFs change -- safe to re-run, drops and
rebuilds the table each time.

Source PDFs (Project Documentation/, gitignored for size):
  part6.pdf                     - MUTCD Part 6, Temporary Traffic Control (204pp)
  B-2011Supplement-adopted.pdf  - NYS MUTCD Supplement (269pp)
  2026_1_stdsht_usc_book_3.pdf  - NYSDOT standard detail sheets (159pp)

part6.pdf and the Supplement are clean prose with section-number headers
(confirmed live via PyMuPDF extraction: patterns like "6A.01", "2G.04")
-- chunked by section. The standard-sheet book is mostly engineering-
drawing labels/dimensions, not prose -- chunked per page instead
(confirmed live: scattered fragment text, no reliable section-header
pattern to key on).
"""
from __future__ import annotations

import re
import sqlite3
from pathlib import Path

import fitz  # PyMuPDF

_REPO_ROOT = Path(__file__).resolve().parents[1]
DOCS_DIR = _REPO_ROOT / "Project Documentation"
INDEX_PATH = _REPO_ROOT / "Data" / "manual-index.sqlite"

SHEET_NUM_RE = re.compile(r"619-\d{3}")

# (source key, filename, chunk mode) -- source key is what search_reference_manual's
# `source` filter matches against.
SOURCES = [
    ("part6", "part6.pdf", "section"),
    ("supplement", "B-2011Supplement-adopted.pdf", "section"),
    ("stdsht", "2026_1_stdsht_usc_book_3.pdf", "page"),
]

# Matches section-number headers like "6A.01", "2G.04", "6D.03" at the
# start of a line.
SECTION_HEADER_RE = re.compile(r"^(\d[A-Z]\.\d\d[A-Za-z]?)\b")

# Cap a section chunk's page span even if no new header appears -- confirmed
# live that the Supplement PDF isn't uniform prose throughout: it switches
# to a "Sign Drawing" appendix (one sign per page, no section numbers) partway
# through, which without a cap absorbed 100+ pages into a single chunk under
# the last real section header, making both the chunk and its page citation
# useless. A capped continuation chunk keeps citations meaningful.
MAX_PAGES_PER_CHUNK = 6


def chunk_by_section(doc: fitz.Document) -> list[dict]:
    """One chunk per detected section-header run, capped at
    MAX_PAGES_PER_CHUNK pages. Text seen before the first header on a
    document (e.g. a title page), or a run with no new header for longer
    than the cap (e.g. an unnumbered appendix), becomes a page-numbered
    fallback/continuation chunk rather than growing unboundedly."""
    chunks: list[dict] = []
    current_heading = ""
    current_text: list[str] = []
    current_start_page = 0
    is_continuation = False

    def flush(end_page: int) -> None:
        if current_text and "".join(current_text).strip():
            heading = current_heading
            if is_continuation and current_heading:
                heading = f"{current_heading} (cont'd, p{current_start_page})"
            chunks.append({
                "heading": heading,
                "text": "\n".join(current_text).strip(),
                "page_start": current_start_page,
                "page_end": end_page,
            })

    for page_idx in range(len(doc)):
        text = doc[page_idx].get_text()
        page_no = page_idx + 1

        if current_text and (page_no - current_start_page) >= MAX_PAGES_PER_CHUNK:
            flush(page_no - 1)
            current_text = []
            current_start_page = page_no
            is_continuation = True

        for line in text.splitlines():
            m = SECTION_HEADER_RE.match(line.strip())
            if m:
                flush(page_no)
                current_heading = m.group(1)
                current_text = [line]
                current_start_page = page_no
                is_continuation = False
            else:
                if not current_text:
                    if not current_heading:
                        current_heading = f"(page {page_no})"
                    current_start_page = page_no
                current_text.append(line)
    flush(len(doc))
    return chunks


def chunk_by_page(doc: fitz.Document) -> list[dict]:
    """One chunk per page. For stdsht drawings, embed the first 619-NNN
    found on the page into the heading so citations read clearly
    (e.g. '619-310 page 79') even though FTS still matches body text."""
    chunks = []
    for page_idx in range(len(doc)):
        text = doc[page_idx].get_text().strip()
        if not text:
            continue
        page_no = page_idx + 1
        m = SHEET_NUM_RE.search(text)
        heading = f"{m.group(0)} page {page_no}" if m else f"page {page_no}"
        chunks.append({
            "heading": heading,
            "text": text,
            "page_start": page_no,
            "page_end": page_no,
        })
    return chunks


def build_index() -> None:
    INDEX_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(INDEX_PATH)
    conn.execute("DROP TABLE IF EXISTS chunks")
    conn.execute(
        "CREATE VIRTUAL TABLE chunks USING fts5("
        "text, heading UNINDEXED, source UNINDEXED, "
        "page_start UNINDEXED, page_end UNINDEXED)"
    )

    total = 0
    for source_key, filename, mode in SOURCES:
        path = DOCS_DIR / filename
        if not path.exists():
            print(f"SKIP {filename} -- not found at {path}")
            continue
        doc = fitz.open(path)
        chunks = chunk_by_section(doc) if mode == "section" else chunk_by_page(doc)
        doc.close()
        conn.executemany(
            "INSERT INTO chunks (text, heading, source, page_start, page_end) "
            "VALUES (?, ?, ?, ?, ?)",
            [(c["text"], c["heading"], source_key, c["page_start"], c["page_end"]) for c in chunks],
        )
        print(f"{filename}: {len(chunks)} chunks ({mode} mode)")
        total += len(chunks)

    conn.commit()
    conn.close()
    print(f"Indexed {total} chunks into {INDEX_PATH}")


if __name__ == "__main__":
    build_index()
