"""
Read-only FTS5 search over Data/manual-index.sqlite, built by
ingest_manuals.py. Returns excerpts with page/source citations so an
answer can point the engineer back to the actual manual -- this project
treats "why is that there" as load-bearing (see the reason= field on
every WZTCBridge op), and manual citations follow the same spirit.
"""
from __future__ import annotations

import re
import sqlite3
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]
INDEX_PATH = _REPO_ROOT / "Data" / "manual-index.sqlite"
DOCS_DIR = _REPO_ROOT / "Project Documentation"

SOURCE_NAMES = {
    "part6": "MUTCD Part 6 (Temporary Traffic Control)",
    "supplement": "NYS MUTCD Supplement",
    "stdsht": "NYSDOT Standard Detail Sheets",
}

# Duplicated from ingest_manuals.py's SOURCES (rather than imported) to keep
# this read-only search module decoupled from that one-time ingestion
# script -- same pattern as each file already keeping its own _REPO_ROOT.
_SOURCE_FILES = {
    "part6": "part6.pdf",
    "supplement": "B-2011Supplement-adopted.pdf",
    "stdsht": "2026_1_stdsht_usc_book_3.pdf",
}

_SELECT = (
    "SELECT heading, source, page_start, page_end, "
    "snippet(chunks, 0, '>>>', '<<<', '...', 20) AS excerpt "
    "FROM chunks WHERE chunks MATCH ?"
)

# Tokens for OR/phrase fallback — drop pure FTS5 operators so a deliberate
# "a OR b" query isn't rewritten into nonsense.
_FTS_OPS = {"AND", "OR", "NOT", "NEAR"}
_TOKEN_RE = re.compile(r'[^\s"\'()]+|"[^"]+"')


def _run_query(conn: sqlite3.Connection, match_expr: str, source: str, max_results: int) -> list[sqlite3.Row]:
    sql = _SELECT
    params: list = [match_expr]
    if source:
        sql += " AND source = ?"
        params.append(source)
    sql += " ORDER BY rank LIMIT ?"
    params.append(max_results)
    return conn.execute(sql, params).fetchall()


def _plain_tokens(query: str) -> list[str]:
    """Split a plain-English query into FTS tokens (keeps quoted phrases)."""
    tokens: list[str] = []
    for m in _TOKEN_RE.finditer(query.strip()):
        tok = m.group(0)
        if tok.upper() in _FTS_OPS:
            continue
        tokens.append(tok)
    return tokens


def _rows_to_dicts(rows: list[sqlite3.Row]) -> list[dict]:
    return [
        {
            "source": r["source"],
            "source_name": SOURCE_NAMES.get(r["source"], r["source"]),
            "heading": r["heading"],
            "page_start": r["page_start"],
            "page_end": r["page_end"],
            "excerpt": r["excerpt"],
        }
        for r in rows
    ]


def search(query: str, source: str = "", max_results: int = 10) -> list[dict]:
    """FTS5 MATCH search across the three manuals. source, if given, must be
    one of part6/supplement/stdsht (empty searches all three).

    If the index file is missing, returns a single INDEX_MISSING diagnostic
    hit (never silent []) so agents know to run ingest_manuals.py.

    query is tried as an FTS5 MATCH expression first (so an agent can use
    AND/OR/"phrase" deliberately). On syntax error, retries as a quoted
    literal. If a multi-token query still returns zero hits, retries with
    OR-joined tokens, then as a quoted phrase — FTS5 default AND is why
    reasonable English + a narrow source= filter often returned empty."""
    if not INDEX_PATH.exists():
        return [{
            "source": "",
            "source_name": "",
            "heading": "INDEX_MISSING",
            "page_start": 0,
            "page_end": 0,
            "excerpt": (
                f"Data/manual-index.sqlite not found at {INDEX_PATH}. "
                "Place the three reference PDFs under Project Documentation/ "
                "and run: python mcp-server/ingest_manuals.py"
            ),
        }]

    conn = sqlite3.connect(INDEX_PATH)
    conn.row_factory = sqlite3.Row
    try:
        rows: list[sqlite3.Row] = []
        try:
            rows = _run_query(conn, query, source, max_results)
        except sqlite3.OperationalError:
            literal = '"' + query.replace('"', '""') + '"'
            try:
                rows = _run_query(conn, literal, source, max_results)
            except sqlite3.OperationalError:
                rows = []

        if not rows:
            tokens = _plain_tokens(query)
            if len(tokens) >= 2:
                or_expr = " OR ".join(tokens)
                try:
                    rows = _run_query(conn, or_expr, source, max_results)
                except sqlite3.OperationalError:
                    rows = []
                if not rows:
                    literal = '"' + query.replace('"', '""') + '"'
                    try:
                        rows = _run_query(conn, literal, source, max_results)
                    except sqlite3.OperationalError:
                        rows = []
    finally:
        conn.close()

    return _rows_to_dicts(rows)


def render_page_image(source: str, page_num: int, out_path: str | Path, dpi: int = 150) -> Path:
    """Renders one page (1-based, matching a search() hit's page_start) of
    the given source manual to a PNG at out_path -- lets a caller show the
    actual manual/sheet page an answer is grounded in, not just the text
    excerpt (2026-08-02 feedback: "could the image show the reference it
    pulled the info from, like the drawing screenshot does").

    Only renders page_start, not the full page_start..page_end span --
    that's the page the matched chunk actually starts on, and stitching a
    multi-page span into one image is unneeded complexity for what's meant
    to be a quick visual aid alongside the text excerpt, not a full-chunk
    reproduction.

    Raises FileNotFoundError if that source's PDF isn't present locally
    (gitignored for size, see Data/README.md) or ValueError for an unknown
    source / out-of-range page -- callers treat this as best-effort and
    catch broadly, so raising precisely here is more useful than silently
    swallowing the problem."""
    import fitz  # PyMuPDF -- already a dependency, see ingest_manuals.py

    filename = _SOURCE_FILES.get(source)
    if filename is None:
        raise ValueError(f"unknown source {source!r} -- expected one of {sorted(_SOURCE_FILES)}")
    pdf_path = DOCS_DIR / filename
    if not pdf_path.exists():
        raise FileNotFoundError(
            f"{pdf_path} not found -- reference PDFs are gitignored locally, see Data/README.md"
        )

    doc = fitz.open(pdf_path)
    try:
        if not (1 <= page_num <= len(doc)):
            raise ValueError(f"page {page_num} out of range for {filename} ({len(doc)} pages)")
        pix = doc[page_num - 1].get_pixmap(dpi=dpi)
        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        pix.save(str(out_path))
    finally:
        doc.close()
    return out_path
