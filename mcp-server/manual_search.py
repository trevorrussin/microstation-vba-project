"""
Read-only FTS5 search over Data/manual-index.sqlite, built by
ingest_manuals.py. Returns excerpts with page/source citations so an
answer can point the engineer back to the actual manual -- this project
treats "why is that there" as load-bearing (see the reason= field on
every WZTCBridge op), and manual citations follow the same spirit.
"""
from __future__ import annotations

import sqlite3
from pathlib import Path

INDEX_PATH = Path(r"c:\repos\microstation-vba-project\Data\manual-index.sqlite")

SOURCE_NAMES = {
    "part6": "MUTCD Part 6 (Temporary Traffic Control)",
    "supplement": "NYS MUTCD Supplement",
    "stdsht": "NYSDOT Standard Detail Sheets",
}

_SELECT = (
    "SELECT heading, source, page_start, page_end, "
    "snippet(chunks, 0, '>>>', '<<<', '...', 20) AS excerpt "
    "FROM chunks WHERE chunks MATCH ?"
)


def _run_query(conn: sqlite3.Connection, match_expr: str, source: str, max_results: int) -> list[sqlite3.Row]:
    sql = _SELECT
    params: list = [match_expr]
    if source:
        sql += " AND source = ?"
        params.append(source)
    sql += " ORDER BY rank LIMIT ?"
    params.append(max_results)
    return conn.execute(sql, params).fetchall()


def search(query: str, source: str = "", max_results: int = 10) -> list[dict]:
    """FTS5 MATCH search across the three manuals. source, if given, must be
    one of part6/supplement/stdsht (empty searches all three). Returns []
    (not an error) if the index doesn't exist yet -- run ingest_manuals.py
    first, or if genuinely nothing matched.

    query is passed through as an FTS5 MATCH expression first (so an agent
    can use FTS5 operators like AND/OR/"phrase" deliberately); if that
    raises a syntax error -- e.g. a bareword query containing a hyphen,
    which FTS5 parses as NOT -- it's retried as a single literal phrase
    (query wrapped in quotes) instead of surfacing a confusing SQL error
    for what's just a plain-English search term."""
    if not INDEX_PATH.exists():
        return []

    conn = sqlite3.connect(INDEX_PATH)
    conn.row_factory = sqlite3.Row
    try:
        try:
            rows = _run_query(conn, query, source, max_results)
        except sqlite3.OperationalError:
            literal = '"' + query.replace('"', '""') + '"'
            rows = _run_query(conn, literal, source, max_results)
    finally:
        conn.close()

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
