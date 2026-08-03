"""Shared PDF word-extraction primitives for authoring/round-tripping sheet
specs (see Data/sheet-specs/AUTHORING.md).

These are the coordinate-grouping techniques that made 619-311 reliable,
factored out of the one-off Bridge/_roundtrip_619311.py so the next sheet
doesn't reinvent them. What is deliberately NOT generalized here: the
per-table comparison logic (which JSON field lines up with which extracted
cell) is different for almost every table shape across the 619 catalog, so
a single blind "diff this table" function would either miss real transcription
errors on unfamiliar layouts or need so many special cases it stops being
trustworthy. Each sheet still gets a short round-trip script under
Bridge/roundtrip/<sheet>.py, but that script should only ever call these
primitives, never re-derive them.
"""
from __future__ import annotations

from collections import defaultdict
from typing import Iterable

Word = tuple  # (x0, y0, x1, y1, text, block, line, word_no) from page.get_text("words")


def words_in_window(words: Iterable[Word], x0: float, y0: float, x1: float, y1: float,
                     pad: float = 2.0) -> list[Word]:
    """Words whose bounding box lies fully inside the window, with a small
    pad on the far edges. Filtering on w[3] <= y1 with no pad silently drops
    any row whose bottom crosses the boundary -- this is the "window eats the
    last row" failure mode AUTHORING.md warns about. Widen the window rather
    than trust a tight fit."""
    return [w for w in words
            if w[0] >= x0 and w[2] <= x1 + pad and w[1] >= y0 and w[3] <= y1 + pad]


def group_rows(words: Iterable[Word], y_tol: float = 3.0) -> list[list[Word]]:
    """Group words into rows by rounded y, sorted left-to-right within a row,
    rows sorted top-to-bottom. A logical row split across two close y-bands
    (seen on 619-311's 45 mph row of table 311-02) needs a coarser y_tol --
    if a row looks half-empty, retry with a larger value (e.g. 8.0) before
    concluding the table is short a row."""
    rows: dict[int, list[Word]] = defaultdict(list)
    for w in words:
        rows[round(w[1] / y_tol)].append(w)
    return [sorted(rows[k], key=lambda w: w[0]) for k in sorted(rows)]


def row_text(row: Iterable[Word]) -> str:
    return " ".join(w[4] for w in row)


def squash(s: str) -> str:
    """Normalize for substring/equality comparison against PDF text, which
    often has inconsistent spacing/punctuation around the same content."""
    return s.replace(".", "").replace(" ", "").upper()


def assert_row_count(rows: list, expected: int, where: str) -> None:
    """Call this after every group_rows() in a round-trip script. A silently
    short table (the notes column dropping its RURAL row, table 311-05 losing
    WARNING FLAG) looks like a clean extraction unless the count is checked."""
    if len(rows) != expected:
        raise AssertionError(
            f"{where}: expected {expected} rows, got {len(rows)} -- "
            f"re-check the window bounds and y_tol before trusting this extraction")
