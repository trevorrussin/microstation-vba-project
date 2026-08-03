"""Dump Family 5 page-2 table windows with row grouping for extraction."""
from __future__ import annotations

import pathlib
import re
from collections import defaultdict

import fitz

ROOT = pathlib.Path(__file__).resolve().parents[1]


def words_page(sheet: int, page_idx: int, derotate: bool = True):
    doc = fitz.open(str(ROOT / f"Bridge/captures/619-{sheet}.pdf"))
    p = doc[page_idx]
    if derotate and p.rotation:
        # get_text with clip on rotated pages: use raw words as stored
        pass
    return list(p.get_text("words")), p.rotation


def dump_region(words, x0, y0, x1, y1, label, y_tol=4.0):
    sel = [w for w in words if x0 <= w[0] < x1 and y0 <= w[1] <= y1]
    rows = defaultdict(list)
    for w in sel:
        rows[round(w[1] / y_tol)].append(w)
    print(f"\n## {label} window=({x0},{y0})-({x1},{y1}) n={len(sel)}")
    for k in sorted(rows):
        toks = [w[4] for w in sorted(rows[k], key=lambda w: w[0])]
        print(f"  y~{k * y_tol:6.1f}: {' | '.join(toks)}")


def find_table_ids(words):
    for w in words:
        if re.fullmatch(r"\d{3}-\d{2}", w[4]) or w[4] == "TABLE":
            nearby = [
                x[4]
                for x in words
                if abs(x[1] - w[1]) < 8 and abs(x[0] - w[0]) < 350
            ]
            print(f"  {w[4]:10s} @({w[0]:6.1f},{w[1]:6.1f}) :: {' '.join(nearby)[:100]}")


def main():
    # ---- 318 page 2 (tables) ----
    print("=" * 60, "318 p2")
    W, rot = words_page(318, 1)
    print("rotation", rot, "words", len(W))
    find_table_ids(W)
    # Broad dumps of likely table regions
    dump_region(W, 50, 20, 700, 200, "318 left-top (taper?)")
    dump_region(W, 50, 180, 700, 320, "318 left-mid (roll?)")
    dump_region(W, 50, 300, 700, 550, "318 left-bot")
    dump_region(W, 700, 20, 1200, 280, "318 right-top (chan/PV?)")
    dump_region(W, 700, 280, 1200, 520, "318 right-mid (signs?)")
    dump_region(W, 700, 520, 1200, 750, "318 right-bot")

    # ---- 319 p2 ----
    print("\n" + "=" * 60, "319 p2")
    W, rot = words_page(319, 1)
    find_table_ids(W)
    dump_region(W, 50, 20, 700, 220, "319 left-top")
    dump_region(W, 700, 280, 1200, 520, "319 signs")

    # ---- 113 ----
    print("\n" + "=" * 60, "113 p1")
    W, rot = words_page(113, 0)
    print("rotation", rot)
    find_table_ids(W)
    dump_region(W, 400, 700, 800, 1200, "113 tables region (rot270 coords)")
    dump_region(W, 50, 700, 400, 1200, "113 left")

    # ---- 211 ----
    print("\n" + "=" * 60, "211 p1")
    W, rot = words_page(211, 0)
    find_table_ids(W)
    dump_region(W, 700, 250, 1200, 550, "211 tables")


if __name__ == "__main__":
    main()
