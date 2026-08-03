"""Extract Family 7 (mobile 110/111/112) table cells into draft JSON."""
from __future__ import annotations

import json
import pathlib
import re
from collections import defaultdict

import fitz

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "Data" / "sheet-specs"


def dump_region(words, x0, x1, y0, y1, label):
    sel = [w for w in words if x0 <= w[0] <= x1 and y0 <= w[1] <= y1]
    rows = defaultdict(list)
    for w in sel:
        rows[round(w[1] / 4.0)].append(w)
    print(f"\n-- {label} window ({x0},{y0})-({x1},{y1}) n={len(sel)} --")
    for k in sorted(rows):
        cells = sorted(rows[k], key=lambda w: w[0])
        print(f"  y~{k*4:6.1f}: " + " | ".join(f"{c[4]}@{c[0]:.0f}" for c in cells))


def all_phrases(pg):
    words = list(pg.get_text("words"))
    body = " ".join(w[4] for w in words)
    return words, body


def extract_111():
    doc = fitz.open(str(ROOT / "Bridge/captures/619-111.pdf"))
    draft = {"sheet": "619-111", "pages": doc.page_count, "pages_detail": []}
    for pi in range(doc.page_count):
        pg = doc[pi]
        words, body = all_phrases(pg)
        print(f"\n========== 111 page {pi} ==========")
        dump_region(words, 900, 1224, 0, 280, "PV table top")
        dump_region(words, 900, 1224, 280, 450, "roll ahead mid")
        dump_region(words, 900, 1224, 450, 700, "sign sizes / notes")
        dump_region(words, 0, 550, 80, 720, "plan left")
        dump_region(words, 550, 920, 200, 550, "notes mid")
        # Sign codes
        for tok in ("NYW8-33", "W20-5R", "W4-2R", "W20-5aR", "W20-5AR", "MOBILE", "500'", "2 MILE"):
            print(f"  has {tok!r}: {tok in body}")
        draft["pages_detail"].append({"page": pi, "body_len": len(body)})
    out = OUT / "_draft_619111_tables.json"
    # Also dump full word grid for right column more carefully
    for pi in range(doc.page_count):
        words = list(doc[pi].get_text("words"))
        dump_region(words, 980, 1224, 0, 792, f"page{pi} full right col")
    return draft


def extract_rotated(sheet: str):
    """110/112 are pageRotation=270 — words come in landscape display coords."""
    doc = fitz.open(str(ROOT / f"Bridge/captures/{sheet}.pdf"))
    print(f"\n\n################ {sheet} pages={doc.page_count} ################")
    draft = {"sheet": sheet, "pages": doc.page_count, "tables": {}, "notes": [], "findings": []}
    for pi in range(doc.page_count):
        pg = doc[pi]
        words = list(pg.get_text("words"))
        print(f"\n===== {sheet} page {pi} rot={pg.rotation} rect={pg.rect} =====")
        # Find extent
        xs = [w[0] for w in words]
        ys = [w[1] for w in words]
        print(f"  x[{min(xs):.0f},{max(xs):.0f}] y[{min(ys):.0f},{max(ys):.0f}] n={len(words)}")
        # Dump all words sorted by y then x in coarse bands — full page
        rows = defaultdict(list)
        for w in words:
            rows[round(w[1] / 6.0)].append(w)
        for k in sorted(rows):
            cells = sorted(rows[k], key=lambda w: w[0])
            line = " ".join(c[4] for c in cells)
            # only print lines that look table-ish or notes
            if any(
                t in line.upper()
                for t in (
                    "TABLE",
                    "ROLL",
                    "PROTECTIVE",
                    "SIGN",
                    "MPH",
                    "FREEWAY",
                    "CLOSURE",
                    "NOTE",
                    "MOBILE",
                    "W8",
                    "W20",
                    "W4",
                    "NYW",
                    "MINIMUM",
                    "MAXIMUM",
                    "MILE",
                    "WARNING",
                    "TMIA",
                    "LANE",
                    "SHOULDER",
                )
            ):
                print(f"  y~{k*6:6.1f}: {line[:200]}")
        body = " ".join(w[4] for w in words)
        for tok in (
            "NYW8-33",
            "W20-5R",
            "W20-5aR",
            "W20-5AR",
            "W4-2R",
            "W8-23",
            "MOBILE",
            "500'",
            "2 MILE",
            "MERGING",
            "SHOULDER TAPER",
            "ROLL AHEAD",
        ):
            print(f"  has {tok!r}: {tok in body}")
    return draft


if __name__ == "__main__":
    extract_111()
    extract_rotated("619-110")
    extract_rotated("619-112")
