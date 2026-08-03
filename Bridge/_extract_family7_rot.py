"""Coordinate-accurate extraction for rotated Family 7 PDFs (110/112)."""
from __future__ import annotations

import json
import pathlib
from collections import defaultdict

import fitz

ROOT = pathlib.Path(__file__).resolve().parent.parent


def upright_words(pg):
    """Return words in upright reading order for any page rotation."""
    # Render text through a derotated transformation by using page.get_text
    # with clip on the mediabox after setting rotation temporarily.
    # PyMuPDF: words coords are in unrotated page space when rotation!=0
    # for many MicroStation exports — try both.
    w0 = list(pg.get_text("words"))
    # Also try after forcing rotation=0 display
    # Create a display list and extract from derotated pixmap? Too heavy.
    # Instead: if rotation==270, swap/transform:
    # unrotated page is 792x1224; display shows landscape.
    if pg.rotation == 0:
        return w0
    # For rot=270: (x',y') display where x' = y_raw?, check extents
    # Observed: x in [22,739], y in [26,1176] — so y is the long axis =
    # landscape width. Treat y as horizontal reading axis when dumping tables
    # that sit at high y (right side of landscape = bottom of portrait?).
    return w0


def dump_by_y(words, y0, y1, x0=0, x1=9999, band=5.0, label=""):
    sel = [w for w in words if y0 <= w[1] <= y1 and x0 <= w[0] <= x1]
    rows = defaultdict(list)
    for w in sel:
        rows[round(w[1] / band)].append(w)
    print(f"\n-- {label} y[{y0},{y1}] x[{x0},{x1}] n={len(sel)} --")
    for k in sorted(rows):
        cells = sorted(rows[k], key=lambda w: w[0])
        print(f"  y~{k*band:6.1f}: " + " | ".join(f"{c[4]}@{c[0]:.0f}" for c in cells))


def extract_110():
    doc = fitz.open(str(ROOT / "Bridge/captures/619-110.pdf"))
    pg = doc[0]
    words = upright_words(pg)
    print("===== 619-110 detailed =====")
    # Tables appear around y 920-1100 from prior dump
    dump_by_y(words, 600, 920, label="notes+PV legend")
    dump_by_y(words, 900, 1180, label="tables bottom")
    # Plan area (lower x? or mid)
    dump_by_y(words, 100, 600, label="plan")
    # Full body tokens of interest
    body = " ".join(w[4] for w in words)
    import re
    for m in re.finditer(r"\d+/\d+(?:/\d+)?", body):
        print("slash:", m.group())
    for m in re.finditer(r"\d+x\d+", body):
        print("size:", m.group())
    for m in re.finditer(r"TABLE\s+\d+-\d+", body):
        print("tableid:", m.group())
    for m in re.finditer(r"619-\d+", body):
        print("sheetref:", m.group())
    for tok in ("750'", "1000'", "1500'", "500'", "1/2", "2 MILE", "MINIMUM", "MAXIMUM", "W8-23"):
        print(f"has {tok}: {tok in body}")


def extract_112():
    doc = fitz.open(str(ROOT / "Bridge/captures/619-112.pdf"))
    for pi in range(doc.page_count):
        words = upright_words(doc[pi])
        print(f"\n===== 619-112 page {pi} detailed =====")
        dump_by_y(words, 600, 920, label="notes+PV")
        dump_by_y(words, 900, 1180, label="tables")
        dump_by_y(words, 100, 600, label="plan")
        body = " ".join(w[4] for w in words)
        import re
        for m in re.finditer(r"TABLE\s+\d+-\d+", body):
            print("tableid:", m.group())
        for m in re.finditer(r"\d+/\d+", body):
            print("slash:", m.group())
        for m in re.finditer(r"\d+x\d+", body):
            print("size:", m.group())
        for m in re.finditer(r"619-\d+", body):
            print("sheetref:", m.group())
        for tok in ("750'", "1000'", "1500'", "500'", "1/2", "2 MILE", "W20-5AR", "W4-2R", "NYW8-33"):
            print(f"has {tok}: {tok in body}")


if __name__ == "__main__":
    extract_110()
    extract_112()
