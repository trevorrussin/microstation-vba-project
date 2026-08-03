"""Round-trip check for Data/sheet-specs/619-302.json: re-extract every table
cell and the 8 printed notes from the PDF's vector text layer and diff
against the JSON. See Bridge/roundtrip/619-311.py for the pattern and
scripts/pdf_table_extract.py for the shared primitives.

Run: python Bridge/roundtrip/619-302.py
"""
from __future__ import annotations

import json
import pathlib
import sys
from collections import defaultdict

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import words_in_window, group_rows, squash, assert_row_count

spec = json.loads((ROOT / "Data/sheet-specs/619-302.json").read_text())
pg = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))[0]
W = pg.get_text("words")
fails = []
# Same cross-sheet-discrepancy cell as 619-011's 011-02 (documented in both
# specs' knownAnomalies) -- 619-302's own printed value (800/20/21) is
# internally consistent and is what this round-trip checks against.


def eq(label, pdf, js):
    if str(pdf) != str(js):
        fails.append(f"{label}: PDF={pdf!r} JSON={js!r}")


t01 = spec["tables"]["302-01"]
t02 = spec["tables"]["302-02"]
t03 = spec["tables"]["302-03"]
t04 = spec["tables"]["302-04"]
t05 = spec["tables"]["302-05"]

# ---- 302-01: PROTECTIVE VEHICLE REQUIREMENTS (4 rows x FREEWAY/ge45/b35to40/le30)
raw01 = words_in_window(W, 940, 105, 1180, 205)
rows01 = defaultdict(list)
for w in raw01:
    rows01[round(w[1] / 8.0)].append(w)
data01 = [sorted(rows01[k], key=lambda w: w[0]) for k in sorted(rows01)]
# keep only rows containing a P/NA/SEE cell value
data01 = [r for r in data01 if any(w[4] in ("P,", "P", "NA", "SEE") for w in r)]
assert_row_count(data01, 4, "302-01")
for r, js in zip(data01, t01["rows"]):
    toks = [w[4] for w in r]
    joined = " ".join(toks)
    # cells are "P, TMIA" / "P" / "SEE NOTE 2" -- split on the 4 columns by
    # re-joining consecutive tokens greedily; simplest robust check is a
    # substring match per expected cell value against the whole row text.
    for col in ("FREEWAY", "ge45", "b35to40", "le30"):
        val = js[col]
        if squash(val) not in squash(joined):
            fails.append(f"302-01 [{js['closureType'][:4]}/{js['exposureCondition'][:4]}] {col}: "
                          f"{val!r} not found in row text {joined!r}")
print("302-01 rows compared:", len(data01))

# ---- 302-02: LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS (8 rows, NOT 9)
raw02 = words_in_window(W, 810, 375, 1190, 480)
rows02 = group_rows(raw02, y_tol=3.0)
assert_row_count(rows02, 8, "302-02")
lw = ["10", "11", "12"]
bands = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
for toks_row, js in zip(rows02, t02["rows"]):
    toks = [w[4] for w in toks_row]
    s = int(toks[0])
    eq("302-02 speed", s, js["speedMph"])
    b = toks[1].split("/")
    eq(f"302-02 {s} buffer ft", b[0], js["longitudinalBufferSpace"]["ft"])
    eq(f"302-02 {s} buffer skips", b[1], js["longitudinalBufferSpace"]["skipLines"])
    for i, w_ in enumerate(lw):
        c = toks[2 + i].split("/")
        e = js["laneTaper"][w_]
        eq(f"302-02 {s} lane{w_} ft", c[0], e["ft"])
        eq(f"302-02 {s} lane{w_} skip", c[1], e["skipLines"])
        eq(f"302-02 {s} lane{w_} dev", c[2], e["devices"])
    for i, bd in enumerate(bands):
        c = toks[5 + i].split("/")
        e = js["shoulderTaper"][bd]
        eq(f"302-02 {s} sh[{bd}] ft", c[0], e["ft"])
        eq(f"302-02 {s} sh[{bd}] skip", c[1], e["skipLines"])
        eq(f"302-02 {s} sh[{bd}] dev", c[2], e["devices"])
print("302-02 speed rows compared:", [t[0].split("/")[0] if "/" in t[0] else t[0] for t in [[w[4] for w in r] for r in rows02]])

# ---- 302-03: ADVANCE WARNING SIGN SPACING (5 rows incl FREEWAY)
def is_distance_num(tok: str) -> bool:
    t = tok.replace(",", "")
    return t.isdigit() and len(t) >= 3

rows03 = group_rows(words_in_window(W, 795, 503, 1012, 604), y_tol=3.0)
data03 = [r for r in rows03 if any(is_distance_num(w[4]) for w in r)]
assert_row_count(data03, 5, "302-03")
for r, js in zip(data03, t03["rows"]):
    nums = [w[4] for w in r if is_distance_num(w[4])]
    eq(f"302-03 {js['roadType']} A", nums[0].replace(",", ""), js["A"])
    eq(f"302-03 {js['roadType']} B", nums[1].replace(",", ""), js["B"])
    eq(f"302-03 {js['roadType']} C", nums[2].replace(",", ""), js["C"])
print("302-03 rows compared:", len(data03))

# ---- 302-04: REQUIRED SIGN SIZES (6 rows incl WARNING FLAG)
rows04 = group_rows(words_in_window(W, 795, 615, 995, 738), y_tol=4.0)
data04 = [r for r in rows04 if any("x" in w[4] for w in r)]
assert_row_count(data04, 6, "302-04")
for r, js in zip(data04, t04["rows"]):
    sizes = [w[4] for w in r if "x" in w[4]]
    if len(sizes) < 2:
        fails.append(f"302-04 {js['signCode']}: expected 2 size cells, got {sizes}")
        continue
    eq(f"302-04 {js['signCode']} non-freeway", sizes[0], js["NON-FREEWAY"])
    eq(f"302-04 {js['signCode']} freeway", sizes[1], js["FREEWAY"])
print("302-04 rows compared:", len(data04))

# ---- 302-05: ROLL AHEAD DISTANCE (3 rows, stationary only)
def is_ratio(tok: str) -> bool:
    a, _, b = tok.partition("/")
    return a.isdigit() and b.isdigit()

rows05 = group_rows(words_in_window(W, 1040, 555, 1189, 600), y_tol=8.0)
data05 = [r for r in rows05 if any(is_ratio(w[4]) for w in r)]
assert_row_count(data05, 3, "302-05")
for r, js in zip(data05, t05["rows"]):
    ratios = [w[4] for w in r if is_ratio(w[4])]
    if len(ratios) < 2:
        fails.append(f"302-05 {js['speedBand']}: expected 2 ratio cells, got {ratios}")
        continue
    eq(f"302-05 {js['speedBand']} min", ratios[0], f"{js['min']['ft']}/{js['min']['skipLines']}")
    eq(f"302-05 {js['speedBand']} max", ratios[1], f"{js['max']['ft']}/{js['max']['skipLines']}")
print("302-05 rows compared:", len(data05))

# ---- notes (8 printed notes). This font renders the >=/<= glyphs as bare
# 'w'/'l' characters in the text layer (same quirk hit throughout 619-011 and
# 619-302's tables) -- collapse ">=" and "<=" to single chars on both sides
# before comparing so Note 8's ">= 8'" matches the PDF's "w 8'".
notes_col = sorted(words_in_window(W, 408, 290, 660, 650), key=lambda w: (round(w[1] / 3), w[0]))
body = " ".join(w[4] for w in notes_col)


def squash_glyphs(s: str) -> str:
    return squash(s).replace(">=", "W").replace("<=", "L")


for n in spec["notes"]["printed"]:
    if squash_glyphs(n) not in squash_glyphs(body):
        fails.append(f"note not found verbatim in PDF: {n[:60]}...")
print("notes compared:", len(spec["notes"]["printed"]))

print()
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
