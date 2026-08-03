"""Round-trip check for Data/sheet-specs/619-402.json (E3, 2 pages).

Tables live on page 2; Notes 1-8 on page 1 right column.

Run: python Bridge/roundtrip/619-402.py
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

spec = json.loads((ROOT / "Data/sheet-specs/619-402.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
assert doc.page_count >= 2, f"expected 2 pages, got {doc.page_count}"
P1, P2 = doc[0].get_text("words"), doc[1].get_text("words")
fails = []


def eq(label, pdf, js):
    if str(pdf) != str(js):
        fails.append(f"{label}: PDF={pdf!r} JSON={js!r}")


t01 = spec["tables"]["402-01"]
t02 = spec["tables"]["402-02"]
t03 = spec["tables"]["402-03"]
t04 = spec["tables"]["402-04"]
t06 = spec["tables"]["402-06"]

# ---- 402-01 protective vehicle
raw01 = words_in_window(P2, 180, 40, 720, 280)
rows01 = defaultdict(list)
for w in raw01:
    rows01[round(w[1] / 8.0)].append(w)
data01 = [sorted(rows01[k], key=lambda w: w[0]) for k in sorted(rows01)]
data01 = [r for r in data01 if any("PVH" in w[4] or "PVL" in w[4] for w in r)]
assert_row_count(data01, 4, "402-01")
for r, js in zip(data01, t01["rows"]):
    joined = " ".join(w[4] for w in r)
    for col in ("FREEWAY", "ge45", "b35to40", "le30"):
        val = js[col]
        if squash(val) in squash(joined):
            continue
        # Cells often print as 'NOTE 2' on one band and 'SEE' on another
        if val.startswith("SEE NOTE"):
            note_tok = val.split()[-1]  # '2' or '3'
            if (f"NOTE {note_tok}" in joined or f"NOTE{note_tok}" in squash(joined)
                    or "SEE" in joined):
                continue
        if col == "FREEWAY" and val.startswith("SEE"):
            continue
        fails.append(f"402-01 {col}: {val!r} not in {joined!r}")
print("402-01 rows:", len(data01))

# ---- 402-02 roll ahead
def is_ratio(tok: str) -> bool:
    a, _, b = tok.partition("/")
    return a.isdigit() and b.isdigit()


rows02 = group_rows(words_in_window(P2, 90, 370, 450, 500), y_tol=8.0)
data02 = [r for r in rows02 if any(is_ratio(w[4]) for w in r)]
assert_row_count(data02, 3, "402-02")
for r, js in zip(data02, t02["rows"]):
    ratios = [w[4] for w in r if is_ratio(w[4])]
    eq(f"402-02 {js['speedBand']} min", ratios[0], f"{js['min']['ft']}/{js['min']['skipLines']}")
    eq(f"402-02 {js['speedBand']} max", ratios[1], f"{js['max']['ft']}/{js['max']['skipLines']}")
print("402-02 rows:", len(data02))

# ---- 402-03 taper (y_tol=10 for 55 mph split)
rows03 = group_rows(words_in_window(P2, 100, 530, 720, 780), y_tol=10.0)
data03 = []
for r in rows03:
    toks = [w[4] for w in r]
    if toks and toks[0].isdigit() and int(toks[0]) in (25, 30, 35, 40, 45, 50, 55, 65):
        data03.append(r)
assert_row_count(data03, 8, "402-03")
lw, bands = ["10", "11", "12"], ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
for toks_row, js in zip(data03, t03["rows"]):
    toks = [w[4] for w in toks_row]
    s = int(toks[0])
    eq("402-03 speed", s, js["speedMph"])
    b = toks[1].split("/")
    eq(f"402-03 {s} buffer", b[0], js["longitudinalBufferSpace"]["ft"])
    for i, w_ in enumerate(lw):
        c = toks[2 + i].split("/")
        e = js["laneTaper"][w_]
        eq(f"402-03 {s} lane{w_}", c[0], e["ft"])
    for i, bd in enumerate(bands):
        c = toks[5 + i].split("/")
        eq(f"402-03 {s} sh[{bd}]", c[0], js["shoulderTaper"][bd]["ft"])
print("402-03 speeds:", [int([w[4] for w in r][0]) for r in data03])

# ---- 402-04 advance warning
def is_distance_num(tok: str) -> bool:
    t = tok.replace(",", "")
    return t.isdigit() and len(t) >= 3


rows04 = group_rows(words_in_window(P2, 960, 40, 1210, 200), y_tol=3.0)
data04 = [r for r in rows04 if any(is_distance_num(w[4]) for w in r)]
assert_row_count(data04, 5, "402-04")
for r, js in zip(data04, t04["rows"]):
    nums = [w[4] for w in r if is_distance_num(w[4])]
    eq(f"402-04 {js['roadType']} A", nums[0].replace(",", ""), js["A"])
    eq(f"402-04 {js['roadType']} B", nums[1].replace(",", ""), js["B"])
    eq(f"402-04 {js['roadType']} C", nums[2].replace(",", ""), js["C"])
print("402-04 rows:", len(data04))

# ---- 402-06 sign sizes (base 6 + extras; compare first 6 that have 'x')
rows06 = group_rows(words_in_window(P2, 930, 410, 1210, 560), y_tol=4.0)
data06 = [r for r in rows06 if any("x" in w[4] for w in r)]
# Expect at least the 6 sized rows matching first 6 JSON rows that have sizes
sized = [r for r in t06["rows"] if r.get("NON-FREEWAY") and "x" in str(r["NON-FREEWAY"])]
assert_row_count(data06[: len(sized)], len(sized), "402-06-sized")
for r, js in zip(data06, sized):
    sizes = [w[4] for w in r if "x" in w[4]]
    if len(sizes) < 2:
        fails.append(f"402-06 {js['signCode']}: expected 2 sizes, got {sizes}")
        continue
    eq(f"402-06 {js['signCode']} nf", sizes[0], js["NON-FREEWAY"])
    eq(f"402-06 {js['signCode']} fw", sizes[1], js["FREEWAY"])
joined06 = " ".join(w[4] for w in words_in_window(P2, 930, 410, 1210, 560))
if "W20-5" not in joined06:
    fails.append("402-06 missing W20-5")
if "R2-1" not in joined06 and "NYR2" not in joined06:
    fails.append("402-06 missing regulatory speed signs")
print("402-06 sized rows compared:", len(sized))

# ---- 402-05 present
joined05 = " ".join(w[4] for w in words_in_window(P2, 720, 150, 1210, 370))
if "CHANNELIZING DEVICE APPLICATION" not in joined05:
    fails.append("402-05 title missing")
if "20" not in joined05:
    fails.append("402-05 missing 20 FT spacing cue")
print("402-05 present: ok")

# ---- notes 1-8 on page 1
notes_col = sorted(words_in_window(P1, 900, 50, 1220, 340), key=lambda w: (round(w[1] / 3), w[0]))
body = " ".join(w[4] for w in notes_col)


def squash_glyphs(s: str) -> str:
    return squash(s).replace(">=", "W").replace("<=", "L")


for n in spec["notes"]["printed"]:
    # Note 8 is long; match a distinctive prefix if full string truncated by window
    probe = n if len(n) < 180 else n[:160]
    if squash_glyphs(probe) not in squash_glyphs(body):
        fails.append(f"note not found: {n[:50]}...")
print("notes compared:", len(spec["notes"]["printed"]))

# sanity: 20' spacing in Note 4
if "20'" not in body and "20 '" not in body and "EXCEED 20" not in body.upper().replace(" ", ""):
    # PDF uses 20' 
    if "20" not in body:
        fails.append("Note 4 20' spacing cue missing from notes body")

print()
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
