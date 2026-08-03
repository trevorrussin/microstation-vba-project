"""Round-trip check for Data/sheet-specs/619-303.json.

Table layout on this sheet differs from 619-302: roll-ahead sits LEFT of
advance-warning (both ~y=300), taper table mid-right (~y=420), sign sizes
bottom (~y=620). Windows below were probed against the live PDF.

Run: python Bridge/roundtrip/619-303.py
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

spec = json.loads((ROOT / "Data/sheet-specs/619-303.json").read_text(encoding="utf-8"))
pg = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))[0]
W = pg.get_text("words")
fails = []


def eq(label, pdf, js):
    if str(pdf) != str(js):
        fails.append(f"{label}: PDF={pdf!r} JSON={js!r}")


t01 = spec["tables"]["303-01"]
t02 = spec["tables"]["303-02"]
t03 = spec["tables"]["303-03"]
t04 = spec["tables"]["303-04"]
t05 = spec["tables"]["303-05"]

# ---- 303-01 protective vehicle (4 data rows)
raw01 = words_in_window(W, 900, 30, 1200, 200)
rows01 = defaultdict(list)
for w in raw01:
    rows01[round(w[1] / 8.0)].append(w)
data01 = [sorted(rows01[k], key=lambda w: w[0]) for k in sorted(rows01)]
data01 = [r for r in data01 if any(w[4] in ("P,", "P", "NA", "SEE") for w in r)]
assert_row_count(data01, 4, "303-01")
for r, js in zip(data01, t01["rows"]):
    joined = " ".join(w[4] for w in r)
    for col in ("FREEWAY", "ge45", "b35to40", "le30"):
        val = js[col]
        if squash(val) not in squash(joined):
            fails.append(
                f"303-01 [{js['closureType'][:4]}/{js['exposureCondition'][:4]}] {col}: "
                f"{val!r} not found in row text {joined!r}")
print("303-01 rows compared:", len(data01))

# ---- 303-02 roll ahead (3 bands)
def is_ratio(tok: str) -> bool:
    a, _, b = tok.partition("/")
    return a.isdigit() and b.isdigit()


rows02 = group_rows(words_in_window(W, 800, 290, 980, 400), y_tol=8.0)
data02 = [r for r in rows02 if any(is_ratio(w[4]) for w in r)]
assert_row_count(data02, 3, "303-02")
for r, js in zip(data02, t02["rows"]):
    ratios = [w[4] for w in r if is_ratio(w[4])]
    if len(ratios) < 2:
        fails.append(f"303-02 {js['speedBand']}: expected 2 ratio cells, got {ratios}")
        continue
    eq(f"303-02 {js['speedBand']} min", ratios[0], f"{js['min']['ft']}/{js['min']['skipLines']}")
    eq(f"303-02 {js['speedBand']} max", ratios[1], f"{js['max']['ft']}/{js['max']['skipLines']}")
print("303-02 rows compared:", len(data02))

# ---- 303-03 advance warning (5 rows; roadType labels sit at x≈982)
def is_distance_num(tok: str) -> bool:
    t = tok.replace(",", "")
    return t.isdigit() and len(t) >= 3


rows03 = group_rows(words_in_window(W, 970, 300, 1200, 410), y_tol=3.0)
data03 = [r for r in rows03 if any(is_distance_num(w[4]) for w in r)]
assert_row_count(data03, 5, "303-03")
for r, js in zip(data03, t03["rows"]):
    nums = [w[4] for w in r if is_distance_num(w[4])]
    eq(f"303-03 {js['roadType']} A", nums[0].replace(",", ""), js["A"])
    eq(f"303-03 {js['roadType']} B", nums[1].replace(",", ""), js["B"])
    eq(f"303-03 {js['roadType']} C", nums[2].replace(",", ""), js["C"])
print("303-03 rows compared:", len(data03))

# ---- 303-04 taper+buffer (8 speeds, no 60)
rows04 = group_rows(words_in_window(W, 800, 450, 1200, 615), y_tol=8.0)
# keep rows whose first token is a speed
data04 = []
for r in rows04:
    toks = [w[4] for w in r]
    if toks and toks[0].isdigit() and int(toks[0]) in (25, 30, 35, 40, 45, 50, 55, 65):
        data04.append(r)
assert_row_count(data04, 8, "303-04")
lw = ["10", "11", "12"]
bands = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
for toks_row, js in zip(data04, t04["rows"]):
    toks = [w[4] for w in toks_row]
    s = int(toks[0])
    eq("303-04 speed", s, js["speedMph"])
    b = toks[1].split("/")
    eq(f"303-04 {s} buffer ft", b[0], js["longitudinalBufferSpace"]["ft"])
    eq(f"303-04 {s} buffer skips", b[1], js["longitudinalBufferSpace"]["skipLines"])
    for i, w_ in enumerate(lw):
        c = toks[2 + i].split("/")
        e = js["laneTaper"][w_]
        eq(f"303-04 {s} lane{w_} ft", c[0], e["ft"])
        eq(f"303-04 {s} lane{w_} skip", c[1], e["skipLines"])
        eq(f"303-04 {s} lane{w_} dev", c[2], e["devices"])
    for i, bd in enumerate(bands):
        c = toks[5 + i].split("/")
        e = js["shoulderTaper"][bd]
        eq(f"303-04 {s} sh[{bd}] ft", c[0], e["ft"])
        eq(f"303-04 {s} sh[{bd}] skip", c[1], e["skipLines"])
        eq(f"303-04 {s} sh[{bd}] dev", c[2], e["devices"])
print("303-04 speeds:", [[w[4] for w in r][0] for r in data04])

# ---- 303-05 sign sizes (6 rows incl WARNING FLAG + W20-5aR)
rows05 = group_rows(words_in_window(W, 800, 610, 1050, 730), y_tol=4.0)
data05 = [r for r in rows05 if any("x" in w[4] for w in r)]
assert_row_count(data05, 6, "303-05")
for r, js in zip(data05, t05["rows"]):
    sizes = [w[4] for w in r if "x" in w[4]]
    if len(sizes) < 2:
        fails.append(f"303-05 {js['signCode']}: expected 2 size cells, got {sizes}")
        continue
    eq(f"303-05 {js['signCode']} non-freeway", sizes[0], js["NON-FREEWAY"])
    eq(f"303-05 {js['signCode']} freeway", sizes[1], js["FREEWAY"])
    # also confirm the sign code token is present for the two-lane mid sign
    joined = " ".join(w[4] for w in r)
    if js["signCode"] == "W20-5aR" and "W20-5aR" not in joined:
        fails.append(f"303-05 expected W20-5aR in row, got {joined!r}")
print("303-05 rows compared:", len(data05), [js["signCode"] for js in t05["rows"]])

# ---- notes (9 printed). PDF renders >= as bare 'w'.
notes_col = sorted(words_in_window(W, 380, 300, 660, 530), key=lambda w: (round(w[1] / 3), w[0]))
body = " ".join(w[4] for w in notes_col)


def squash_glyphs(s: str) -> str:
    return squash(s).replace(">=", "W").replace("<=", "L")


for n in spec["notes"]["printed"]:
    if squash_glyphs(n) not in squash_glyphs(body):
        fails.append(f"note not found verbatim in PDF: {n[:60]}...")
print("notes compared:", len(spec["notes"]["printed"]))

# sanity: 2L and W20-5aR must appear on the plan
plan = " ".join(w[4] for w in W)
if "2L" not in plan:
    fails.append("plan text missing '2L' dimension label")
if "W20-5aR" not in plan:
    fails.append("plan text missing W20-5aR")

print()
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
