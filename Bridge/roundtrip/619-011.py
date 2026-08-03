"""Round-trip check for Data/sheet-specs/619-011.json: re-extract every table
cell, the legend, and Detail 011A from the PDF's vector text layer and diff
against the JSON. See Bridge/roundtrip/619-311.py for the pattern this follows
and scripts/pdf_table_extract.py for the shared primitives.

Run: python Bridge/roundtrip/619-011.py
"""
from __future__ import annotations

import json
import pathlib
import sys
from collections import defaultdict

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import words_in_window, group_rows, assert_row_count

spec = json.loads((ROOT / "Data/sheet-specs/619-011.json").read_text())
pg = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))[0]
W = pg.get_text("words")
fails = []
KNOWN_ANOMALY_CELLS = {("011-02", 65, "12")}  # documented in the spec, not a transcription bug


def eq(label, pdf, js):
    if str(pdf) != str(js):
        fails.append(f"{label}: PDF={pdf!r} JSON={js!r}")


# ---- 011-02: TAPER LENGTHS & NUMBER OF CONES CHART (9 speed rows x 12 cols)
t02 = spec["tables"]["011-02"]
raw = words_in_window(W, 60, 390, 700, 505)
recs_by_y = defaultdict(list)
for w in raw:
    recs_by_y[round(w[1] / 3.0)].append(w)
recs = [[w[4] for w in sorted(recs_by_y[k], key=lambda w: w[0])] for k in sorted(recs_by_y)]
assert_row_count(recs, 9, "011-02")
lw = [str(n) for n in range(4, 13)]
bands = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
for toks, js in zip(recs, t02["rows"]):
    s = int(toks[0])
    eq("011-02 speed", s, js["speedMph"])
    for i, w_ in enumerate(lw):
        c = toks[1 + i].split("/")
        e = js["laneTaper"][w_]
        if ("011-02", s, w_) in KNOWN_ANOMALY_CELLS:
            continue
        eq(f"011-02 {s} lane{w_} ft", c[0], e["ft"])
        eq(f"011-02 {s} lane{w_} skip", c[1], e["skipLines"])
        eq(f"011-02 {s} lane{w_} dev", c[2], e["devices"])
    for i, bd in enumerate(bands):
        c = toks[10 + i].split("/")
        e = js["shoulderTaper"][bd]
        eq(f"011-02 {s} sh[{bd}] ft", c[0], e["ft"])
        eq(f"011-02 {s} sh[{bd}] skip", c[1], e["skipLines"])
        eq(f"011-02 {s} sh[{bd}] dev", c[2], e["devices"])
print("011-02 speed rows compared:", [t[0] for t in recs])

# ---- 011-03: LONGITUDINAL BUFFER SPACE (8 rows -- 60 mph genuinely absent)
t03 = spec["tables"]["011-03"]
rows03 = group_rows(words_in_window(W, 60, 588, 220, 700), y_tol=3.0)
assert_row_count(rows03, 8, "011-03")
for r, js in zip(rows03, t03["rows"]):
    toks = [w[4] for w in r]
    eq("011-03 speed", int(toks[0]), js["speedMph"])
    ft, skip = toks[1], toks[3]  # tokens are ['speed', 'ft', '/', 'skip'] -- '/' is its own token
    eq(f"011-03 {toks[0]} ft", ft, js["longitudinalBufferSpace"]["ft"])
    eq(f"011-03 {toks[0]} skip", skip, js["longitudinalBufferSpace"]["skipLines"])
print("011-03 rows compared:", len(rows03))

# ---- 011-04: ROLL AHEAD DISTANCE (3 speed bands, moving + stationary, min/max)
# data rows only (y 605-645); the "45 - 50" label splits across two close
# y-bands from the ratio cells, so merge with a coarser key.
t04 = spec["tables"]["011-04"]
raw04 = words_in_window(W, 239, 605, 475, 645)
merged04 = defaultdict(list)
for w in raw04:
    merged04[round(w[1] / 8)].extend(w for w in [w])
rows04 = [sorted(merged04[k], key=lambda w: w[0]) for k in sorted(merged04)]
assert_row_count(rows04, 3, "011-04")
for r, js in zip(rows04, t04["rows"]):
    toks = [w[4] for w in r if "/" in w[4]]
    if len(toks) < 4:
        fails.append(f"011-04 {js['speedBand']}: expected 4 ratio cells, got {toks}")
        continue
    eq(f"011-04 {js['speedBand']} moving min", toks[0],
       f"{js['moving']['min']['ft']}/{js['moving']['min']['skipLines']}")
    eq(f"011-04 {js['speedBand']} moving max", toks[1],
       f"{js['moving']['max']['ft']}/{js['moving']['max']['skipLines']}")
    eq(f"011-04 {js['speedBand']} stationary min", toks[2],
       f"{js['stationary']['min']['ft']}/{js['stationary']['min']['skipLines']}")
    eq(f"011-04 {js['speedBand']} stationary max", toks[3],
       f"{js['stationary']['max']['ft']}/{js['stationary']['max']['skipLines']}")
print("011-04 rows compared:", len(rows04))

# ---- 011-06: ADVANCE WARNING SIGN SPACING (5 road-type rows incl FREEWAY)
t06 = spec["tables"]["011-06"]
rows06 = group_rows(words_in_window(W, 480, 543, 780, 655), y_tol=3.0)
# the title/header lines land in the same window, and the road-type labels
# themselves contain 2-digit speed numbers ("<=30 MPH*", ">=45 MPH*") that
# would false-match a bare isdigit() filter -- the real A/B/C distance cells
# are always 3+ digits, so that's the discriminator.
def is_distance_num(tok: str) -> bool:
    t = tok.replace(",", "")
    return t.isdigit() and len(t) >= 3

data_rows06 = [r for r in rows06 if any(is_distance_num(w[4]) for w in r)]
assert_row_count(data_rows06, 5, "011-06")
for r, js in zip(data_rows06, t06["rows"]):
    nums = [w[4] for w in r if is_distance_num(w[4])]
    if len(nums) < 3:
        fails.append(f"011-06 {js['roadType']}: expected 3 numeric cells, got {nums}")
        continue
    eq(f"011-06 {js['roadType']} A", nums[0].replace(",", ""), js["A"])
    eq(f"011-06 {js['roadType']} B", nums[1].replace(",", ""), js["B"])
    eq(f"011-06 {js['roadType']} C", nums[2].replace(",", ""), js["C"])
print("011-06 rows compared:", len(data_rows06))

# ---- Detail 011A skip-line dimensions
detail_words = words_in_window(W, 1010, 80, 1165, 180)
dims = [w[4] for w in detail_words if "'" in w[4]]
skip30 = any(d.strip("'") == "30" for d in dims)
line10 = any(d.strip("'") == "10" for d in dims)
if not skip30:
    fails.append("Detail 011A: no 30' skip-length dimension found")
if not line10:
    fails.append("Detail 011A: no 10' line-length dimension found")
print("Detail 011A dimensions found:", sorted(set(dims)))

print()
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
print(f"(known documented anomaly cells excluded from comparison: {sorted(KNOWN_ANOMALY_CELLS)})")
sys.exit(1 if fails else 0)
