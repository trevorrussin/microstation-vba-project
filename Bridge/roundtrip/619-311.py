"""Round-trip check for Data/sheet-specs/619-311.json: re-extract every table
cell and all five notes from the PDF's vector text layer and diff them
against the JSON. This is the only unfakeable gate in the three-layer
validation (see scripts/validate_sheet_spec.py) -- structural and invariant
checks can pass on a self-consistent but wrong transcription; only comparing
back against the source PDF catches that.

Uses the shared primitives in scripts/pdf_table_extract.py. The per-table
comparison logic below is NOT generic across sheets -- every 619 table has
its own column layout, and a blind generic differ would either miss real
transcription errors on an unfamiliar layout or need so many special cases
it stops being trustworthy (see that module's docstring). Every sheet gets
its own short script here, under this naming convention, built the same way.

Run: python Bridge/roundtrip/619-311.py
"""
from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import words_in_window, group_rows, squash, assert_row_count

spec = json.loads((ROOT / "Data/sheet-specs/619-311.json").read_text())
pg = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))[0]
W = pg.get_text("words")
fails = []


def eq(label, pdf, js):
    if str(pdf) != str(js):
        fails.append(f"{label}: PDF={pdf!r} JSON={js!r}")


t02 = spec["tables"][spec["tableRoles"]["taperAndBuffer"]]
t03 = spec["tables"][spec["tableRoles"]["advanceWarningSpacing"]]
t04 = spec["tables"][spec["tableRoles"]["rollAheadDistance"]]
t05 = spec["tables"][spec["tableRoles"]["signSizes"]]
t01 = spec["tables"][spec["tableRoles"]["protectiveVehicle"]]

# ---- 311-02 (rows wrap across two y-bands at 45 mph, so merge on a coarse key)
lw = ["10", "11", "12"]
bands = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
raw_rows = group_rows(words_in_window(W, 780, 400, 1224, 500), y_tol=3.0)
# re-merge with a coarser key to join split rows before counting
from collections import defaultdict
merged = defaultdict(list)
for r in raw_rows:
    merged[round(r[0][1] / 8)].extend(r)
recs = [[w[4] for w in sorted(merged[k], key=lambda w: w[0])] for k in sorted(merged)]
assert_row_count(recs, 7, "311-02")
for toks, js in zip(recs, t02["rows"]):
    s = int(toks[0])
    eq("311-02 speed", s, js["speedMph"])
    b = toks[1].split("/")
    eq(f"311-02 {s} buffer ft", b[0], js["longitudinalBufferSpace"]["ft"])
    eq(f"311-02 {s} buffer skips", b[1], js["longitudinalBufferSpace"]["skipLines"])
    for i, w_ in enumerate(lw):
        c = toks[2 + i].split("/")
        e = js["laneTaper"][w_]
        eq(f"311-02 {s} lane{w_} ft", c[0], e["ft"])
        eq(f"311-02 {s} lane{w_} skip", c[1], e["skipLines"])
        eq(f"311-02 {s} lane{w_} dev", c[2], e["devices"])
    for i, bd in enumerate(bands):
        c = toks[5 + i].split("/")
        e = js["shoulderTaper"][bd]
        eq(f"311-02 {s} sh[{bd}] ft", c[0], e["ft"])
        eq(f"311-02 {s} sh[{bd}] skip", c[1], e["skipLines"])
        eq(f"311-02 {s} sh[{bd}] dev", c[2], e["devices"])
print("311-02 speed rows compared:", [t[0] for t in recs])

# ---- 311-03
t3rows = group_rows(words_in_window(W, 780, 555, 1030, 596))
assert_row_count(t3rows, 4, "311-03")
for r, js in zip(t3rows, t03["rows"]):
    nums = [w[4] for w in r if w[4].isdigit() and len(w[4]) == 3]
    eq("311-03 A", nums[0], js["A"])
    eq("311-03 B", nums[1], js["B"])
    eq("311-03 C", nums[2], js["C"])
    legend = squash(" ".join(w[4] for w in r if w[0] > 940))
    for k in ("XX", "YY"):
        if squash(js[k]) not in legend:
            fails.append(f"311-03 {k}: {js[k]!r} not found in PDF legend {legend!r}")
print("311-03 rows compared:", len(t3rows))

# ---- 311-04
t4rows = group_rows(words_in_window(W, 1030, 573, 1224, 612))
assert_row_count(t4rows, 3, "311-04")
for r, js in zip(t4rows, t04["rows"]):
    cells = [w[4] for w in r if "/" in w[4]]
    eq("311-04 min", cells[0], f"{js['min']['ft']}/{js['min']['skipLines']}")
    eq("311-04 max", cells[1], f"{js['max']['ft']}/{js['max']['skipLines']}")
print("311-04 rows compared:", len(t4rows))

# ---- 311-05 (W4-2R's code and its sizes land on adjacent y-bands)
sizes_by_y, name_by_y = {}, {}
for r in group_rows(words_in_window(W, 740, 665, 1010, 745)):
    y = round(r[0][1])
    s = [w[4] for w in r if "x" in w[4]]
    n = " ".join(w[4] for w in r if "x" not in w[4])
    if s:
        sizes_by_y[y] = s
    if n:
        name_by_y[y] = n
assert_row_count(list(name_by_y), 6, "311-05")
for js in t05["rows"]:
    ys = [y for y, n in name_by_y.items() if js["signCode"] in n]
    if not ys:
        fails.append(f"311-05 {js['signCode']}: no PDF label row")
        continue
    near = min(sizes_by_y, key=lambda y: abs(y - ys[0]))
    if abs(near - ys[0]) > 2:
        fails.append(f"311-05 {js['signCode']}: no sizes within 2pt of label")
        continue
    s = sizes_by_y[near]
    eq(f"311-05 {js['signCode']} non-freeway", s[0], js["NON-FREEWAY"])
    eq(f"311-05 {js['signCode']} freeway", s[1], js["FREEWAY"])

# ---- 311-01
cols = [(990, 1050, "ge45"), (1050, 1122, "b35to40"), (1122, 1200, "le30")]
for y, js in zip([110.7, 139.5, 168.2, 196.0], t01["rows"]):
    sel = sorted((w for w in W if abs(w[1] - y) < 2.5 and w[0] > 985), key=lambda w: w[0])
    for x0, x1, cid in cols:
        txt = " ".join(w[4] for w in sel if x0 <= w[0] < x1).strip().rstrip(",")
        eq(f"311-01 [{js['closureType'][:6]}/{js['exposureCondition'][:12]}] {cid}", txt, js[cid])
print("311-01 cells compared: 12")

# ---- notes verbatim. The notes column shares y-bands with plan callouts in
# the column to its left, so a flat word join interleaves the two.
notes_col = sorted(words_in_window(W, 405, 585, 660, 750), key=lambda w: (round(w[1] / 3), w[0]))
body = " ".join(w[4] for w in notes_col)
for n in spec["notes"]["printed"]:
    if squash(n) not in squash(body):
        fails.append(f"note not found verbatim in PDF: {n[:60]}...")
print("notes compared:", len(spec["notes"]["printed"]))

print()
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
