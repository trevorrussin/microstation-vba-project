"""Compare Family 1 sibling taper cells against 619-311 with coarse y-merge."""
from __future__ import annotations

import json
import pathlib
import sys
from collections import defaultdict

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import words_in_window, group_rows, row_text

ref = json.loads((ROOT / "Data/sheet-specs/619-311.json").read_text(encoding="utf-8"))
t02 = ref["tables"]["311-02"]["rows"]
bands = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
lw = ["10", "11", "12"]


def extract(pdf, page, box, ytol=10):
    W = fitz.open(pdf)[page].get_text("words")
    raw = group_rows(words_in_window(W, *box), y_tol=3.0)
    merged = defaultdict(list)
    for r in raw:
        merged[round(r[0][1] / ytol)].extend(r)
    rows = []
    for k in sorted(merged):
        toks = [w[4] for w in sorted(merged[k], key=lambda w: w[0])]
        if toks and toks[0].isdigit() and int(toks[0]) in (25, 30, 35, 40, 45, 50, 55):
            cells = [t for t in toks if "/" in t]
            rows.append((int(toks[0]), cells))
    return rows


def cmp(name, rows, has_sh=True):
    fails = 0
    for (spd, cells), js in zip(rows, t02):
        exp = [f"{js['longitudinalBufferSpace']['ft']}/{js['longitudinalBufferSpace']['skipLines']}"]
        for w in lw:
            e = js["laneTaper"][w]
            exp.append(f"{e['ft']}/{e['skipLines']}/{e['devices']}")
        if has_sh:
            for b in bands:
                e = js["shoulderTaper"][b]
                exp.append(f"{e['ft']}/{e['skipLines']}/{e['devices']}")
        if cells != exp:
            fails += 1
            print(name, spd, "DIFF pdf", cells)
            print("         exp", exp)
    print(name, "rows", len(rows), "fails", fails)


base = ROOT / "Bridge/captures"
cmp("317", extract(base / "619-317.pdf", 1, (100, 160, 780, 350)))
cmp("325", extract(base / "619-325.pdf", 1, (100, 450, 780, 630)))
cmp("414", extract(base / "619-414.pdf", 1, (100, 148, 780, 340)))
cmp("423", extract(base / "619-423.pdf", 1, (100, 143, 780, 330)))
cmp("523", extract(base / "619-523.pdf", 1, (100, 143, 780, 330)))
cmp("312", extract(base / "619-312.pdf", 1, (90, 590, 400, 720)), has_sh=False)

# Sign size codes by y alignment for a few sheets
for sheet, page, box in [
    ("317", 1, (980, 340, 1224, 520)),
    ("414", 1, (980, 300, 1224, 520)),
    ("312", 1, (400, 520, 650, 720)),
    ("202", 0, (990, 340, 1224, 450)),
    ("523", 1, (980, 270, 1224, 480)),
]:
    print(f"\n=== sizes {sheet} ===")
    W = fitz.open(base / f"619-{sheet}.pdf")[page].get_text("words")
    for r in group_rows(words_in_window(W, *box), y_tol=4):
        print(row_text(r)[:120])
