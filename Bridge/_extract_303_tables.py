"""Extract all 619-303 tables into a draft JSON fragment."""
from __future__ import annotations

import json
import pathlib
import re
import sys
from collections import defaultdict

import fitz

sys.path.insert(0, "scripts")
from pdf_table_extract import assert_row_count, group_rows, row_text, words_in_window

pg = fitz.open("Bridge/captures/619-303.pdf")[0]
W = pg.get_text("words")


def parse_triplet(tok: str) -> dict:
    # 120/3/4 or 155/4 (buffer sometimes 2-part)
    parts = tok.split("/")
    if len(parts) == 3:
        return {"ft": int(parts[0]), "skipLines": int(parts[1]), "devices": int(parts[2])}
    if len(parts) == 2:
        return {"ft": int(parts[0]), "skipLines": int(parts[1])}
    raise ValueError(tok)


# ---- 303-04 taper+buffer (same shape as 302-02) ----
ww = words_in_window(W, 740, 490, 1220, 610, pad=6)
rows = group_rows(ww, y_tol=5)
data_rows = []
for r in rows:
    toks = [w[4] for w in r]
    if toks and toks[0].isdigit() and int(toks[0]) in (25, 30, 35, 40, 45, 50, 55, 65):
        data_rows.append(toks)
assert_row_count(data_rows, 8, "303-04 speed rows")
t304 = []
for toks in data_rows:
    # speed, buffer, lane10, lane11, lane12, sh<=4, sh5-7, sh>=8
    speed = int(toks[0])
    # tokens may be split; re-join by reading row_text
    rt = " ".join(toks)
    nums = re.findall(r"\d+(?:/\d+){1,2}", rt)
    # first is speed alone sometimes already consumed — nums should be buffer + 3 lane + 3 shoulder = 7
    # row_text like: 25 155/4 120/3/4 120/3/4 120/3/4 40/1/2 40/1/2 40/1/2
    all_toks = re.findall(r"\d+(?:/\d+)*", rt)
    # all_toks[0]=speed, [1]=buffer 2-part, [2:5]=lane triplets, [5:8]=shoulder
    assert all_toks[0] == str(speed), all_toks
    buf = parse_triplet(all_toks[1])
    lane = {str(w): parse_triplet(all_toks[i]) for i, w in enumerate((10, 11, 12), start=2)}
    sh = {
        "<= 4 ft": parse_triplet(all_toks[5]),
        "5 - 7 ft": parse_triplet(all_toks[6]),
        ">= 8 ft": parse_triplet(all_toks[7]),
    }
    # buffer is longitudinalBufferSpace — devices optional
    t304.append({
        "speedMph": speed,
        "longitudinalBufferSpace": {"ft": buf["ft"], "skipLines": buf["skipLines"]},
        "laneTaper": lane,
        "shoulderTaper": sh,
    })
print("303-04 OK", len(t304), "e.g. 65/12", t304[-1]["laneTaper"]["12"])

# ---- 303-03 advance warning ----
ww = words_in_window(W, 900, 300, 1220, 420, pad=4)
rows = group_rows(ww, y_tol=4)
print("\n303-03 rows:")
for r in rows:
    print(" ", row_text(r)[:110])

# ---- 303-02 roll ahead ----
ww = words_in_window(W, 780, 295, 920, 420, pad=4)
rows = group_rows(ww, y_tol=4)
print("\n303-02 rows:")
for r in rows:
    print(" ", row_text(r)[:110])

# ---- 303-01 protective vehicle ----
ww = words_in_window(W, 740, 30, 1220, 250, pad=4)
rows = group_rows(ww, y_tol=5)
print("\n303-01 rows:")
for r in rows:
    print(" ", row_text(r)[:130])

# ---- 303-05 ----
sign_rows = [
    {"signCode": "G20-2", "NON-FREEWAY": "36x18", "FREEWAY": "48x24"},
    {"signCode": "NYW8-33", "NON-FREEWAY": "48x24", "FREEWAY": "48x24"},
    {"signCode": "W4-2R", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
    {"signCode": "W20-1", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
    {"signCode": "W20-5aR", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
    {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
]

out = {
    "tableRoles": {
        "protectiveVehicle": "303-01",
        "rollAheadDistance": "303-02",
        "advanceWarningSpacing": "303-03",
        "taperAndBuffer": "303-04",
        "signSizes": "303-05",
        "note": "Numbering trap vs 302: here 02=roll ahead, 04=taper+buffer, 05=sign sizes (302 had 05=roll ahead, 02=taper+buffer, 04=sign sizes).",
    },
    "taperAndBufferRows": t304,
    "signSizeRows": sign_rows,
}
pathlib.Path("Data/sheet-specs/_draft_619303_tables.json").write_text(
    json.dumps(out, indent=2), encoding="utf-8")
print("\nwrote draft with 303-04 + 303-05; printouts above for 01/02/03 manual encode")
