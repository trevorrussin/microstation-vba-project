"""Round-trip for 619-504.json (long-term barrier, no PV/roll-ahead)."""
from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import words_in_window, group_rows, squash, assert_row_count

spec = json.loads((ROOT / "Data/sheet-specs/619-504.json").read_text(encoding="utf-8"))
ref402 = json.loads((ROOT / "Data/sheet-specs/619-402.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
P1, P2 = doc[0].get_text("words"), doc[1].get_text("words")
body = " ".join(w[4] for pg in doc for w in pg.get_text("words"))
fails = []


def eq(label, pdf, js):
    if str(pdf) != str(js):
        fails.append(f"{label}: PDF={pdf!r} JSON={js!r}")


# Identical siblings
if spec["tables"]["504-02"]["rows"] != ref402["tables"]["402-03"]["rows"]:
    fails.append("504-02 != 402-03")
else:
    print("504-02 == 402-03")
if spec["tables"]["504-01"]["rows"] != ref402["tables"]["402-04"]["rows"]:
    fails.append("504-01 != 402-04")
else:
    print("504-01 == 402-04")

if "800/20/21" not in body:
    fails.append("missing 800/20/21")
if "W20-5a" in body:
    fails.append("unexpected W20-5a on one-lane long-term sheet")
if "W20-5" not in body:
    fails.append("missing W20-5")
if "FLARE" not in body.upper() and "504-03" not in body:
    fails.append("missing flare / 504-03 cue")
if "POSITIVE" not in body.upper() and "BARRIER" not in body.upper():
    fails.append("missing barrier language")

# No PV / roll-ahead roles
roles = spec["tableRoles"]
if roles.get("protectiveVehicle") or roles.get("rollAheadDistance"):
    fails.append("504 must not declare PV or rollAhead roles")

# Sign sizes W20-5
codes = {r["signCode"] for r in spec["tables"]["504-05"]["rows"]}
if "W20-5" not in codes and "W20-5R" not in codes:
    fails.append(f"504-05 missing W20-5: {codes}")
print("504-05 codes:", sorted(codes))

# Notes phrases
phrases = [
    "LONG-TERM IS STATIONARY WORK",
    "MORE THAN 3",
    "TEMPORARY POSITIVE BARRIER SHALL NOT BE PLACED ALONG THE MERGING TAPER",
    "MOVABLE BARRIER",
    "W20-5L",
]
sb = squash(body)
for p in phrases:
    if squash(p) not in sb:
        fails.append(f"phrase missing: {p}")
print("phrases ok:", len(phrases))

# Order table: no roll ahead / buffer
labels = [r.get("label") or r.get("signCode") for al in spec["orderTable"]["alignments"] for r in al["rows"]]
if "ROLL AHEAD DISTANCE" in labels or "BUFFER SPACE" in labels:
    fails.append(f"orderTable must not include roll/buffer: {labels}")
if "MERGING TAPER" not in labels:
    fails.append("orderTable missing MERGING TAPER")
print("order labels:", labels)

# Flare table present on p2
joined03 = " ".join(w[4] for w in words_in_window(P2, 80, 350, 500, 500))
if "FLARE" not in joined03.upper():
    fails.append("504-03 flare title missing in window")
else:
    print("504-03 present")

print()
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
