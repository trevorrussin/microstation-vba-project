"""Round-trip for 619-041.json (Family 4 mowing / parkway-adjacent)."""
from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import squash

spec = json.loads((ROOT / "Data/sheet-specs/619-041.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
body = " ".join(w[4] for pg in doc for w in pg.get_text("words"))
sb = squash(body)
fails = []

roles = spec["tableRoles"]
if roles.get("taperAndBuffer") or roles.get("advanceWarningSpacing"):
    fails.append("041 must not declare taperAndBuffer or advanceWarningSpacing")

for tok in ("W8-23", "36x36", "48x48", "NON-FREEWAY", "200/5", "280/7", "120/3", "MOVING"):
    if tok not in body and squash(tok) not in sb:
        fails.append(f"missing {tok}")
if "MERGING TAPER" in body:
    fails.append("unexpected MERGING TAPER")

# Roll ahead 3 bands
ra = spec["tables"]["041-02"]["rows"]
if len(ra) != 3:
    fails.append(f"041-02 expected 3 rows, got {len(ra)}")
if ra[2]["maxMph"] != 40 or ra[2]["min"]["ft"] != 120:
    fails.append(f"<=40 band unexpected {ra[2]}")

# PV speed bands present
if "speedBands" not in spec["tables"]["041-01"]:
    fails.append("041-01 missing speedBands")
pv = spec["tables"]["041-01"]["rows"]
if not all(r.get("ge45") for r in pv):
    fails.append("041-01 rows missing ge45 cells")

for p in ("5 MINUTES", "619-201", "40 FEET", "NON-PEAK"):
    if squash(p) not in sb:
        fails.append(f"note phrase missing: {p}")

labels = [r.get("label") or r.get("signCode") for al in spec["orderTable"]["alignments"] for r in al["rows"]]
if labels != ["ROLL AHEAD DISTANCE", "W8-23"]:
    fails.append(f"unexpected order: {labels}")

size_codes = {r["signCode"] for r in spec["tables"]["041-03"]["rows"]}
sign_codes = {s["signCode"] for s in spec["signs"]["items"]}
if size_codes != sign_codes:
    fails.append(f"sign/size mismatch {size_codes} vs {sign_codes}")

print("order", labels)
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
