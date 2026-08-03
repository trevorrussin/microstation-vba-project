"""Round-trip for 619-212.json (Family 4 short-duration parkway)."""
from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import squash

spec = json.loads((ROOT / "Data/sheet-specs/619-212.json").read_text(encoding="utf-8"))
ref302 = json.loads((ROOT / "Data/sheet-specs/619-302.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
body = " ".join(w[4] for pg in doc for w in pg.get_text("words"))
sb = squash(body)
fails = []

speeds = [r["speedMph"] for r in spec["tables"]["212-03"]["rows"]]
if speeds != [45, 50, 55, 65]:
    fails.append(f"212-03 speeds {speeds}")

for row in spec["tables"]["212-03"]["rows"]:
    r302 = next(r for r in ref302["tables"]["302-02"]["rows"] if r["speedMph"] == row["speedMph"])
    if row["shoulderTaper"] != r302["shoulderTaper"]:
        fails.append(f"shoulderTaper mismatch {row['speedMph']}")
    if row["laneTaper"] != r302["laneTaper"]:
        fails.append(f"laneTaper mismatch {row['speedMph']}")

for tok in ("NYW8-33", "W20-1", "W20-5R", "W4-2R", "500'", "1500'", "SHORT DURATION", "PARKWAY"):
    if tok not in body and squash(tok) not in sb:
        fails.append(f"missing {tok}")
if "MERGING TAPER" in body:
    fails.append("unexpected MERGING TAPER on plan text")
if spec["tableRoles"].get("advanceWarningSpacing"):
    fails.append("must not declare advanceWarningSpacing")

for p in ("SHORT DURATION IS WORK", "OPERATOR(S) SHALL REMAIN", "NO WORKERS, EQUIPMENT"):
    if squash(p) not in sb:
        fails.append(f"note phrase missing: {p}")

labels = [r.get("label") or r.get("signCode") for al in spec["orderTable"]["alignments"] for r in al["rows"]]
if labels != ["ROLL AHEAD DISTANCE", "SHOULDER TAPER", "W4-2R", "W20-5R", "W20-1"]:
    fails.append(f"unexpected order: {labels}")
if "MERGING TAPER" in labels or "BUFFER SPACE" in labels:
    fails.append("order must omit merging/buffer")

# gap values
gaps = {z["id"]: z["lengthSource"].get("fixedFt") for z in spec["corridor"]["zones"] if z["kind"] == "gap"}
if gaps.get("gapA") != 500 or gaps.get("gapB") != 1500 or gaps.get("gapC") != 2640:
    fails.append(f"unexpected gaps {gaps}")

size_codes = {r["signCode"] for r in spec["tables"]["212-04"]["rows"]}
sign_codes = {s["signCode"] for s in spec["signs"]["items"]}
if size_codes != sign_codes:
    fails.append(f"sign/size mismatch {size_codes} vs {sign_codes}")

print("order", labels)
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
