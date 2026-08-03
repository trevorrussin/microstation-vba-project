"""Round-trip for 619-114.json (Family 4 mobile parkway)."""
from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import squash

spec = json.loads((ROOT / "Data/sheet-specs/619-114.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
body = " ".join(w[4] for pg in doc for w in pg.get_text("words"))
sb = squash(body)
fails = []

roles = spec["tableRoles"]
if roles.get("taperAndBuffer") or roles.get("advanceWarningSpacing"):
    fails.append("114 must not declare taperAndBuffer or advanceWarningSpacing")

for tok in ("NYW8-33", "W20-5R", "MOBILE", "PARKWAY", "500'", "200/5", "280/7", "160/4", "240/6"):
    if tok not in body and squash(tok) not in sb:
        fails.append(f"missing {tok}")
if "MERGING" in body and "MERGING TAPER" in body:
    fails.append("unexpected MERGING TAPER")

# Roll ahead moving values
ra = spec["tables"]["114-02"]["rows"]
if ra[0]["min"]["ft"] != 200 or ra[0]["max"]["ft"] != 280:
    fails.append(f"roll ahead >=55 unexpected {ra[0]}")
if ra[1]["min"]["ft"] != 160 or ra[1]["max"]["ft"] != 240:
    fails.append(f"roll ahead 45-50 unexpected {ra[1]}")

# PV NA rows
pv = spec["tables"]["114-01"]["rows"]
na_rows = [r for r in pv if r["FREEWAY"] == "NA"]
if len(na_rows) != 2:
    fails.append(f"expected 2 NA PV rows, got {len(na_rows)}")

for p in ("MOBILE WORK IS WORK", "15 MINUTE", "619-212"):
    if squash(p) not in sb:
        fails.append(f"note phrase missing: {p}")

labels = [r.get("label") or r.get("signCode") for al in spec["orderTable"]["alignments"] for r in al["rows"]]
if labels != ["ROLL AHEAD DISTANCE", "W20-5R"]:
    fails.append(f"unexpected order: {labels}")

size_codes = {r["signCode"] for r in spec["tables"]["114-03"]["rows"]}
sign_codes = {s["signCode"] for s in spec["signs"]["items"]}
if size_codes != sign_codes:
    fails.append(f"sign/size mismatch {size_codes} vs {sign_codes}")

print("order", labels)
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
