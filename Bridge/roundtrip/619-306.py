"""Round-trip for 619-306.json (Family 4 parkway reference)."""
from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import squash

spec = json.loads((ROOT / "Data/sheet-specs/619-306.json").read_text(encoding="utf-8"))
ref302 = json.loads((ROOT / "Data/sheet-specs/619-302.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
body = " ".join(w[4] for pg in doc for w in pg.get_text("words"))
sb = squash(body)
fails = []

# Speeds
speeds = [r["speedMph"] for r in spec["tables"]["306-03"]["rows"]]
if speeds != [45, 50, 55, 65]:
    fails.append(f"306-03 speeds {speeds}")

# Overlap with 302
for row in spec["tables"]["306-03"]["rows"]:
    r302 = next(r for r in ref302["tables"]["302-02"]["rows"] if r["speedMph"] == row["speedMph"])
    if row["longitudinalBufferSpace"] != r302["longitudinalBufferSpace"]:
        fails.append(f"buffer mismatch {row['speedMph']}")
    if row["laneTaper"] != r302["laneTaper"]:
        fails.append(f"laneTaper mismatch {row['speedMph']}")
    if row["shoulderTaper"] != r302["shoulderTaper"]:
        fails.append(f"shoulderTaper mismatch {row['speedMph']}")

# Tokens
for tok in ("W20-1", "W20-5R", "W4-2R", "G20-2", "1000'", "1500'", "2640'", "MERGING", "PARKWAY"):
    if tok not in body:
        fails.append(f"missing {tok}")
if "NYW8-33" in body:
    fails.append("unexpected NYW8-33 on 306")

# No AW role
if spec["tableRoles"].get("advanceWarningSpacing"):
    fails.append("must not declare advanceWarningSpacing")

# Roll ahead ratios
for part in ("120/3", "200/5", "80/2", "160/4"):
    if part not in body:
        fails.append(f"missing roll-ahead {part}")

# Notes
for p in ("SHORT-TERM STATIONARY", "PROTECTIVE VEHICLE(S) SHALL MAINTAIN", "NO WORKERS, EQUIPMENT"):
    if squash(p) not in sb:
        fails.append(f"note phrase missing: {p}")
if len(spec["notes"]["printed"]) != 3:
    fails.append(f"expected 3 notes, got {len(spec['notes']['printed'])}")

# Order
labels = [r.get("label") or r.get("signCode") for al in spec["orderTable"]["alignments"] for r in al["rows"]]
if labels[:3] != ["ROLL AHEAD DISTANCE", "BUFFER SPACE", "MERGING TAPER"]:
    fails.append(f"unexpected upstream start: {labels[:3]}")
if "SHOULDER TAPER" in labels:
    fails.append("order must not include SHOULDER TAPER")
if "G20-2" not in labels:
    fails.append("missing G20-2 downstream")

# Sign sizes sync
size_codes = {r["signCode"] for r in spec["tables"]["306-04"]["rows"]}
sign_codes = {s["signCode"] for s in spec["signs"]["items"]}
if size_codes != sign_codes:
    fails.append(f"sign/size mismatch {size_codes} vs {sign_codes}")

print("order", labels)
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
