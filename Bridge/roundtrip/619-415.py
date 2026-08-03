"""Round-trip for 619-415.json (Family 3 intermediate ramp-approach shoulder)."""
from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import squash

spec = json.loads((ROOT / "Data/sheet-specs/619-415.json").read_text(encoding="utf-8"))
ref301 = json.loads((ROOT / "Data/sheet-specs/619-301.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
body = " ".join(w[4] for pg in doc for w in pg.get_text("words"))
sb = squash(body)
fails = []

for tok in ("INTERMEDIATE", "RAMP APPROACH", "W21-5aR", "W21-5bR", "NYR9-11",
            "SHALL NOT EXCEED 20", "LATERAL", "1320"):
    if squash(tok) not in sb and tok not in body:
        fails.append(f"missing {tok}")

# Table role numbering trap: 415-01 = taper (not PV)
if spec["tableRoles"].get("taperAndBuffer") != "415-01":
    fails.append(f"taperAndBuffer should be 415-01, got {spec['tableRoles'].get('taperAndBuffer')}")
if spec["tableRoles"].get("protectiveVehicle") != "415-03":
    fails.append(f"protectiveVehicle should be 415-03")

taper_id = "415-01"
for row in spec["tables"][taper_id]["rows"]:
    s = row["speedMph"]
    r301 = next(r for r in ref301["tables"]["301-03"]["rows"] if r["speedMph"] == s)
    if row["longitudinalBufferSpace"] != r301["longitudinalBufferSpace"]:
        fails.append(f"buffer mismatch at {s}")
    if row["shoulderTaper"] != r301["shoulderTaper"]:
        fails.append(f"shoulder mismatch at {s}")
    if "lateralShiftTaper" not in row:
        fails.append(f"415-01 missing lateralShiftTaper at {s}")

codes = {r["signCode"] for r in spec["tables"]["415-05"]["rows"]}
if "NYR9-11" not in codes:
    fails.append(f"415-05 missing NYR9-11: {codes}")

labels = [r.get("label") or r.get("signCode") for al in spec["orderTable"]["alignments"] for r in al["rows"]]
if labels[:3] != ["ROLL AHEAD DISTANCE", "BUFFER SPACE", "SHOULDER TAPER"]:
    fails.append(f"unexpected upstream start: {labels[:3]}")
if "W21-5aR" not in labels:
    fails.append(f"order missing W21-5aR: {labels}")
if "MERGING TAPER" in labels:
    fails.append("order has MERGING TAPER")

if spec["tableRoles"].get("advanceWarningSpacing"):
    fails.append("must not declare advanceWarningSpacing")

print("order", labels)
print("roles", {k: v for k, v in spec["tableRoles"].items() if k != "note"})
print()
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
