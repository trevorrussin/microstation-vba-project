"""Round-trip for 619-501.json (Family 3 long-term barrier shoulder)."""
from __future__ import annotations

import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import fitz
from pdf_table_extract import squash

spec = json.loads((ROOT / "Data/sheet-specs/619-501.json").read_text(encoding="utf-8"))
ref301 = json.loads((ROOT / "Data/sheet-specs/619-301.json").read_text(encoding="utf-8"))
doc = fitz.open(str(ROOT / spec["sheet"]["localPdf"]))
body = " ".join(w[4] for pg in doc for w in pg.get_text("words"))
sb = squash(body)
fails = []

roles = spec["tableRoles"]
if roles.get("protectiveVehicle") or roles.get("rollAheadDistance"):
    fails.append("501 must not declare PV or rollAhead roles")
if roles.get("positiveBarrierFlareRates") != "501-03":
    fails.append(f"flare role should be 501-03, got {roles.get('positiveBarrierFlareRates')}")

for tok in ("LONG-TERM", "POSITIVE BARRIER", "FLARE", "W21-5aR", "W21-5bR",
            "NYR9-11", "MORE THAN 3", "LEFT SHOULDER"):
    if squash(tok) not in sb and tok not in body:
        fails.append(f"missing {tok}")

# Shoulder/buffer overlap with 301 where both have the row
for row in spec["tables"]["501-01"]["rows"]:
    s = row["speedMph"]
    try:
        r301 = next(r for r in ref301["tables"]["301-03"]["rows"] if r["speedMph"] == s)
    except StopIteration:
        continue
    if "shoulderTaper" in row:
        # 501 may include laneTaper; compare shoulder bands that exist in both
        for b, e in r301["shoulderTaper"].items():
            if b in row["shoulderTaper"] and row["shoulderTaper"][b] != e:
                fails.append(f"shoulder {b} mismatch at {s}")

labels = [r.get("label") or r.get("signCode") for al in spec["orderTable"]["alignments"] for r in al["rows"]]
if "ROLL AHEAD DISTANCE" in labels or "BUFFER SPACE" in labels:
    fails.append(f"long-term order must omit roll/buffer: {labels}")
if labels[0] != "SHOULDER TAPER":
    fails.append(f"order should start with SHOULDER TAPER: {labels}")
if "MERGING TAPER" in labels:
    fails.append("order has MERGING TAPER")
if "W21-5aR" not in labels:
    fails.append(f"order missing W21-5aR: {labels}")

# Corridor has positiveBarrier, not protectiveVehicle
zone_ids = {z["id"] for z in spec["corridor"]["zones"]}
if "positiveBarrier" not in zone_ids:
    fails.append("corridor missing positiveBarrier")
if "protectiveVehicle" in zone_ids or "rollAheadDistance" in zone_ids:
    fails.append("corridor should not have PV/roll zones")

print("order", labels)
print("zones", sorted(zone_ids))
print()
print("ROUND-TRIP FAILURES:", len(fails))
for f in fails:
    print("  ", f)
sys.exit(1 if fails else 0)
