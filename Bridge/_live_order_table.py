"""Live test: build the 619-311 order table through the bridge with the
sheet spec driving stations/signs. Requires MicroStation open and WZTCBridge
polling (or VBA RUN via the bridge client)."""
from __future__ import annotations

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "mcp-server"))

import sheet_spec  # noqa: E402
from bridge_client import Bridge  # noqa: E402

spec = sheet_spec.load("619-311")
resolved = sheet_spec.resolve(
    spec, 45, 12, "12 ft", "URBAN", "LANE CLOSURE", "WORKERS ON FOOT")
payload = sheet_spec.order_table_rows(spec, resolved)

non_sign = "|".join(payload["nonSignRows"])
sign = "|".join(payload["signRows"])
overrides = "|".join([
    f"bufferSpace={resolved['bufferFt']}",
    f"mergingTaper={resolved['laneTaper']['ft']}",
    f"shoulderTapers={resolved['shoulderTaper']['ft']}",
    f"rollAhead={resolved['rollAheadFt']['min']}",
    f"laneTaperSkips={resolved['laneTaper']['skipLines']}",
    f"shoulderTaperSkips={resolved['shoulderTaper']['skipLines']}",
    f"laneTaperDevices={resolved['laneTaper']['devices']}",
    f"shoulderTaperDevices={resolved['shoulderTaper']['devices']}",
])

print("calling BUILD_WZTC_ORDER_TABLE...")
print(f"  nonSign: {non_sign}")
print(f"  signs:   {sign}")
print(f"  overs:   {overrides}")

bridge = Bridge()
resp = bridge.call(
    "BUILD_WZTC_ORDER_TABLE",
    category="Multilane Undivided",
    sheetNum="619-311",
    speed=45,
    roadType="Non-Freeway",
    laneWidth=12,
    shoulderWidth="12 ft",
    signRowsTSV=sign,
    nonSignRowsTSV=non_sign,
    spacingOverridesTSV=overrides,
)

print("\nstatus:", resp.get("status") or resp.get("ok") or resp)
rows = resp.get("rows") or []
print(f"rows returned: {len(rows)}")
print(f"{'align':<6} {'#':>3} {'type':<10} {'label':<28} {'spacing':>8} {'size':<14} {'side'}")
for r in rows:
    if isinstance(r, dict):
        print(f"{r.get('alignName', r.get('alignIdx','')):<6} "
              f"{str(r.get('rowNum','')):>3} "
              f"{r.get('type',''):<10} "
              f"{str(r.get('label','')):<28} "
              f"{str(r.get('spacing','')):>8} "
              f"{str(r.get('size','')):<14} "
              f"{r.get('side','')}")
    else:
        print(" ", r)

# Expected shape for 619-311:
# Upstream: RAD, Buffer, Lane Taper, W04-02R, W20-05RA, W20-01RF
# Downstream: Downstream Taper, G20-02
# NO Vehicle Space / temp barrier / box-corr beam
expect_up = ["ROLL AHEAD DISTANCE", "BUFFER SPACE", "LANE TAPER",
             "W04-02R", "W20-05RA", "W20-01RF"]
expect_dn = ["DOWNSTREAM TAPER", "G20-02"]
forbidden = ["VEHICLE SPACE", "TEMP BARRIER", "BOX/CORR", "SHOULDER TAPER"]

labels = []
for r in rows:
    if isinstance(r, dict):
        labels.append((str(r.get("alignIdx")), str(r.get("label", "")).upper()))
    elif isinstance(r, (list, tuple)) and len(r) >= 5:
        labels.append((str(r[0]), str(r[4]).upper()))

print("\n=== checks ===")
up = [lab for a, lab in labels if a in ("1", "Upstream")]
dn = [lab for a, lab in labels if a in ("2", "Downstream")]
# alignIdx may be numeric in rows - also check alignName
if not up and not dn:
    up = [lab for a, lab in labels if a == "1"]
    dn = [lab for a, lab in labels if a == "2"]
    if not up:
        # dict rows with alignIdx as int-ish
        for r in rows:
            if not isinstance(r, dict):
                continue
            a = str(r.get("alignIdx", ""))
            lab = str(r.get("label", "")).upper()
            if a == "1":
                up.append(lab)
            elif a == "2":
                dn.append(lab)

print("upstream labels:", up)
print("downstream labels:", dn)

def ok(name, cond):
    print(f"  [{'PASS' if cond else 'FAIL'}] {name}")

ok("no Vehicle Space / barrier phantoms",
   not any(any(f in lab for f in forbidden) for _, lab in labels))
ok("upstream has 6 rows (3 non-sign + 3 signs)", len(up) == 6)
ok("W20-01RF present (URBAN 45 -> 1000 FT legend)",
   any("W20-01RF" in lab for lab in up))
ok("W20-05RA present (URBAN -> AHEAD legend)",
   any("W20-05RA" in lab for lab in up))
ok("shoulder taper NOT a sequential station",
   not any("SHOULDER" in lab for lab in up))
