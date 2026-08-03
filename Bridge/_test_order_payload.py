"""Prove the exact bridge payload build_wztc_order_table will send, without
MicroStation. Stubs the bridge so the call is captured instead of executed."""
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "mcp-server"))

import wztc_ops  # noqa: E402


class FakeBridge:
    def __init__(self):
        self.last = None

    def call(self, op, **kw):
        self.last = (op, kw)
        return {"ok": True, "rows": [], "op": op}


fake = FakeBridge()
wztc_ops._bridge = fake
wztc_ops._ok_or_raise = lambda resp, name: dict(resp)

print("=== 619-311, 45 mph, 12 ft lane, 12 ft shoulder, URBAN, no sign_rows ===")
resp = wztc_ops.build_wztc_order_table(
    speed=45, road_type="Non-Freeway", lane_width=12, shoulder_width="12 ft",
    sheet_num="619-311", area_type="URBAN",
    closure_type="LANE CLOSURE", exposure_condition="WORKERS ON FOOT")

op, kw = fake.last
print(f"op = {op}")
for k in ("sheetNum", "speed", "roadType", "laneWidth", "shoulderWidth"):
    print(f"  {k:20s} {kw[k]}")
print(f"  {'nonSignRowsTSV':20s} {kw['nonSignRowsTSV']}")
print(f"  {'signRowsTSV':20s} {kw['signRowsTSV']}")
print(f"  {'spacingOverridesTSV':20s} {kw['spacingOverridesTSV']}")

print(f"\nspecDriven = {resp['specDriven']}   shoulder band = {resp['shoulderBandUsed']}")
print(f"legends    = {resp['signLegends']}")
print("station walk shown to the engineer:")
for w in resp["stationWalk"]:
    print(f"   {w['alignName']:<10} {str(w['item']):<24} "
          f"len={w['lengthFt']:>6g}  sta={w['stationFt']:>6g}")

print("\n=== no spec for this sheet -> generic fallback ===")
resp2 = wztc_ops.build_wztc_order_table(
    speed=45, road_type="Non-Freeway", lane_width=12, shoulder_width="12 ft",
    sign_rows=[{"align_idx": 1, "sign_num": "W20-01RA"}],
    sheet_num="619-999", area_type="URBAN")
op, kw = fake.last
print(f"specDriven = {resp2['specDriven']}")
print(f"  nonSignRowsTSV      = {kw['nonSignRowsTSV']!r}  (empty -> VBA uses its defaults)")
print(f"  spacingOverridesTSV = {kw['spacingOverridesTSV']!r}")

print("\n=== spec sheet but caller forgot area_type ===")
try:
    wztc_ops.build_wztc_order_table(
        speed=45, road_type="Non-Freeway", lane_width=12, shoulder_width="12 ft",
        sheet_num="619-311")
    print("  NO ERROR RAISED <-- BAD")
except ValueError as e:
    print(f"  raises: {e}")

print("\n=== spec sheet at 65 mph (outside the sheet's tables) ===")
try:
    wztc_ops.build_wztc_order_table(
        speed=65, road_type="Non-Freeway", lane_width=12, shoulder_width="12 ft",
        sheet_num="619-311", area_type="URBAN")
    print("  NO ERROR RAISED <-- BAD")
except Exception as e:
    print(f"  raises: {e}")
