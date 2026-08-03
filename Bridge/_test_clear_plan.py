"""Live prove CLEAR_PLAN_ELEMENTS is idempotent: clear once, clear again
(deleted=0), and the Python re-place gate fires without clear_prior."""
from __future__ import annotations

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "mcp-server"))

import wztc_ops  # noqa: E402
from bridge_client import Bridge  # noqa: E402

wztc_ops._bridge = Bridge()

print("=== clear #1 (should delete the 3:17 AM plan geometry still in journal) ===")
r1 = wztc_ops.clear_plan_elements(keep_alignments=True)
print("status:", r1.get("status"), "deleted:", r1.get("deleted"),
      "clearedReqCount:", r1.get("clearedReqCount"), "note:", r1.get("note"))

print("\n=== clear #2 (idempotent: deleted=0) ===")
r2 = wztc_ops.clear_plan_elements(keep_alignments=True)
print("status:", r2.get("status"), "deleted:", r2.get("deleted"),
      "clearedReqCount:", r2.get("clearedReqCount"))

print("\n=== re-place gate (simulate stations already placed) ===")
wztc_ops._PLAN_SESSION["stations_placed_aligns"] = {1}
try:
    wztc_ops.place_order_table_stations(1, reset_session=True)
    print("  NO ERROR <-- BAD")
except ValueError as e:
    print("  refuses:", e)

print("\n=== clear_prior path resets the gate without calling place (stub check) ===")
wztc_ops._PLAN_SESSION["stations_placed_aligns"] = {1}
# clear_prior calls clear then place — place will fail if no order table /
# alignment, but clear should run and empty the set first. Catch whatever place raises.
try:
    wztc_ops.place_order_table_stations(1, reset_session=True, clear_prior=True)
    print("  place returned OK (alignment+table already present)")
except Exception as e:
    print(f"  after clear_prior, place failed as expected without live align: {type(e).__name__}: {e}")
print("  stations_placed_aligns after clear_prior attempt:",
      wztc_ops._PLAN_SESSION["stations_placed_aligns"])
