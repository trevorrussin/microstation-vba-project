"""List Default ARC/LINE near Roll Ahead tips; frame tip close-ups."""
from __future__ import annotations

import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from bridge_client import chat_bridge  # noqa: E402
import view_capture  # noqa: E402
import wztc_ops as ops  # noqa: E402

ops.set_bridge(chat_bridge)
OUT = ROOT / "Bridge" / "captures"

spots = {
    "ra_east": (78570, 287965),
    "ra_west": (78445, 287930),
    "ra_mid": (78510, 287945),
    "dn_mid": (78655, 288090),
}

for name, (x, y) in spots.items():
    r = chat_bridge.call("FIND_ELEMENTS_NEAR", x=x, y=y, radius=40, reason=name)
    rows = [
        row
        for row in (r.get("rows") or [])
        if row.get("level") == "Default"
        and row.get("type") in ("LINE", "ARC", "DIMENSION")
    ]
    print(f"=== {name} ({x},{y}) default dim-like={len(rows)}")
    for row in sorted(rows, key=lambda z: float(z.get("distanceFt") or 0))[:20]:
        rl = float(row.get("rangeLowX", 0))
        rh = float(row.get("rangeHighX", 0))
        rb = float(row.get("rangeLowY", 0))
        rt = float(row.get("rangeHighY", 0))
        span = ((rh - rl) ** 2 + (rt - rb) ** 2) ** 0.5
        print(
            f"  {row.get('type')} id={row.get('elementId')} span={span:.1f} "
            f"dist={row.get('distanceFt')} "
            f"rng=({rl:.1f},{rb:.1f})-({rh:.1f},{rt:.1f})"
        )

for name, cx, cy, w, h in (
    ("qa_311_curve_ra_tips", 78510, 287950, 220, 140),
    ("qa_311_curve_dn_tips", 78655, 288090, 160, 140),
):
    view_capture.navigate_view(cx, cy, w, h, view_num=1)
    time.sleep(0.4)
    dest = OUT / f"{name}.png"
    shutil.copy2(view_capture.capture_microstation(), dest)
    print("saved", dest.name)
