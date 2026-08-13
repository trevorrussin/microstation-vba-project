"""Inspect Default-level dim graphics near Roll Ahead on curved 619-311."""
from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from bridge_client import chat_bridge  # noqa: E402

r = chat_bridge.call(
    "FIND_ELEMENTS_NEAR", x=78520, y=287960, radius=120, reason="find dims arcs"
)
rows = r.get("rows") or []
print("types", Counter(row.get("type") for row in rows))
for t in ("DIMENSION", "ARC", "ELLIPSE", "TEXT", "SHAPE"):
    hits = [row for row in rows if row.get("type") == t]
    print(t, len(hits))
    for row in hits[:10]:
        print(
            " ",
            row.get("elementId"),
            row.get("level"),
            row.get("cx"),
            row.get("cy"),
            "dist",
            row.get("distanceFt"),
        )

defaults = [
    row
    for row in rows
    if row.get("level") == "Default"
    and row.get("type") in ("LINE", "ARC", "DIMENSION")
]
print("Default dim-like", len(defaults))
for row in sorted(defaults, key=lambda x: float(x.get("distanceFt") or 0))[:40]:
    rl = float(row.get("rangeLowX", 0))
    rh = float(row.get("rangeHighX", 0))
    rb = float(row.get("rangeLowY", 0))
    rt = float(row.get("rangeHighY", 0))
    span = ((rh - rl) ** 2 + (rt - rb) ** 2) ** 0.5
    print(
        f"  {row.get('type')} id={row.get('elementId')} span~{span:.1f} "
        f"cx={row.get('cx')} cy={row.get('cy')} dist={row.get('distanceFt')}"
    )
