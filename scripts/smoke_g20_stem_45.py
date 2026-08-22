"""Smoke: place G20-2 on a 45-degree outward ray and measure stem alignment."""
from __future__ import annotations

import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(r"c:\repos\microstation-vba-project\mcp-server")))

from bridge_client import Bridge  # noqa: E402
import wztc_ops as ops  # noqa: E402

ops.set_bridge(Bridge())

TIP = (200000.0, 350000.0)
DIR = (math.sqrt(0.5), -math.sqrt(0.5))  # 45° SE — failing case


def f(r, *ks):
    for k in ks:
        if r.get(k) not in (None, ""):
            try:
                return float(r[k])
            except Exception:
                pass
    return None


def main() -> int:
    print(f"placing G20-2 at {TIP} dir={DIR}", flush=True)
    r = ops.place_sign(
        sign_num="G20-2",
        road_type="NON-FREEWAY",
        side="One Side",
        pt1x=TIP[0],
        pt1y=TIP[1],
        pt1z=0.0,
        dir1x=DIR[0],
        dir1y=DIR[1],
        one_off=True,
        post_angle_deg=-45.0,
        reason="45deg stem smoke",
    )
    print("place_sign", r, flush=True)

    mid = (TIP[0] + DIR[0] * 40, TIP[1] + DIR[1] * 40)
    cells = ops.find_elements_near(mid[0], mid[1], 120, "CELL", force=True)
    post = face = None
    for c in cells:
        name = (c.get("cellName") or "").upper()
        if "TWZSGN" in name:
            post = c
        if "G20" in name:
            face = c
    print("post", None if not post else {k: post.get(k) for k in ("elementId", "cellName", "cx", "cy")})
    print("face", None if not face else {k: face.get(k) for k in ("elementId", "cellName", "cx", "cy")})
    if not post or not face:
        return 1

    px, py = f(post, "cx"), f(post, "cy")
    gx, gy = f(face, "cx"), f(face, "cy")
    dx, dy = gx - px, gy - py
    n = math.hypot(dx, dy) or 1.0
    ux, uy = dx / n, dy / n
    print(f"post->face=({ux:.3f},{uy:.3f}) expected=({DIR[0]:.3f},{DIR[1]:.3f})", flush=True)

    lines = ops.find_elements_near(0.5 * (px + gx), 0.5 * (py + gy), 80, "LINE", force=True)
    best = None
    for ln in lines:
        if (ln.get("level") or "") != "SF_P":
            continue
        x1, y1 = f(ln, "rangeLowX"), f(ln, "rangeLowY")
        x2, y2 = f(ln, "rangeHighX"), f(ln, "rangeHighY")
        L = math.hypot(x2 - x1, y2 - y1)
        if L < 30 or L > 90:
            continue
        lx, ly = (x2 - x1) / L, (y2 - y1) / L
        align = abs(lx * ux + ly * uy)
        rec = (align, round(L, 2), ln.get("elementId"),
               round(x1, 1), round(y1, 1), round(x2, 1), round(y2, 1))
        if best is None or align > best[0]:
            best = rec
    print("best SF_P stem", best, flush=True)
    if best is None or best[0] < 0.85:
        print("FAIL: stem not aligned with post->face", flush=True)
        return 2
    print("PASS", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
