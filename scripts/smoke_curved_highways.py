"""Live smoke: place S-curve versions of every highway strip type.

Requires MicroStation + bridge. Offsets each type in Y so they don't stack.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from bridge_client import chat_bridge  # noqa: E402
import wztc_ops as ops  # noqa: E402

ops.set_bridge(chat_bridge)


def _s(base_x: float, base_y: float):
    return [
        [base_x + 0.0, base_y],
        [base_x + 120.0, base_y],
        [base_x + 220.0, base_y + 50.0],
        [base_x + 320.0, base_y + 50.0],
        [base_x + 420.0, base_y],
    ]


def main() -> int:
    # Fresh band south of earlier smokes; wide spacing so packs don't visually merge.
    origin_x, origin_y = 51000.0, 298500.0
    gap = 300.0
    results = []

    r = ops.place_lane_highway(
        2, vertices=_s(origin_x, origin_y),
        shoulder_width_ft=8.0, reason="smoke curved one-way",
    )
    results.append(("one_way", r))

    r = ops.place_two_way_highway(
        4, vertices=_s(origin_x, origin_y - gap),
        shoulder_width_ft=8.0, reason="smoke curved two-way",
    )
    results.append(("two_way", r))

    r = ops.place_divided_highway(
        2, median_width_ft=20.0,
        vertices=_s(origin_x, origin_y - 2 * gap),
        shoulder_width_ft=8.0, reason="smoke curved divided",
    )
    results.append(("divided", r))

    r = ops.place_twlt_highway(
        2, twlt_width_ft=12.0,
        vertices=_s(origin_x, origin_y - 3 * gap),
        reason="smoke curved twlt",
    )
    results.append(("twlt", r))

    r = ops.place_ramp_gore(
        mainline_lanes=3, ramp_angle_deg=18.0,
        gore_station_ft=200.0, ramp_length_ft=100.0,
        vertices=_s(origin_x, origin_y - 4 * gap),
        reason="smoke curved ramp gore",
    )
    results.append(("ramp_gore", r))

    ok = True
    for name, res in results:
        st = res.get("status")
        note = res.get("note") or ""
        placed = res.get("placedCount") or len(res.get("placed") or [])
        print(f"{name}: {st} placed={placed} {note}")
        if st != "OK":
            ok = False
            print("  errors:", res.get("errors"))

    if ok:
        try:
            import view_capture
            # Frame the full stack (one-way top → gore bottom)
            cy = origin_y - 2 * gap
            view_capture.navigate_view(
                origin_x + 210.0, cy, 900.0, 1400.0, view_num=1,
            )
            print(f"view framed near ({origin_x + 210:.0f}, {cy:.0f})")
        except Exception as e:
            print(f"view navigate skipped: {e}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
