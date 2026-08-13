"""Live smoke: every highway strip type on several curve geometries.

Layouts (columns left→right), each column stacks road types top→bottom:
  L-bend | C-curve | gentle-S | reverse-S | hairpin
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from bridge_client import chat_bridge  # noqa: E402
import wztc_ops as ops  # noqa: E402

ops.set_bridge(chat_bridge)


def _l_bend(bx: float, by: float) -> list[list[float]]:
    """Single left bend (L / elbow)."""
    return [
        [bx + 0.0, by],
        [bx + 250.0, by],
        [bx + 400.0, by + 120.0],
    ]


def _c_curve(bx: float, by: float) -> list[list[float]]:
    """C-shaped open curve (polyline approximation of a wide arc)."""
    pts: list[list[float]] = []
    # 180° arc via chords, opening to the right
    for i in range(9):
        ang = math.pi * (i / 8.0)  # 0 → π
        pts.append([bx + 200.0 * math.sin(ang), by + 200.0 * (1.0 - math.cos(ang))])
    return pts


def _gentle_s(bx: float, by: float) -> list[list[float]]:
    """Long gentle S (shallow offsets)."""
    return [
        [bx + 0.0, by],
        [bx + 150.0, by],
        [bx + 280.0, by + 25.0],
        [bx + 400.0, by + 25.0],
        [bx + 530.0, by],
    ]


def _reverse_s(bx: float, by: float) -> list[list[float]]:
    """Classic reverse S (sharper than gentle)."""
    return [
        [bx + 0.0, by],
        [bx + 120.0, by],
        [bx + 220.0, by + 60.0],
        [bx + 320.0, by + 60.0],
        [bx + 420.0, by],
    ]


def _hairpin(bx: float, by: float) -> list[list[float]]:
    """Near-U turn (tight polyline hairpin)."""
    return [
        [bx + 0.0, by],
        [bx + 200.0, by],
        [bx + 280.0, by + 40.0],
        [bx + 280.0, by + 160.0],
        [bx + 200.0, by + 200.0],
        [bx + 0.0, by + 200.0],
    ]


CURVES = [
    ("L-bend", _l_bend),
    ("C-curve", _c_curve),
    ("gentle-S", _gentle_s),
    ("reverse-S", _reverse_s),
    ("hairpin", _hairpin),
]


def _place_all_types(name: str, verts: list[list[float]], reason_prefix: str) -> list[tuple[str, dict]]:
    """Place one-way / two-way / divided / TWLT / ramp-gore on the same path shape.

    Offset each type in Y so they don't stack on top of each other within a column.
    """
    # Tall shapes (C / hairpin) need a large vertical pitch.
    ys = [p[1] for p in verts]
    shape_h = max(ys) - min(ys)
    gap = max(550.0, shape_h + 200.0)
    out: list[tuple[str, dict]] = []

    def shift(dy: float) -> list[list[float]]:
        return [[p[0], p[1] + dy] for p in verts]

    out.append((
        f"{name}/one_way",
        ops.place_lane_highway(
            2, vertices=shift(0.0), shoulder_width_ft=8.0,
            reason=f"{reason_prefix} one-way",
        ),
    ))
    out.append((
        f"{name}/two_way",
        ops.place_two_way_highway(
            4, vertices=shift(-gap), shoulder_width_ft=8.0,
            reason=f"{reason_prefix} two-way",
        ),
    ))
    out.append((
        f"{name}/divided",
        ops.place_divided_highway(
            2, median_width_ft=20.0, vertices=shift(-2 * gap),
            shoulder_width_ft=8.0, reason=f"{reason_prefix} divided",
        ),
    ))
    out.append((
        f"{name}/twlt",
        ops.place_twlt_highway(
            2, twlt_width_ft=12.0, vertices=shift(-3 * gap),
            reason=f"{reason_prefix} twlt",
        ),
    ))
    out.append((
        f"{name}/ramp_gore",
        ops.place_ramp_gore(
            mainline_lanes=3, ramp_angle_deg=18.0,
            gore_station_ft=180.0, ramp_length_ft=90.0,
            vertices=shift(-4 * gap),
            reason=f"{reason_prefix} ramp gore",
        ),
    ))
    return out


def main() -> int:
    # Farther east; taller vertical pitch for C/hairpin
    origin_x, origin_y = 56000.0, 298000.0
    col_gap = 800.0
    results: list[tuple[str, dict]] = []

    for i, (cname, builder) in enumerate(CURVES):
        bx = origin_x + i * col_gap
        verts = builder(bx, origin_y)
        results.extend(_place_all_types(cname, verts, f"smoke {cname}"))

    ok = True
    for name, res in results:
        st = res.get("status")
        placed = res.get("placedCount") or len(res.get("placed") or [])
        print(f"{name}: {st} placed={placed} {res.get('note') or ''}")
        if st != "OK":
            ok = False
            print("  errors:", res.get("errors"))

    if ok:
        try:
            import view_capture
            cx = origin_x + 2 * col_gap + 200.0
            cy = origin_y - 1200.0
            view_capture.navigate_view(cx, cy, 4000.0, 3200.0, view_num=1)
            print(f"view framed near ({cx:.0f}, {cy:.0f}) — 5 curve cols × 5 road types")
        except Exception as e:
            print(f"view navigate skipped: {e}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
