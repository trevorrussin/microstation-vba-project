"""Station->XY interpolation over an alignment's raw path segments, done in
pure Python instead of one MicroStation bridge round trip per point.

Placement-plan compiler, Stage 1 (see Data/sheet-specs/STATUS.md). Fetch an
alignment's segments once via wztc_ops.get_alignment_vertices(), then call
station_to_xy() as many times as needed locally -- zero further MicroStation
calls for geometry. This mirrors PerpPlacement.GetPointAndTangent's algorithm
exactly (Modules/PerpPlacement.bas:1185-1274) so the two stay in agreement:
same segment walk, same straight-line interpolation, same arc-length
parametrization, same tangent normalization.

Unit-testable without MicroStation -- see the __main__ block below.
"""
from __future__ import annotations

import math
from dataclasses import dataclass


class AlignmentGeometryError(Exception):
    pass


@dataclass
class PathSegment:
    is_arc: bool
    sx: float
    sy: float
    sz: float
    ex: float
    ey: float
    ez: float
    seg_len: float
    cx: float = 0.0
    cy: float = 0.0
    radius: float = 0.0
    start_angle: float = 0.0
    sweep_angle: float = 0.0


def parse_vertices(rows: list[dict]) -> list[PathSegment]:
    """rows as returned by wztc_ops.get_alignment_vertices() -- string-valued
    dicts with keys segIndex/isArc/sx/sy/sz/ex/ey/ez/segLen/cx/cy/radius/
    startAngle/sweepAngle, already in path order (start of alignment to end,
    per Modules/PerpPlacement.bas's GetAlignmentVertices)."""
    segments = []
    for r in rows:
        segments.append(PathSegment(
            is_arc=str(r.get("isArc", "N")).strip().upper() == "Y",
            sx=float(r["sx"]), sy=float(r["sy"]), sz=float(r["sz"]),
            ex=float(r["ex"]), ey=float(r["ey"]), ez=float(r["ez"]),
            seg_len=float(r["segLen"]),
            cx=float(r.get("cx", 0.0)), cy=float(r.get("cy", 0.0)),
            radius=float(r.get("radius", 0.0)),
            start_angle=float(r.get("startAngle", 0.0)),
            sweep_angle=float(r.get("sweepAngle", 0.0)),
        ))
    return segments


def total_length(segments: list[PathSegment]) -> float:
    return sum(s.seg_len for s in segments)


def station_to_xy(segments: list[PathSegment], station: float) -> tuple[float, float, float, float]:
    """Returns (x, y, tangent_x, tangent_y) at `station` (design units,
    normally ft) along the path. Clamps to [0, total_length] rather than
    raising -- same as GetPointAndTangent, which silently clamps an
    out-of-range station to the nearest path end (flagged as a real
    placement-accuracy risk in WZTCQuery.StationToPoint's own comments;
    callers here should diff `station` against total_length() themselves
    if they need to know whether clamping happened)."""
    if not segments:
        raise AlignmentGeometryError("no segments -- alignment has no path")

    total = total_length(segments)
    dist = station
    if dist < 0:
        dist = 0.0
    if dist > total:
        dist = total

    cum_len = 0.0
    for seg in segments:
        seg_end = cum_len + seg.seg_len
        if dist <= seg_end + 0.00001:
            t = dist - cum_len
            if t < 0:
                t = 0.0

            if not seg.is_arc:
                l_len = seg.seg_len if seg.seg_len >= 0.000001 else 0.000001
                tdx = (seg.ex - seg.sx) / l_len
                tdy = (seg.ey - seg.sy) / l_len
                x = seg.sx + t * tdx
                y = seg.sy + t * tdy
                tan_x, tan_y = tdx, tdy
            else:
                r = seg.radius
                sa = seg.start_angle
                sw = seg.sweep_angle
                if abs(sw) > 0.000001 and r > 0.000001:
                    theta = sa + (t / r) * (1.0 if sw >= 0 else -1.0)
                else:
                    theta = sa
                x = seg.cx + r * math.cos(theta)
                y = seg.cy + r * math.sin(theta)
                sw_sign = 1.0 if sw >= 0 else -1.0
                tan_x = -math.sin(theta) * sw_sign
                tan_y = math.cos(theta) * sw_sign

            mag = math.hypot(tan_x, tan_y)
            if mag > 0.000001:
                tan_x /= mag
                tan_y /= mag
            return (x, y, tan_x, tan_y)

        cum_len = seg_end

    # Fell through (shouldn't happen given the clamp above) -- return the
    # last segment's end point, same fallback GetPointAndTangent uses.
    last = segments[-1]
    return (last.ex, last.ey, 0.0, 0.0)


if __name__ == "__main__":
    # Synthetic self-test, no MicroStation required.
    def approx(a: float, b: float, tol: float = 1e-6) -> bool:
        return abs(a - b) <= tol

    # Two straight segments: (0,0)->(100,0)->(100,150), matching the live
    # smoke test run against GET_ALIGNMENT_VERTICES during Stage 1
    # development (Bridge journal, 2026-08-03: alignIdx=9 test alignment).
    straight = [
        PathSegment(False, 0, 0, 0, 100, 0, 0, 100.0),
        PathSegment(False, 100, 0, 0, 100, 150, 0, 150.0),
    ]
    assert approx(total_length(straight), 250.0)

    x, y, tx, ty = station_to_xy(straight, 0)
    assert approx(x, 0) and approx(y, 0) and approx(tx, 1) and approx(ty, 0), (x, y, tx, ty)

    x, y, tx, ty = station_to_xy(straight, 50)
    assert approx(x, 50) and approx(y, 0), (x, y)

    x, y, tx, ty = station_to_xy(straight, 100)
    assert approx(x, 100) and approx(y, 0), (x, y)

    x, y, tx, ty = station_to_xy(straight, 175)
    assert approx(x, 100) and approx(y, 75) and approx(tx, 0) and approx(ty, 1), (x, y, tx, ty)

    x, y, tx, ty = station_to_xy(straight, 250)
    assert approx(x, 100) and approx(y, 150), (x, y)

    # Clamping past either end
    x, y, _, _ = station_to_xy(straight, -50)
    assert approx(x, 0) and approx(y, 0), (x, y)
    x, y, _, _ = station_to_xy(straight, 999)
    assert approx(x, 100) and approx(y, 150), (x, y)

    # Quarter-circle arc, center (0,0), radius 100, start angle 0, sweep +90deg
    # (CCW): station 0 -> (100, 0); station at quarter length -> (0, 100).
    quarter = math.pi / 2
    arc_len = 100.0 * quarter
    arc = [PathSegment(True, 100, 0, 0, 0, 100, 0, arc_len,
                        cx=0, cy=0, radius=100, start_angle=0, sweep_angle=quarter)]
    x, y, tx, ty = station_to_xy(arc, 0)
    assert approx(x, 100) and approx(y, 0), (x, y)
    assert approx(tx, 0) and approx(ty, 1), (tx, ty)  # tangent CCW at angle 0 is +Y

    x, y, tx, ty = station_to_xy(arc, arc_len)
    assert approx(x, 0, tol=1e-4) and approx(y, 100, tol=1e-4), (x, y)

    x, y, tx, ty = station_to_xy(arc, arc_len / 2)
    expected = 100.0 * math.cos(quarter / 2)
    assert approx(x, expected, tol=1e-4) and approx(y, expected, tol=1e-4), (x, y)

    # Mixed straight-then-arc path, station continuity across the seam
    mixed = [
        PathSegment(False, 0, 0, 0, 100, 0, 0, 100.0),
        PathSegment(True, 100, 0, 0, 200, 100, 0, arc_len,
                    cx=100, cy=100, radius=100, start_angle=-quarter, sweep_angle=quarter),
    ]
    x, y, _, _ = station_to_xy(mixed, 100.0)
    assert approx(x, 100) and approx(y, 0), (x, y)  # exactly at the seam
    x, y, _, _ = station_to_xy(mixed, 100.0 + arc_len)
    assert approx(x, 200, tol=1e-4) and approx(y, 100, tol=1e-4), (x, y)

    print("alignment_geometry: all self-tests passed")
