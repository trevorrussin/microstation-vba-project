"""Small 2D geometry helpers for operation contracts. Pure functions, no MicroStation access.

Points are [x, y] or [x, y, z]. Tolerances are in design units (ft); MicroStation stores 10,000 UOR per ft, so
0.02 ft is far above round-trip noise and far below any engineering-relevant error.
"""
from __future__ import annotations

import math

TOL = 0.02          # ft
ANG_TOL = 0.05      # degrees


def sub(a, b):
    return (a[0] - b[0], a[1] - b[1])


def dist(a, b):
    return math.hypot(a[0] - b[0], a[1] - b[1])


def length(vs):
    return sum(dist(vs[i], vs[i + 1]) for i in range(len(vs) - 1))


def rotate(p, o, deg):
    r = math.radians(deg)
    c, s = math.cos(r), math.sin(r)
    dx, dy = p[0] - o[0], p[1] - o[1]
    return (o[0] + c * dx - s * dy, o[1] + s * dx + c * dy)


def mirror(p, a, b):
    ux, uy = b[0] - a[0], b[1] - a[1]
    n = math.hypot(ux, uy)
    ux, uy = ux / n, uy / n
    dx, dy = p[0] - a[0], p[1] - a[1]
    t = dx * ux + dy * uy
    fx, fy = a[0] + t * ux, a[1] + t * uy
    return (2 * fx - p[0], 2 * fy - p[1])


def project_on_line(p, a, b):
    """-> (foot point, signed perpendicular distance (+ = left of a->b), parameter t along a->b in ft)."""
    ux, uy = b[0] - a[0], b[1] - a[1]
    n = math.hypot(ux, uy)
    ux, uy = ux / n, uy / n
    dx, dy = p[0] - a[0], p[1] - a[1]
    t = dx * ux + dy * uy
    return (a[0] + t * ux, a[1] + t * uy), (-uy * dx + ux * dy), t


def heading_deg(a, b):
    return math.degrees(math.atan2(b[1] - a[1], b[0] - a[0]))


def ang_diff(a, b):
    d = (a - b + 180.0) % 360.0 - 180.0
    return abs(d)


def same_pts(va, vb, tol=TOL):
    """Pairwise equal (same count, each within tol)."""
    if len(va) != len(vb):
        return False, f"vertex count {len(va)} vs {len(vb)}"
    worst = max((dist(a, b) for a, b in zip(va, vb)), default=0.0)
    return worst <= tol, f"max vertex error {worst:.4f} ft"


def bbox(vs):
    xs = [v[0] for v in vs]
    ys = [v[1] for v in vs]
    return min(xs), min(ys), max(xs), max(ys)


def bearing_to_azimuth(text: str) -> float:
    """'N29°54'00"E' / 'S 57 08 20 W' / 'N01:33:00:W' -> azimuth in degrees clockwise from north (0-360).

    As-built plans give quadrant bearings; ORD key-ins accept N01:33:00:W. Raises ValueError on anything else."""
    import re
    m = re.match(r"^\s*([NS])\s*(\d{1,3})\D+(\d{1,2})(?:\D+(\d{1,2}(?:\.\d+)?))?\D*([EW])\s*$", text.strip(), re.I)
    if not m:
        raise ValueError(f"not a quadrant bearing: {text!r}")
    ns, d, mi, se, ew = m.group(1).upper(), float(m.group(2)), float(m.group(3)), float(m.group(4) or 0), m.group(5).upper()
    ang = d + mi / 60.0 + se / 3600.0
    if ang > 90.0:
        raise ValueError(f"quadrant bearing angle {ang} exceeds 90 degrees")
    if ns == "N":
        return ang if ew == "E" else 360.0 - ang
    return 180.0 - ang if ew == "E" else 180.0 + ang


def azimuth_to_math_deg(az: float) -> float:
    """Azimuth (clockwise from north) -> standard math angle (counterclockwise from +X / east), 0-360."""
    return (90.0 - az) % 360.0


def nearest_on_polyline(p, pts):
    """-> (distance, signed distance (+ = left of the nearest segment's direction), segment index) from p to the polyline."""
    best = None
    for i in range(len(pts) - 1):
        a, b = pts[i], pts[i + 1]
        ux, uy = b[0] - a[0], b[1] - a[1]
        n2 = ux * ux + uy * uy
        if n2 < 1e-18:
            continue
        t = max(0.0, min(1.0, ((p[0] - a[0]) * ux + (p[1] - a[1]) * uy) / n2))
        fx, fy = a[0] + t * ux, a[1] + t * uy
        d = math.hypot(p[0] - fx, p[1] - fy)
        if best is None or d < best[0] - 1e-12:
            side = ux * (p[1] - a[1]) - uy * (p[0] - a[0])  # cross > 0: p is left of a->b
            best = (d, d if side >= 0 else -d, i)
    if best is None:
        raise ValueError("polyline has no non-degenerate segment")
    return best


def sample_polyline(pts, per_segment=4):
    """Vertices plus evenly spaced interior points of every segment."""
    out = []
    for i in range(len(pts) - 1):
        a, b = pts[i], pts[i + 1]
        for k in range(per_segment):
            f = k / per_segment
            out.append((a[0] + f * (b[0] - a[0]), a[1] + f * (b[1] - a[1])))
    out.append((pts[-1][0], pts[-1][1]))
    return out


def dist_box_polyline(box, pts):
    """Minimum distance between an axis-aligned box (lx, ly, hx, hy) and a polyline; 0 when they touch or overlap."""
    lx, ly, hx, hy = box
    corners = [(lx, ly), (hx, ly), (hx, hy), (lx, hy)]
    best = float("inf")
    for i in range(len(pts) - 1):
        a, b = pts[i], pts[i + 1]
        if (lx <= a[0] <= hx and ly <= a[1] <= hy) or (lx <= b[0] <= hx and ly <= b[1] <= hy):
            return 0.0
        for j in range(4):
            c, d = corners[j], corners[(j + 1) % 4]
            if _segments_cross(a, b, c, d):
                return 0.0
        for q in corners:  # box corners to the segment
            best = min(best, nearest_on_polyline(q, [a, b])[0])
        for q in (a, b):   # segment ends to the box
            dx = max(lx - q[0], 0.0, q[0] - hx)
            dy = max(ly - q[1], 0.0, q[1] - hy)
            best = min(best, math.hypot(dx, dy))
    return best


def _segments_cross(a, b, c, d):
    def cr(o, p, q):
        return (p[0] - o[0]) * (q[1] - o[1]) - (p[1] - o[1]) * (q[0] - o[0])
    d1, d2, d3, d4 = cr(c, d, a), cr(c, d, b), cr(a, b, c), cr(a, b, d)
    return ((d1 > 0) != (d2 > 0)) and ((d3 > 0) != (d4 > 0))


def y_at_x(pts, x):
    """Y of the polyline where it crosses the vertical line at x (None if it does not span x). First crossing."""
    for i in range(len(pts) - 1):
        a, b = pts[i], pts[i + 1]
        lo, hi = min(a[0], b[0]), max(a[0], b[0])
        if lo - 1e-9 <= x <= hi + 1e-9:
            if abs(b[0] - a[0]) < 1e-12:
                return a[1]
            return a[1] + (b[1] - a[1]) * (x - a[0]) / (b[0] - a[0])
    return None
