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
