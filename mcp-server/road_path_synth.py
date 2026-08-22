"""Generate a road centreline/edge path from plain-English parameters.

Fills the gap behind "just build me a curved highway": every existing tool
either picks an EXISTING path (corridor_path, find_reference_linework,
get_element_vertices) or needs the engineer to click one. Nothing produced a
path, so a vague request collapsed into a long click session -- or into
inventing coordinates, which the prompt forbids.

Pure geometry: no MicroStation, no bridge. Output feeds straight into the
striping catalog's `vertices=` and the sheet build's `path_vertices=`.

Vertex budget matters. A live PLACE_POLYLINE of ~5k vertices was rejected by
VBA (2026-08-13), so sampling defaults to 25 ft and the result is capped via
corridor_path.downsample_polyline.
"""
from __future__ import annotations

import math
from typing import Optional

import corridor_path as cp

# Fraction of total length spent inside bends (rest is straight lead-in /
# between / lead-out). Half reads as a real highway curve rather than a
# hairpin or a barely-perceptible drift.
_ARC_LENGTH_FRACTION = 0.5
_DEFAULT_BEND_SWEEP_DEG = 45.0
_MIN_RADIUS_FT = 100.0

KINDS = ("straight", "c_curve", "s_curve", "l_bend", "n_bend")

_COMPASS = [
    (0.0, "east"), (45.0, "northeast"), (90.0, "north"), (135.0, "northwest"),
    (180.0, "west"), (225.0, "southwest"), (270.0, "south"), (315.0, "southeast"),
]


def bearing_word(bearing_deg: float) -> str:
    """Nearest compass word for a math-convention bearing (0 = +X = east)."""
    b = bearing_deg % 360.0
    best = min(_COMPASS, key=lambda c: min(abs(b - c[0]), 360.0 - abs(b - c[0])))
    return best[1]


def infer_kind(bends: Optional[int], kind: str = "") -> str:
    """Map a bend count to a shape when the engineer named one but not the other.

    "two bends" is an S; "a bend" is a simple curve; "no bends" is straight.
    """
    k = (kind or "").strip().lower().replace("-", "_").replace(" ", "_")
    if k in KINDS:
        return k
    aliases = {
        "s": "s_curve", "s_curve": "s_curve", "reverse_s": "s_curve",
        "c": "c_curve", "curve": "c_curve", "curved": "c_curve",
        "l": "l_bend", "corner": "l_bend", "bend": "c_curve",
    }
    if k in aliases:
        return aliases[k]
    if bends is None:
        return "c_curve" if k else "straight"
    if bends <= 0:
        return "straight"
    if bends == 1:
        return "c_curve"
    if bends == 2:
        return "s_curve"
    return "n_bend"


def bend_count(kind: str, bends: Optional[int] = None) -> int:
    if bends is not None and bends > 0:
        return int(bends)
    return {"straight": 0, "c_curve": 1, "l_bend": 1, "s_curve": 2, "n_bend": 3}.get(kind, 1)


def _segments(kind: str, length_ft: float, n_bends: int,
              radius_ft: Optional[float]) -> tuple[list[tuple], float, dict]:
    """Build ('straight', L) / ('arc', L, turn_sign) runs totalling length_ft.

    turn_sign +1 curves left of travel, -1 right. Returns (segments, radius,
    assumptions-applied).
    """
    assumed: dict = {}
    if n_bends <= 0:
        return ([("straight", length_ft)], 0.0, assumed)

    arc_total = length_ft * _ARC_LENGTH_FRACTION
    arc_each = arc_total / n_bends
    if radius_ft is None or float(radius_ft) <= 0:
        sweep = math.radians(_DEFAULT_BEND_SWEEP_DEG)
        radius = max(_MIN_RADIUS_FT, arc_each / sweep)
        assumed["radiusFt"] = round(radius, 1)
        assumed["bendSweepDeg"] = _DEFAULT_BEND_SWEEP_DEG
    else:
        radius = max(_MIN_RADIUS_FT, float(radius_ft))
        if radius != float(radius_ft):
            assumed["radiusClampedFt"] = radius

    straight_total = max(0.0, length_ft - arc_each * n_bends)
    n_straights = n_bends + 1
    straight_each = straight_total / n_straights if n_straights else 0.0

    segs: list[tuple] = []
    # L-bend is a sharp 90 deg corner, not a gentle drift: long approach, tight
    # filleted corner, long exit. Deriving its radius from the 45 deg default
    # produced a 1910 ft "corner" on a 3000 ft run, which reads as a C-curve.
    if kind == "l_bend":
        if radius_ft is None or float(radius_ft) <= 0:
            radius = max(_MIN_RADIUS_FT, length_ft * 0.10)
            assumed["radiusFt"] = round(radius, 1)
            assumed["bendSweepDeg"] = 90.0
        corner = radius * (math.pi / 2.0)
        leg = max(0.0, (length_ft - corner) / 2.0)
        return ([("straight", leg), ("arc", corner, 1.0), ("straight", leg)],
                radius, assumed)

    turn = 1.0
    for i in range(n_bends):
        if straight_each > 0:
            segs.append(("straight", straight_each))
        segs.append(("arc", arc_each, turn))
        # S and N-bend alternate direction; a C-curve keeps turning one way.
        if kind in ("s_curve", "n_bend"):
            turn = -turn
    if straight_each > 0:
        segs.append(("straight", straight_each))
    return (segs, radius, assumed)


def _walk(segments: list[tuple], start_x: float, start_y: float,
          bearing_deg: float, radius: float, step_ft: float) -> list[list[float]]:
    """Sample the segment run into vertices."""
    pts: list[list[float]] = [[float(start_x), float(start_y)]]
    x, y = float(start_x), float(start_y)
    head = math.radians(bearing_deg)
    for seg in segments:
        if seg[0] == "straight":
            L = float(seg[1])
            if L <= 1e-9:
                continue
            n = max(1, int(math.ceil(L / max(step_ft, 1.0))))
            for i in range(1, n + 1):
                d = L * i / n
                pts.append([x + d * math.cos(head), y + d * math.sin(head)])
            x, y = pts[-1][0], pts[-1][1]
        else:
            L = float(seg[1])
            turn = float(seg[2])
            if L <= 1e-9 or radius <= 0:
                continue
            sweep = (L / radius) * turn
            # Arc centre is 90 deg off the heading, on the turn side.
            cx = x + radius * math.cos(head + turn * math.pi / 2.0)
            cy = y + radius * math.sin(head + turn * math.pi / 2.0)
            a0 = math.atan2(y - cy, x - cx)
            n = max(1, int(math.ceil(L / max(step_ft, 1.0))))
            for i in range(1, n + 1):
                a = a0 + sweep * i / n
                pts.append([cx + radius * math.cos(a), cy + radius * math.sin(a)])
            x, y = pts[-1][0], pts[-1][1]
            head += sweep
    return pts


def synthesize_path(length_ft: float,
                    kind: str = "",
                    bends: Optional[int] = None,
                    start_x: float = 0.0, start_y: float = 0.0,
                    bearing_deg: float = 0.0,
                    radius_ft: Optional[float] = None,
                    step_ft: float = 25.0,
                    max_vertices: Optional[int] = None) -> dict:
    """Generate a road path from what the engineer actually said.

    Returns vertices plus every assumption applied, so the caller can state
    them instead of defaulting silently.
    """
    if float(length_ft) <= 0:
        raise ValueError("synthesize_path needs length_ft > 0")
    resolved_kind = infer_kind(bends, kind)
    n_bends = bend_count(resolved_kind, bends)
    segs, radius, assumed = _segments(resolved_kind, float(length_ft), n_bends, radius_ft)
    pts = _walk(segs, start_x, start_y, bearing_deg, radius, step_ft)
    cap = int(max_vertices) if max_vertices else cp.MAX_PATH_VERTS
    verts = cp.downsample_polyline(pts, max_n=cap)
    actual = cp.polyline_length(verts)
    if not kind:
        assumed["kind"] = resolved_kind
    if bends is None and n_bends:
        assumed["bends"] = n_bends
    return {
        "status": "OK",
        "kind": resolved_kind,
        "bends": n_bends,
        "radiusFt": round(radius, 1) if radius else 0.0,
        "requestedLengthFt": round(float(length_ft), 1),
        "actualLengthFt": round(actual, 1),
        "bearingDeg": float(bearing_deg),
        "vertices": verts,
        "vertexCount": len(verts),
        "assumedDefaults": assumed,
        "description": describe_path(resolved_kind, actual, n_bends, radius,
                                     start_x, start_y, bearing_deg),
    }


def describe_path(kind: str, length_ft: float, n_bends: int, radius_ft: float,
                  start_x: float, start_y: float, bearing_deg: float) -> str:
    """One sentence the agent can read back for a yes/no confirmation."""
    shape = {
        "straight": "straight run",
        "c_curve": "constant-radius curve",
        "s_curve": "reverse-S",
        "l_bend": "L-bend",
        "n_bend": f"{n_bends}-bend alignment",
    }.get(kind, kind)
    bend_txt = ""
    if n_bends and radius_ft:
        word = {1: "one", 2: "two", 3: "three"}.get(n_bends, str(n_bends))
        bend_txt = f" with {word} {radius_ft:.0f} ft-radius bend{'s' if n_bends > 1 else ''}"
    return (f"{shape}, {length_ft:.0f} ft{bend_txt}, from "
            f"({start_x:.0f}, {start_y:.0f}) heading {bearing_word(bearing_deg)}")
