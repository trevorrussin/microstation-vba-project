"""Reclassify INNER curb segments using engineer-placed yellow flares.

Authority (user defs):
  Corner    = along/under yellow PROP-D-ROAD-RAMPS triangles (flares)
  Depressed = between the two flares of each ramp along the inner face
  Straight  = rest of the inner curb (beyond flares)

Never touches OUTER C-TOPO-FEAT / outer FL handles.
Only re-layers existing AcDbLine on the three steel-faced layers.
"""
from __future__ import annotations

import math
import pythoncom
from win32com.client import Dispatch

LAYER_STRAIGHT = "Straight Steel Faced Concrete Curb (18'' Deep)"
LAYER_DEPRESSED = "Depressed Steel Faced Concrete Curb (18'' Deep)"
LAYER_CORNER = "Corner Steel Faced Concrete Curb (18'' Deep)"
LAYER_FLARE = "PROP-D-ROAD-RAMPS"
COLORS = {LAYER_STRAIGHT: 1, LAYER_DEPRESSED: 3, LAYER_CORNER: 5}

# Island clusters for pairing flares + depressed span between them
ISLANDS = [
    {"name": "NW", "center": (1028470.0, 187478.0), "r": 30.0},
    {"name": "MidN", "center": (1028501.0, 187485.0), "r": 22.0},
    {"name": "MidS", "center": (1028525.0, 187428.0), "r": 22.0},
    {"name": "E", "center": (1028589.0, 187445.0), "r": 22.0},
    {"name": "N", "center": (1028564.0, 187503.0), "r": 25.0},
    {"name": "SW", "center": (1028486.0, 187421.0), "r": 25.0},
]

# Outer FL handles — never modify these entities
OUTER_HANDLES = {"13D80A", "13D25C", "13D7DC", "13CEC7", "13D2EA", "13D7BC"}

CORNER_PAD = 0.75  # ft — treat as Corner if mid within this of flare tri


def dist(a, b):
    return math.hypot(b[0] - a[0], b[1] - a[1])


def point_in_tri(p, a, b, c, eps=1e-9):
    """Barycentric inside test."""
    x, y = p
    x1, y1 = a
    x2, y2 = b
    x3, y3 = c
    denom = (y2 - y3) * (x1 - x3) + (x3 - x2) * (y1 - y3)
    if abs(denom) < eps:
        return False
    a_ = ((y2 - y3) * (x - x3) + (x3 - x2) * (y - y3)) / denom
    b_ = ((y3 - y1) * (x - x3) + (x1 - x3) * (y - y3)) / denom
    c_ = 1.0 - a_ - b_
    return a_ >= -1e-6 and b_ >= -1e-6 and c_ >= -1e-6


def dist_to_tri(p, tri):
    if point_in_tri(p, tri[0], tri[1], tri[2]):
        return 0.0
    best = 1e9
    for i in range(3):
        a, b = tri[i], tri[(i + 1) % 3]
        ax, ay = a
        bx, by = b
        dx, dy = bx - ax, by - ay
        L2 = dx * dx + dy * dy
        t = 0.0 if L2 < 1e-12 else max(0.0, min(1.0, ((p[0] - ax) * dx + (p[1] - ay) * dy) / L2))
        qx, qy = ax + t * dx, ay + t * dy
        best = min(best, math.hypot(p[0] - qx, p[1] - qy))
    return best


def solid_tri(e):
    pts = []
    for pi in range(4):
        try:
            p = e.Coordinate(pi)
            pts.append((float(p[0]), float(p[1])))
        except Exception:
            break
    u = []
    for p in pts:
        if not any(dist(p, q) < 1e-4 for q in u):
            u.append(p)
    if len(u) < 3:
        return None
    # area filter — skip degenerate
    a, b, c = u[0], u[1], u[2]
    area = abs((a[0] * (b[1] - c[1]) + b[0] * (c[1] - a[1]) + c[0] * (a[1] - b[1])) / 2)
    if area < 0.2:
        return None
    return (u[0], u[1], u[2])


def ensure_layer(doc, name, color):
    layers = doc.Layers
    lyr = None
    for i in range(layers.Count):
        try:
            if layers.Item(i).Name == name:
                lyr = layers.Item(i)
                break
        except Exception:
            continue
    if lyr is None:
        lyr = layers.Add(name)
    lyr.Color = color
    return lyr


def collect_flares(doc):
    ms = doc.ModelSpace
    flares = []  # (cx, cy, tri, handle)
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbSolid" or e.Layer != LAYER_FLARE:
                continue
            tri = solid_tri(e)
            if not tri:
                continue
            cx = sum(p[0] for p in tri) / 3
            cy = sum(p[1] for p in tri) / 3
            flares.append((cx, cy, tri, e.Handle))
        except Exception:
            continue
    return flares


def island_flare_pairs(flares):
    """For each island, the flares near it (expect ~2+)."""
    out = {}
    for isl in ISLANDS:
        near = [
            f
            for f in flares
            if dist((f[0], f[1]), isl["center"]) <= isl["r"]
        ]
        out[isl["name"]] = near
    return out


def angle_of(origin, p):
    return math.atan2(p[1] - origin[1], p[0] - origin[0])


def depressed_wedge(island_center, flare_a, flare_b):
    """Return True if point is angularly between the two flare centroids."""
    # Use flare centroids relative to island center
    return None  # handled via station on segment mid vs flare angles


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace

    for name, color in COLORS.items():
        ensure_layer(doc, name, color)

    flares = collect_flares(doc)
    print(f"flares usable: {len(flares)}")
    by_isl = island_flare_pairs(flares)
    for name, fls in by_isl.items():
        print(f"  {name}: {len(fls)} flares {[h for *_, h in fls]}")

    # Precompute per-island angular span between outermost two flares
    isl_spans = {}
    for isl in ISLANDS:
        fls = by_isl[isl["name"]]
        if len(fls) < 2:
            continue
        angs = sorted(angle_of(isl["center"], (f[0], f[1])) for f in fls)
        # pair with largest gap complementary = depressed opening
        # Use min/max angle of flares; depressed = between them on shorter arc
        a0, a1 = angs[0], angs[-1]
        # also consider wrap
        isl_spans[isl["name"]] = {
            "center": isl["center"],
            "r": isl["r"],
            "a0": a0,
            "a1": a1,
            "flares": fls,
        }

    def classify_mid(mx, my):
        p = (mx, my)
        # Corner if near any flare triangle
        for _cx, _cy, tri, _h in flares:
            if dist_to_tri(p, tri) <= CORNER_PAD:
                return LAYER_CORNER
        # Depressed if inside an island's flare-to-flare angular wedge
        # AND closer to island than outer ring — and NOT beyond flares
        for name, sp in isl_spans.items():
            cx, cy = sp["center"]
            if dist(p, (cx, cy)) > sp["r"]:
                continue
            ang = angle_of((cx, cy), p)
            a0, a1 = sp["a0"], sp["a1"]
            # normalize ang into [a0, a1] if a1-a0 is the short span
            span = a1 - a0
            if span <= math.pi:
                if a0 - 0.15 <= ang <= a1 + 0.15:
                    # exclude if outside the radial band of flares
                    return LAYER_DEPRESSED
            else:
                # wrap case: depressed is outside [a0,a1]
                if ang <= a0 + 0.15 or ang >= a1 - 0.15:
                    return LAYER_DEPRESSED
        return LAYER_STRAIGHT

    # Re-layer all classified curb lines
    moved = {LAYER_STRAIGHT: 0.0, LAYER_DEPRESSED: 0.0, LAYER_CORNER: 0.0}
    counts = {LAYER_STRAIGHT: 0, LAYER_DEPRESSED: 0, LAYER_CORNER: 0}
    changed = 0
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Handle in OUTER_HANDLES:
                continue
            if e.ObjectName != "AcDbLine":
                continue
            if e.Layer not in COLORS:
                continue
            s, en = e.StartPoint, e.EndPoint
            mx = 0.5 * (float(s[0]) + float(en[0]))
            my = 0.5 * (float(s[1]) + float(en[1]))
            L = math.hypot(float(en[0]) - float(s[0]), float(en[1]) - float(s[1]))
            new_lyr = classify_mid(mx, my)
            if e.Layer != new_lyr:
                e.Layer = new_lyr
                e.Color = 256
                changed += 1
            moved[new_lyr] += L
            counts[new_lyr] += 1
        except Exception:
            continue

    try:
        doc.Regen(1)
    except Exception:
        pass

    print(f"\nchanged layers: {changed}")
    print("NEW LF:")
    for lyr in (LAYER_STRAIGHT, LAYER_DEPRESSED, LAYER_CORNER):
        print(f"  {counts[lyr]:4d} segs  {moved[lyr]:8.1f} ft  {lyr}")
    print("DONE")


if __name__ == "__main__":
    main()
