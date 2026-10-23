"""Explode curb island C-TOPO-FEAT into classified steel-faced curb segments.

Uses 3d polylines 13DA9A+13DAAA+13DA8B (same path as AECC feature lines;
Civil feature lines cannot be vertex-sampled in this AutoCAD session).
Deletes matching feature lines 13D822/13D824 afterward to avoid duplicates.
"""
from __future__ import annotations

import math
import pythoncom
from win32com.client import Dispatch, VARIANT

SRC = ["13DA9A", "13DAAA", "13DA8B"]
FL_TO_DELETE = ["13D822", "13D824"]
FLARE_HANDLES = ["13DABE", "13DABD", "13D9E6", "13D9E7"]
HEX_HANDLES = ["13DAB9", "13D9E2"]

LAYER_STRAIGHT = "Straight Steel Faced Concrete Curb (18'' Deep)"
LAYER_DEPRESSED = "Depressed Steel Faced Concrete Curb (18'' Deep)"
LAYER_CORNER = "Corner Steel Faced Concrete Curb (18'' Deep)"
COLORS = {LAYER_STRAIGHT: 1, LAYER_DEPRESSED: 3, LAYER_CORNER: 5}
PAD = 0.15


def vpt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def inside(box, x, y):
    return box[0] <= x <= box[2] and box[1] <= y <= box[3]


def classify(mx, my, flares, hexes):
    if any(inside(b, mx, my) for b in flares):
        return LAYER_CORNER
    if any(inside(b, mx, my) for b in hexes):
        return LAYER_DEPRESSED
    return LAYER_STRAIGHT


def poly_pts(e):
    coords = list(e.Coordinates)
    step = 3 if e.ObjectName == "AcDb3dPolyline" else 2
    pts = []
    for i in range(0, len(coords), step):
        z = float(coords[i + 2]) if step == 3 else 0.0
        pts.append((float(coords[i]), float(coords[i + 1]), z))
    out = []
    for p in pts:
        if not out or math.hypot(p[0] - out[-1][0], p[1] - out[-1][1]) > 1e-6:
            out.append(p)
    return out


def ensure_layers(doc):
    for name, col in COLORS.items():
        found = None
        for i in range(doc.Layers.Count):
            if doc.Layers.Item(i).Name == name:
                found = doc.Layers.Item(i)
                break
        if found is None:
            found = doc.Layers.Add(name)
        found.Color = col
        print(f"layer {name!r} color={col}")


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    print("file", doc.FullName)
    ensure_layers(doc)

    # clear prior classified lines
    to_del = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName == "AcDbLine" and e.Layer in COLORS:
                to_del.append(e.Handle)
        except Exception:
            continue
    for h in to_del:
        try:
            doc.HandleToObject(h).Delete()
        except Exception:
            pass
    print(f"cleared prior classified lines: {len(to_del)}")

    flares = []
    for h in FLARE_HANDLES:
        e = doc.HandleToObject(h)
        mn, mx = e.GetBoundingBox()
        flares.append(
            (
                float(mn[0]) - PAD,
                float(mn[1]) - PAD,
                float(mx[0]) + PAD,
                float(mx[1]) + PAD,
            )
        )
    hexes = []
    for h in HEX_HANDLES:
        e = doc.HandleToObject(h)
        mn, mx = e.GetBoundingBox()
        hexes.append(
            (
                float(mn[0]) - PAD,
                float(mn[1]) - PAD,
                float(mx[0]) + PAD,
                float(mx[1]) + PAD,
            )
        )

    # Build connected chain
    chain = []
    src_els = []
    for h in SRC:
        e = doc.HandleToObject(h)
        src_els.append(e)
        pts = poly_pts(e)
        print(f"src {h} n={len(pts)}")
        if not chain:
            chain.extend(pts)
            continue
        end = chain[-1]
        if math.hypot(pts[-1][0] - end[0], pts[-1][1] - end[1]) < math.hypot(
            pts[0][0] - end[0], pts[0][1] - end[1]
        ):
            pts = list(reversed(pts))
        chain.extend(pts)
    dedup = []
    for p in chain:
        if not dedup or math.hypot(p[0] - dedup[-1][0], p[1] - dedup[-1][1]) > 1e-6:
            dedup.append(p)
    print(f"chain n={len(dedup)}")

    totals = {k: 0.0 for k in COLORS}
    counts = {k: 0 for k in COLORS}

    for i in range(len(dedup) - 1):
        a, b = dedup[i], dedup[i + 1]
        length = math.hypot(b[0] - a[0], b[1] - a[1])
        if length < 1e-4:
            continue
        mx, my = 0.5 * (a[0] + b[0]), 0.5 * (a[1] + b[1])
        layer = classify(mx, my, flares, hexes)
        line = ms.AddLine(vpt(a[0], a[1], a[2]), vpt(b[0], b[1], b[2]))
        line.Layer = layer
        line.Color = 256
        totals[layer] += length
        counts[layer] += 1
        print(f"  {length:6.3f} ft  {layer.split()[0]:9s}  mid=({mx:.2f},{my:.2f})")

    for e in src_els:
        h = e.Handle
        try:
            e.Delete()
            print("deleted source", h)
        except Exception as ex:
            print("del src fail", h, ex)

    for h in FL_TO_DELETE:
        try:
            doc.HandleToObject(h).Delete()
            print("deleted feature line", h)
        except Exception as ex:
            print("del FL fail", h, ex)

    doc.Regen(1)
    print("\n===== QUANTITIES (linear feet) =====")
    for lyr in (LAYER_STRAIGHT, LAYER_DEPRESSED, LAYER_CORNER):
        print(f"{lyr}")
        print(f"  segments={counts[lyr]}   LF = {totals[lyr]:.2f}")
    print(f"TOTAL LF = {sum(totals.values()):.2f}")


if __name__ == "__main__":
    main()
