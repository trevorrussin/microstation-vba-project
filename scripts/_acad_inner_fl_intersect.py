"""Extract AECC feature-line points via IntersectWith grid, then classify."""
from __future__ import annotations

import math
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

FL_HANDLES = ["13D822", "13D824"]
FLARE_HANDLES = ["13DABE", "13DABD", "13D9E6", "13D9E7"]
HEX_HANDLES = ["13DAB9", "13D9E2"]

LAYER_STRAIGHT = "Straight Steel Faced Concrete Curb (18'' Deep)"
LAYER_DEPRESSED = "Depressed Steel Faced Concrete Curb (18'' Deep)"
LAYER_CORNER = "Corner Steel Faced Concrete Curb (18'' Deep)"
COLORS = {LAYER_STRAIGHT: 1, LAYER_DEPRESSED: 3, LAYER_CORNER: 5}
PAD = 0.15
GRID = 0.20  # ft


def vpt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def bbox_of(doc, handle, pad=0.0):
    e = doc.HandleToObject(handle)
    mn, mx = e.GetBoundingBox()
    return (
        float(mn[0]) - pad,
        float(mn[1]) - pad,
        float(mx[0]) + pad,
        float(mx[1]) + pad,
        e,
    )


def inside(box, x, y):
    return box[0] <= x <= box[2] and box[1] <= y <= box[3]


def classify(mx, my, flares, hexes):
    if any(inside(b, mx, my) for b in flares):
        return LAYER_CORNER
    if any(inside(b, mx, my) for b in hexes):
        return LAYER_DEPRESSED
    return LAYER_STRAIGHT


def sample_fl(ms, fl, lx, ly, hx, hy):
    """Return unique intersection points along feature line via grid of cut lines."""
    pts = []
    temps = []
    # horizontal cuts
    y = ly
    while y <= hy + 1e-9:
        line = ms.AddLine(vpt(lx - 1, y), vpt(hx + 1, y))
        temps.append(line)
        try:
            inter = fl.IntersectWith(line, 0)
            if inter is not None:
                arr = list(inter)
                for i in range(0, len(arr), 3):
                    pts.append((float(arr[i]), float(arr[i + 1]), float(arr[i + 2])))
        except Exception:
            pass
        y += GRID
    # vertical cuts
    x = lx
    while x <= hx + 1e-9:
        line = ms.AddLine(vpt(x, ly - 1), vpt(x, hy + 1))
        temps.append(line)
        try:
            inter = fl.IntersectWith(line, 0)
            if inter is not None:
                arr = list(inter)
                for i in range(0, len(arr), 3):
                    pts.append((float(arr[i]), float(arr[i + 1]), float(arr[i + 2])))
        except Exception:
            pass
        x += GRID

    for t in temps:
        try:
            t.Delete()
        except Exception:
            pass

    # dedupe
    pts = sorted(pts, key=lambda p: (round(p[0], 4), round(p[1], 4)))
    out = []
    for p in pts:
        if not out or math.hypot(p[0] - out[-1][0], p[1] - out[-1][1]) > 0.05:
            out.append(p)
    return out


def order_polyline(pts):
    """Greedy nearest-neighbor order from endpoint with extreme coordinate."""
    if len(pts) < 2:
        return pts
    # start at point with max Y then max X (north end) — matches island curb direction
    start = max(pts, key=lambda p: (p[1], p[0]))
    remaining = pts[:]
    remaining.remove(start)
    ordered = [start]
    while remaining:
        last = ordered[-1]
        nxt = min(remaining, key=lambda p: math.hypot(p[0] - last[0], p[1] - last[1]))
        remaining.remove(nxt)
        ordered.append(nxt)
    return ordered


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


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    print("file", doc.FullName)
    print("ms", ms.Count)

    ensure_layers(doc)

    # Remove any existing classified lines from earlier attempt
    to_del = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName == "AcDbLine" and e.Layer in COLORS:
                to_del.append(e.Handle)
        except Exception:
            continue
    print(f"clear prior classified: {len(to_del)}")
    for h in to_del:
        try:
            doc.HandleToObject(h).Delete()
        except Exception:
            pass

    flares = []
    for h in FLARE_HANDLES:
        lx, ly, hx, hy, _ = bbox_of(doc, h, PAD)
        flares.append((lx, ly, hx, hy))
    hexes = []
    for h in HEX_HANDLES:
        lx, ly, hx, hy, _ = bbox_of(doc, h, PAD)
        hexes.append((lx, ly, hx, hy))

    totals = {k: 0.0 for k in COLORS}
    counts = {k: 0 for k in COLORS}

    for h in FL_HANDLES:
        lx, ly, hx, hy, fl = bbox_of(doc, h, 0.5)
        print(f"\nSampling {h} bbox=[{lx:.2f},{ly:.2f}]-[{hx:.2f},{hy:.2f}]")
        raw = sample_fl(ms, fl, lx, ly, hx, hy)
        print(f"  raw intersections={len(raw)}")
        ordered = order_polyline(raw)
        print(f"  ordered={len(ordered)}")
        if len(ordered) < 2:
            print("  SKIP too few points")
            continue

        for i in range(len(ordered) - 1):
            a, b = ordered[i], ordered[i + 1]
            length = math.hypot(b[0] - a[0], b[1] - a[1])
            if length < 0.05 or length > 5.0:
                # skip tiny or jump artifacts from NN ordering
                continue
            mx = 0.5 * (a[0] + b[0])
            my = 0.5 * (a[1] + b[1])
            layer = classify(mx, my, flares, hexes)
            line = ms.AddLine(vpt(a[0], a[1], a[2]), vpt(b[0], b[1], b[2]))
            line.Layer = layer
            line.Color = 256
            totals[layer] += length
            counts[layer] += 1
            print(f"  seg {length:5.2f} {layer.split()[0]:9s} mid=({mx:.2f},{my:.2f})")

        # Delete original feature line (exploded into segs)
        try:
            fl.Delete()
            print(f"  deleted FL {h}")
        except Exception as ex:
            print(f"  delete FL fail {h}: {ex}")

    doc.Regen(1)
    print("\n===== QUANTITIES (LF) INNER FEATURE LINES =====")
    for lyr in COLORS:
        print(f"  {lyr}")
        print(f"    n={counts[lyr]}  LF={totals[lyr]:.2f}")
    print(f"  TOTAL LF={sum(totals.values()):.2f}")


if __name__ == "__main__":
    main()
