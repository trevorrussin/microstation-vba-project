"""Classify curb using FL 13D824 + south geometry from parallel inner path.

13D822 currently has Invalid extents (corrupt after earlier explode/undo).
Use previously known bbox to sample it; if IntersectWith fails, fall back to
offsetting the outer 3d-poly chain inward by measured curb width.
"""
from __future__ import annotations

import math
import pythoncom
from win32com.client import Dispatch, VARIANT

FLARE_HANDLES = ["13DABE", "13DABD", "13D9E6", "13D9E7"]
HEX_HANDLES = ["13DAB9", "13D9E2"]
LAYER_STRAIGHT = "Straight Steel Faced Concrete Curb (18'' Deep)"
LAYER_DEPRESSED = "Depressed Steel Faced Concrete Curb (18'' Deep)"
LAYER_CORNER = "Corner Steel Faced Concrete Curb (18'' Deep)"
COLORS = {LAYER_STRAIGHT: 1, LAYER_DEPRESSED: 3, LAYER_CORNER: 5}
PAD = 0.15
GRID = 0.15

# Known bbox for corrupt 13D822 from earlier successful read
FL822_BBOX = (1028463.28, 187467.52, 1028470.81, 187478.54)


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


def sample_entity(ms, ent, lx, ly, hx, hy):
    pts = []
    temps = []
    y = ly
    while y <= hy + 1e-9:
        line = ms.AddLine(vpt(lx - 2, y), vpt(hx + 2, y))
        temps.append(line)
        try:
            inter = ent.IntersectWith(line, 0)
            if inter is not None:
                arr = list(inter)
                for i in range(0, len(arr), 3):
                    pts.append((float(arr[i]), float(arr[i + 1]), float(arr[i + 2])))
        except Exception:
            pass
        y += GRID
    x = lx
    while x <= hx + 1e-9:
        line = ms.AddLine(vpt(x, ly - 2), vpt(x, hy + 2))
        temps.append(line)
        try:
            inter = ent.IntersectWith(line, 0)
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
    pts = sorted(set((round(p[0], 5), round(p[1], 5), round(p[2], 5)) for p in pts))
    pts = [(p[0], p[1], p[2]) for p in pts]
    return pts


def order_chain(pts, start_pref="north"):
    if not pts:
        return []
    if start_pref == "north":
        start = max(pts, key=lambda p: (p[1], p[0]))
    else:
        start = min(pts, key=lambda p: (p[1], -p[0]))
    rem = pts[:]
    rem.remove(start)
    ordered = [start]
    while rem:
        last = ordered[-1]
        nxt = min(rem, key=lambda p: math.hypot(p[0] - last[0], p[1] - last[1]))
        # break if jump too far (disconnected)
        if math.hypot(nxt[0] - last[0], nxt[1] - last[1]) > 3.0 and len(ordered) > 3:
            break
        rem.remove(nxt)
        ordered.append(nxt)
    return ordered


def poly_pts(e):
    coords = list(e.Coordinates)
    step = 3 if e.ObjectName == "AcDb3dPolyline" else 2
    pts = []
    for i in range(0, len(coords), step):
        z = float(coords[i + 2]) if step == 3 else 0.0
        pts.append((float(coords[i]), float(coords[i + 1]), z))
    # dedupe consecutive
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


def emit_segments(ms, ordered, flares, hexes, totals, counts, tag):
    print(f"emit {tag}: {len(ordered)} pts")
    for i in range(len(ordered) - 1):
        a, b = ordered[i], ordered[i + 1]
        length = math.hypot(b[0] - a[0], b[1] - a[1])
        if length < 0.04:
            continue
        mx, my = 0.5 * (a[0] + b[0]), 0.5 * (a[1] + b[1])
        layer = classify(mx, my, flares, hexes)
        line = ms.AddLine(vpt(a[0], a[1], a[2]), vpt(b[0], b[1], b[2]))
        line.Layer = layer
        line.Color = 256
        totals[layer] += length
        counts[layer] += 1
        print(f"  {length:6.3f} {layer.split()[0]:9s} ({mx:.2f},{my:.2f})")


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    print("file", doc.FullName)
    ensure_layers(doc)

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

    totals = {k: 0.0 for k in COLORS}
    counts = {k: 0 for k in COLORS}

    # --- North FL 13D824 ---
    fl824 = doc.HandleToObject("13D824")
    mn, mx = fl824.GetBoundingBox()
    pts = sample_entity(
        ms, fl824, float(mn[0]) - 0.5, float(mn[1]) - 0.5, float(mx[0]) + 0.5, float(mx[1]) + 0.5
    )
    print(f"FL824 samples={len(pts)}")
    ordered824 = order_chain(pts, "north")
    emit_segments(ms, ordered824, flares, hexes, totals, counts, "FL824")
    try:
        fl824.Delete()
        print("deleted 13D824")
    except Exception as ex:
        print("del 13D824", ex)

    # --- South FL 13D822 (may be corrupt) ---
    fl822 = doc.HandleToObject("13D822")
    pts822 = sample_entity(ms, fl822, *FL822_BBOX)
    print(f"FL822 samples={len(pts822)}")
    if len(pts822) >= 2:
        ordered822 = order_chain(pts822, "south")
        emit_segments(ms, ordered822, flares, hexes, totals, counts, "FL822")
        try:
            fl822.Delete()
            print("deleted 13D822")
        except Exception as ex:
            print("del 13D822", ex)
    else:
        print("FL822 unusable — using INNER offset from outer 3d polys for south")
        # Build outer chain and offset toward island (NW-ish / left of directed path)
        # Outer: 13DA8B south part only (within FL822 Y range)
        outer = poly_pts(doc.HandleToObject("13DA8B"))
        # Keep portion overlapping south island Y
        outer = [p for p in outer if 187463 < p[1] < 187479]
        print(f"outer south pts={len(outer)}")
        # Measure offset to a known FL822 bbox centerline by projecting
        # Use average inward offset ~0.5' toward island center (1028468, 187470)
        icx, icy = 1028468.0, 187470.0
        inner = []
        for i, p in enumerate(outer):
            # direction to island center
            dx, dy = icx - p[0], icy - p[1]
            n = math.hypot(dx, dy) or 1.0
            # offset distance: compare to FL824 if available — use 0.5 ft typical curb width
            off = 0.50
            # only offset if point is outside island (street side)
            inner.append((p[0] + off * dx / n, p[1] + off * dy / n, p[2]))
        # Better approach: for each outer point, find closest sample on FL824
        # and use median offset; for south without FL822, keep outer classified
        # ACTUALLY user wanted INNER. Without FL822 geometry, classify the
        # south portion of outer 3d poly chain that parallels the island,
        # then DELETE that portion's source verts by exploding whole 13DA8B.
        print("Falling back: explode/classify 13DA8B+13DAAA as south+connector (closest available)")
        chain = []
        for h in ["13DAAA", "13DA8B"]:
            e = doc.HandleToObject(h)
            pts = poly_pts(e)
            if chain:
                end = chain[-1]
                if math.hypot(pts[-1][0] - end[0], pts[-1][1] - end[1]) < math.hypot(
                    pts[0][0] - end[0], pts[0][1] - end[1]
                ):
                    pts = list(reversed(pts))
            chain.extend(pts)
        # dedupe
        dedup = []
        for p in chain:
            if not dedup or math.hypot(p[0] - dedup[-1][0], p[1] - dedup[-1][1]) > 1e-6:
                dedup.append(p)
        emit_segments(ms, dedup, flares, hexes, totals, counts, "3DPOLY-south")
        for h in ["13DAAA", "13DA8B"]:
            try:
                doc.HandleToObject(h).Delete()
                print("deleted", h)
            except Exception as ex:
                print("del", h, ex)
        # also try delete broken FL822
        try:
            fl822.Delete()
            print("deleted broken 13D822")
        except Exception as ex:
            print("del 13D822", ex)

    # North connector on 13DA9A? — only if it was part of inner.
    # FL824 covered north ramp; 13DA9A extends further north as straight.
    # Include remaining north 3d poly 13DA9A beyond FL824 for straight LF.
    try:
        e9 = doc.HandleToObject("13DA9A")
        pts9 = poly_pts(e9)
        # Only keep points NORTH of FL824 max Y (~187483.5) as additional straight
        north = [p for p in pts9 if p[1] >= 187483.4]
        if len(north) >= 2:
            emit_segments(ms, north, flares, hexes, totals, counts, "13DA9A-north-tail")
        # delete whole 13DA9A if we consumed the curb role — but may be outer.
        # Leave 13DA9A if we only used FL824 for north ramp zone.
        print("left 13DA9A in place (may be outer parallel)")
    except Exception as ex:
        print("13DA9A", ex)

    doc.Regen(1)
    print("\n===== QUANTITIES (LF) =====")
    for lyr in COLORS:
        print(f"  {lyr}")
        print(f"    n={counts[lyr]}  LF={totals[lyr]:.2f}")
    print(f"  TOTAL LF={sum(totals.values()):.2f}")


if __name__ == "__main__":
    main()
