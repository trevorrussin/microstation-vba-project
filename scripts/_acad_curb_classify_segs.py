"""Explode inner curb 3d-polylines into classified steel-faced curb segments.

Source: C-TOPO-FEAT 3d polylines 13DA9A + 13DAAA + 13DA8B (street-face curb
along the top-left curb island / ramps). Civil AECC_FEATURE_LINE objects on the
same layer (13D822/13D824) cannot be exploded to vertices via COM.

Classification (by segment midpoint):
  Corner    – inside a yellow PROP-D-ROAD-RAMPS solid (flare)
  Depressed – inside a C-P-WARNIG STRIP hex outline (ramp pad), not corner
  Straight  – everywhere else along this curb (beyond flares)
"""
from __future__ import annotations

import math
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

SRC_HANDLES = ["13DA9A", "13DAAA", "13DA8B"]
FLARE_HANDLES = ["13DABE", "13DABD", "13D9E6", "13D9E7"]
HEX_HANDLES = ["13DAB9", "13D9E2"]

# AutoCAD layer names cannot contain " — use 18'' for inches
LAYER_STRAIGHT = "Straight Steel Faced Concrete Curb (18'' Deep)"
LAYER_DEPRESSED = "Depressed Steel Faced Concrete Curb (18'' Deep)"
LAYER_CORNER = "Corner Steel Faced Concrete Curb (18'' Deep)"

# Distinct ACI colors
COLOR_STRAIGHT = 1   # red
COLOR_DEPRESSED = 3  # green
COLOR_CORNER = 5     # blue

PAD = 0.15  # ft pad on flare/hex boxes for midpoint test


def dedupe_pts(pts, tol=1e-6):
    out = []
    for p in pts:
        if not out or math.hypot(p[0] - out[-1][0], p[1] - out[-1][1]) > tol:
            out.append(p)
    return out


def poly_pts(e):
    coords = list(e.Coordinates)
    step = 3 if e.ObjectName == "AcDb3dPolyline" else 2
    pts = []
    for i in range(0, len(coords), step):
        z = float(coords[i + 2]) if step == 3 else 0.0
        pts.append((float(coords[i]), float(coords[i + 1]), z))
    return dedupe_pts(pts)


def bbox_of(doc, handle, pad=0.0):
    e = doc.HandleToObject(handle)
    mn, mx = e.GetBoundingBox()
    return (
        float(mn[0]) - pad,
        float(mn[1]) - pad,
        float(mx[0]) + pad,
        float(mx[1]) + pad,
    )


def inside(box, x, y):
    return box[0] <= x <= box[2] and box[1] <= y <= box[3]


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
    print(f"layer OK: {name!r} color={color}")
    return lyr


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    print("file", doc.FullName)

    # Build ordered chain: 13DA9A (north) -> 13DAAA -> 13DA8B (south)
    # Orient so they connect end-to-end.
    pieces = []
    for h in SRC_HANDLES:
        e = doc.HandleToObject(h)
        pts = poly_pts(e)
        pieces.append({"handle": h, "pts": pts, "el": e})
        print(f"src {h} n={len(pts)} start={pts[0][:2]} end={pts[-1][:2]}")

    # Connect: find order by matching endpoints
    # Expected: 13DA9A ends near 13DAAA start; 13DAAA ends near 13DA8B start
    chain = []
    for p in pieces:
        pts = p["pts"]
        if not chain:
            chain.extend(pts)
            continue
        # attach in orientation that matches current chain end
        end = chain[-1]
        d0 = math.hypot(pts[0][0] - end[0], pts[0][1] - end[1])
        d1 = math.hypot(pts[-1][0] - end[0], pts[-1][1] - end[1])
        if d1 < d0:
            pts = list(reversed(pts))
            d0 = d1
        if d0 > 1.0:
            print(f"WARN gap {d0:.3f} before {p['handle']}")
        chain.extend(pts)
    chain = dedupe_pts(chain)
    print(f"chain vertices={len(chain)}")

    flares = [bbox_of(doc, h, PAD) for h in FLARE_HANDLES]
    hexes = [bbox_of(doc, h, PAD) for h in HEX_HANDLES]
    for i, b in enumerate(flares):
        print(f"flare[{i}] {FLARE_HANDLES[i]} {b}")
    for i, b in enumerate(hexes):
        print(f"hex[{i}] {HEX_HANDLES[i]} {b}")

    ensure_layer(doc, LAYER_STRAIGHT, COLOR_STRAIGHT)
    ensure_layer(doc, LAYER_DEPRESSED, COLOR_DEPRESSED)
    ensure_layer(doc, LAYER_CORNER, COLOR_CORNER)

    totals = {LAYER_STRAIGHT: 0.0, LAYER_DEPRESSED: 0.0, LAYER_CORNER: 0.0}
    counts = {LAYER_STRAIGHT: 0, LAYER_DEPRESSED: 0, LAYER_CORNER: 0}
    created = []

    for i in range(len(chain) - 1):
        p0, p1 = chain[i], chain[i + 1]
        length = math.hypot(p1[0] - p0[0], p1[1] - p0[1])
        if length < 1e-4:
            continue
        mx = 0.5 * (p0[0] + p1[0])
        my = 0.5 * (p0[1] + p1[1])

        if any(inside(b, mx, my) for b in flares):
            layer = LAYER_CORNER
        elif any(inside(b, mx, my) for b in hexes):
            layer = LAYER_DEPRESSED
        else:
            layer = LAYER_STRAIGHT

        # 3D line preserving Z — AutoCAD COM needs VARIANT point arrays
        sp = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (p0[0], p0[1], p0[2]))
        ep = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (p1[0], p1[1], p1[2]))
        line = ms.AddLine(sp, ep)
        line.Layer = layer
        line.Color = 256  # ByLayer
        created.append((line.Handle, layer, length, mx, my))
        totals[layer] += length
        counts[layer] += 1
        print(
            f"  seg {i}: {length:6.3f} ft -> {layer.split()[0]:9s} "
            f"mid=({mx:.2f},{my:.2f}) id={line.Handle}"
        )

    # Remove source 3d polylines (exploded)
    for p in pieces:
        try:
            p["el"].Delete()
            print(f"deleted source {p['handle']}")
        except Exception as ex:
            print(f"delete fail {p['handle']}: {ex}")

    doc.Regen(1)
    time.sleep(0.2)

    print("\n===== QUANTITIES (linear feet) =====")
    for lyr in (LAYER_STRAIGHT, LAYER_DEPRESSED, LAYER_CORNER):
        print(f"  {lyr}")
        print(f"    segments={counts[lyr]}  LF={totals[lyr]:.2f}")
    print(f"  TOTAL LF={sum(totals.values()):.2f}")
    print(f"created={len(created)} lines")


if __name__ == "__main__":
    main()
