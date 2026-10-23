"""Find all explodable curb polylines/arcs for location 01 islands."""
from __future__ import annotations

import math
import pythoncom
from win32com.client import Dispatch


def dump_poly(e, label):
    print(f"\n{label} {e.Handle} {e.ObjectName} lyr={e.Layer}")
    if e.ObjectName == "AcDbArc":
        c = e.Center
        print(
            f"  center=({float(c[0]):.2f},{float(c[1]):.2f}) r={float(e.Radius):.2f} "
            f"startAng={math.degrees(float(e.StartAngle)):.1f} "
            f"endAng={math.degrees(float(e.EndAngle)):.1f} "
            f"arcLen={float(e.ArcLength):.2f}"
        )
        return
    try:
        coords = list(e.Coordinates)
    except Exception as ex:
        print("  no coords", ex)
        return
    step = 3 if e.ObjectName == "AcDb3dPolyline" else 2
    pts = []
    for i in range(0, len(coords), step):
        z = float(coords[i + 2]) if step == 3 else 0.0
        pts.append((float(coords[i]), float(coords[i + 1]), z))
    # dedupe consec
    out = []
    for p in pts:
        if not out or math.hypot(p[0] - out[-1][0], p[1] - out[-1][1]) > 1e-6:
            out.append(p)
    print(f"  n={len(out)} len={getattr(e,'Length',None)}")
    for i, p in enumerate(out):
        print(f"    [{i}] ({p[0]:.3f},{p[1]:.3f},{p[2]:.3f})")


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace

    # All 3dpolys on C-TOPO-FEAT
    print("=== C-TOPO-FEAT 3d/lw polys & arcs in loc01 ===")
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer != "C-TOPO-FEAT":
                continue
            if e.ObjectName not in ("AcDb3dPolyline", "AcDbPolyline", "AcDbArc", "AcDbLine"):
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if not (1028430 < cx < 1028630 and 187390 < cy < 187540):
                continue
            dump_poly(e, "TOPO")
        except Exception:
            continue

    # Key PRO-S-CURB-LINE near each ramp cluster
    print("\n=== PRO-S-CURB-LINE polys/arcs near ramps ===")
    handles = [
        "13D986",
        "138A44",
        "13A7EA",
        "13C903",
        "138B86",
        "138B6B",
        "13C226",
        "13383E",
        "13C115",
        "13C116",
        "13C17C",
        "13C17D",
    ]
    for h in handles:
        try:
            dump_poly(doc.HandleToObject(h), "CURB")
        except Exception as ex:
            print(h, ex)


if __name__ == "__main__":
    main()
