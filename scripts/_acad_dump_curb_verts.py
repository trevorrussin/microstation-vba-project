"""Dump vertices for curb candidates + yellow flare solids near island."""
from __future__ import annotations

import math
import pythoncom
from win32com.client import Dispatch

HANDLES_3DPOLY = ["13DA8B", "13DA6F", "13DA4F", "13DA9A", "13DAAA"]
HANDLES_FL = ["13D822", "13D824", "13D425", "13D2EA", "13D2EE", "13D42D"]
HANDLES_SOLIDS = ["13D9E7", "13D9E6", "13DABD", "13DABE"]
HANDLES_HEX = ["13DAB9", "13D9E2"]  # warning strip outlines


def get_ent(doc, handle):
    return doc.HandleToObject(handle)


def dump_poly(e, label):
    print(f"\n=== {label} handle={e.Handle} type={e.ObjectName} layer={e.Layer} ===")
    try:
        print(f"  color={e.Color} len={getattr(e, 'Length', '?')}")
    except Exception:
        pass
    # Coordinates for LW/3d poly
    try:
        coords = list(e.Coordinates)
        # 3d poly: x,y,z triples; lwpoly: x,y pairs
        step = 3 if e.ObjectName == "AcDb3dPolyline" else 2
        pts = []
        for i in range(0, len(coords), step):
            pts.append((float(coords[i]), float(coords[i + 1]), float(coords[i + 2]) if step == 3 else 0.0))
        print(f"  nverts={len(pts)}")
        for i, p in enumerate(pts):
            print(f"    [{i}] ({p[0]:.4f}, {p[1]:.4f}, {p[2]:.4f})")
        # segment lengths
        total = 0.0
        for i in range(len(pts) - 1):
            d = math.hypot(pts[i + 1][0] - pts[i][0], pts[i + 1][1] - pts[i][1])
            total += d
            print(f"    seg {i}->{i+1} len={d:.4f}")
        print(f"  planar_len={total:.4f}")
        return pts
    except Exception as ex:
        print("  coords fail", ex)

    # Feature line: try GetPoints
    for meth in ("GetPoints", "Points", "Coordinates"):
        if hasattr(e, meth):
            try:
                print(f"  trying {meth}")
                val = getattr(e, meth)
                if callable(val):
                    # GetPoints may need an enum arg
                    try:
                        val = val(0)
                    except Exception:
                        try:
                            val = val()
                        except Exception as e2:
                            print(f"    call fail {e2}")
                            continue
                print(f"    type={type(val)} repr={str(val)[:200]}")
            except Exception as e2:
                print(f"  {meth} fail {e2}")
    return None


def dump_solid(e):
    print(f"\n=== SOLID {e.Handle} layer={e.Layer} color={e.Color} ===")
    pts = []
    for i in range(4):
        try:
            p = e.Coordinate(i)
            pts.append((float(p[0]), float(p[1])))
            print(f"  p{i}=({p[0]:.4f},{p[1]:.4f})")
        except Exception as ex:
            print(f"  p{i} fail {ex}")
    if len(pts) >= 3:
        cx = sum(p[0] for p in pts[:3]) / 3
        cy = sum(p[1] for p in pts[:3]) / 3
        print(f"  centroid~({cx:.4f},{cy:.4f})")
    return pts


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument

    for h in HANDLES_3DPOLY:
        try:
            dump_poly(get_ent(doc, h), "3DPOLY")
        except Exception as e:
            print(h, e)

    for h in HANDLES_FL:
        try:
            dump_poly(get_ent(doc, h), "FEATLINE")
        except Exception as e:
            print(h, e)

    for h in HANDLES_SOLIDS:
        try:
            dump_solid(get_ent(doc, h))
        except Exception as e:
            print(h, e)

    for h in HANDLES_HEX:
        try:
            dump_poly(get_ent(doc, h), "HEX_OUTLINE")
        except Exception as e:
            print(h, e)


if __name__ == "__main__":
    main()
