"""Explode copies of island feature lines to inspect geometry; identify inner curb."""
from __future__ import annotations

import math
import pythoncom
from win32com.client import Dispatch, VARIANT
import win32com.client


def planar_len(pts):
    t = 0.0
    for i in range(len(pts) - 1):
        t += math.hypot(pts[i + 1][0] - pts[i][0], pts[i + 1][1] - pts[i][1])
    return t


def coords_of(e):
    et = e.ObjectName
    if et in ("AcDbLine",):
        s, en = e.StartPoint, e.EndPoint
        return [(float(s[0]), float(s[1]), float(s[2])), (float(en[0]), float(en[1]), float(en[2]))]
    if et in ("AcDbArc",):
        # sample ends
        s, en = e.StartPoint, e.EndPoint
        return [(float(s[0]), float(s[1]), 0), (float(en[0]), float(en[1]), 0)]
    if "Poly" in et:
        coords = list(e.Coordinates)
        step = 3 if et == "AcDb3dPolyline" else 2
        pts = []
        for i in range(0, len(coords), step):
            z = float(coords[i + 2]) if step == 3 else 0.0
            pts.append((float(coords[i]), float(coords[i + 1]), z))
        return pts
    return []


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace

    # TrueColor of candidates
    for h in ["13D822", "13D824", "13DA8B", "13DA9A", "13DAAA", "13DA4F", "13DA6F"]:
        e = doc.HandleToObject(h)
        try:
            tc = e.TrueColor
            print(f"{h} {e.ObjectName} layer={e.Layer} ColorIndex={e.Color} RGB=({tc.Red},{tc.Green},{tc.Blue})")
        except Exception as ex:
            print(h, ex)

    # Explode COPIES of feature lines (don't destroy originals yet)
    for h in ["13D822", "13D824"]:
        e = doc.HandleToObject(h)
        mn, mx = e.GetBoundingBox()
        print(f"\n--- explode copy of {h} bbox=[{mn[0]:.2f},{mn[1]:.2f}]-[{mx[0]:.2f},{mx[1]:.2f}] ---")
        # Copy in place
        try:
            copied = e.Copy()
        except Exception as ex:
            print("Copy fail", ex)
            continue
        # Copy may return variant array
        if isinstance(copied, (list, tuple)):
            copied = copied[0]
        try:
            # Explode returns array of objects
            result = win32com.client.VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY | pythoncom.VT_BYREF, [])
            # AcadEntity.Explode() returns SafeArray of objects
            exploded = copied.Explode()
            print("exploded type", type(exploded), exploded)
            items = list(exploded) if exploded is not None else []
            print(f"  n pieces={len(items)}")
            for p in items:
                pts = coords_of(p)
                plen = planar_len(pts) if pts else None
                print(
                    f"  piece {p.Handle} {p.ObjectName} layer={p.Layer} "
                    f"nverts={len(pts)} plen={plen}"
                )
                for i, pt in enumerate(pts[:8]):
                    print(f"    [{i}] ({pt[0]:.3f},{pt[1]:.3f},{pt[2]:.3f})")
                if len(pts) > 8:
                    print(f"    ... +{len(pts)-8} more")
            # cleanup exploded copies + leftover
            for p in items:
                try:
                    p.Delete()
                except Exception:
                    pass
            try:
                copied.Delete()
            except Exception:
                pass
        except Exception as ex:
            print("Explode fail", ex)
            try:
                copied.Delete()
            except Exception:
                pass

    # Flare solid extents for classification zones
    print("\n=== Flare zones (yellow solids) ===")
    for h in ["13DABE", "13DABD", "13D9E6", "13D9E7"]:
        e = doc.HandleToObject(h)
        mn, mx = e.GetBoundingBox()
        print(
            f"  {h} [{float(mn[0]):.3f},{float(mn[1]):.3f}]-"
            f"[{float(mx[0]):.3f},{float(mx[1]):.3f}]"
        )

    print("\n=== Hex ramp pads ===")
    for h in ["13DAB9", "13D9E2"]:
        e = doc.HandleToObject(h)
        mn, mx = e.GetBoundingBox()
        print(
            f"  {h} [{float(mn[0]):.3f},{float(mn[1]):.3f}]-"
            f"[{float(mx[0]):.3f},{float(mx[1]):.3f}]"
        )


if __name__ == "__main__":
    main()
