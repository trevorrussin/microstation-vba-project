"""Dump NW steel segs vs rules; compare N/SW segment layers along station."""
from __future__ import annotations

import math
import pythoncom
from win32com.client import GetObject

STEEL = {
    "Straight Steel Faced Concrete Curb (18'' Deep)": "Straight",
    "Depressed Steel Faced Concrete Curb (18'' Deep)": "Depressed",
    "Corner Steel Faced Concrete Curb (18'' Deep)": "Corner",
}
NW = (1028470.0, 187478.0)


def main():
    pythoncom.CoInitialize()
    doc = GetObject(None, "AutoCAD.Application").ActiveDocument
    ms = doc.ModelSpace
    print("file", doc.Name)
    print("RULES (engineer / NW authority):")
    print("  Depressed = ON the ped ramp from outer edge of one flare to outer")
    print("              edge of the other flare (includes flares + DWS between).")
    print("  Corner    = NOT in a depressed span, AND the curb neighborhood is")
    print("              curved (~5 ft path/chord + heading turn).")
    print("  Straight  = NOT in a depressed span, AND essentially colinear.")
    print()

    by = {k: [] for k in STEEL}
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer not in STEEL or e.ObjectName != "AcDbLine":
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if math.hypot(cx - NW[0], cy - NW[1]) > 32:
                continue
            a = e.StartPoint
            b = e.EndPoint
            L = float(e.Length)
            by[e.Layer].append((L, float(a[0]), float(a[1]), float(b[0]), float(b[1])))
        except Exception:
            continue

    print("=== NW steel (r<=32) — what 'correct' looks like ===")
    for lyr, short in [
        ("Straight Steel Faced Concrete Curb (18'' Deep)", "Straight"),
        ("Depressed Steel Faced Concrete Curb (18'' Deep)", "Depressed"),
        ("Corner Steel Faced Concrete Curb (18'' Deep)", "Corner"),
    ]:
        segs = by[lyr]
        tot = sum(s[0] for s in segs)
        print(f"  {short:9s}  n={len(segs):3d}  LF={tot:7.2f}")
        # longest few
        for L, x0, y0, x1, y1 in sorted(segs, key=lambda t: -t[0])[:5]:
            print(f"    L={L:6.2f}  ({x0:.1f},{y0:.1f})->({x1:.1f},{y1:.1f})")


if __name__ == "__main__":
    main()
