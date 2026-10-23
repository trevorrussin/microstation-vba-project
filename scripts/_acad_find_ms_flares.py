"""Find PROP-D-ROAD-RAMPS solids near Mid/E for depressed measure."""
from __future__ import annotations

import math
import pythoncom
from win32com.client import GetObject

CENTERS = {
    "MidN": (1028501.0, 187485.0),
    "MidS": (1028525.0, 187428.0),
    "E": (1028589.0, 187445.0),
    "N": (1028564.0, 187503.0),
    "SW": (1028486.0, 187421.0),
    "NW": (1028470.0, 187478.0),
}


def main():
    pythoncom.CoInitialize()
    doc = GetObject(None, "AutoCAD.Application").ActiveDocument
    ms = doc.ModelSpace
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbSolid":
                continue
            if e.Layer != "PROP-D-ROAD-RAMPS":
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            best = None
            for name, c in CENTERS.items():
                d = math.hypot(cx - c[0], cy - c[1])
                if best is None or d < best[0]:
                    best = (d, name)
            if best and best[0] < 40:
                print(f"solid {e.Handle} near {best[1]} d={best[0]:.1f} c=({cx:.1f},{cy:.1f})")
        except Exception:
            continue


if __name__ == "__main__":
    main()
