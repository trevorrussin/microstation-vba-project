"""Inventory Loc01 curb sources + steel for classify work."""
from __future__ import annotations

import pythoncom
from win32com.client import GetObject

STEEL = [
    "Straight Steel Faced Concrete Curb (18'' Deep)",
    "Depressed Steel Faced Concrete Curb (18'' Deep)",
    "Corner Steel Faced Concrete Curb (18'' Deep)",
]
HANDLES = [
    ("13D986", "NW-3d"),
    ("13D822", "NW-in"),
    ("13D824", "NW-in"),
    ("13D263", "N-in"),
    ("13D265", "N-in"),
    ("13D25E", "N-in"),
    ("13D282", "N-in"),
    ("13D7DB", "SW-in"),
    ("13D7E0", "SW-in"),
    ("13D7E8", "SW-in"),
    ("13CED4", "E-in"),
    ("13CED2", "E-in"),
    ("13D425", "MidN-in"),
    ("13D42D", "MidN-in"),
    ("13D2EE", "MidN-in"),
    ("13D7C4", "MidS-in"),
    ("13D7CC", "MidS-in"),
    ("13D7BF", "MidS-in"),
    ("13D80A", "OUT-NW"),
    ("13D25C", "OUT-N"),
    ("13D7DC", "OUT-SW"),
    ("13CEC7", "OUT-E"),
    ("13D2EA", "OUT-MidN"),
    ("13D7BC", "OUT-MidS"),
]
RAMPS = ["13CDD7", "13D60F", "13CDC0", "13D4BC", "13D72E"]


def main():
    pythoncom.CoInitialize()
    doc = GetObject(None, "AutoCAD.Application").ActiveDocument
    print("file", doc.FullName)
    print("Saved", doc.Saved)
    for h, n in HANDLES:
        try:
            e = doc.HandleToObject(h)
            vis = getattr(e, "Visible", "?")
            print(f"OK {n:10s} {h} {e.ObjectName} lyr={e.Layer} Vis={vis}")
        except Exception:
            print(f"MISS {n:10s} {h}")
    for h in RAMPS:
        try:
            e = doc.HandleToObject(h)
            print(f"RAMP {h} {e.Name}")
        except Exception:
            print(f"RAMP MISS {h}")
    ms = doc.ModelSpace
    for lyr in STEEL:
        n = 0
        L = 0.0
        for i in range(ms.Count):
            try:
                e = ms.Item(i)
                if e.Layer != lyr or e.ObjectName != "AcDbLine":
                    continue
                mn, mx = e.GetBoundingBox()
                cx = (float(mn[0]) + float(mx[0])) / 2
                cy = (float(mn[1]) + float(mx[1])) / 2
                if not (1028440 < cx < 1028620 and 187390 < cy < 187530):
                    continue
                n += 1
                L += float(e.Length)
            except Exception:
                continue
        print(f"STEEL {n:4d} {L:8.2f} ft  {lyr}")


if __name__ == "__main__":
    main()
