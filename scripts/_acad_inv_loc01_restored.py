"""Inventory Loc01 after reopen: outers, inners, PRO-S, DWS, flares."""
from __future__ import annotations

import math
import pythoncom
from win32com.client import Dispatch

OUTERS = {
    "13D80A": "NW",
    "13D25C": "N",
    "13D7DC": "SW",
    "13CEC7": "E",
    "13D2EA": "MidN",
    "13D7BC": "MidS",
}


def main():
    pythoncom.CoInitialize()
    doc = Dispatch("AutoCAD.Application").ActiveDocument
    ms = doc.ModelSpace
    print("file", doc.Name, "Saved", doc.Saved)

    for h, name in OUTERS.items():
        try:
            e = doc.HandleToObject(h)
            print(f"OUTER {name} {h} {e.ObjectName} layer={e.Layer}")
        except Exception:
            print(f"OUTER {name} {h} MISSING")

    print("--- C-TOPO-FEAT near Loc01 ---")
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer != "C-TOPO-FEAT":
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if not (1028440 < cx < 1028620 and 187390 < cy < 187530):
                continue
            tag = OUTERS.get(e.Handle, "inner?")
            print(f"  {e.Handle} {e.ObjectName} {tag} c=({cx:.1f},{cy:.1f})")
        except Exception:
            continue

    print("--- PRO-S-CURB near Loc01 ---")
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if "PRO-S-CURB" not in (e.Layer or ""):
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if not (1028440 < cx < 1028620 and 187390 < cy < 187530):
                continue
            print(f"  {e.Handle} {e.ObjectName} c=({cx:.1f},{cy:.1f})")
        except Exception:
            continue

    # trail / DWS / PED_RAMP
    print("--- DWS/PED/hatch-ish near Loc01 ---")
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            lyr = (e.Layer or "").upper()
            ot = e.ObjectName
            if not any(k in lyr for k in ("WARN", "PED", "RAMP", "DWS", "DETECT")):
                if "Hatch" not in ot and "Block" not in ot:
                    continue
                if "WARN" not in lyr and "RAMP" not in lyr and "PED" not in lyr:
                    continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if not (1028440 < cx < 1028620 and 187390 < cy < 187530):
                continue
            print(f"  {e.Handle} {ot} layer={e.Layer} c=({cx:.1f},{cy:.1f})")
        except Exception:
            continue

    n_sol = 0
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer == "PROP-D-ROAD-RAMPS" and e.ObjectName == "AcDbSolid":
                n_sol += 1
        except Exception:
            continue
    print("PROP solids", n_sol)

    for name, color in [
        ("Straight Steel Faced Concrete Curb (18'' Deep)", 1),
        ("Depressed Steel Faced Concrete Curb (18'' Deep)", 3),
        ("Corner Steel Faced Concrete Curb (18'' Deep)", 5),
    ]:
        try:
            lyr = doc.Layers.Item(name)
            print(f"layer OK {name!r} color={lyr.Color}")
        except Exception:
            print(f"layer MISSING {name!r}")


if __name__ == "__main__":
    main()
