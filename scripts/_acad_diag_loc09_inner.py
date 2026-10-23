"""Diagnose Loc09 inner vs outer; list candidate curb layers."""
from __future__ import annotations

import math
import pythoncom
from win32com.client import GetObject, VARIANT

PATH = r"c:\repos\microstation-vba-project\Screenshots\HWP24QC1_Location_09_SGD.dwg"
STEEL = [
    "Straight Steel Faced Concrete Curb (18'' Deep)",
    "Depressed Steel Faced Concrete Curb (18'' Deep)",
    "Corner Steel Faced Concrete Curb (18'' Deep)",
]


def ss_layer(doc, name, layer):
    try:
        doc.SelectionSets.Item(name).Delete()
    except Exception:
        pass
    ss = doc.SelectionSets.Add(name)
    ft = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [8])
    fd = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [layer])
    ss.Select(5, None, None, ft, fd)
    return ss


def main():
    pythoncom.CoInitialize()
    acad = GetObject(None, "AutoCAD.Application")
    doc = None
    for i in range(acad.Documents.Count):
        d = acad.Documents.Item(i)
        if "09" in d.Name and "SGD" in d.Name.upper():
            doc = d
            break
    if doc is None:
        doc = acad.Documents.Open(PATH, False)
    print("FILE", doc.Name)

    print("=== STEEL ===")
    for lyr in STEEL:
        ss = ss_layer(doc, "S", lyr)
        L = 0.0
        for i in range(ss.Count):
            try:
                L += float(ss.Item(i).Length)
            except Exception:
                pass
        print(f"  n={ss.Count:4d} L={L:8.2f}  {lyr!r}")
        ss.Delete()

    print("=== layer table (curb/topo/feat/prop/steel) ===")
    for i in range(doc.Layers.Count):
        n = doc.Layers.Item(i).Name
        u = n.upper()
        if any(k in u for k in ("STEEL", "CURB", "TOPO", "FEAT", "PROP", "STRAIGHT", "DEPRESS", "CORNER")):
            print(" ", repr(n))

    # Compare hidden 1544B (wrong outer we used) vs 15474 (likely true inner)
    print("=== handles 1544B vs 15474 vs 15482 ===")
    for h in ("1544B", "15474", "15482", "1548A", "15480"):
        try:
            e = doc.HandleToObject(h)
            mn, mx = e.GetBoundingBox()
            cx = (mn[0] + mx[0]) / 2
            cy = (mn[1] + mx[1]) / 2
            diag = math.hypot(mx[0] - mn[0], mx[1] - mn[1])
            print(
                f"  {h} Visible={e.Visible} Layer={e.Layer!r} "
                f"diag={diag:.1f} c=({cx:.1f},{cy:.1f})"
            )
        except Exception as ex:
            print(f"  {h} ERR {ex}")

    # Ramp centers at classified island
    print("=== distance of each FL to nearest ramp at this island ===")
    ramps = []
    ss = ss_layer(doc, "P", "PROP-D-ROAD-RAMPS")
    for i in range(ss.Count):
        e = ss.Item(i)
        if e.ObjectName != "AcDbBlockReference":
            continue
        mn, mx = e.GetBoundingBox()
        ramps.append(((mn[0] + mx[0]) / 2, (mn[1] + mx[1]) / 2, e.Name, e.Handle))
    ss.Delete()
    # only ramps near 1029250,187705
    local = [r for r in ramps if math.hypot(r[0] - 1029250, r[1] - 187705) < 40]
    print(" local ramps", [(r[2], r[3], r[0], r[1]) for r in local])
    for h in ("1544B", "15474", "15482"):
        e = doc.HandleToObject(h)
        mn, mx = e.GetBoundingBox()
        cx = (mn[0] + mx[0]) / 2
        cy = (mn[1] + mx[1]) / 2
        rd = min(math.hypot(cx - r[0], cy - r[1]) for r in local)
        print(f"  {h} d_ramp={rd:.2f} diag={math.hypot(mx[0]-mn[0], mx[1]-mn[1]):.1f}")


if __name__ == "__main__":
    main()
