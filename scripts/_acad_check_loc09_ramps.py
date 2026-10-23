"""Force-open Loc09; count PROP-D ramps; list hidden C-TOPO-FEAT; confirm no ramp deletes."""
from __future__ import annotations

import math
import pythoncom
from win32com.client import GetObject, VARIANT

PATH = r"c:\repos\microstation-vba-project\Screenshots\HWP24QC1_Location_09_SGD.dwg"


def main():
    pythoncom.CoInitialize()
    acad = GetObject(None, "AutoCAD.Application")
    acad.Visible = True
    doc = None
    for i in range(acad.Documents.Count):
        d = acad.Documents.Item(i)
        if d.Name.lower() == "hwp24qc1_location_09_sgd.dwg":
            doc = d
            break
    if doc is None:
        print("opening", PATH)
        doc = acad.Documents.Open(PATH, False)
    doc.Activate()
    print("ACTIVE", doc.Name, doc.FullName)

    ms = doc.ModelSpace
    ramps = []
    fls = []
    steel_n = 0
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            lyr = e.Layer or ""
            on = e.ObjectName or ""
        except Exception:
            continue
        if lyr == "PROP-D-ROAD-RAMPS" and on == "AcDbBlockReference":
            mn, mx = e.GetBoundingBox()
            ramps.append(
                {
                    "h": e.Handle,
                    "name": e.Name,
                    "vis": bool(e.Visible),
                    "cx": (mn[0] + mx[0]) / 2,
                    "cy": (mn[1] + mx[1]) / 2,
                }
            )
        if lyr == "C-TOPO-FEAT" and on == "AeccDbFeatureLine":
            mn, mx = e.GetBoundingBox()
            fls.append(
                {
                    "h": e.Handle,
                    "vis": bool(e.Visible),
                    "diag": math.hypot(mx[0] - mn[0], mx[1] - mn[1]),
                    "cx": (mn[0] + mx[0]) / 2,
                    "cy": (mn[1] + mx[1]) / 2,
                }
            )
        if "Steel Faced Concrete Curb" in lyr and on == "AcDbLine":
            steel_n += 1

    print(f"PROP-D-ROAD-RAMPS blocks: {len(ramps)}")
    for r in sorted(ramps, key=lambda t: (-t["cy"], t["cx"])):
        print(
            f"  {r['h']} {r['name']:8s} vis={r['vis']} c=({r['cx']:.1f},{r['cy']:.1f})"
        )
    hidden = [f for f in fls if not f["vis"]]
    print(f"C-TOPO-FEAT FLs: {len(fls)}  hidden={len(hidden)}  steel_lines={steel_n}")
    for f in hidden:
        print(f"  HIDDEN FL {f['h']} diag={f['diag']:.1f} c=({f['cx']:.1f},{f['cy']:.1f})")
    print(
        "POLICY: batch only Deletes AcDbLine on Straight/Depressed/Corner; "
        "INNER FLs are Visible=False only; PROP-D ramps are never targeted."
    )


if __name__ == "__main__":
    main()
