"""Inventory all curb ramps (DWS + flares) and existing classified curb lines."""
from __future__ import annotations

import pythoncom
from win32com.client import Dispatch


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    print("file", doc.FullName)
    print("ms", ms.Count)

    warn, flares, classified, topo = [], [], [], []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            lyr = e.Layer
            u = lyr.upper()
            mn, mx = e.GetBoundingBox()
            info = {
                "h": e.Handle,
                "t": e.ObjectName,
                "lyr": lyr,
                "lx": float(mn[0]),
                "ly": float(mn[1]),
                "hx": float(mx[0]),
                "hy": float(mx[1]),
                "cx": (float(mn[0]) + float(mx[0])) / 2,
                "cy": (float(mn[1]) + float(mx[1])) / 2,
                "w": float(mx[0]) - float(mn[0]),
                "hgt": float(mx[1]) - float(mn[1]),
            }
            if "WARN" in u:
                warn.append(info)
            elif lyr == "PROP-D-ROAD-RAMPS" and e.ObjectName == "AcDbSolid":
                flares.append(info)
            elif "Steel Faced Concrete Curb" in lyr:
                classified.append(info)
            elif lyr == "C-TOPO-FEAT" and e.ObjectName in (
                "AcDb3dPolyline",
                "AeccDbFeatureLine",
                "AcDbPolyline",
            ):
                topo.append(info)
        except Exception:
            continue

    print(
        f"warn={len(warn)} flares={len(flares)} "
        f"classified={len(classified)} topo={len(topo)}"
    )
    print("\n--- WARN HATCHES ---")
    hatches = [w for w in warn if w["t"] == "AcDbHatch"]
    for w in sorted(hatches, key=lambda x: (x["cy"], x["cx"])):
        print(
            f"  {w['h']} cx={w['cx']:.1f} cy={w['cy']:.1f} "
            f"[{w['lx']:.1f},{w['ly']:.1f}]-[{w['hx']:.1f},{w['hy']:.1f}] "
            f"{w['w']:.1f}x{w['hgt']:.1f}"
        )
    print(f"\n--- FLARES ({len(flares)}) ---")
    for f in sorted(flares, key=lambda x: (x["cy"], x["cx"])):
        print(
            f"  {f['h']} cx={f['cx']:.1f} cy={f['cy']:.1f} "
            f"[{f['lx']:.1f},{f['ly']:.1f}]-[{f['hx']:.1f},{f['hy']:.1f}]"
        )
    print(f"\nclassified curb lines: {len(classified)}")
    print(f"C-TOPO-FEAT curb-ish: {len(topo)}")
    for t in sorted(topo, key=lambda x: (x["cy"], x["cx"]))[:40]:
        print(
            f"  {t['h']} {t['t']} cx={t['cx']:.1f} cy={t['cy']:.1f} "
            f"{t['w']:.1f}x{t['hgt']:.1f}"
        )


if __name__ == "__main__":
    main()
