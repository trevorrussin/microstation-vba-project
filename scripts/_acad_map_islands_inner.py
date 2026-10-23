"""Map loc01 curb islands: inner vs outer FLs, all ramps."""
from __future__ import annotations

import math
import pythoncom
from win32com.client import Dispatch


def xform(br, x, y):
    ins = br.InsertionPoint
    ang = float(br.Rotation)
    sx, sy = float(br.XScaleFactor), float(br.YScaleFactor)
    ca, sa = math.cos(ang), math.sin(ang)
    return float(ins[0]) + x * sx * ca - y * sy * sa, float(ins[1]) + x * sx * sa + y * sy * ca


def ramp_centers(doc):
    ms = doc.ModelSpace
    out = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbBlockReference":
                continue
            nm = e.Name or ""
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if not (1028400 < cx < 1028650 and 187380 < cy < 187560):
                continue
            # detect DWS / flare presence
            flares = 0
            dws = False
            try:
                blk = doc.Blocks.Item(e.Name)
                for j in range(blk.Count):
                    be = blk.Item(j)
                    if be.ObjectName == "AcDbSolid":
                        flares += 1
                    if be.ObjectName == "AcDbHatch" and "WARN" in (be.Layer or "").upper():
                        dws = True
                    if be.ObjectName == "AcDbBlockReference":
                        # nested
                        pass
            except Exception:
                pass
            out.append(
                {
                    "handle": e.Handle,
                    "name": nm,
                    "cx": cx,
                    "cy": cy,
                    "flares": flares,
                    "dws": dws,
                }
            )
        except Exception:
            continue
    return out


def fl_list(doc):
    ms = doc.ModelSpace
    out = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer != "C-TOPO-FEAT":
                continue
            if e.ObjectName not in (
                "AeccDbFeatureLine",
                "AcDb3dPolyline",
                "AcDbPolyline",
            ):
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if not (1028400 < cx < 1028650 and 187380 < cy < 187560):
                continue
            w = float(mx[0]) - float(mn[0])
            h = float(mx[1]) - float(mn[1])
            out.append(
                {
                    "handle": e.Handle,
                    "type": e.ObjectName,
                    "cx": cx,
                    "cy": cy,
                    "w": w,
                    "h": h,
                    "diag": math.hypot(w, h),
                }
            )
        except Exception:
            continue
    return out


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    print("count", doc.ModelSpace.Count)

    ramps = ramp_centers(doc)
    print(f"\n=== RAMPS n={len(ramps)} ===")
    for r in sorted(ramps, key=lambda t: (t["cy"], t["cx"]), reverse=True):
        print(
            f"  {r['handle']} {r['name']:16s} c=({r['cx']:.1f},{r['cy']:.1f}) "
            f"flares={r['flares']} dws={r['dws']}"
        )

    fls = fl_list(doc)
    print(f"\n=== C-TOPO-FEAT n={len(fls)} ===")
    for f in sorted(fls, key=lambda t: (-t["diag"], t["cy"])):
        print(
            f"  {f['handle']} {f['type']:22s} c=({f['cx']:.1f},{f['cy']:.1f}) "
            f"wh=({f['w']:.1f},{f['h']:.1f}) diag={f['diag']:.1f}"
        )

    # Cluster ramps into islands (distance < 35 ft)
    unused = list(ramps)
    islands = []
    while unused:
        seed = unused.pop(0)
        island = [seed]
        changed = True
        while changed:
            changed = False
            for r in list(unused):
                if any(math.hypot(r["cx"] - s["cx"], r["cy"] - s["cy"]) < 35 for s in island):
                    island.append(r)
                    unused.remove(r)
                    changed = True
        islands.append(island)

    print(f"\n=== ISLANDS n={len(islands)} ===")
    for ii, island in enumerate(islands):
        icx = sum(r["cx"] for r in island) / len(island)
        icy = sum(r["cy"] for r in island) / len(island)
        names = ", ".join(f"{r['name']}({r['handle']})" for r in island)
        print(f"\nIsland {ii+1} center=({icx:.1f},{icy:.1f}) ramps=[{names}]")
        # FLs near island, prefer those whose bbox covers island
        cands = []
        for f in fls:
            d = math.hypot(f["cx"] - icx, f["cy"] - icy)
            if d > 45:
                continue
            cands.append((d, f))
        cands.sort(key=lambda t: t[0])
        # Pair similar-sized parallel FLs: larger diag = outer loop often
        for d, f in cands[:8]:
            # distance from FL center to nearest ramp center (smaller => more inner toward ramps)
            rd = min(math.hypot(f["cx"] - r["cx"], f["cy"] - r["cy"]) for r in island)
            print(
                f"  FL {f['handle']} d_island={d:.1f} d_ramp={rd:.1f} "
                f"wh=({f['w']:.1f},{f['h']:.1f}) {f['type']}"
            )


if __name__ == "__main__":
    main()
