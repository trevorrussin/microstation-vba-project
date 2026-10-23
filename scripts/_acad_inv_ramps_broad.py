"""Broader scan for DWS / flares / ramp geometry across the DWG."""
from __future__ import annotations

import pythoncom
from win32com.client import Dispatch


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace

    layers = {}
    hex_hatches = []
    solids = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            lyr = e.Layer
            layers[lyr] = layers.get(lyr, 0) + 1
            if e.ObjectName == "AcDbHatch":
                try:
                    pat = e.PatternName
                except Exception:
                    pat = "?"
                if pat and ("HEX" in pat.upper() or "WARN" in lyr.upper() or "DWS" in lyr.upper()):
                    mn, mx = e.GetBoundingBox()
                    hex_hatches.append(
                        (
                            e.Handle,
                            lyr,
                            pat,
                            (float(mn[0]) + float(mx[0])) / 2,
                            (float(mn[1]) + float(mx[1])) / 2,
                            float(mx[0]) - float(mn[0]),
                            float(mx[1]) - float(mn[1]),
                        )
                    )
            if e.ObjectName == "AcDbSolid":
                mn, mx = e.GetBoundingBox()
                w = float(mx[0]) - float(mn[0])
                h = float(mx[1]) - float(mn[1])
                if max(w, h) < 20 and min(w, h) > 0.5:
                    solids.append(
                        (
                            e.Handle,
                            lyr,
                            e.Color,
                            (float(mn[0]) + float(mx[0])) / 2,
                            (float(mn[1]) + float(mx[1])) / 2,
                            w,
                            h,
                        )
                    )
        except Exception:
            continue

    print("layers with WARN/RAMP/FLARE/CURB/HEX:")
    for L, n in sorted(layers.items()):
        u = L.upper()
        if any(k in u for k in ("WARN", "RAMP", "FLARE", "CURB", "DWS", "DETECT")):
            print(f"  {n:4d}  {L}")

    print(f"\nHEX-ish hatches: {len(hex_hatches)}")
    for h in sorted(hex_hatches, key=lambda x: (x[4], x[3])):
        print(f"  {h[0]} lyr={h[1]} pat={h[2]} cx={h[3]:.1f} cy={h[4]:.1f} {h[5]:.1f}x{h[6]:.1f}")

    print(f"\nsmall solids: {len(solids)}")
    by_lyr = {}
    for s in solids:
        by_lyr.setdefault(s[1], []).append(s)
    for L, items in sorted(by_lyr.items()):
        print(f"  {L}: {len(items)}")
        for s in items[:8]:
            print(f"    {s[0]} col={s[2]} cx={s[3]:.1f} cy={s[4]:.1f} {s[5]:.1f}x{s[6]:.1f}")
        if len(items) > 8:
            print(f"    ... +{len(items)-8}")


if __name__ == "__main__":
    main()
