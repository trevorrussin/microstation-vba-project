"""Diagnose Loc10: why PROP-D-ROAD-RAMPS was 0; where flares/ramps live."""
from __future__ import annotations

import math
from collections import Counter, defaultdict

import pythoncom
from win32com.client import GetObject, VARIANT

RAMP_LAYER = "PROP-D-ROAD-RAMPS"
STEEL = [
    "Straight Steel Faced Concrete Curb (18'' Deep)",
    "Depressed Steel Faced Concrete Curb (18'' Deep)",
    "Corner Steel Faced Concrete Curb (18'' Deep)",
]


def main() -> None:
    pythoncom.CoInitialize()
    doc = GetObject(None, "AutoCAD.Application").ActiveDocument
    print("file", doc.FullName)
    ms = doc.ModelSpace

    # Layer catalog: blocks + names hinting ramp
    by_layer_blocks = Counter()
    rampish = []
    layer_names = set()
    n_solid_in_ms = 0
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            lyr = e.Layer or ""
            layer_names.add(lyr)
            on = e.ObjectName or ""
            if on == "AcDbBlockReference":
                by_layer_blocks[lyr] += 1
                name = ""
                try:
                    name = e.Name or ""
                except Exception:
                    pass
                u = name.upper()
                if (
                    "RAMP" in u
                    or "U2" in u
                    or lyr.upper().startswith("PROP-D")
                    or "ROAD-RAMP" in lyr.upper()
                    or "PED" in u
                ):
                    cx = cy = None
                    try:
                        mn, mx = e.GetBoundingBox()
                        cx = (float(mn[0]) + float(mx[0])) / 2
                        cy = (float(mn[1]) + float(mx[1])) / 2
                    except Exception:
                        pass
                    rampish.append((e.Handle, lyr, name, on, cx, cy))
            if on == "AcDbSolid":
                n_solid_in_ms += 1
        except Exception:
            continue

    print(f"\nBlock refs by layer (top):")
    for lyr, n in by_layer_blocks.most_common(25):
        print(f"  {n:4d}  {lyr}")

    print(f"\nRamp-ish block refs ({len(rampish)}):")
    for h, lyr, name, on, cx, cy in rampish[:40]:
        print(f"  {h} lyr={lyr} name={name} xy=({cx},{cy})")

    # Exact PROP-D-ROAD-RAMPS via SelectionSet + ModelSpace
    try:
        try:
            doc.SelectionSets.Item("LOC10_R").Delete()
        except Exception:
            pass
        ss = doc.SelectionSets.Add("LOC10_R")
        ft = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [8])
        fd = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [RAMP_LAYER])
        ss.Select(5, None, None, ft, fd)
        print(f"\nSS {RAMP_LAYER} Count={ss.Count}")
        for i in range(min(ss.Count, 15)):
            e = ss.Item(i)
            print(f"  {e.Handle} {e.ObjectName} name={getattr(e,'Name',None)}")
        ss.Delete()
    except Exception as ex:
        print("SS FAIL", ex)

    # Layers containing RAMP / PROP-D
    print("\nLayers matching RAMP/PROP-D/CURB:")
    for lyr in sorted(layer_names):
        u = lyr.upper()
        if any(k in u for k in ("RAMP", "PROP-D", "CURB", "FLARE", "DWS", "WARN")):
            print(" ", lyr)

    # Steel LF near any rampish
    print("\nSteel segs by layer:")
    steel = Counter()
    steel_lf = defaultdict(float)
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer not in STEEL or e.ObjectName != "AcDbLine":
                continue
            try:
                from win32com.client.dynamic import Dispatch

                L = float(Dispatch(e).Length)
            except Exception:
                continue
            steel[e.Layer] += 1
            steel_lf[e.Layer] += L
        except Exception:
            continue
    for lyr in STEEL:
        print(f"  {steel[lyr]:4d}  {steel_lf[lyr]:8.1f} ft  {lyr}")

    print(f"\nModelSpace solids={n_solid_in_ms}")


if __name__ == "__main__":
    main()
