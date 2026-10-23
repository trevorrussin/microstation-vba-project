"""Loc10: inventory ped-ramp / curb-drop geometry for Depressed measure."""
from __future__ import annotations

import pythoncom
from win32com.client import GetObject, VARIANT

LAYERS = ("S-CURB-DROP", "D-ROAD-MISC", "S-ROAD-MARKINGS", "0")
BLOCK_HINTS = ("PEDRAMP", "PED RAMP", "RAMP", "DROP", "DWS", "FLARE")


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


def describe_block(doc, br, limit_ents=12):
    name = br.Name
    blk = doc.Blocks.Item(name)
    types = {}
    layers = {}
    solids = []
    for j in range(blk.Count):
        try:
            be = blk.Item(j)
            on = be.ObjectName
            types[on] = types.get(on, 0) + 1
            layers[be.Layer or "?"] = layers.get(be.Layer or "?", 0) + 1
            if on == "AcDbSolid" and len(solids) < 4:
                pts = []
                for pi in range(4):
                    try:
                        p = be.Coordinate(pi)
                        pts.append((float(p[0]), float(p[1])))
                    except Exception:
                        break
                solids.append(pts)
        except Exception:
            continue
    return name, types, layers, solids, blk.Count


def main():
    pythoncom.CoInitialize()
    doc = GetObject(None, "AutoCAD.Application").ActiveDocument
    print("file", doc.Name)

    for lyr in LAYERS:
        try:
            ss = ss_layer(doc, f"L10_{lyr[:8]}", lyr)
        except Exception as ex:
            print(f"SS {lyr} FAIL {ex}")
            continue
        print(f"\n=== layer {lyr} count={ss.Count} ===")
        shown = 0
        for i in range(ss.Count):
            try:
                e = ss.Item(i)
                if e.ObjectName != "AcDbBlockReference":
                    continue
                name = e.Name or ""
                u = name.upper()
                interesting = any(h in u for h in BLOCK_HINTS) or lyr == "S-CURB-DROP"
                if not interesting:
                    continue
                mn, mx = e.GetBoundingBox()
                cx = (float(mn[0]) + float(mx[0])) / 2
                cy = (float(mn[1]) + float(mx[1])) / 2
                w = float(mx[0]) - float(mn[0])
                h = float(mx[1]) - float(mn[1])
                bname, types, layers, solids, nent = describe_block(doc, e)
                print(
                    f"  {e.Handle} blk={bname} xy=({cx:.1f},{cy:.1f}) "
                    f"box={w:.1f}x{h:.1f} ents={nent} solids={len(solids)}"
                )
                print(f"    types={types}")
                if shown < 3:
                    print(f"    layers={dict(list(layers.items())[:10])}")
                shown += 1
                if shown >= 12:
                    break
            except Exception as ex:
                print("  item fail", ex)
        ss.Delete()

    # All inserts whose block name looks like ped ramp (any layer)
    print("\n=== any-layer PEDRAMP inserts ===")
    try:
        doc.SelectionSets.Item("L10_ALL").Delete()
    except Exception:
        pass
    # filter by block name is harder; scan Blocks then GetSelection? 
    # Use typed selection DXF 2 = block name
    for bname in ("PEDRAMP", "5-PEDRAMP_m2", "BORDER PED RAMP NEW"):
        try:
            doc.SelectionSets.Item("L10_BN").Delete()
        except Exception:
            pass
        ss = doc.SelectionSets.Add("L10_BN")
        ft = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0, 2])
        fd = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["INSERT", bname])
        try:
            ss.Select(5, None, None, ft, fd)
            print(f"  INSERT {bname}: {ss.Count}")
            for i in range(min(ss.Count, 8)):
                e = ss.Item(i)
                mn, mx = e.GetBoundingBox()
                cx = (float(mn[0]) + float(mx[0])) / 2
                cy = (float(mn[1]) + float(mx[1])) / 2
                print(f"    {e.Handle} lyr={e.Layer} xy=({cx:.1f},{cy:.1f})")
            ss.Delete()
        except Exception as ex:
            print(f"  INSERT {bname} FAIL {ex}")
            try:
                ss.Delete()
            except Exception:
                pass


if __name__ == "__main__":
    main()
