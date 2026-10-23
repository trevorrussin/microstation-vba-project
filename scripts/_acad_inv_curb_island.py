"""Inventory curb island: C_TOP_FEAT_LINE + yellow flares near current view."""
from __future__ import annotations

import math
import sys

import pythoncom
from win32com.client import Dispatch, dynamic


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    print("file", doc.FullName)
    print("activeLayer", doc.ActiveLayer.Name)

    # current view
    try:
        vp = doc.ActiveViewport  # may fail in model
    except Exception:
        pass
    # use viewctr sysvars
    vc = doc.GetVariable("VIEWCTR")
    vh = float(doc.GetVariable("VIEWSIZE"))
    print(f"VIEWCTR=({vc[0]:.3f},{vc[1]:.3f}) VIEWSIZE={vh:.3f}")
    cx, cy = float(vc[0]), float(vc[1])
    half = vh * 1.2

    # Scan model space for candidates near view
    curb_lines = []
    yellow_tri = []
    hatches = []
    layers_seen = set()

    for i in range(ms.Count):
        try:
            e = ms.Item(i)
        except Exception:
            continue
        try:
            layer = e.Layer
            layers_seen.add(layer)
            et = e.ObjectName
            # bbox
            try:
                mn, mx = e.GetBoundingBox()
                bx = (float(mn[0]) + float(mx[0])) / 2
                by = (float(mn[1]) + float(mx[1])) / 2
                if abs(bx - cx) > half * 2 or abs(by - cy) > half * 2:
                    continue
            except Exception:
                continue

            if layer.upper() == "C_TOP_FEAT_LINE" or "TOP_FEAT" in layer.upper():
                curb_lines.append(
                    {
                        "handle": e.Handle,
                        "type": et,
                        "layer": layer,
                        "mn": (float(mn[0]), float(mn[1])),
                        "mx": (float(mx[0]), float(mx[1])),
                        "cx": bx,
                        "cy": by,
                        "w": float(mx[0]) - float(mn[0]),
                        "h": float(mx[1]) - float(mn[1]),
                        "closed": getattr(e, "Closed", None),
                        "color": e.Color,
                        "len": getattr(e, "Length", None),
                    }
                )
            # yellow triangles - often solids or hatches color 2
            col = e.Color
            if et in ("AcDbSolid", "AcDbHatch", "AcDbPolyline", "AcDb2dPolyline"):
                if col == 2 or (hasattr(e, "TrueColor") and False):
                    yellow_tri.append(
                        {
                            "handle": e.Handle,
                            "type": et,
                            "layer": layer,
                            "color": col,
                            "mn": (float(mn[0]), float(mn[1])),
                            "mx": (float(mx[0]), float(mx[1])),
                            "cx": bx,
                            "cy": by,
                        }
                    )
            if et == "AcDbHatch":
                hatches.append(
                    {
                        "handle": e.Handle,
                        "layer": layer,
                        "pattern": getattr(e, "PatternName", "?"),
                        "mn": (float(mn[0]), float(mn[1])),
                        "mx": (float(mx[0]), float(mx[1])),
                        "cx": bx,
                        "cy": by,
                    }
                )
        except Exception:
            continue

    print(f"\nC_TOP_FEAT / TOP_FEAT near view: {len(curb_lines)}")
    for c in sorted(curb_lines, key=lambda x: x["w"] * x["h"]):
        print(
            f"  {c['handle']} {c['type']} layer={c['layer']} "
            f"w={c['w']:.2f} h={c['h']:.2f} closed={c['closed']} "
            f"len={c['len']} bbox=[{c['mn'][0]:.2f},{c['mn'][1]:.2f}]-"
            f"[{c['mx'][0]:.2f},{c['mx'][1]:.2f}]"
        )

    print(f"\nyellow-ish color2 near view: {len(yellow_tri)}")
    for y in yellow_tri[:30]:
        print(
            f"  {y['handle']} {y['type']} layer={y['layer']} "
            f"cx={y['cx']:.2f} cy={y['cy']:.2f} "
            f"[{y['mn'][0]:.2f},{y['mn'][1]:.2f}]-[{y['mx'][0]:.2f},{y['mx'][1]:.2f}]"
        )

    print(f"\nhatches near view: {len(hatches)}")
    for h in hatches[:20]:
        print(
            f"  {h['handle']} layer={h['layer']} pat={h['pattern']} "
            f"cx={h['cx']:.2f} cy={h['cy']:.2f}"
        )

    # related layer names
    curb_layers = [L for L in layers_seen if any(
        k in L.upper() for k in ("CURB", "STEEL", "FEAT", "RAMP", "DEPRESS", "FLARE")
    )]
    print(f"\nrelated layers ({len(curb_layers)}):")
    for L in sorted(curb_layers):
        print(" ", L)


if __name__ == "__main__":
    main()
