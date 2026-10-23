"""Deeper inventory of curb-related entities near AutoCAD view."""
from __future__ import annotations

import pythoncom
from win32com.client import Dispatch

LAYER_HINTS = (
    "TOPO",
    "FEAT",
    "CURB",
    "RAMP",
    "WARN",
    "FLARE",
    "STEEL",
)


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    vc = doc.GetVariable("VIEWCTR")
    vh = float(doc.GetVariable("VIEWSIZE"))
    cx, cy = float(vc[0]), float(vc[1])
    half = max(vh * 2.5, 40)

    hits = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            layer = e.Layer
            if not any(k in layer.upper() for k in LAYER_HINTS):
                continue
            mn, mx = e.GetBoundingBox()
            bx = (float(mn[0]) + float(mx[0])) / 2
            by = (float(mn[1]) + float(mx[1])) / 2
            if abs(bx - cx) > half or abs(by - cy) > half:
                continue
            w = float(mx[0]) - float(mn[0])
            h = float(mx[1]) - float(mn[1])
            et = e.ObjectName
            info = {
                "handle": e.Handle,
                "type": et,
                "layer": layer,
                "color": e.Color,
                "w": w,
                "h": h,
                "cx": bx,
                "cy": by,
                "mn": (float(mn[0]), float(mn[1])),
                "mx": (float(mx[0]), float(mx[1])),
            }
            if "Poly" in et or et == "AcDbLine" or et == "AcDbArc":
                try:
                    info["len"] = float(e.Length) if hasattr(e, "Length") else None
                except Exception:
                    info["len"] = None
                try:
                    info["closed"] = bool(e.Closed)
                except Exception:
                    info["closed"] = None
                try:
                    n = int(e.Coordinates.__len__() // 2) if hasattr(e, "Coordinates") else None
                    info["nverts"] = n
                except Exception:
                    info["nverts"] = None
            if et == "AcDbHatch":
                info["pattern"] = e.PatternName
            if et == "AcDbSolid":
                try:
                    pts = []
                    for pi in range(4):
                        p = e.Coordinate(pi)
                        pts.append((float(p[0]), float(p[1])))
                    info["pts"] = pts
                except Exception:
                    pass
            hits.append(info)
        except Exception:
            continue

    print(f"hits near view: {len(hits)}")
    # group by layer
    by_layer = {}
    for h in hits:
        by_layer.setdefault(h["layer"], []).append(h)
    for layer, items in sorted(by_layer.items()):
        print(f"\n=== {layer} ({len(items)}) ===")
        for it in sorted(items, key=lambda x: (x["type"], -x["w"])):
            extra = ""
            if "len" in it and it["len"] is not None:
                extra += f" len={it['len']:.2f}"
            if "nverts" in it and it["nverts"] is not None:
                extra += f" nverts={it['nverts']}"
            if "closed" in it and it["closed"] is not None:
                extra += f" closed={it['closed']}"
            if "pattern" in it:
                extra += f" pat={it['pattern']}"
            print(
                f"  {it['handle']} {it['type']} col={it['color']} "
                f"w={it['w']:.2f} h={it['h']:.2f} "
                f"cx={it['cx']:.2f} cy={it['cy']:.2f}{extra}"
            )


if __name__ == "__main__":
    main()
