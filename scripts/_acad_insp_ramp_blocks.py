"""Inspect PROP-D-ROAD-RAMPS anonymous block contents (flares + DWS)."""
from __future__ import annotations

import math
import pythoncom
from win32com.client import Dispatch

HANDLES = ["13CDD7", "13D60F", "13CDC0", "13D4BC", "13D72E", "133817"]


def xform_pt(br, x, y, z=0.0):
    """Rough transform: insertion + rotation + scale (2D)."""
    ins = br.InsertionPoint
    ang = float(br.Rotation)  # radians
    sx = float(br.XScaleFactor)
    sy = float(br.YScaleFactor)
    ca, sa = math.cos(ang), math.sin(ang)
    xr, yr = x * sx, y * sy
    return (
        float(ins[0]) + xr * ca - yr * sa,
        float(ins[1]) + xr * sa + yr * ca,
        float(ins[2]) + z,
    )


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument

    for h in HANDLES:
        try:
            br = doc.HandleToObject(h)
        except Exception as ex:
            print(h, "MISSING", ex)
            continue
        print(
            f"\n=== {h} block={br.Name!r} lyr={br.Layer} "
            f"ins=({float(br.InsertionPoint[0]):.2f},{float(br.InsertionPoint[1]):.2f}) "
            f"rot={math.degrees(float(br.Rotation)):.1f} "
            f"scale=({float(br.XScaleFactor):.2f},{float(br.YScaleFactor):.2f}) ==="
        )
        blk = doc.Blocks.Item(br.Name)
        for i in range(blk.Count):
            try:
                e = blk.Item(i)
                et = e.ObjectName
                lyr = e.Layer
                extra = ""
                if et == "AcDbHatch":
                    extra = f" pat={e.PatternName} area={getattr(e,'Area',None)}"
                if et == "AcDbSolid":
                    pts = []
                    for pi in range(4):
                        try:
                            p = e.Coordinate(pi)
                            wp = xform_pt(br, float(p[0]), float(p[1]), float(p[2]) if len(p) > 2 else 0)
                            pts.append(wp)
                        except Exception:
                            break
                    extra = f" pts_wcs={[(round(p[0],2),round(p[1],2)) for p in pts[:4]]}"
                if "Poly" in et or et == "AcDbLine":
                    try:
                        extra += f" len={e.Length:.2f}"
                    except Exception:
                        pass
                # bbox in block space then rough
                try:
                    mn, mx = e.GetBoundingBox()
                    # GetBoundingBox on nested may already be WCS for some - try both
                    extra += f" bb=[{float(mn[0]):.2f},{float(mn[1]):.2f}]-[{float(mx[0]):.2f},{float(mx[1]):.2f}]"
                except Exception:
                    pass
                print(f"  [{i}] {et} lyr={lyr!r} col={e.Color}{extra}")
            except Exception as ex:
                print(f"  [{i}] ERR {ex}")


if __name__ == "__main__":
    main()
