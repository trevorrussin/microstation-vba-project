"""Map location-01 curb sources and all ramp flare spans."""
from __future__ import annotations

import math
import pythoncom
from win32com.client import Dispatch

RAMP_BLOCKS = ["13CDD7", "13D60F", "13CDC0", "13D4BC", "13D72E"]  # *U24..*U29


def xform(br, x, y, z=0.0):
    ins = br.InsertionPoint
    ang = float(br.Rotation)
    sx, sy = float(br.XScaleFactor), float(br.YScaleFactor)
    ca, sa = math.cos(ang), math.sin(ang)
    xr, yr = x * sx, y * sy
    return (float(ins[0]) + xr * ca - yr * sa, float(ins[1]) + xr * sa + yr * ca, float(ins[2]) + z)


def solid_pts_wcs(br, solid):
    pts = []
    for i in range(4):
        try:
            p = solid.Coordinate(i)
            pts.append(xform(br, float(p[0]), float(p[1]), float(p[2]) if len(list(p)) > 2 else 0))
        except Exception:
            break
    # solids often duplicate p0 as p2
    uniq = []
    for p in pts:
        if not uniq or math.hypot(p[0] - uniq[-1][0], p[1] - uniq[-1][1]) > 1e-6:
            # also skip if equals first when closing
            if uniq and math.hypot(p[0] - uniq[0][0], p[1] - uniq[0][1]) < 1e-6:
                continue
            uniq.append(p)
    return uniq


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace

    print("=== RAMP FLARE SPANS ===")
    for h in RAMP_BLOCKS:
        br = doc.HandleToObject(h)
        blk = doc.Blocks.Item(br.Name)
        flares = []
        dws_bb = None
        for i in range(blk.Count):
            e = blk.Item(i)
            if e.ObjectName == "AcDbSolid":
                pts = solid_pts_wcs(br, e)
                xs = [p[0] for p in pts]
                ys = [p[1] for p in pts]
                flares.append(
                    {
                        "pts": pts,
                        "lx": min(xs),
                        "ly": min(ys),
                        "hx": max(xs),
                        "hy": max(ys),
                        "cx": sum(xs) / len(xs),
                        "cy": sum(ys) / len(ys),
                    }
                )
            if e.ObjectName == "AcDbHatch" and "WARN" in e.Layer.upper():
                # hatch bbox may be block-space; transform corners
                try:
                    mn, mx = e.GetBoundingBox()
                    # For hatch inside block, GetBoundingBox often returns WCS already when queried via... 
                    # safer: transform block-space if values look local
                    if abs(float(mn[0])) < 50 and abs(float(mn[1])) < 50:
                        c1 = xform(br, float(mn[0]), float(mn[1]))
                        c2 = xform(br, float(mx[0]), float(mx[1]))
                        c3 = xform(br, float(mn[0]), float(mx[1]))
                        c4 = xform(br, float(mx[0]), float(mn[1]))
                        xs = [c1[0], c2[0], c3[0], c4[0]]
                        ys = [c1[1], c2[1], c3[1], c4[1]]
                        dws_bb = (min(xs), min(ys), max(xs), max(ys), float(e.Area))
                    else:
                        dws_bb = (
                            float(mn[0]),
                            float(mn[1]),
                            float(mx[0]),
                            float(mx[1]),
                            float(e.Area),
                        )
                except Exception as ex:
                    dws_bb = ("err", str(ex))
        print(f"\n{h} {br.Name}")
        for i, f in enumerate(flares):
            print(
                f"  flare{i} cx={f['cx']:.2f} cy={f['cy']:.2f} "
                f"[{f['lx']:.2f},{f['ly']:.2f}]-[{f['hx']:.2f},{f['hy']:.2f}] pts={[(round(p[0],2),round(p[1],2)) for p in f['pts']]}"
            )
        print(f"  dws={dws_bb}")
        if len(flares) == 2:
            # depressed span AABB union of both flares
            lx = min(flares[0]["lx"], flares[1]["lx"])
            ly = min(flares[0]["ly"], flares[1]["ly"])
            hx = max(flares[0]["hx"], flares[1]["hx"])
            hy = max(flares[0]["hy"], flares[1]["hy"])
            print(f"  DEPRESSED_UNION_BB [{lx:.2f},{ly:.2f}]-[{hx:.2f},{hy:.2f}]")

    # Curb candidates near loc01
    print("\n=== C-TOPO-FEAT / PRO-S-CURB near loc01 ===")
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer not in ("C-TOPO-FEAT", "PRO-S-CURB-LINE", "S-CURB-LINE"):
                continue
            if e.ObjectName not in (
                "AcDb3dPolyline",
                "AcDbPolyline",
                "AeccDbFeatureLine",
                "AcDbLine",
                "AcDbArc",
            ):
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if not (1028440 < cx < 1028620 and 187400 < cy < 187530):
                continue
            w = float(mx[0]) - float(mn[0])
            h = float(mx[1]) - float(mn[1])
            ln = None
            try:
                ln = float(e.Length)
            except Exception:
                pass
            print(
                f"  {e.Handle} {e.ObjectName} lyr={e.Layer} "
                f"cx={cx:.1f} cy={cy:.1f} {w:.1f}x{h:.1f} len={ln}"
            )
        except Exception:
            continue

    # any other PROP-D-ROAD-RAMPS?
    print("\n=== all PROP-D-ROAD-RAMPS refs ===")
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer == "PROP-D-ROAD-RAMPS":
                mn, mx = e.GetBoundingBox()
                print(
                    f"  {e.Handle} {e.ObjectName} name={getattr(e,'Name',None)} "
                    f"[{float(mn[0]):.1f},{float(mn[1]):.1f}]-[{float(mx[0]):.1f},{float(mx[1]):.1f}]"
                )
        except Exception:
            continue


if __name__ == "__main__":
    main()
