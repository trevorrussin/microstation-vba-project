"""Measure NW reference flares vs outer/inner/PRO-S lines; compare auto flares."""
from __future__ import annotations

import math
import os
import time
import pythoncom
from win32com.client import Dispatch

PTS = r"c:\repos\microstation-vba-project\scripts\_fl_curve_pts.txt"


def sample_fl(doc, handle, step=0.4):
    if os.path.exists(PTS):
        try:
            os.remove(PTS)
        except Exception:
            pass
    path = PTS.replace("\\", "/")
    lisp = (
        f'(progn (vl-load-com) (setq e (handent "{handle}")) '
        f'(if (null e) (progn (princ "NO") (princ)) (progn '
        f'(setq len (vlax-curve-getDistAtParam e (vlax-curve-getEndParam e))) '
        f'(setq f (open "{path}" "w")) (setq d 0.0) '
        f'(while (<= d (+ len 1e-6)) '
        f'(setq p (vlax-curve-getPointAtDist e (min d len))) '
        f'(if p (write-line (strcat (rtos (car p) 2 5) "," (rtos (cadr p) 2 5)) f)) '
        f'(setq d (+ d {step:.2f}))) (close f) (princ "OK"))) (princ)) '
    )
    doc.SendCommand(lisp)
    for _ in range(40):
        time.sleep(0.15)
        if os.path.exists(PTS) and os.path.getsize(PTS) > 10:
            break
    pts = []
    with open(PTS) as f:
        for line in f:
            a = line.strip().split(",")
            if len(a) >= 2:
                pts.append((float(a[0]), float(a[1])))
    return pts


def d_poly(px, py, poly):
    best = 1e9
    for i in range(len(poly) - 1):
        ax, ay = poly[i]
        bx, by = poly[i + 1]
        dx, dy = bx - ax, by - ay
        L2 = dx * dx + dy * dy
        t = 0.0 if L2 < 1e-12 else max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / L2))
        best = min(best, math.hypot(px - (ax + t * dx), py - (ay + t * dy)))
    return best


def solid_tris(doc):
    ms = doc.ModelSpace
    out = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbSolid" or e.Layer != "PROP-D-ROAD-RAMPS":
                continue
            pts = []
            for pi in range(4):
                try:
                    p = e.Coordinate(pi)
                    pts.append((float(p[0]), float(p[1])))
                except Exception:
                    break
            u = []
            for p in pts:
                if not any(math.hypot(p[0] - q[0], p[1] - q[1]) < 1e-4 for q in u):
                    u.append(p)
            if len(u) >= 3:
                cx = sum(p[0] for p in u) / len(u)
                cy = sum(p[1] for p in u) / len(u)
                out.append((e.Handle, cx, cy, u))
        except Exception:
            continue
    return out


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    try:
        doc.SendCommand(chr(27) + chr(27))
    except Exception:
        pass
    time.sleep(0.3)

    outer_nw = sample_fl(doc, "13D80A")
    print("outer NW", len(outer_nw))

    # PRO-S-CURB-LINE entities near NW
    ms = doc.ModelSpace
    pros = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if "PRO-S-CURB" not in (e.Layer or "").upper() and e.Layer != "PRO-S-CURB-LINE":
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if math.hypot(cx - 1028470, cy - 187478) > 40:
                continue
            pros.append((e.Handle, e.ObjectName, e.Layer, cx, cy))
        except Exception:
            continue
    print("PRO-S near NW", pros[:10])

    # sample nearest PRO curb if feature line / poly
    pro_pts = []
    for h, ot, lyr, cx, cy in pros[:5]:
        try:
            e = doc.HandleToObject(h)
            if "Poly" in ot or "Feature" in ot or "Line" in ot:
                pts = sample_fl(doc, h, step=0.5)
                if pts:
                    pro_pts = pts
                    print("using PRO sample", h, len(pts))
                    break
        except Exception:
            continue

    # classified depressed near NW as inner proxy
    inner = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if not e.Layer.startswith("Depressed"):
                continue
            if e.ObjectName != "AcDbLine":
                continue
            s, en = e.StartPoint, e.EndPoint
            mx = 0.5 * (float(s[0]) + float(en[0]))
            my = 0.5 * (float(s[1]) + float(en[1]))
            if math.hypot(mx - 1028470, my - 187478) < 30:
                inner.append((float(s[0]), float(s[1])))
                inner.append((float(en[0]), float(en[1])))
        except Exception:
            continue
    print("depressed pts NW", len(inner) // 2)

    tris = solid_tris(doc)
    print("total PROP solids", len(tris))
    for h, cx, cy, u in tris:
        region = "NW" if math.hypot(cx - 1028470, cy - 187478) < 25 else (
            "MidS" if math.hypot(cx - 1028528, cy - 187425) < 25 else (
                "MidN" if math.hypot(cx - 1028505, cy - 187492) < 25 else (
                    "E" if math.hypot(cx - 1028589, cy - 187445) < 25 else "other"
                )
            )
        )
        print(f"\n{h} {region} c=({cx:.2f},{cy:.2f})")
        edges = []
        for a, b in zip(u, u[1:] + u[:1]):
            edges.append(math.hypot(b[0] - a[0], b[1] - a[1]))
        print("  edges", [round(x, 3) for x in edges])
        for j, p in enumerate(u):
            do = d_poly(p[0], p[1], outer_nw) if outer_nw else None
            di = d_poly(p[0], p[1], inner) if inner else None
            dp = d_poly(p[0], p[1], pro_pts) if pro_pts else None
            print(f"  v{j} ({p[0]:.3f},{p[1]:.3f}) d_outer={do} d_inner={di} d_pro={dp}")


if __name__ == "__main__":
    main()
