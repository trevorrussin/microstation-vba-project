"""Verify rebuilt flares sit on their island OUTER like NW."""
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
        f'(setq len (vlax-curve-getDistAtParam e (vlax-curve-getEndParam e))) '
        f'(setq f (open "{path}" "w")) (setq d 0.0) '
        f'(while (<= d (+ len 1e-6)) '
        f'(setq p (vlax-curve-getPointAtDist e (min d len))) '
        f'(if p (write-line (strcat (rtos (car p) 2 5) "," (rtos (cadr p) 2 5)) f)) '
        f'(setq d (+ d {step:.2f}))) (close f) (princ)) '
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


def dpoly(px, py, poly):
    best = 1e9
    for i in range(len(poly) - 1):
        ax, ay = poly[i]
        bx, by = poly[i + 1]
        dx, dy = bx - ax, by - ay
        L2 = dx * dx + dy * dy
        t = 0.0 if L2 < 1e-12 else max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / L2))
        best = min(best, math.hypot(px - (ax + t * dx), py - (ay + t * dy)))
    return best


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    try:
        doc.SendCommand(chr(27) + chr(27))
    except Exception:
        pass
    time.sleep(0.2)

    checks = [
        ("NW", "13D80A", 1028470, 187478),
        ("MidS", "13D7BC", 1028528, 187425),
        ("MidN", "13D2EA", 1028505, 187492),
        ("E", "13CEC7", 1028589, 187445),
    ]
    for name, oh, cx0, cy0 in checks:
        outer = sample_fl(doc, oh)
        print(f"=== {name} outer n={len(outer)} ===")
        for i in range(ms.Count):
            try:
                e = ms.Item(i)
                if e.ObjectName != "AcDbSolid" or e.Layer != "PROP-D-ROAD-RAMPS":
                    continue
                mn, mx = e.GetBoundingBox()
                cx = (float(mn[0]) + float(mx[0])) / 2
                cy = (float(mn[1]) + float(mx[1])) / 2
                if math.hypot(cx - cx0, cy - cy0) > 25:
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
                ds = [dpoly(p[0], p[1], outer) for p in u]
                on = sum(1 for d in ds if d < 0.25)
                print(f"  {e.Handle} d_outer={[[round(d, 3) for d in ds]]} on_outer={on}/3")
            except Exception:
                continue


if __name__ == "__main__":
    main()
