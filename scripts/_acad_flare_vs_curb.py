"""Analyze flare vs outer/inner; tighten corner thresholds."""
from __future__ import annotations

import math
import os
import time
import pythoncom
from win32com.client import Dispatch

PTS = r"c:\repos\microstation-vba-project\scripts\_fl_curve_pts.txt"


def dist_pt_seg(px, py, ax, ay, bx, by):
    dx, dy = bx - ax, by - ay
    L2 = dx * dx + dy * dy
    t = 0.0 if L2 < 1e-12 else max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / L2))
    qx, qy = ax + t * dx, ay + t * dy
    return math.hypot(px - qx, py - qy)


def sample_fl(doc, handle, step=0.5):
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
        f'(if p (write-line (strcat (rtos (car p) 2 4) "," (rtos (cadr p) 2 4)) f)) '
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
            parts = line.strip().split(",")
            if len(parts) >= 2:
                pts.append((float(parts[0]), float(parts[1])))
    return pts


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    try:
        doc.SendCommand(chr(27) + chr(27))
    except Exception:
        pass
    time.sleep(0.3)

    flare_pts = [
        (1028562.006, 187502.329),
        (1028567.193, 187509.131),
        (1028568.876, 187505.679),
        (1028564.198, 187497.835),
        (1028571.067, 187501.185),
        (1028572.786, 187497.661),
    ]
    opts = sample_fl(doc, "13D25C")
    print("outer N samples", len(opts))

    steel = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if "Steel Faced" not in (e.Layer or ""):
                continue
            if e.ObjectName != "AcDbLine":
                continue
            s, en = e.StartPoint, e.EndPoint
            mx = 0.5 * (float(s[0]) + float(en[0]))
            my = 0.5 * (float(s[1]) + float(en[1]))
            if math.hypot(mx - 1028564, my - 187503) < 45:
                steel.append(
                    (
                        e.Layer.split()[0],
                        float(s[0]),
                        float(s[1]),
                        float(en[0]),
                        float(en[1]),
                    )
                )
        except Exception:
            continue
    print("steel near N", len(steel))

    for i, p in enumerate(flare_pts):
        best_o = 1e9
        for j in range(len(opts) - 1):
            d = dist_pt_seg(p[0], p[1], opts[j][0], opts[j][1], opts[j + 1][0], opts[j + 1][1])
            best_o = min(best_o, d)
        best_s, best_lyr = 1e9, None
        for lyr, ax, ay, bx, by in steel:
            d = dist_pt_seg(p[0], p[1], ax, ay, bx, by)
            if d < best_s:
                best_s, best_lyr = d, lyr
        print(
            f"flare_pt{i} d_outer={best_o:.3f} d_steel={best_s:.3f} ({best_lyr})"
        )

    # Corner aggressiveness: for MidS classified, count how many Corner segs have large R
    print("\n=== MidS Corner local radius sample ===")
    mid = []
    for lyr, ax, ay, bx, by in []:
        pass
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer.startswith("Corner") and e.ObjectName == "AcDbLine":
                s, en = e.StartPoint, e.EndPoint
                mx = 0.5 * (float(s[0]) + float(en[0]))
                my = 0.5 * (float(s[1]) + float(en[1]))
                if math.hypot(mx - 1028525, my - 187425) < 35:
                    mid.append((float(s[0]), float(s[1]), float(en[0]), float(en[1]), float(e.Length)))
        except Exception:
            continue
    print("corner segs MidS", len(mid), "LF", sum(m[4] for m in mid))


if __name__ == "__main__":
    main()
