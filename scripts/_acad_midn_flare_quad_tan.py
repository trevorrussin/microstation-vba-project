"""MidN flares: outer base opposite tan PRO-S edge (correct AddSolid pairing).

AddSolid(p1,p2,p3,p4) = edge p1-p2 opposite edge p3-p4.
Blue-sketch pocket: short edge on outer, long edge along tan past green arrow.
"""
from __future__ import annotations

import math
import os
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

PTS = r"c:\repos\microstation-vba-project\scripts\_fl_curve_pts.txt"
LAYER_FLARE = "PROP-D-ROAD-RAMPS"
OUTER = "13D2EA"
TRAIL = "13C0E2"
PRO_H = "13A7EA"
MIDN = (1028501.0, 187485.0)


def vpt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def dist(a, b):
    return math.hypot(b[0] - a[0], b[1] - a[1])


def sample_fl(doc, handle, step=0.25):
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
        f'(if p (write-line (strcat (rtos (car p) 2 6) "," (rtos (cadr p) 2 6)) f)) '
        f'(setq d (+ d {step:.3f}))) (close f) (princ)) '
    )
    doc.SendCommand(lisp)
    for _ in range(50):
        time.sleep(0.15)
        if os.path.exists(PTS) and os.path.getsize(PTS) > 10:
            break
    pts = []
    with open(PTS) as f:
        for line in f:
            a = line.strip().split(",")
            if len(a) >= 2:
                pts.append((float(a[0]), float(a[1])))
    out = []
    for p in pts:
        if not out or dist(out[-1], p) > 1e-4:
            out.append(p)
    return out


def stations(pts):
    s = [0.0]
    for i in range(1, len(pts)):
        s.append(s[-1] + dist(pts[i - 1], pts[i]))
    return s


def project_station(pts, svals, x, y):
    best_d, best_s = 1e9, 0.0
    for i in range(len(pts) - 1):
        ax, ay = pts[i]
        bx, by = pts[i + 1]
        dx, dy = bx - ax, by - ay
        L2 = dx * dx + dy * dy
        t = 0.0 if L2 < 1e-12 else max(0.0, min(1.0, ((x - ax) * dx + (y - ay) * dy) / L2))
        px, py = ax + t * dx, ay + t * dy
        d = math.hypot(x - px, y - py)
        if d < best_d:
            best_d = d
            best_s = svals[i] + t * (svals[i + 1] - svals[i])
    return best_s, best_d


def point_at(pts, svals, s_target):
    if s_target <= svals[0]:
        return pts[0]
    if s_target >= svals[-1]:
        return pts[-1]
    for i in range(len(pts) - 1):
        if svals[i] <= s_target <= svals[i + 1]:
            span = svals[i + 1] - svals[i]
            t = 0.0 if span < 1e-12 else (s_target - svals[i]) / span
            return (
                pts[i][0] + t * (pts[i + 1][0] - pts[i][0]),
                pts[i][1] + t * (pts[i + 1][1] - pts[i][1]),
            )
    return pts[-1]


def tangent_at(pts, svals, s):
    p = point_at(pts, svals, s)
    p2 = point_at(pts, svals, min(s + 0.2, svals[-1]))
    if dist(p, p2) < 1e-6:
        p2 = point_at(pts, svals, max(s - 0.2, 0))
    tx, ty = p2[0] - p[0], p2[1] - p[1]
    L = math.hypot(tx, ty) or 1.0
    return p, (tx / L, ty / L)


def delete_midn(doc):
    """Delete every PROP-D solid in MidN pocket (incl. leftover tip tris)."""
    ms = doc.ModelSpace
    to_del = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbSolid" or e.Layer != LAYER_FLARE:
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            # MidN pocket only — keep true NW (west of 1028488)
            if cx < 1028488:
                continue
            if math.hypot(cx - MIDN[0], cy - MIDN[1]) < 30:
                to_del.append(e.Handle)
        except Exception:
            continue
    for h in to_del:
        try:
            doc.HandleToObject(h).Delete()
        except Exception:
            pass
    return len(to_del)


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    try:
        doc.SendCommand(chr(27) + chr(27))
        time.sleep(0.2)
    except Exception:
        pass

    print("deleted", delete_midn(doc))

    # PRO-S 13A7EA: [3] north tip → [2] → [1] → [0] SE far
    epro = doc.HandleToObject(PRO_H)
    c = list(epro.Coordinates)
    pro = [(float(c[i]), float(c[i + 1])) for i in range(0, len(c), 2)]
    print("PRO verts", pro)

    outer = sample_fl(doc, OUTER)
    os_ = stations(outer)

    e = doc.HandleToObject(TRAIL)
    mn, mx = e.GetBoundingBox()
    dws = [
        (float(mn[0]), float(mn[1])),
        (float(mx[0]), float(mn[1])),
        (float(mx[0]), float(mx[1])),
        (float(mn[0]), float(mx[1])),
    ]
    hits = []
    for p in dws:
        s, d = project_station(outer, os_, p[0], p[1])
        if d < 12:
            hits.append(s)
    L = os_[-1]
    start = [s for s in hits if s <= 25]
    end = [s for s in hits if s >= L - 25]
    ends = [min(end), max(start)] if start and end else [min(hits), max(hits)]
    print("ends", ends)

    s_north = max(ends)
    s_south = min(ends)
    half = 1.4

    # North: blue-sketch triangle — tip at PRO[3] (35.01), left on outer,
    # far right on tan past green arrow (PRO[1]).
    Po, tang = tangent_at(outer, os_, s_north)
    A = (Po[0] - tang[0] * half, Po[1] - tang[1] * half)
    B = (Po[0] + tang[0] * half, Po[1] + tang[1] * half)
    tip = pro[3]
    far = pro[1]  # past green arrow along tan
    print(
        f"north TRI tip=({tip[0]:.2f},{tip[1]:.2f}) "
        f"outer=({A[0]:.2f},{A[1]:.2f}) tanFar=({far[0]:.2f},{far[1]:.2f})"
    )
    # Triangle: tip + outer point + far tan (p4=p3)
    sol = ms.AddSolid(vpt(*tip), vpt(*A), vpt(*far), vpt(*far))
    sol.Layer = LAYER_FLARE
    sol.Color = 2

    # Also fill as quad outer A-B opposite full tan tip→far if triangle undershoots
    # Prefer one solid matching blue: tip-A-far. If A is too close to tip, use B.
    if dist(tip, A) < 0.5:
        sol.Delete()
        sol = ms.AddSolid(vpt(*tip), vpt(*B), vpt(*far), vpt(*far))
        sol.Layer = LAYER_FLARE
        sol.Color = 2
        print("north used outer B (A too close to tip)")

    # South: outer base opposite tan PRO[1]→PRO[0]
    Po2, tang2 = tangent_at(outer, os_, s_south)
    A2 = (Po2[0] - tang2[0] * half, Po2[1] - tang2[1] * half)
    B2 = (Po2[0] + tang2[0] * half, Po2[1] + tang2[1] * half)
    print(
        f"south QUAD outer=({A2[0]:.2f},{A2[1]:.2f})-({B2[0]:.2f},{B2[1]:.2f}) "
        f"tan=({pro[1][0]:.2f},{pro[1][1]:.2f})-({pro[0][0]:.2f},{pro[0][1]:.2f})"
    )
    # Correct pairing: edge A2-B2 opposite edge pro[1]-pro[0]
    sol2 = ms.AddSolid(vpt(*A2), vpt(*B2), vpt(*pro[1]), vpt(*pro[0]))
    sol2.Layer = LAYER_FLARE
    sol2.Color = 2

    doc.Regen(1)
    doc.SendCommand("._ZOOM _C 1028502,187483 16 ")
    print("DONE — MidN north tip→tan past green arrow; AddSolid edges fixed")


if __name__ == "__main__":
    main()
