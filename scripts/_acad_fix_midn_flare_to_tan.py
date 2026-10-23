"""Fix MidN flare to fill blue-sketch pocket: tip on PRO-S tan curve past green arrow."""
from __future__ import annotations

import math
import os
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

PTS = r"c:\repos\microstation-vba-project\scripts\_fl_curve_pts.txt"
LAYER_FLARE = "PROP-D-ROAD-RAMPS"
# MidN only — the zoomed island in the screenshot
OUTER = "13D2EA"
TRAIL = "13C0E2"
PRO_HANDLES = ["13A7EA", "13C903"]  # the tan PRO-S curb around MidN
DWS_CENTER = (1028501.1, 187482.6)
# delete all auto flares near MidN (keep NW < 30 from 1028470,187478)
MIDN_CENTER = (1028501.0, 187485.0)


def vpt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def dist(a, b):
    return math.hypot(b[0] - a[0], b[1] - a[1])


def poly_coords(e):
    if e.ObjectName == "AcDbArc":
        c = e.Center
        cx, cy = float(c[0]), float(c[1])
        r = float(e.Radius)
        a0, a1 = float(e.StartAngle), float(e.EndAngle)
        if a1 < a0:
            a1 += 2 * math.pi
        n = max(16, int(r * (a1 - a0) / 0.2) + 1)
        return [
            (
                cx + r * math.cos(a0 + (a1 - a0) * i / n),
                cy + r * math.sin(a0 + (a1 - a0) * i / n),
            )
            for i in range(n + 1)
        ]
    if "Poly" in e.ObjectName:
        c = list(e.Coordinates)
        return [(float(c[i]), float(c[i + 1])) for i in range(0, len(c), 2)]
    if e.ObjectName == "AcDbLine":
        s, en = e.StartPoint, e.EndPoint
        return [(float(s[0]), float(s[1])), (float(en[0]), float(en[1]))]
    return []


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


def closest_on_poly(px, py, poly):
    best_d, best_pt = 1e9, poly[0]
    for i in range(len(poly) - 1):
        ax, ay = poly[i]
        bx, by = poly[i + 1]
        dx, dy = bx - ax, by - ay
        L2 = dx * dx + dy * dy
        t = 0.0 if L2 < 1e-12 else max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / L2))
        qx, qy = ax + t * dx, ay + t * dy
        d = math.hypot(px - qx, py - qy)
        if d < best_d:
            best_d, best_pt = d, (qx, qy)
    return best_pt, best_d


def d_poly(px, py, poly):
    return closest_on_poly(px, py, poly)[1]


def tri_area(a, b, c):
    return abs((b[0] - a[0]) * (c[1] - a[1]) - (b[1] - a[1]) * (c[0] - a[0])) / 2.0


def delete_midn_flares(doc):
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
            if math.hypot(cx - 1028470, cy - 187478) < 28:
                continue  # NW
            if math.hypot(cx - MIDN_CENTER[0], cy - MIDN_CENTER[1]) < 25:
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

    print("deleted MidN flares", delete_midn_flares(doc))

    # PRO-S polylines (the tan curb the blue sketch follows)
    pro = []
    for h in PRO_HANDLES:
        e = doc.HandleToObject(h)
        pts = poly_coords(e)
        print(f"PRO {h} n={len(pts)}")
        for i, p in enumerate(pts):
            print(f"  [{i}] ({p[0]:.3f},{p[1]:.3f})")
        pro.extend(pts)

    outer = sample_fl(doc, OUTER)
    print("outer n", len(outer))
    os_ = stations(outer)

    # DWS ends on outer
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
    print("dws stations", hits)
    L = os_[-1]
    # wrap ends
    start = [s for s in hits if s <= 25]
    end = [s for s in hits if s >= L - 25]
    if start and end:
        ends = [min(end), max(start)]
    else:
        ends = [min(hits), max(hits)]
    print("flare ends", ends)

    # Filter PRO: must be meaningfully off the outer (the visible tan pocket edge)
    pro_off = [p for p in pro if d_poly(p[0], p[1], outer) >= 0.75]
    print("pro_off", len(pro_off), "of", len(pro))
    for p in pro_off:
        print(f"  pro_off ({p[0]:.3f},{p[1]:.3f}) d_outer={d_poly(p[0],p[1],outer):.3f}")

    half = 1.35  # slightly wider base like NW ~2.7
    for s_end in ends:
        Po, tang = tangent_at(outer, os_, s_end)
        A = (Po[0] - tang[0] * half, Po[1] - tang[1] * half)
        B = (Po[0] + tang[0] * half, Po[1] + tang[1] * half)

        # Tip = furthest-right point on PRO-S pocket edge (follows tan curve
        # past the green arrow, matching the hand-drawn blue triangle).
        best = None
        for q in pro_off:
            d = dist(Po, q)
            if d < 3.0 or d > 18.0:
                continue
            area = tri_area(A, B, q)
            if area < 2.0:
                continue
            # strongly prefer larger X (right side of pocket / tan curve)
            score = (q[0] - Po[0]) * 2.0 + area * 0.15
            if best is None or score > best[0]:
                best = (score, area, d, q)

        if best is None:
            # absolute farthest pro_off within 14ft with positive east component
            for q in pro_off:
                d = dist(Po, q)
                if d < 3 or d > 14:
                    continue
                if q[0] < Po[0] - 1:
                    continue
                if best is None or d > best[2]:
                    best = (0, 0, d, q)

        if best is None:
            print("NO TIP for", s_end)
            continue
        Pp = best[3]
        print(
            f"flare s={s_end:.2f} A=({A[0]:.2f},{A[1]:.2f}) B=({B[0]:.2f},{B[1]:.2f}) "
            f"tip=({Pp[0]:.2f},{Pp[1]:.2f}) area={best[1]:.2f} d={best[2]:.2f} "
            f"d_tip_pro={d_poly(Pp[0],Pp[1],pro):.4f}"
        )
        sol = ms.AddSolid(vpt(*A), vpt(*B), vpt(*Pp), vpt(*Pp))
        sol.Layer = LAYER_FLARE
        sol.Color = 2

    doc.Regen(1)
    print("DONE MidN")


if __name__ == "__main__":
    main()
