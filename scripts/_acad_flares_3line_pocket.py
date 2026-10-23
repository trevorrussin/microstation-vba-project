"""Place flares as 3-line pockets: OUTER + INNER + PRO-S (not on-outer).

NW reference kept. MidS/MidN/E rebuilt so the triangle is contained by:
  - C-TOPO outer
  - inner curb (offset from outer / steel face)
  - PRO-S curb that is offset from outer (the diagonal tan), not the
    PRO-S that coincides with the outer face.
"""
from __future__ import annotations

import math
import os
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

PTS = r"c:\repos\microstation-vba-project\scripts\_fl_curve_pts.txt"
LAYER_FLARE = "PROP-D-ROAD-RAMPS"
INNER_OFFSET = 2.0
PRO_OFF_OUTER_MIN = 1.25  # PRO-S pts must be this far from outer

JOBS = [
    {
        "name": "MidS",
        "outer": "13D7BC",
        "center": (1028522.3, 187429.3),  # DWS center
        "trail": "138B88",
        "pro_pref": ["13C116", "13C115", "138B64", "13C226", "138B86", "138B6B"],
    },
    {
        "name": "MidN",
        "outer": "13D2EA",
        "center": (1028501.0, 187485.0),
        "trail": "13C0E2",
        "pro_pref": [],
    },
    {
        "name": "E",
        "outer": "13CEC7",
        "center": (1028589.0, 187445.0),
        "trail": "133817",
        "pro_pref": [],
    },
]


def vpt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def dist(a, b):
    return math.hypot(b[0] - a[0], b[1] - a[1])


def sample_fl(doc, handle, step=0.3):
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
        f'(if p (write-line (strcat (rtos (car p) 2 6) "," (rtos (cadr p) 2 6)) f)) '
        f'(setq d (+ d {step:.3f}))) (close f) (princ "OK"))) (princ)) '
    )
    doc.SendCommand(lisp)
    for _ in range(50):
        time.sleep(0.18)
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


def entity_pts(doc, handle):
    e = doc.HandleToObject(handle)
    if e.ObjectName == "AcDbArc":
        c = e.Center
        cx, cy = float(c[0]), float(c[1])
        r = float(e.Radius)
        a0, a1 = float(e.StartAngle), float(e.EndAngle)
        if a1 < a0:
            a1 += 2 * math.pi
        n = max(12, int(r * (a1 - a0) / 0.25) + 1)
        return [
            (
                cx + r * math.cos(a0 + (a1 - a0) * i / n),
                cy + r * math.sin(a0 + (a1 - a0) * i / n),
            )
            for i in range(n + 1)
        ]
    if e.ObjectName == "AcDbLine":
        s, en = e.StartPoint, e.EndPoint
        return [(float(s[0]), float(s[1])), (float(en[0]), float(en[1]))]
    if "Poly" in e.ObjectName:
        coords = list(e.Coordinates)
        return [(float(coords[i]), float(coords[i + 1])) for i in range(0, len(coords), 2)]
    return sample_fl(doc, handle)


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
    p2 = point_at(pts, svals, min(s + 0.25, svals[-1]))
    if dist(p, p2) < 1e-6:
        p2 = point_at(pts, svals, max(s - 0.25, 0))
    tx, ty = p2[0] - p[0], p2[1] - p[1]
    L = math.hypot(tx, ty) or 1.0
    return p, (tx / L, ty / L)


def offset_toward(pts, center, offset):
    cx, cy = center
    out = []
    n = len(pts)
    for i in range(n):
        if i == 0:
            tx, ty = pts[1][0] - pts[0][0], pts[1][1] - pts[0][1]
        elif i == n - 1:
            tx, ty = pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]
        else:
            tx = pts[i + 1][0] - pts[i - 1][0]
            ty = pts[i + 1][1] - pts[i - 1][1]
        L = math.hypot(tx, ty)
        if L < 1e-9:
            out.append(pts[i])
            continue
        nx, ny = -ty / L, tx / L
        if nx * (cx - pts[i][0]) + ny * (cy - pts[i][1]) < 0:
            nx, ny = -nx, -ny
        out.append((pts[i][0] + offset * nx, pts[i][1] + offset * ny))
    return out


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
            best_d = d
            best_pt = (qx, qy)
    return best_pt, best_d


def d_to_poly(px, py, poly):
    return closest_on_poly(px, py, poly)[1]


def gather_pro(doc, center, pref):
    pts = []
    seen = set()
    for h in pref:
        try:
            pts.extend(entity_pts(doc, h))
            seen.add(h)
        except Exception:
            pass
    ms = doc.ModelSpace
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if "PRO-S-CURB" not in (e.Layer or "").upper():
                continue
            if e.Handle in seen:
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if math.hypot(cx - center[0], cy - center[1]) > 40:
                continue
            pts.extend(entity_pts(doc, e.Handle))
        except Exception:
            continue
    return pts


def filter_pro_off_outer(pro_pts, outer):
    """Keep only PRO-S geometry that is NOT sitting on the outer curb face."""
    if not pro_pts or len(outer) < 2:
        return []
    kept = []
    for p in pro_pts:
        if d_to_poly(p[0], p[1], outer) >= PRO_OFF_OUTER_MIN:
            kept.append(p)
    return kept


def trail_bbox(doc, handle):
    e = doc.HandleToObject(handle)
    mn, mx = e.GetBoundingBox()
    return [
        (float(mn[0]), float(mn[1])),
        (float(mx[0]), float(mn[1])),
        (float(mx[0]), float(mx[1])),
        (float(mn[0]), float(mx[1])),
    ]


def depressed_ends(pts, svals, dws, center):
    L = svals[-1]
    hits = []
    for p in dws:
        s, d = project_station(pts, svals, p[0], p[1])
        if d <= 10:
            hits.append(s)
    if len(hits) < 2:
        return None
    start = [s for s in hits if s <= 22]
    end = [s for s in hits if s >= L - 22]
    if start and end and (max(hits) - min(hits) > 25):
        return (min(end), max(start), True)
    lo, hi = min(hits), max(hits)
    if hi - lo > 35:
        cs, _ = project_station(pts, svals, center[0], center[1])
        near = [s for s in hits if abs(s - cs) < 20]
        if len(near) >= 2:
            lo, hi = min(near), max(near)
    if hi - lo < 0.5:
        return None
    return (lo, hi, False)


def delete_auto_flares(doc, center, radius=28.0):
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
            if math.hypot(cx - 1028470, cy - 187478) < 30:
                continue
            if math.hypot(cx - center[0], cy - center[1]) <= radius:
                to_del.append(e.Handle)
        except Exception:
            continue
    for h in to_del:
        try:
            doc.HandleToObject(h).Delete()
        except Exception:
            pass
    return len(to_del)


def place_flare_3line(ms, outer, outer_s, inner, s_end, dws_center, pro_off):
    """NW-style: short edge ON OUTER, tip on offset PRO-S (fills pocket to diagonal)."""
    Po, tang = tangent_at(outer, outer_s, s_end)
    half = 1.25  # ~NW short-edge half (full ~2.5)
    A = (Po[0] - tang[0] * half, Po[1] - tang[1] * half)
    B = (Po[0] + tang[0] * half, Po[1] + tang[1] * half)
    Pi, _ = closest_on_poly(Po[0], Po[1], inner)  # for inland hint only

    def tri_area(a, b, c):
        return abs(
            (b[0] - a[0]) * (c[1] - a[1]) - (b[1] - a[1]) * (c[0] - a[0])
        ) / 2.0

    inland = (Pi[0] - Po[0], Pi[1] - Po[1])
    inL = math.hypot(*inland) or 1.0
    inland = (inland[0] / inL, inland[1] / inL)

    best = None
    for q in pro_off:
        d = dist(Po, q)
        if d < 2.5 or d > 9.0:
            continue
        if abs(q[0] - Po[0]) > 9.0 or abs(q[1] - Po[1]) > 9.0:
            continue
        if dist(q, dws_center) < dist(Po, dws_center) - 1.0:
            continue
        # tip should be inland of outer (toward island / PRO-S pocket)
        if (q[0] - Po[0]) * inland[0] + (q[1] - Po[1]) * inland[1] < 0.5:
            continue
        area = tri_area(A, B, q)
        if area < 1.5:
            continue
        if best is None or area > best[0]:
            best = (area, d, q)

    if best is None:
        # fallback: farthest local pro_off inland
        for q in pro_off:
            d = dist(Po, q)
            if d < 2.0 or d > 10:
                continue
            if (q[0] - Po[0]) * inland[0] + (q[1] - Po[1]) * inland[1] < 0:
                continue
            if best is None or d > best[1]:
                best = (0, d, q)

    if best is None:
        Pp = (Po[0] + inland[0] * 5, Po[1] + inland[1] * 5)
        how = "fallback5"
    else:
        Pp = best[2]
        how = f"area={best[0]:.2f} d={best[1]:.2f}"

    sol = ms.AddSolid(vpt(*A), vpt(*B), vpt(*Pp), vpt(*Pp))
    sol.Layer = LAYER_FLARE
    sol.Color = 2
    print(
        f"    s={s_end:.2f} base A=({A[0]:.2f},{A[1]:.2f}) B=({B[0]:.2f},{B[1]:.2f}) "
        f"tip=({Pp[0]:.2f},{Pp[1]:.2f}) {how}"
    )
    return sol


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

    for job in JOBS:
        name = job["name"]
        print(f"\n=== {name} ===")
        n = delete_auto_flares(doc, job["center"])
        print(f"deleted {n}")
        outer = sample_fl(doc, job["outer"])
        print(f"outer n={len(outer)}")
        if len(outer) < 5:
            continue
        outer_s = stations(outer)
        inner = offset_toward(outer, job["center"], INNER_OFFSET)
        pro_all = gather_pro(doc, job["center"], job.get("pro_pref") or [])
        pro_off = filter_pro_off_outer(pro_all, outer)
        print(f"pro_all={len(pro_all)} pro_off_outer={len(pro_off)}")
        try:
            dws = trail_bbox(doc, job["trail"])
        except Exception as ex:
            print("trail", ex)
            continue
        # DWS center from bbox
        dws_c = (
            sum(p[0] for p in dws) / 4,
            sum(p[1] for p in dws) / 4,
        )
        span = depressed_ends(outer, outer_s, dws, dws_c)
        if not span:
            print("no span")
            continue
        lo, hi, wrap = span
        print(f"span [{lo:.2f}..{hi:.2f}] wrap={wrap} dws_c=({dws_c[0]:.1f},{dws_c[1]:.1f})")
        place_flare_3line(ms, outer, outer_s, inner, lo, dws_c, pro_off)
        place_flare_3line(ms, outer, outer_s, inner, hi, dws_c, pro_off)

    doc.Regen(1)
    print("\nDONE — 3-line flares (outer + inner + offset PRO-S)")


if __name__ == "__main__":
    main()
