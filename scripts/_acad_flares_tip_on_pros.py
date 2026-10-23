"""Rebuild flares: base on OUTER, tip ON nearest PRO-S-CURB-LINE (ray hit).

Matches user sketch — yellow fills pocket to the diagonal tan PRO-S line.
Leaves NW reference flares alone.
"""
from __future__ import annotations

import math
import os
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

PTS = r"c:\repos\microstation-vba-project\scripts\_fl_curve_pts.txt"
LAYER_FLARE = "PROP-D-ROAD-RAMPS"
FLARE_BASE = 2.5

JOBS = [
    {
        "name": "MidS",
        "outer": "13D7BC",
        "center": (1028525.0, 187428.0),
        "trail": "138B88",
        "pro_handles": ["13C226", "138B86", "138B6B", "13C115", "13C116", "138B64"],
    },
    {
        "name": "MidN",
        "outer": "13D2EA",
        "center": (1028505.0, 187492.0),
        "trail": "13C0E2",
        "pro_handles": [],  # auto-find
    },
    {
        "name": "E",
        "outer": "13CEC7",
        "center": (1028589.0, 187445.0),
        "trail": "133817",
        "pro_handles": [],
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


def sample_entity_pts(doc, handle):
    """Sample any curve-like entity via vlax; fallback for arcs/polys."""
    try:
        e = doc.HandleToObject(handle)
    except Exception:
        return []
    # Arc: sample by angle
    if e.ObjectName == "AcDbArc":
        c = e.Center
        cx, cy = float(c[0]), float(c[1])
        r = float(e.Radius)
        a0, a1 = float(e.StartAngle), float(e.EndAngle)
        if a1 < a0:
            a1 += 2 * math.pi
        n = max(8, int(r * (a1 - a0) / 0.3) + 1)
        return [
            (
                cx + r * math.cos(a0 + (a1 - a0) * i / n),
                cy + r * math.sin(a0 + (a1 - a0) * i / n),
            )
            for i in range(n + 1)
        ]
    if e.ObjectName in ("AcDbLine",):
        s, en = e.StartPoint, e.EndPoint
        return [(float(s[0]), float(s[1])), (float(en[0]), float(en[1]))]
    if "Poly" in e.ObjectName:
        try:
            coords = list(e.Coordinates)
            step = 2
            pts = [(float(coords[i]), float(coords[i + 1])) for i in range(0, len(coords), step)]
            return pts
        except Exception:
            pass
    # feature line / others
    return sample_fl(doc, handle, step=0.35)


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


def ray_hit_poly(origin, direction, poly, max_dist=20.0):
    """First intersection of ray origin+t*dir (t>0) with polyline segments."""
    ox, oy = origin
    dx, dy = direction
    best_t = None
    best_pt = None
    for i in range(len(poly) - 1):
        ax, ay = poly[i]
        bx, by = poly[i + 1]
        # solve ox+t*dx = ax+u*(bx-ax), oy+t*dy = ay+u*(by-ay)
        ex, ey = bx - ax, by - ay
        det = dx * ey - dy * ex
        if abs(det) < 1e-12:
            continue
        t = ((ax - ox) * ey - (ay - oy) * ex) / det
        u = ((ax - ox) * dy - (ay - oy) * dx) / det
        if t > 1e-6 and 0.0 <= u <= 1.0 and t <= max_dist:
            if best_t is None or t < best_t:
                best_t = t
                best_pt = (ox + t * dx, oy + t * dy)
    return best_pt, best_t


def gather_pro_pts(doc, center, handles):
    pts = []
    used = set()
    for h in handles:
        try:
            p = sample_entity_pts(doc, h)
            if p:
                pts.extend(p)
                used.add(h)
        except Exception:
            continue
    # also auto-scan nearby PRO-S
    ms = doc.ModelSpace
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if "PRO-S-CURB" not in (e.Layer or "").upper():
                continue
            if e.Handle in used:
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if math.hypot(cx - center[0], cy - center[1]) > 35:
                continue
            p = sample_entity_pts(doc, e.Handle)
            if p:
                pts.extend(p)
        except Exception:
            continue
    # dedupe chain-ish
    out = []
    for p in pts:
        if not out or dist(out[-1], p) > 0.05:
            out.append(p)
    return out


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


def delete_auto_flares(doc, center, radius=25.0):
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
                continue  # keep NW
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


def place_flare(ms, outer, svals, s_end, center, pro_pts):
    p, tang = tangent_at(outer, svals, s_end)
    half = FLARE_BASE * 0.5
    a = (p[0] - tang[0] * half, p[1] - tang[1] * half)
    b = (p[0] + tang[0] * half, p[1] + tang[1] * half)
    # inland toward island center
    to_c = (center[0] - p[0], center[1] - p[1])
    L = math.hypot(*to_c) or 1.0
    nx, ny = to_c[0] / L, to_c[1] / L

    tip = None
    hit_how = None
    if pro_pts and len(pro_pts) >= 2:
        hit, t = ray_hit_poly(p, (nx, ny), pro_pts, max_dist=18.0)
        if hit is not None and t is not None and 1.0 <= t <= 14.0:
            tip = hit
            hit_how = f"ray t={t:.2f}"
        else:
            # closest PRO-S point that is inland of base (dot with nx,ny > 0)
            best = None
            for i in range(len(pro_pts)):
                q = pro_pts[i]
                if (q[0] - p[0]) * nx + (q[1] - p[1]) * ny < 0.5:
                    continue
                d = dist(p, q)
                if d < 1.0 or d > 16:
                    continue
                if best is None or d < best[0]:
                    best = (d, q)
            if best:
                tip = best[1]
                hit_how = f"closest inland d={best[0]:.2f}"
            else:
                # snap closest overall then pull
                q, d = closest_on_poly(p[0] + nx * 5, p[1] + ny * 5, pro_pts)
                tip = q
                hit_how = f"fallback closest d={d:.2f}"

    if tip is None:
        tip = (p[0] + nx * 5.0, p[1] + ny * 5.0)
        hit_how = "no-pro fallback 5ft"

    sol = ms.AddSolid(vpt(*a), vpt(*b), vpt(*tip), vpt(*tip))
    sol.Layer = LAYER_FLARE
    sol.Color = 2
    # verify tip on pro
    d_pro = closest_on_poly(tip[0], tip[1], pro_pts)[1] if pro_pts else None
    print(
        f"    flare s={s_end:.2f} tip=({tip[0]:.3f},{tip[1]:.3f}) "
        f"{hit_how} d_pro={d_pro}"
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
        print(f"deleted {n} old")
        outer = sample_fl(doc, job["outer"])
        print(f"outer n={len(outer)}")
        if len(outer) < 5:
            continue
        svals = stations(outer)
        pro = gather_pro_pts(doc, job["center"], job.get("pro_handles") or [])
        print(f"pro pts n={len(pro)}")
        try:
            dws = trail_bbox(doc, job["trail"])
        except Exception as ex:
            print("trail", ex)
            continue
        span = depressed_ends(outer, svals, dws, job["center"])
        if not span:
            print("no span")
            continue
        lo, hi, wrap = span
        print(f"span [{lo:.2f}..{hi:.2f}] wrap={wrap}")
        place_flare(ms, outer, svals, lo, job["center"], pro)
        place_flare(ms, outer, svals, hi, job["center"], pro)

    doc.Regen(1)
    print("\nDONE — tips snapped to PRO-S")


if __name__ == "__main__":
    main()
