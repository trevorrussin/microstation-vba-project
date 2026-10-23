"""Rebuild auto flares like NW reference: short edge ON outer topo, tip inland.

Deletes prior auto-flares at E / MidN / MidS and places new PROP-D-ROAD-RAMPS
solids matching NW geometry (base on outer, tip ~5 ft toward island).
Does not touch NW reference solids or OUTER feature lines.
"""
from __future__ import annotations

import math
import os
time = __import__("time")
import pythoncom
from win32com.client import Dispatch, VARIANT

PTS = r"c:\repos\microstation-vba-project\scripts\_fl_curve_pts.txt"
LAYER_FLARE = "PROP-D-ROAD-RAMPS"

# Islands that need flares rebuilt (NW already correct)
JOBS = [
    {"name": "E", "outer": "13CEC7", "center": (1028589.0, 187445.0), "trail": "133817"},
    {"name": "MidN", "outer": "13D2EA", "center": (1028505.0, 187492.0), "trail": "13C0E2"},
    {"name": "MidS", "outer": "13D7BC", "center": (1028528.0, 187425.0), "trail": "138B88"},
]

# NW reference metrics
FLARE_BASE = 2.5  # short edge along outer (NW ~2.2-2.7)
FLARE_DEPTH = 5.0  # tip inland from outer (NW tip d_outer ~4-6.5)


def vpt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def dist(a, b):
    return math.hypot(b[0] - a[0], b[1] - a[1])


def sample_fl(doc, handle, step=0.35):
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
        time.sleep(0.2)
        if os.path.exists(PTS) and os.path.getsize(PTS) > 10:
            break
    pts = []
    with open(PTS) as f:
        for line in f:
            a = line.strip().split(",")
            if len(a) >= 2:
                pts.append((float(a[0]), float(a[1])))
    # dedupe
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
    best_d, best_s, best_i, best_t = 1e9, 0.0, 0, 0.0
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
            best_i, best_t = i, t
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


def trail_bbox(doc, handle):
    e = doc.HandleToObject(handle)
    mn, mx = e.GetBoundingBox()
    return [
        (float(mn[0]), float(mn[1])),
        (float(mx[0]), float(mn[1])),
        (float(mx[0]), float(mx[1])),
        (float(mn[0]), float(mx[1])),
    ]


def depressed_ends_on_outer(pts, svals, dws_corners, center):
    """Project DWS/trail to outer; return (lo, hi, wrap) stations on outer."""
    L = svals[-1]
    hits = []
    for p in dws_corners:
        s, d = project_station(pts, svals, p[0], p[1])
        if d <= 8.0:  # trail hatch can sit a bit off outer
            hits.append(s)
    if len(hits) < 2:
        return None
    start = [s for s in hits if s <= 20]
    end = [s for s in hits if s >= L - 20]
    if start and end and any(s <= 20 for s in hits) and any(s >= L - 20 for s in hits):
        if max(hits) - min(hits) > 25:
            return (min(end), max(start), True)
    lo, hi = min(hits), max(hits)
    if hi - lo > 30:
        # use cluster nearest center projection
        cs, _ = project_station(pts, svals, center[0], center[1])
        near = [s for s in hits if abs(s - cs) < 18]
        if len(near) >= 2:
            lo, hi = min(near), max(near)
    if hi - lo < 0.5:
        return None
    return (lo, hi, False)


def delete_auto_flares(doc, center, radius=22.0):
    """Delete PROP solids near island that are NOT the NW originals (far from NW)."""
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
            # never delete NW reference flares
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


def place_flare_on_outer(ms, pts, svals, s_end, center):
    """NW-style: short edge on OUTER at s_end, tip inland toward island center."""
    p, tang = tangent_at(pts, svals, s_end)
    half = FLARE_BASE * 0.5
    a = (p[0] - tang[0] * half, p[1] - tang[1] * half)
    b = (p[0] + tang[0] * half, p[1] + tang[1] * half)
    # inland normal toward center
    to_c = (center[0] - p[0], center[1] - p[1])
    L = math.hypot(*to_c) or 1.0
    nx, ny = to_c[0] / L, to_c[1] / L
    tip = (p[0] + nx * FLARE_DEPTH, p[1] + ny * FLARE_DEPTH)
    sol = ms.AddSolid(vpt(*a), vpt(*b), vpt(*tip), vpt(*tip))
    sol.Layer = LAYER_FLARE
    sol.Color = 2
    print(
        f"    flare s={s_end:.2f} base=({a[0]:.2f},{a[1]:.2f})-({b[0]:.2f},{b[1]:.2f}) "
        f"tip=({tip[0]:.2f},{tip[1]:.2f})"
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
        try:
            doc.HandleToObject(job["outer"])
        except Exception:
            print("missing outer", job["outer"])
            continue
        n = delete_auto_flares(doc, job["center"])
        print(f"deleted old auto flares: {n}")

        outer = sample_fl(doc, job["outer"])
        print(f"outer n={len(outer)}")
        if len(outer) < 5:
            continue
        svals = stations(outer)
        try:
            dws = trail_bbox(doc, job["trail"])
        except Exception as ex:
            print("trail fail", ex)
            continue
        span = depressed_ends_on_outer(outer, svals, dws, job["center"])
        if not span:
            print("no span on outer — skip")
            continue
        lo, hi, wrap = span
        slen = (svals[-1] - lo) + hi if wrap else (hi - lo)
        print(f"span [{lo:.2f}..{hi:.2f}] wrap={wrap} len={slen:.2f}")
        place_flare_on_outer(ms, outer, svals, lo, job["center"])
        place_flare_on_outer(ms, outer, svals, hi, job["center"])

    doc.Regen(1)
    print("\nDONE — NW flares untouched; E/MidN/MidS rebuilt on OUTER")


if __name__ == "__main__":
    main()
