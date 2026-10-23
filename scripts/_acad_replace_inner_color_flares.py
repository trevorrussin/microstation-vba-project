"""Replace green INNER C-TOPO-FEAT with color-coded curb segs (same path).

Does NOT offset/rebuild geometry. Does NOT leave a second parallel line.
1) Sample existing INNER FL verts
2) Place AcDbLine segs at those exact coords on Straight/Depressed/Corner
3) Delete the original INNER FL (replace green topo — never delete OUTER)
4) Place yellow flares independently (outer + PRO-S pocket / NW-style)
"""
from __future__ import annotations

import math
import os
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

PTS = r"c:\repos\microstation-vba-project\scripts\_fl_curve_pts.txt"
LAYER_FEAT = "C-TOPO-FEAT"
LAYER_FLARE = "PROP-D-ROAD-RAMPS"
LAYER_STRAIGHT = "Straight Steel Faced Concrete Curb (18'' Deep)"
LAYER_DEPRESSED = "Depressed Steel Faced Concrete Curb (18'' Deep)"
LAYER_CORNER = "Corner Steel Faced Concrete Curb (18'' Deep)"
COLORS = {LAYER_STRAIGHT: 1, LAYER_DEPRESSED: 3, LAYER_CORNER: 5}

# OUTER — never delete / never re-layer
OUTER_HANDLES = {"13D80A", "13D25C", "13D7DC", "13CEC7", "13D2EA", "13D7BC"}

ISLANDS = [
    {
        "name": "NW",
        "outer": "13D80A",
        "inner": ["13D822", "13D824"],
        "center": (1028470.0, 187478.0),
        "trail": None,
        "pro": [],
        "prop_blocks": ["13CDD7", "13D60F"],
        "nw_style": True,
    },
    {
        "name": "MidN",
        "outer": "13D2EA",
        "inner": ["13D425", "13D42D", "13D2EE"],
        "center": (1028501.0, 187485.0),
        "trail": "13C0E2",
        "pro": ["13A7EA", "13C903"],
        "nw_style": False,
    },
    {
        "name": "MidS",
        "outer": "13D7BC",
        "inner": ["13D7C4", "13D7CC", "13D7BF"],
        "center": (1028525.0, 187428.0),
        "trail": "138B88",
        "pro": ["13C116", "13C115", "138B64", "13C226", "138B86", "138B6B"],
        "nw_style": False,
    },
    {
        "name": "E",
        "outer": "13CEC7",
        "inner": ["13CED4", "13CED2"],
        "center": (1028589.0, 187445.0),
        "trail": "133817",
        "pro": ["13383E"],
        "nw_style": False,
    },
    {
        "name": "N",
        "outer": "13D25C",
        "inner": ["13D263", "13D265", "13D25E", "13D282"],
        "center": (1028564.0, 187503.0),
        "trail": None,
        "pro": [],
        "prop_blocks": ["13CDC0"],
        "nw_style": True,
    },
    {
        "name": "SW",
        "outer": "13D7DC",
        "inner": ["13D7DB", "13D7E0", "13D7E8"],
        "center": (1028486.0, 187421.0),
        "trail": None,
        "pro": ["13C17C", "13C17D"],
        "prop_blocks": ["13D4BC", "13D72E"],
        "nw_style": False,
    },
]

CORNER_PAD = 0.75
PRO_OFF_MIN = 1.0
FLARE_BASE = 2.5
FLARE_DEPTH = 5.0
MIN_AREA = 0.4


def vpt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def dist(a, b):
    return math.hypot(b[0] - a[0], b[1] - a[1])


def ensure_layer(doc, name, color):
    layers = doc.Layers
    lyr = None
    for i in range(layers.Count):
        try:
            if layers.Item(i).Name == name:
                lyr = layers.Item(i)
                break
        except Exception:
            continue
    if lyr is None:
        lyr = layers.Add(name)
    lyr.Color = color
    print(f"layer OK {name!r} color={color}")
    return lyr


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
    for _ in range(60):
        time.sleep(0.12)
        if os.path.exists(PTS) and os.path.getsize(PTS) > 10:
            break
    pts = []
    try:
        with open(PTS) as f:
            for line in f:
                a = line.strip().split(",")
                if len(a) >= 2:
                    pts.append((float(a[0]), float(a[1])))
    except Exception:
        return []
    out = []
    for p in pts:
        if not out or dist(out[-1], p) > 1e-4:
            out.append(p)
    return out


def entity_pts(doc, handle):
    try:
        e = doc.HandleToObject(handle)
    except Exception:
        return []
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
    if e.ObjectName == "AcDbLine":
        s, en = e.StartPoint, e.EndPoint
        return [(float(s[0]), float(s[1])), (float(en[0]), float(en[1]))]
    if "Poly" in e.ObjectName:
        try:
            coords = list(e.Coordinates)
            step = 3 if e.ObjectName == "AcDb3dPolyline" else 2
            return [
                (float(coords[i]), float(coords[i + 1]))
                for i in range(0, len(coords), step)
            ]
        except Exception:
            pass
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


def tri_area(a, b, c):
    return abs(
        (a[0] * (b[1] - c[1]) + b[0] * (c[1] - a[1]) + c[0] * (a[1] - b[1])) / 2
    )


def bbox_pts(doc, handle):
    e = doc.HandleToObject(handle)
    mn, mx = e.GetBoundingBox()
    return [
        (float(mn[0]), float(mn[1])),
        (float(mx[0]), float(mn[1])),
        (float(mx[0]), float(mx[1])),
        (float(mn[0]), float(mx[1])),
    ]


def dws_ends(outer, os_, doc, isl):
    hits = []
    sources = []
    if isl.get("trail"):
        sources.append(isl["trail"])
    sources.extend(isl.get("prop_blocks") or [])
    for h in sources:
        try:
            for p in bbox_pts(doc, h):
                s, d = project_station(outer, os_, p[0], p[1])
                if d < 14:
                    hits.append(s)
        except Exception:
            continue
    if len(hits) < 2:
        cs, _ = project_station(outer, os_, isl["center"][0], isl["center"][1])
        return max(0.5, cs - 6), min(os_[-1] - 0.5, cs + 6)
    L = os_[-1]
    start = [s for s in hits if s <= 25]
    end = [s for s in hits if s >= L - 25]
    if start and end and (max(hits) - min(hits) > 20):
        return min(end), max(start)
    return min(hits), max(hits)


def gather_pro(doc, isl, outer):
    pts = []
    for h in isl.get("pro") or []:
        pts.extend(entity_pts(doc, h))
    ms = doc.ModelSpace
    cx0, cy0 = isl["center"]
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if "PRO-S-CURB" not in (e.Layer or "").upper():
                continue
            if e.ObjectName == "AcDbBlockReference":
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if math.hypot(cx - cx0, cy - cy0) > 35:
                continue
            pts.extend(entity_pts(doc, e.Handle))
        except Exception:
            continue
    kept = []
    for p in pts:
        if closest_on_poly(p[0], p[1], outer)[1] >= PRO_OFF_MIN:
            kept.append(p)
    return kept


def place_solid(ms, a, b, c):
    if tri_area(a, b, c) < MIN_AREA:
        return None
    if dist(a, b) < 0.15 or dist(b, c) < 0.15 or dist(a, c) < 0.15:
        return None
    sol = ms.AddSolid(vpt(*a), vpt(*b), vpt(*c), vpt(*c))
    sol.Layer = LAYER_FLARE
    sol.Color = 2
    return sol


def nw_style_flare(ms, outer, os_, s_end, center):
    Po, tang = tangent_at(outer, os_, s_end)
    half = FLARE_BASE * 0.5
    a = (Po[0] - tang[0] * half, Po[1] - tang[1] * half)
    b = (Po[0] + tang[0] * half, Po[1] + tang[1] * half)
    to_c = (center[0] - Po[0], center[1] - Po[1])
    L = math.hypot(*to_c) or 1.0
    tip = (Po[0] + to_c[0] / L * FLARE_DEPTH, Po[1] + to_c[1] / L * FLARE_DEPTH)
    return place_solid(ms, a, b, tip)


def pocket_flare(ms, outer, os_, s_end, pro_pts, center):
    if len(pro_pts) < 2:
        return nw_style_flare(ms, outer, os_, s_end, center)
    Po, tang = tangent_at(outer, os_, s_end)
    half = FLARE_BASE * 0.55
    a = (Po[0] - tang[0] * half, Po[1] - tang[1] * half)
    b = (Po[0] + tang[0] * half, Po[1] + tang[1] * half)
    cands = []
    for p in pro_pts:
        s, d = project_station(outer, os_, p[0], p[1])
        if d < PRO_OFF_MIN:
            continue
        ds = abs(s - s_end)
        ds = min(ds, os_[-1] - ds)
        if ds > 14:
            continue
        cands.append(p)
    if not cands:
        scored = sorted(
            ((closest_on_poly(p[0], p[1], outer)[1], dist(p, Po), p) for p in pro_pts),
            key=lambda t: (t[1] - 0.3 * t[0]),
        )
        cands = [t[2] for t in scored[:12] if t[0] >= PRO_OFF_MIN]
    best = None
    best_area = 0.0
    for tip in cands:
        for base_a, base_b in ((a, b), (Po, a), (Po, b)):
            ar = tri_area(base_a, base_b, tip)
            if ar > best_area:
                best_area = ar
                best = (base_a, base_b, tip)
    if not best or best_area < MIN_AREA:
        return nw_style_flare(ms, outer, os_, s_end, center)
    return place_solid(ms, *best)


def delete_auto_flares_near(doc, center, radius=30.0):
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


def delete_wrong_extra_steel_lines(doc):
    """Remove AcDbLines on steel layers from the mistaken 'add new line' run."""
    ms = doc.ModelSpace
    to_del = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbLine":
                continue
            if e.Layer not in COLORS:
                continue
            to_del.append(e.Handle)
        except Exception:
            continue
    for h in to_del:
        try:
            doc.HandleToObject(h).Delete()
        except Exception:
            pass
    return len(to_del)


def merge_inner_pts(doc, handles):
    pieces = []
    alive = []
    for h in handles:
        try:
            e = doc.HandleToObject(h)
            if e.Handle in OUTER_HANDLES:
                continue
            pts = sample_fl(doc, h, step=0.35)
            if len(pts) >= 2:
                pieces.append(pts)
                alive.append(h)
                print(f"    sample inner {h} n={len(pts)}")
        except Exception as ex:
            print(f"    inner {h} FAIL {ex}")
    if not pieces:
        return [], []
    chain = list(pieces[0])
    used = {0}
    while len(used) < len(pieces):
        end = chain[-1]
        best_i, best_d, rev = None, 1e9, False
        for i, pts in enumerate(pieces):
            if i in used:
                continue
            d0 = dist(end, pts[0])
            d1 = dist(end, pts[-1])
            if d0 < best_d:
                best_d, best_i, rev = d0, i, False
            if d1 < best_d:
                best_d, best_i, rev = d1, i, True
        if best_i is None:
            break
        pts = list(reversed(pieces[best_i])) if rev else list(pieces[best_i])
        chain.extend(pts)
        used.add(best_i)
    out = []
    for p in chain:
        if not out or dist(out[-1], p) > 0.05:
            out.append(p)
    return out, alive


def point_in_tri(p, a, b, c):
    x, y = p
    x1, y1 = a
    x2, y2 = b
    x3, y3 = c
    denom = (y2 - y3) * (x1 - x3) + (x3 - x2) * (y1 - y3)
    if abs(denom) < 1e-12:
        return False
    aa = ((y2 - y3) * (x - x3) + (x3 - x2) * (y - y3)) / denom
    bb = ((y3 - y1) * (x - x3) + (x1 - x3) * (y - y3)) / denom
    cc = 1.0 - aa - bb
    return aa >= -1e-6 and bb >= -1e-6 and cc >= -1e-6


def dist_to_tri(p, tri):
    if point_in_tri(p, *tri):
        return 0.0
    best = 1e9
    for i in range(3):
        a, b = tri[i], tri[(i + 1) % 3]
        ax, ay = a
        bx, by = b
        dx, dy = bx - ax, by - ay
        L2 = dx * dx + dy * dy
        t = 0.0 if L2 < 1e-12 else max(
            0.0, min(1.0, ((p[0] - ax) * dx + (p[1] - ay) * dy) / L2)
        )
        best = min(best, math.hypot(p[0] - (ax + t * dx), p[1] - (ay + t * dy)))
    return best


def collect_flares(doc, center, r=35.0):
    ms = doc.ModelSpace
    out = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbSolid" or e.Layer != LAYER_FLARE:
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
                if not any(dist(p, q) < 1e-4 for q in u):
                    u.append(p)
            if len(u) < 3 or tri_area(u[0], u[1], u[2]) < MIN_AREA:
                continue
            cx = sum(p[0] for p in u[:3]) / 3
            cy = sum(p[1] for p in u[:3]) / 3
            if math.hypot(cx - center[0], cy - center[1]) > r:
                continue
            out.append((cx, cy, (u[0], u[1], u[2])))
        except Exception:
            continue
    return out


def replace_inner_with_colored(ms, doc, inner_pts, flares, center, delete_handles):
    """Draw colored segs on same path, then delete original green INNER FLs."""
    totals = {LAYER_STRAIGHT: 0.0, LAYER_DEPRESSED: 0.0, LAYER_CORNER: 0.0}
    counts = {LAYER_STRAIGHT: 0, LAYER_DEPRESSED: 0, LAYER_CORNER: 0}
    if len(inner_pts) < 2:
        return totals, counts

    dep = False
    a0 = a1 = None
    if len(flares) >= 2:
        angs = sorted(math.atan2(f[1] - center[1], f[0] - center[0]) for f in flares)
        a0, a1 = angs[0], angs[-1]
        dep = True

    for i in range(len(inner_pts) - 1):
        a, b = inner_pts[i], inner_pts[i + 1]
        L = dist(a, b)
        if L < 0.02:
            continue
        mx, my = 0.5 * (a[0] + b[0]), 0.5 * (a[1] + b[1])
        lyr = LAYER_STRAIGHT
        for _cx, _cy, tri in flares:
            if dist_to_tri((mx, my), tri) <= CORNER_PAD:
                lyr = LAYER_CORNER
                break
        if lyr == LAYER_STRAIGHT and dep:
            ang = math.atan2(my - center[1], mx - center[0])
            span = a1 - a0
            near = math.hypot(mx - center[0], my - center[1]) < 28
            if near:
                if span <= math.pi:
                    if a0 - 0.2 <= ang <= a1 + 0.2:
                        lyr = LAYER_DEPRESSED
                elif ang <= a0 + 0.2 or ang >= a1 - 0.2:
                    lyr = LAYER_DEPRESSED
        line = ms.AddLine(vpt(*a), vpt(*b))
        line.Layer = lyr
        line.Color = 256  # ByLayer
        totals[lyr] += L
        counts[lyr] += 1

    # REPLACE green inner feat — delete originals (never outers)
    deleted = 0
    for h in delete_handles:
        if h in OUTER_HANDLES:
            print(f"    REFUSE delete OUTER {h}")
            continue
        try:
            e = doc.HandleToObject(h)
            if e.Layer != LAYER_FEAT:
                print(f"    skip delete {h} layer={e.Layer}")
                continue
            e.Delete()
            deleted += 1
        except Exception as ex:
            print(f"    delete {h} fail {ex}")
    print(f"    replaced: deleted {deleted} green INNER FL(s)")
    return totals, counts


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    try:
        doc.SendCommand(chr(27) + chr(27))
        time.sleep(0.3)
    except Exception:
        pass

    for name, color in COLORS.items():
        ensure_layer(doc, name, color)

    print("cleanup mistaken extra steel lines", delete_wrong_extra_steel_lines(doc))

    grand = {LAYER_STRAIGHT: 0.0, LAYER_DEPRESSED: 0.0, LAYER_CORNER: 0.0}
    grand_n = {LAYER_STRAIGHT: 0, LAYER_DEPRESSED: 0, LAYER_CORNER: 0}

    for isl in ISLANDS:
        name = isl["name"]
        print(f"\n=== {name} ===")
        try:
            outer = sample_fl(doc, isl["outer"], step=0.3)
        except Exception as ex:
            print("  outer FAIL", ex)
            continue
        if len(outer) < 5:
            print("  outer too short")
            continue
        os_ = stations(outer)
        print(f"  outer n={len(outer)} L={os_[-1]:.1f}")

        n = delete_auto_flares_near(doc, isl["center"], 30)
        print(f"  deleted old flares {n}")
        lo, hi = dws_ends(outer, os_, doc, isl)
        print(f"  ends s={lo:.2f},{hi:.2f}")
        pro = gather_pro(doc, isl, outer)
        print(f"  PRO off-outer pts {len(pro)}")
        for s_end, tag in ((hi, "hi"), (lo, "lo")):
            if isl.get("nw_style") and len(pro) < 3:
                sol = nw_style_flare(ms, outer, os_, s_end, isl["center"])
            else:
                sol = pocket_flare(ms, outer, os_, s_end, pro, isl["center"])
            if sol:
                print(f"  +flare {tag} s={s_end:.2f} {sol.Handle}")
            else:
                print(f"  SKIP flare {tag}")

        flares = collect_flares(doc, isl["center"])
        print(f"  flares now {len(flares)}")

        inner, alive = merge_inner_pts(doc, isl["inner"])
        print(f"  inner path n={len(inner)} from {alive}")
        if len(inner) < 2:
            continue
        totals, counts = replace_inner_with_colored(
            ms, doc, inner, flares, isl["center"], alive
        )
        for lyr in COLORS:
            grand[lyr] += totals[lyr]
            grand_n[lyr] += counts[lyr]
            print(f"  {lyr.split()[0]:9s} {counts[lyr]:3d} {totals[lyr]:7.1f} ft")

    try:
        doc.Regen(1)
    except Exception:
        pass

    print("\n=== TOTAL LF ===")
    for lyr in (LAYER_STRAIGHT, LAYER_DEPRESSED, LAYER_CORNER):
        print(f"  {grand_n[lyr]:4d} segs  {grand[lyr]:8.1f} ft  {lyr}")
    print("DONE — green INNER replaced in place; flares independent")


if __name__ == "__main__":
    main()
