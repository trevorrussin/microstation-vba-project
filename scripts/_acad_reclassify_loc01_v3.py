"""Classify loc01 curb: Depressed=flare-to-flare; Corner=curved; Straight=linear.

Corner detection: each elementary segment inherits curvature from a station
window along the parent poly (path/chord + cumulative heading turn). Small
chord segments that sit on a larger arc become Corner; colinear runs become
Straight. Split vertices inserted at flare/DWS projections.
"""
from __future__ import annotations

import math
import os
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

# PROP-D ramp anonymous blocks + PED_RAMP at loc01
RAMP_HANDLES = ["13CDD7", "13D60F", "13CDC0", "13D4BC", "13D72E", "133817", "13C6D0"]

# Known explodable inner curb (NW island)
TOPO_3DPOLY = ["13D986"]

# Inner-ish feature lines near each island (prefer smaller FLs along face)
FL_HANDLES = [
    "13D822",
    "13D824",
    "13DA05",
    "13DA12",  # NW near U24/U28
    "13D263",
    "13D265",
    "13DA1F",  # N near U25
    "13D7E0",
    "13D7DB",
    "13DA28",
    "13DA33",  # S near U27/U29
    "13CED4",
    "13CEC7",  # E near PED_RAMP / U21
]

LAYER_STRAIGHT = "Straight Steel Faced Concrete Curb (18'' Deep)"
LAYER_DEPRESSED = "Depressed Steel Faced Concrete Curb (18'' Deep)"
LAYER_CORNER = "Corner Steel Faced Concrete Curb (18'' Deep)"
COLORS = {LAYER_STRAIGHT: 1, LAYER_DEPRESSED: 3, LAYER_CORNER: 5}

CURB_NEAR = 2.5
# Neighborhood along poly used to judge larger curvature from small segs
CURVE_WIN_FT = 5.0
# path/chord above this => curved (gentle arc of small chords)
PATH_CHORD_CURVE = 1.006
# absolute cumulative turn (deg) in window => curved
TURN_CURVE_DEG = 6.0
# single-vertex kink still counts if large
KINK_DEG = 4.0

PTS_FILE = r"c:\repos\microstation-vba-project\scripts\_fl_curve_pts.txt"


def vpt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def xform(br, x, y, z=0.0):
    ins = br.InsertionPoint
    ang = float(br.Rotation)
    sx, sy = float(br.XScaleFactor), float(br.YScaleFactor)
    ca, sa = math.cos(ang), math.sin(ang)
    xr, yr = x * sx, y * sy
    return (
        float(ins[0]) + xr * ca - yr * sa,
        float(ins[1]) + xr * sa + yr * ca,
        float(ins[2]) + z,
    )


def dist(a, b):
    return math.hypot(b[0] - a[0], b[1] - a[1])


def heading(a, b):
    return math.atan2(b[1] - a[1], b[0] - a[0])


def ang_diff(a, b):
    d = b - a
    while d > math.pi:
        d -= 2 * math.pi
    while d < -math.pi:
        d += 2 * math.pi
    return d


def dedupe(pts, tol=1e-5):
    out = []
    for p in pts:
        if not out or dist(out[-1], p) > tol:
            out.append(p)
    return out


def poly_pts(e):
    if e.ObjectName == "AcDbArc":
        c = e.Center
        cx, cy = float(c[0]), float(c[1])
        cz = float(c[2]) if len(list(c)) > 2 else 0.0
        r = float(e.Radius)
        a0 = float(e.StartAngle)
        a1 = float(e.EndAngle)
        if a1 < a0:
            a1 += 2 * math.pi
        arc_len = r * (a1 - a0)
        n = max(8, int(arc_len / 0.25) + 1)
        return [
            (
                cx + r * math.cos(a0 + (a1 - a0) * i / n),
                cy + r * math.sin(a0 + (a1 - a0) * i / n),
                cz,
            )
            for i in range(n + 1)
        ]
    if e.ObjectName == "AcDbLine":
        s, en = e.StartPoint, e.EndPoint
        return [
            (float(s[0]), float(s[1]), float(s[2])),
            (float(en[0]), float(en[1]), float(en[2])),
        ]
    coords = list(e.Coordinates)
    step = 3 if e.ObjectName == "AcDb3dPolyline" else 2
    pts = []
    for i in range(0, len(coords), step):
        z = float(coords[i + 2]) if step == 3 else 0.0
        pts.append((float(coords[i]), float(coords[i + 1]), z))
    return dedupe(pts)


def stations(pts):
    s = [0.0]
    for i in range(1, len(pts)):
        s.append(s[-1] + dist(pts[i - 1], pts[i]))
    return s


def project_station(pts, svals, x, y):
    best_d, best_s, best_pt = 1e9, 0.0, pts[0]
    for i in range(len(pts) - 1):
        ax, ay = pts[i][0], pts[i][1]
        bx, by = pts[i + 1][0], pts[i + 1][1]
        dx, dy = bx - ax, by - ay
        L2 = dx * dx + dy * dy
        t = 0.0 if L2 < 1e-12 else max(0.0, min(1.0, ((x - ax) * dx + (y - ay) * dy) / L2))
        px, py = ax + t * dx, ay + t * dy
        d = math.hypot(x - px, y - py)
        if d < best_d:
            best_d = d
            best_s = svals[i] + t * (svals[i + 1] - svals[i])
            z = pts[i][2] + t * (pts[i + 1][2] - pts[i][2])
            best_pt = (px, py, z)
    return best_s, best_d, best_pt


def point_at_station(pts, svals, s_target):
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
                pts[i][2] + t * (pts[i + 1][2] - pts[i][2]),
            )
    return pts[-1]


def insert_stations(pts, svals, new_s):
    all_s = sorted(set(round(s, 5) for s in list(svals) + list(new_s)))
    out = []
    for s in all_s:
        hit = None
        for i, sv in enumerate(svals):
            if abs(sv - s) < 1e-4:
                hit = pts[i]
                break
        if hit is None:
            hit = point_at_station(pts, svals, s)
        if not out or dist(out[-1], hit) > 1e-6:
            out.append(hit)
    return out


def window_is_curved(pts, svals, i_seg):
    """True if neighborhood around segment i looks like a larger curve."""
    n = len(pts)
    if n < 3:
        return False
    sm = 0.5 * (svals[i_seg] + svals[i_seg + 1])
    lo = i_seg
    hi = i_seg + 1
    while lo > 0 and sm - svals[lo - 1] <= CURVE_WIN_FT:
        lo -= 1
    while hi < n - 1 and svals[hi + 1] - sm <= CURVE_WIN_FT:
        hi += 1
    if hi - lo < 2:
        # expand to at least 3 verts if available
        while hi - lo < 2 and (lo > 0 or hi < n - 1):
            if lo > 0:
                lo -= 1
            if hi < n - 1:
                hi += 1
    path = svals[hi] - svals[lo]
    chord = dist(pts[lo], pts[hi])
    if path < 0.15:
        return False
    ratio = path / max(chord, 1e-9)
    turn = 0.0
    max_kink = 0.0
    for k in range(lo + 1, hi):
        kdeg = abs(math.degrees(ang_diff(heading(pts[k - 1], pts[k]), heading(pts[k], pts[k + 1]))))
        turn += kdeg
        max_kink = max(max_kink, kdeg)
    return (ratio >= PATH_CHORD_CURVE) or (turn >= TURN_CURVE_DEG) or (max_kink >= KINK_DEG)


def ramp_flare_data(doc, handle):
    br = doc.HandleToObject(handle)
    blk = doc.Blocks.Item(br.Name)
    flares, dws_pts, area = [], [], None
    for i in range(blk.Count):
        e = blk.Item(i)
        if e.ObjectName == "AcDbSolid":
            pts = []
            for pi in range(4):
                try:
                    p = e.Coordinate(pi)
                    pts.append(xform(br, float(p[0]), float(p[1])))
                except Exception:
                    break
            uniq = []
            for p in pts:
                if not any(math.hypot(p[0] - q[0], p[1] - q[1]) < 1e-5 for q in uniq):
                    uniq.append(p)
            if len(uniq) >= 3:
                flares.append(uniq)
        if e.ObjectName == "AcDbHatch" and "WARN" in (e.Layer or "").upper():
            area = float(e.Area)
            mn, mx = e.GetBoundingBox()
            corners = [
                (float(mn[0]), float(mn[1])),
                (float(mx[0]), float(mn[1])),
                (float(mx[0]), float(mx[1])),
                (float(mn[0]), float(mx[1])),
            ]
            # local vs WCS
            if abs(float(mn[0])) < 50 and abs(float(mn[1])) < 50:
                dws_pts = [xform(br, c[0], c[1])[:2] for c in corners]
            else:
                dws_pts = corners
    return {"handle": handle, "name": br.Name, "flares": flares, "dws": dws_pts, "area": area}


def depressed_span(pts, svals, ramp):
    """Flare-to-flare: min/max station of all flare + DWS vertices near curb."""
    curb_s = []
    for flare in ramp["flares"]:
        for p in flare:
            s, d, _ = project_station(pts, svals, p[0], p[1])
            if d <= CURB_NEAR:
                curb_s.append(s)
    for p in ramp["dws"]:
        s, d, _ = project_station(pts, svals, p[0], p[1])
        if d <= CURB_NEAR + 0.75:
            curb_s.append(s)
    if len(curb_s) < 2:
        return None
    return (min(curb_s), max(curb_s))


def ensure_layers(doc):
    for name, col in COLORS.items():
        found = None
        for i in range(doc.Layers.Count):
            if doc.Layers.Item(i).Name == name:
                found = doc.Layers.Item(i)
                break
        if found is None:
            found = doc.Layers.Add(name)
        found.Color = col


def clear_classified(ms, doc):
    to_del = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer in COLORS:
                to_del.append(e.Handle)
        except Exception:
            continue
    for h in to_del:
        try:
            doc.HandleToObject(h).Delete()
        except Exception:
            pass
    return len(to_del)


def classify_clean(ms, pts, ramps, totals, counts, tag):
    svals = stations(pts)
    spans = []
    split_s = []
    for ramp in ramps:
        span = depressed_span(pts, svals, ramp)
        if span and span[1] - span[0] >= 0.4:
            spans.append((span[0], span[1], ramp["handle"]))
            split_s.extend([span[0], span[1]])
            print(
                f"  [{tag}] depressed {ramp['handle']} "
                f"[{span[0]:.2f}..{span[1]:.2f}] = {span[1]-span[0]:.2f} ft"
            )
    if not spans:
        print(f"  [{tag}] no flare-to-flare span — skip")
        return False

    new_pts = insert_stations(pts, svals, split_s)
    new_s = stations(new_pts)

    def in_dep(sm):
        return any(a - 1e-6 <= sm <= b + 1e-6 for a, b, _ in spans)

    # Map original-ish curvature using new_pts (same shape, extra verts at flares)
    for i in range(len(new_pts) - 1):
        a, b = new_pts[i], new_pts[i + 1]
        L = dist(a, b)
        if L < 1e-4:
            continue
        sm = 0.5 * (new_s[i] + new_s[i + 1])
        if in_dep(sm):
            layer = LAYER_DEPRESSED
            label = "Depressed"
        elif window_is_curved(new_pts, new_s, i):
            layer = LAYER_CORNER
            label = "Corner"
        else:
            layer = LAYER_STRAIGHT
            label = "Straight"
        line = ms.AddLine(vpt(*a), vpt(*b))
        line.Layer = layer
        line.Color = 256
        totals[layer] += L
        counts[layer] += 1
        print(f"    {L:6.3f} {label}")
    return True


def sample_fl_vlax(doc, handle, step=0.25):
    """Sample AECC feature line via vlax-curve → text file (no explode)."""
    if os.path.exists(PTS_FILE):
        try:
            os.remove(PTS_FILE)
        except Exception:
            pass
    path_esc = PTS_FILE.replace("\\", "/")
    lisp = (
        f'(progn (vl-load-com) (setq e (handent "{handle}")) '
        f'(if (null e) (progn (princ "NOENT") (princ)) '
        f'(progn (setq len (vlax-curve-getDistAtParam e (vlax-curve-getEndParam e))) '
        f'(setq f (open "{path_esc}" "w")) '
        f'(setq d 0.0) '
        f'(while (<= d (+ len 1e-6)) '
        f'(setq p (vlax-curve-getPointAtDist e (min d len))) '
        f'(if p (write-line (strcat (rtos (car p) 2 6) "," (rtos (cadr p) 2 6) "," '
        f'(rtos (if (caddr p) (caddr p) 0.0) 2 6)) f)) '
        f'(setq d (+ d {step:.3f}))) '
        f'(close f) (princ "OK"))) (princ)) '
    )
    try:
        doc.SendCommand(lisp)
    except Exception as ex:
        print("  vlax SendCommand fail", ex)
        return []
    # wait for file
    for _ in range(40):
        time.sleep(0.25)
        if os.path.exists(PTS_FILE) and os.path.getsize(PTS_FILE) > 10:
            break
    else:
        print(f"  vlax sample timeout for {handle}")
        return []
    pts = []
    with open(PTS_FILE, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line or "," not in line:
                continue
            parts = line.split(",")
            if len(parts) < 2:
                continue
            x, y = float(parts[0]), float(parts[1])
            z = float(parts[2]) if len(parts) > 2 else 0.0
            pts.append((x, y, z))
    return dedupe(pts, tol=1e-4)


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    print("file", doc.FullName)
    try:
        doc.SendCommand(chr(27) + chr(27))
        time.sleep(0.2)
    except Exception:
        pass

    ensure_layers(doc)
    print("cleared", clear_classified(ms, doc))

    ramps = []
    for h in RAMP_HANDLES:
        try:
            r = ramp_flare_data(doc, h)
            ramps.append(r)
            print(f"ramp {h} {r['name']} flares={len(r['flares'])} dws={len(r['dws'])} area={r['area']}")
        except Exception as ex:
            print(f"ramp {h} FAIL {ex}")

    totals = {k: 0.0 for k in COLORS}
    counts = {k: 0 for k in COLORS}
    used_sources = set()

    # 1) Explodable 3d polys
    for h in TOPO_3DPOLY:
        try:
            e = doc.HandleToObject(h)
        except Exception:
            print(f"missing 3dpoly {h}")
            continue
        pts = poly_pts(e)
        print(f"\n=== 3dpoly {h} n={len(pts)} L={stations(pts)[-1]:.2f} ===")
        ok = classify_clean(ms, pts, ramps, totals, counts, h)
        if ok:
            used_sources.add(h)
            try:
                e.Delete()
                print("deleted", h)
            except Exception as ex:
                print("del", ex)

    # 2) Feature lines via vlax sample (other islands + any leftover)
    for h in FL_HANDLES:
        if h in used_sources:
            continue
        try:
            doc.HandleToObject(h)
        except Exception:
            print(f"missing FL {h}")
            continue
        print(f"\n=== sample FL {h} ===")
        pts = sample_fl_vlax(doc, h, step=0.30)
        print(f"  sampled n={len(pts)}")
        if len(pts) < 3:
            continue
        ok = classify_clean(ms, pts, ramps, totals, counts, h)
        if ok:
            used_sources.add(h)
            try:
                doc.HandleToObject(h).Delete()
                print("deleted FL", h)
            except Exception:
                pass

    # Cleanup NW duplicate FLs if still present after 3dpoly path
    for h in ["13D822", "13D824"]:
        try:
            doc.HandleToObject(h).Delete()
            print("deleted leftover FL", h)
        except Exception:
            pass

    doc.Regen(1)
    print("\n===== QUANTITIES (LF) =====")
    for lyr in (LAYER_STRAIGHT, LAYER_DEPRESSED, LAYER_CORNER):
        print(f"{lyr}\n  n={counts[lyr]}  LF={totals[lyr]:.2f}")
    print(f"TOTAL LF={sum(totals.values()):.2f}")
    print("\n===== DWS SF =====")
    for r in ramps:
        a = r["area"]
        print(f"  {r['name']} ({r['handle']}): {a:.2f}" if a else f"  {r['name']}: n/a")
    areas = [r["area"] for r in ramps if r["area"]]
    print(f"TOTAL DWS SF={sum(areas):.2f}")
    print("DONE sources=", sorted(used_sources))


if __name__ == "__main__":
    main()
