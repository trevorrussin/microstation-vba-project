"""v6: rebuild INNER from OUTER offset; strict Corner; add missing flares.

- Never modify/delete OUTER FLs.
- Inner path = outer sampled + offset toward island/ramp (~parallel gap).
- Corner only where local radius < R_CORNER_MAX (real curb returns).
- Long/gentle arcs stay Straight (matches top-left behavior).
- Depressed = flare-to-flare; split verts at flare curb-edge.
- Missing yellow triangles placed on PROP-D-ROAD-RAMPS between
  outer topo, PRO-S curb, and inner curb edge.
"""
from __future__ import annotations

import math
import os
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

# Island defs: outer handle (untouched), approx island center, ramp handles / hatch
ISLANDS = [
    {
        "name": "NW",
        "outer": "13D80A",
        "center": (1028470.0, 187478.0),
        "prop_ramps": [],  # blocks gone; MS solids may exist
        "trail": None,
        "ms_flare_hint": True,
    },
    {
        "name": "N",
        "outer": "13D25C",
        "center": (1028564.0, 187503.0),
        "prop_ramps": ["13CDC0"],
        "trail": None,
    },
    {
        "name": "SW",
        "outer": "13D7DC",
        "center": (1028486.0, 187421.0),
        "prop_ramps": ["13D4BC", "13D72E"],
        "trail": None,
    },
    {
        "name": "E",
        "outer": "13CEC7",
        "center": (1028589.0, 187445.0),
        "prop_ramps": [],
        "trail": "133817",  # PED_RAMP block bbox
        "need_flares": True,
    },
    {
        "name": "MidN",
        "outer": "13D2EA",
        "center": (1028505.0, 187492.0),
        "prop_ramps": [],
        "trail": "13C0E2",
        "need_flares": True,
    },
    {
        "name": "MidS",
        "outer": "13D7BC",
        "center": (1028528.0, 187425.0),
        "prop_ramps": [],
        "trail": "138B88",
        "need_flares": True,
    },
]

OUTER_HANDLES = {isl["outer"] for isl in ISLANDS}

LAYER_STRAIGHT = "Straight Steel Faced Concrete Curb (18'' Deep)"
LAYER_DEPRESSED = "Depressed Steel Faced Concrete Curb (18'' Deep)"
LAYER_CORNER = "Corner Steel Faced Concrete Curb (18'' Deep)"
LAYER_FLARE = "PROP-D-ROAD-RAMPS"
COLORS = {LAYER_STRAIGHT: 1, LAYER_DEPRESSED: 3, LAYER_CORNER: 5}

# Parallel offset from outer toward island (ft) — measured ~2 ft NW/N
INNER_OFFSET = 2.0
# Real curb-return radius cutoff: larger R stays Straight
R_CORNER_MAX = 20.0
CURVE_WIN_FT = 8.0
MIN_CORNER_TURN = 20.0  # deg total turn in window
CURB_NEAR = 3.0
# Only classify street-face curb near ramps (not full island back)
FACE_BUFFER_FT = 22.0
DECIMATE_TURN_DEG = 1.5
DECIMATE_MAX_SPACING = 2.5
FLARE_ALONG = 3.5  # short edge along curb (~existing flares)
FLARE_DEPTH = 5.0  # how far tip goes toward PRO-S / island

PTS_FILE = r"c:\repos\microstation-vba-project\scripts\_fl_curve_pts.txt"


def vpt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


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


def dedupe(pts, tol=1e-4):
    out = []
    for p in pts:
        if not out or dist(out[-1], p) > tol:
            out.append(p)
    return out


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
            z = pts[i][2] + t * (pts[i + 1][2] - pts[i][2]) if len(pts[i]) > 2 else 0.0
            best_pt = (px, py, z if len(pts[i]) > 2 else 0.0)
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
                (pts[i][2] + t * (pts[i + 1][2] - pts[i][2])) if len(pts[i]) > 2 else 0.0,
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
            out.append(hit if len(hit) == 3 else (hit[0], hit[1], 0.0))
    return out


def circumradius(p0, p1, p2):
    a = dist(p1, p2)
    b = dist(p0, p2)
    c = dist(p0, p1)
    if a < 1e-9 or b < 1e-9 or c < 1e-9:
        return 1e9
    area2 = abs(
        (p1[0] - p0[0]) * (p2[1] - p0[1]) - (p1[1] - p0[1]) * (p2[0] - p0[0])
    )
    if area2 < 1e-10:
        return 1e9
    return (a * b * c) / (2.0 * area2)


def decimate_poly(pts):
    """Keep verts at turns or every DECIMATE_MAX_SPACING — like sparse 3dpoly."""
    if len(pts) < 3:
        return pts
    keep = [0]
    acc = 0.0
    for i in range(1, len(pts) - 1):
        acc += dist(pts[i - 1], pts[i])
        turn = abs(
            math.degrees(
                ang_diff(heading(pts[i - 1], pts[i]), heading(pts[i], pts[i + 1]))
            )
        )
        if turn >= DECIMATE_TURN_DEG or acc >= DECIMATE_MAX_SPACING:
            keep.append(i)
            acc = 0.0
    keep.append(len(pts) - 1)
    # unique sorted
    keep = sorted(set(keep))
    return [pts[i] for i in keep]


def window_is_corner(pts, svals, i_seg):
    """Corner only if local fit radius is a tight curb return + real turn."""
    n = len(pts)
    if n < 3:
        return False
    sm = 0.5 * (svals[i_seg] + svals[i_seg + 1])
    lo, hi = i_seg, i_seg + 1
    while lo > 0 and sm - svals[lo - 1] <= CURVE_WIN_FT:
        lo -= 1
    while hi < n - 1 and svals[hi + 1] - sm <= CURVE_WIN_FT:
        hi += 1
    while hi - lo < 2 and (lo > 0 or hi < n - 1):
        if lo > 0:
            lo -= 1
        if hi < n - 1:
            hi += 1
    min_r = 1e9
    for k in range(lo, hi - 1):
        r = circumradius(pts[k], pts[k + 1], pts[k + 2])
        min_r = min(min_r, r)
    if min_r >= R_CORNER_MAX:
        return False
    turn = 0.0
    for k in range(lo + 1, hi):
        turn += abs(
            math.degrees(
                ang_diff(heading(pts[k - 1], pts[k]), heading(pts[k], pts[k + 1]))
            )
        )
    return turn >= MIN_CORNER_TURN


def near_ramp_face(sm, spans, total_len):
    """True if station is on street face near a depressed span."""
    for lo, hi, wrap, _ in spans:
        if wrap:
            # depressed: s>=lo or s<=hi; face buffer expands those
            if sm >= lo - FACE_BUFFER_FT or sm <= hi + FACE_BUFFER_FT:
                return True
            # also wrap across end
            if lo - FACE_BUFFER_FT < 0 and sm >= total_len + (lo - FACE_BUFFER_FT):
                return True
            if hi + FACE_BUFFER_FT > total_len and sm <= (hi + FACE_BUFFER_FT) - total_len:
                return True
        else:
            if lo - FACE_BUFFER_FT <= sm <= hi + FACE_BUFFER_FT:
                return True
    # between two non-wrap spans on same island (corner between twin ramps)
    nonwrap = [(a, b) for a, b, w, _ in spans if not w]
    nonwrap.sort()
    for i in range(len(nonwrap) - 1):
        a1, b1 = nonwrap[i]
        a2, b2 = nonwrap[i + 1]
        if b1 <= sm <= a2 and (a2 - b1) < 40:
            return True
    return False


def sample_fl_vlax(doc, handle, step=0.35):
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
        f'(setq f (open "{path_esc}" "w")) (setq d 0.0) '
        f'(while (<= d (+ len 1e-6)) '
        f'(setq p (vlax-curve-getPointAtDist e (min d len))) '
        f'(if p (write-line (strcat (rtos (car p) 2 6) "," (rtos (cadr p) 2 6) "," '
        f'(rtos (if (caddr p) (caddr p) 0.0) 2 6)) f)) '
        f'(setq d (+ d {step:.3f}))) (close f) (princ "OK"))) (princ)) '
    )
    doc.SendCommand(lisp)
    for _ in range(50):
        time.sleep(0.2)
        if os.path.exists(PTS_FILE) and os.path.getsize(PTS_FILE) > 10:
            break
    else:
        return []
    pts = []
    with open(PTS_FILE, encoding="utf-8", errors="ignore") as f:
        for line in f:
            parts = line.strip().split(",")
            if len(parts) < 2:
                continue
            pts.append(
                (
                    float(parts[0]),
                    float(parts[1]),
                    float(parts[2]) if len(parts) > 2 else 0.0,
                )
            )
    return dedupe(pts)


def offset_toward(pts, center, offset):
    """Offset polyline toward center (island)."""
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
        # left normal
        nx, ny = -ty / L, tx / L
        # pick direction toward center
        to_c_x, to_c_y = cx - pts[i][0], cy - pts[i][1]
        if nx * to_c_x + ny * to_c_y < 0:
            nx, ny = -nx, -ny
        z = pts[i][2] if len(pts[i]) > 2 else 0.0
        out.append((pts[i][0] + offset * nx, pts[i][1] + offset * ny, z))
    return dedupe(out)


def xform(br, x, y):
    ins = br.InsertionPoint
    ang = float(br.Rotation)
    sx, sy = float(br.XScaleFactor), float(br.YScaleFactor)
    ca, sa = math.cos(ang), math.sin(ang)
    return (
        float(ins[0]) + x * sx * ca - y * sy * sa,
        float(ins[1]) + x * sx * sa + y * sy * ca,
    )


def load_prop_flares(doc, handle):
    br = doc.HandleToObject(handle)
    blk = doc.Blocks.Item(br.Name)
    flares, dws, area = [], [], None
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
            u = []
            for p in pts:
                if not any(math.hypot(p[0] - q[0], p[1] - q[1]) < 1e-4 for q in u):
                    u.append(p)
            if len(u) >= 3:
                flares.append(u)
        if e.ObjectName == "AcDbHatch" and "WARN" in (e.Layer or "").upper():
            area = float(e.Area)
            mn, mx = e.GetBoundingBox()
            corners = [
                (float(mn[0]), float(mn[1])),
                (float(mx[0]), float(mn[1])),
                (float(mx[0]), float(mx[1])),
                (float(mn[0]), float(mx[1])),
            ]
            if abs(float(mn[0])) < 50:
                dws = [xform(br, c[0], c[1]) for c in corners]
            else:
                dws = corners
    return {"handle": handle, "name": br.Name, "flares": flares, "dws": dws, "area": area}


def load_ms_flares(doc):
    """Existing modelspace PROP-D solids (NW exploded flares)."""
    ms = doc.ModelSpace
    flares = []
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
                if not any(math.hypot(p[0] - q[0], p[1] - q[1]) < 1e-4 for q in u):
                    u.append(p)
            if len(u) >= 3:
                flares.append(u)
        except Exception:
            continue
    return flares


def trail_dws(doc, handle):
    e = doc.HandleToObject(handle)
    if e.ObjectName == "AcDbBlockReference":
        mn, mx = e.GetBoundingBox()
    else:
        mn, mx = e.GetBoundingBox()
    corners = [
        (float(mn[0]), float(mn[1])),
        (float(mx[0]), float(mn[1])),
        (float(mx[0]), float(mx[1])),
        (float(mn[0]), float(mx[1])),
    ]
    area = float(getattr(e, "Area", 0) or 0) or None
    return {"handle": handle, "name": "trail", "flares": [], "dws": corners, "area": area}


def depressed_span(pts, svals, ramp):
    L = svals[-1]
    curb_s = []
    for flare in ramp.get("flares") or []:
        for p in flare:
            s, d, _ = project_station(pts, svals, p[0], p[1])
            if d <= CURB_NEAR:
                curb_s.append(s)
    for p in ramp.get("dws") or []:
        s, d, _ = project_station(pts, svals, p[0], p[1])
        if d <= CURB_NEAR + 1.0:
            curb_s.append(s)
    if len(curb_s) < 2:
        return None

    dws_s = []
    for p in ramp.get("dws") or []:
        s, d, _ = project_station(pts, svals, p[0], p[1])
        if d <= CURB_NEAR + 1.0:
            dws_s.append(s)

    start_pts = [s for s in curb_s if s <= 20]
    end_pts = [s for s in curb_s if s >= L - 20]
    if start_pts and end_pts and dws_s:
        if any(s <= 20 for s in dws_s) and any(s >= L - 20 for s in dws_s):
            lo, hi = min(end_pts), max(start_pts)
            wlen = (L - lo) + hi
            if 0.4 <= wlen <= 28:
                return (lo, hi, True)

    if dws_s and max(dws_s) - min(dws_s) <= 28:
        if not ramp.get("flares"):
            lo, hi = min(dws_s), max(dws_s)
            if hi - lo >= 0.4:
                return (lo, hi, False)
        mid = sum(dws_s) / len(dws_s)
        tight = [s for s in curb_s if abs(s - mid) <= 14]
        if len(tight) >= 2:
            lo, hi = min(tight), max(tight)
            if 0.4 <= hi - lo <= 28:
                return (lo, hi, False)

    curb_s = sorted(curb_s)
    best = None
    i = 0
    while i < len(curb_s):
        j = i
        while j + 1 < len(curb_s) and curb_s[j + 1] - curb_s[j] < 8:
            j += 1
        cl = curb_s[i : j + 1]
        if len(cl) >= 2 and cl[-1] - cl[0] <= 28:
            if best is None or len(cl) > best[0]:
                best = (len(cl), cl[0], cl[-1])
        i = j + 1
    if not best:
        return None
    return (best[1], best[2], False)


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
    # flare layer should already exist
    try:
        doc.Layers.Item(LAYER_FLARE)
    except Exception:
        doc.Layers.Add(LAYER_FLARE)


def clear_classified(doc):
    ms = doc.ModelSpace
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


def tan_normal_at(pts, svals, s):
    p = point_at_station(pts, svals, s)
    # small step for tangent
    p2 = point_at_station(pts, svals, min(s + 0.2, svals[-1]))
    if dist(p, p2) < 1e-6:
        p2 = point_at_station(pts, svals, max(s - 0.2, svals[0]))
    tx, ty = p2[0] - p[0], p2[1] - p[1]
    L = math.hypot(tx, ty)
    if L < 1e-9:
        return p, (1.0, 0.0), (0.0, 1.0)
    tx, ty = tx / L, ty / L
    return p, (tx, ty), (-ty, tx)


def delete_nearby_ms_flares(doc, center, radius=18.0, keep_handles=None):
    """Remove previously auto-added flare solids near an island (not NW originals)."""
    keep_handles = keep_handles or set()
    ms = doc.ModelSpace
    to_del = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbSolid" or e.Layer != LAYER_FLARE:
                continue
            if e.Handle in keep_handles:
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


def place_flare_solid(ms, p_curb_a, p_curb_b, tip):
    """AcDbSolid needs 3 or 4 corners; triangle as solid with 4th=3rd."""
    s = ms.AddSolid(vpt(*p_curb_a[:2]), vpt(*p_curb_b[:2]), vpt(*tip[:2]), vpt(*tip[:2]))
    s.Layer = LAYER_FLARE
    s.Color = 2  # yellow
    return s


def make_flares_at_span_ends(ms, pts, svals, lo, hi, wrap, center):
    """Build two yellow triangles at depressed span ends along inner curb."""
    L = svals[-1]
    ends = []
    if wrap:
        ends = [lo, hi]  # lo near end, hi near start
    else:
        ends = [lo, hi]
    created = []
    for s_end in ends:
        p, tang, norm = tan_normal_at(pts, svals, s_end)
        # short edge along curb centered at end, length FLARE_ALONG
        half = FLARE_ALONG * 0.5
        a = (p[0] - tang[0] * half, p[1] - tang[1] * half, 0.0)
        b = (p[0] + tang[0] * half, p[1] + tang[1] * half, 0.0)
        # tip toward island center (and slightly inland from curb)
        to_c = (center[0] - p[0], center[1] - p[1])
        tcl = math.hypot(*to_c)
        if tcl < 1e-9:
            nx, ny = norm
        else:
            nx, ny = to_c[0] / tcl, to_c[1] / tcl
        tip = (p[0] + nx * FLARE_DEPTH, p[1] + ny * FLARE_DEPTH, 0.0)
        sol = place_flare_solid(ms, a, b, tip)
        created.append([a[:2], b[:2], tip[:2]])
        print(f"    +flare at s={s_end:.2f} tip=({tip[0]:.2f},{tip[1]:.2f})")
    return created


def classify_island(ms, pts, ramps, totals, counts, tag):
    pts = decimate_poly(pts)
    svals = stations(pts)
    print(f"  [{tag}] decimated n={len(pts)} L={svals[-1]:.2f}")
    spans = []
    split_s = []
    used = []
    for ramp in ramps:
        span = depressed_span(pts, svals, ramp)
        if not span:
            continue
        lo, hi, wrap = span
        slen = (svals[-1] - lo) + hi if wrap else (hi - lo)
        if slen < 0.4 or slen > 30:
            continue
        spans.append((lo, hi, wrap, ramp))
        split_s.extend([lo, hi])
        used.append(ramp)
        print(
            f"  [{tag}] depressed {ramp.get('name')} "
            f"[{lo:.2f}..{hi:.2f}]{' wrap' if wrap else ''} = {slen:.2f}"
        )

    if not spans:
        print(f"  [{tag}] no depressed — skip")
        return False, [], [], [], []

    new_pts = insert_stations(pts, svals, split_s)
    new_s = stations(new_pts)
    total_len = new_s[-1]

    def in_dep(sm):
        for a, b, wrap, _ in spans:
            if wrap:
                if sm >= a - 1e-6 or sm <= b + 1e-6:
                    return True
            elif a - 1e-6 <= sm <= b + 1e-6:
                return True
        return False

    ns = nd = nc = 0
    for i in range(len(new_pts) - 1):
        a, b = new_pts[i], new_pts[i + 1]
        L = dist(a, b)
        if L < 1e-4:
            continue
        sm = 0.5 * (new_s[i] + new_s[i + 1])
        if not near_ramp_face(sm, spans, total_len):
            continue
        if in_dep(sm):
            layer, label = LAYER_DEPRESSED, "Depressed"
            nd += 1
        elif window_is_corner(new_pts, new_s, i):
            layer, label = LAYER_CORNER, "Corner"
            nc += 1
        else:
            layer, label = LAYER_STRAIGHT, "Straight"
            ns += 1
        line = ms.AddLine(
            vpt(a[0], a[1], a[2] if len(a) > 2 else 0),
            vpt(b[0], b[1], b[2] if len(b) > 2 else 0),
        )
        line.Layer = layer
        line.Color = 256
        totals[layer] += L
        counts[layer] += 1
        if L >= 1.5:
            print(f"    {L:6.3f} {label}")
    print(f"  [{tag}] Straight={ns} Depressed={nd} Corner={nc}")
    return True, spans, used, new_pts, new_s


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

    for h in OUTER_HANDLES:
        try:
            doc.HandleToObject(h)
        except Exception:
            print("ABORT missing outer", h)
            return

    ensure_layers(doc)
    print("cleared classified", clear_classified(doc))

    ms_flares = load_ms_flares(doc)
    print("existing MS flares", len(ms_flares))

    totals = {k: 0.0 for k in COLORS}
    counts = {k: 0 for k in COLORS}

    for isl in ISLANDS:
        name = isl["name"]
        print(f"\n======== {name} outer={isl['outer']} ========")
        outer_pts = sample_fl_vlax(doc, isl["outer"])
        print(f"  outer samples n={len(outer_pts)}")
        if len(outer_pts) < 5:
            continue
        inner = offset_toward(outer_pts, isl["center"], INNER_OFFSET)
        print(f"  inner offset n={len(inner)} L={stations(inner)[-1]:.2f}")

        ramps = []
        for rh in isl.get("prop_ramps") or []:
            try:
                ramps.append(load_prop_flares(doc, rh))
            except Exception as ex:
                print("  prop fail", rh, ex)
        if isl.get("ms_flare_hint") and ms_flares:
            near = []
            for fl in ms_flares:
                cx = sum(p[0] for p in fl) / len(fl)
                cy = sum(p[1] for p in fl) / len(fl)
                if math.hypot(cx - isl["center"][0], cy - isl["center"][1]) < 25:
                    near.append((cy, fl))
            near.sort(key=lambda t: t[0])
            # split into two ramp groups (lower / upper) of 2 flares each when possible
            groups = []
            if len(near) >= 4:
                groups = [near[:2], near[2:4]]
            elif len(near) >= 2:
                groups = [near[: len(near) // 2], near[len(near) // 2 :]]
            else:
                groups = [near] if near else []
            for gi, grp in enumerate(groups):
                fls = [fl for _, fl in grp]
                dws = []
                for fl in fls:
                    dws.extend(fl)
                ramps.append(
                    {
                        "handle": f"ms-nw-{gi}",
                        "name": f"NW-ms-{gi}",
                        "flares": fls,
                        "dws": dws,
                        "area": None,
                    }
                )
            print(f"  using {len(near)} MS flares in {len(groups)} groups")
        if isl.get("trail"):
            try:
                ramps.append(trail_dws(doc, isl["trail"]))
            except Exception as ex:
                print("  trail fail", ex)

        ok, spans, used, face_pts, face_s = classify_island(
            ms, inner, ramps, totals, counts, name
        )
        if not ok:
            continue

        # Add missing flares when requested and ramp has no flare solids
        if isl.get("need_flares"):
            n_del = delete_nearby_ms_flares(doc, isl["center"], radius=20.0)
            if n_del:
                print(f"  cleared {n_del} old auto-flares")
            for lo, hi, wrap, ramp in spans:
                if ramp.get("flares"):
                    continue
                print(f"  adding flares for {ramp.get('name')}")
                make_flares_at_span_ends(
                    ms, face_pts, face_s, lo, hi, wrap, isl["center"]
                )

    print("\n=== OUTER intact ===")
    for h in sorted(OUTER_HANDLES):
        try:
            doc.HandleToObject(h)
            print("OK", h)
        except Exception:
            print("MISSING", h)

    doc.Regen(1)
    print("\n===== LF =====")
    for lyr in (LAYER_STRAIGHT, LAYER_DEPRESSED, LAYER_CORNER):
        print(f"{lyr.split()[0]:9s} n={counts[lyr]} LF={totals[lyr]:.2f}")
    print(f"TOTAL {sum(totals.values()):.2f}")
    print("DONE")


if __name__ == "__main__":
    main()
