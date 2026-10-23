"""Loc01 curb classify v4 — one full curb path per island.

Rules:
  Depressed = flare-to-flare (includes flares) per ped ramp
  Corner    = larger curvature detected from small segment turns
              (path/chord + cumulative heading in a station window)
  Straight  = essentially colinear outside depressed

Skips short ramp-face feature lines (wholly inside one depressed span).
"""
from __future__ import annotations

import math
import os
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

# PROP-D anonymous ramps with flare solids
RAMP_HANDLES = ["13CDD7", "13D60F", "13CDC0", "13D4BC", "13D72E"]

# One primary source per island (full curb, not short face FLs)
# NW: explodable 3dpoly; others: long AECC FL sampled via vlax-curve
SOURCES = [
    ("3dpoly", "13D986"),  # NW U24+U28
    ("fl", "13D25C"),  # N U25
    ("fl", "13D7DC"),  # S U27+U29
    ("fl", "13CEC7"),  # E PED area (if spans exist)
]

LAYER_STRAIGHT = "Straight Steel Faced Concrete Curb (18'' Deep)"
LAYER_DEPRESSED = "Depressed Steel Faced Concrete Curb (18'' Deep)"
LAYER_CORNER = "Corner Steel Faced Concrete Curb (18'' Deep)"
COLORS = {LAYER_STRAIGHT: 1, LAYER_DEPRESSED: 3, LAYER_CORNER: 5}

CURB_NEAR = 2.5
CURVE_WIN_FT = 5.0
PATH_CHORD_CURVE = 1.006
TURN_CURVE_DEG = 6.0
KINK_DEG = 4.0
# Reject sources that are almost entirely one depressed span (short face FLs)
MAX_SINGLE_SPAN_FRAC = 0.85
MIN_SOURCE_LEN = 20.0

PTS_FILE = r"c:\repos\microstation-vba-project\scripts\_fl_curve_pts.txt"
# FLs that duplicate the NW 3dpoly path — delete after NW classified
NW_DUP_FL = ["13D822", "13D824", "13D80A"]


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
    best_d, best_s = 1e9, 0.0
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
    return best_s, best_d


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
    path = svals[hi] - svals[lo]
    if path < 0.15:
        return False
    chord = dist(pts[lo], pts[hi])
    ratio = path / max(chord, 1e-9)
    turn = 0.0
    max_kink = 0.0
    for k in range(lo + 1, hi):
        kdeg = abs(
            math.degrees(
                ang_diff(heading(pts[k - 1], pts[k]), heading(pts[k], pts[k + 1]))
            )
        )
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
            if abs(float(mn[0])) < 50 and abs(float(mn[1])) < 50:
                dws_pts = [xform(br, c[0], c[1])[:2] for c in corners]
            else:
                dws_pts = corners
    return {"handle": handle, "name": br.Name, "flares": flares, "dws": dws_pts, "area": area}


def _near_stations(pts, svals, ramp):
    """Stations of flare+DWS verts within CURB_NEAR of this curb."""
    out = []
    for flare in ramp["flares"]:
        for p in flare:
            s, d = project_station(pts, svals, p[0], p[1])
            if d <= CURB_NEAR:
                out.append(s)
    for p in ramp["dws"]:
        s, d = project_station(pts, svals, p[0], p[1])
        if d <= CURB_NEAR + 0.75:
            out.append(s)
    return out


def depressed_span(pts, svals, ramp):
    """Flare-to-flare near this ramp.

    Returns (lo, hi, wrap) where wrap=True means depressed if
    s >= lo OR s <= hi (ramp sits on an open-polyline seam).
    """
    L = svals[-1]
    curb_s = _near_stations(pts, svals, ramp)
    if len(curb_s) < 2:
        return None

    start_pts = [s for s in curb_s if s <= 20.0]
    end_pts = [s for s in curb_s if s >= L - 20.0]
    mid_pts = [s for s in curb_s if 20.0 < s < L - 20.0]

    # Seam wrap: hits near both ends, little/nothing in middle of those hits
    if start_pts and end_pts and (not mid_pts or (max(mid_pts) - min(mid_pts) > 30)):
        # confirm DWS also wraps (same ramp at seam), not opposite-side noise
        dws_s = []
        for p in ramp["dws"]:
            s, d = project_station(pts, svals, p[0], p[1])
            if d <= CURB_NEAR + 0.75:
                dws_s.append(s)
        dws_start = any(s <= 20 for s in dws_s)
        dws_end = any(s >= L - 20 for s in dws_s)
        if dws_start and dws_end:
            lo = min(end_pts)  # from here to end
            hi = max(start_pts)  # and start to here
            wrap_len = (L - lo) + hi
            if 0.4 <= wrap_len <= 25.0:
                return (lo, hi, True)

    # Linear cluster: prefer points near DWS mid (circular-safe mid via unwrap)
    dws_s = []
    for p in ramp["dws"]:
        s, d = project_station(pts, svals, p[0], p[1])
        if d <= CURB_NEAR + 0.75:
            dws_s.append(s)
    if dws_s:
        # if DWS itself wraps, handled above; else use mean of non-wrapping DWS
        if max(dws_s) - min(dws_s) <= 25.0:
            dws_mid = sum(dws_s) / len(dws_s)
            SPAN_WIN = 14.0
            tight = [s for s in curb_s if abs(s - dws_mid) <= SPAN_WIN]
            if len(tight) >= 2:
                lo, hi = min(tight), max(tight)
                if 0.4 <= hi - lo <= 25.0:
                    return (lo, hi, False)

    # Fallback: densest contiguous cluster (gap < 8 ft), width <= 25
    curb_s = sorted(curb_s)
    best = None
    i = 0
    while i < len(curb_s):
        j = i
        while j + 1 < len(curb_s) and curb_s[j + 1] - curb_s[j] < 8.0:
            j += 1
        cluster = curb_s[i : j + 1]
        if len(cluster) >= 2:
            width = cluster[-1] - cluster[0]
            if width <= 25.0 and (best is None or len(cluster) > best[0]):
                best = (len(cluster), cluster[0], cluster[-1])
        i = j + 1
    if best is None:
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


def sample_fl_vlax(doc, handle, step=0.30):
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
        print("  SendCommand fail", ex)
        return []
    for _ in range(50):
        time.sleep(0.2)
        if os.path.exists(PTS_FILE) and os.path.getsize(PTS_FILE) > 10:
            break
    else:
        print(f"  sample timeout {handle}")
        return []
    pts = []
    with open(PTS_FILE, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line or "," not in line:
                continue
            parts = line.split(",")
            pts.append(
                (
                    float(parts[0]),
                    float(parts[1]),
                    float(parts[2]) if len(parts) > 2 else 0.0,
                )
            )
    return dedupe(pts, tol=1e-4)


def classify_clean(ms, pts, ramps, totals, counts, tag, claimed):
    svals = stations(pts)
    total_len = svals[-1]
    if total_len < MIN_SOURCE_LEN:
        print(f"  [{tag}] too short L={total_len:.2f} — skip")
        return False, []

    spans = []  # (lo, hi, wrap, handle)
    split_s = []
    used_ramps = []
    for ramp in ramps:
        if ramp["handle"] in claimed:
            continue
        span = depressed_span(pts, svals, ramp)
        if not span:
            continue
        lo, hi, wrap = span
        if wrap:
            slen = (total_len - lo) + hi
        else:
            slen = hi - lo
        if slen < 0.4:
            continue
        spans.append((lo, hi, wrap, ramp["handle"]))
        split_s.extend([lo, hi])
        used_ramps.append(ramp["handle"])
        wtag = " wrap" if wrap else ""
        print(
            f"  [{tag}] depressed {ramp['handle']} "
            f"[{lo:.2f}..{hi:.2f}]{wtag} = {slen:.2f} ft"
        )

    if not spans:
        print(f"  [{tag}] no flare-to-flare span — skip")
        return False, []

    # Reject short face FL: one non-wrap span covers almost entire source
    if len(spans) == 1:
        a, b, wrap, _ = spans[0]
        slen = (total_len - a) + b if wrap else (b - a)
        if (not wrap) and slen / total_len >= MAX_SINGLE_SPAN_FRAC:
            print(
                f"  [{tag}] reject face-FL (span {slen/total_len:.0%} of L={total_len:.1f})"
            )
            return False, []

    new_pts = insert_stations(pts, svals, split_s)
    new_s = stations(new_pts)

    def in_dep(sm):
        for a, b, wrap, _ in spans:
            if wrap:
                if sm >= a - 1e-6 or sm <= b + 1e-6:
                    return True
            elif a - 1e-6 <= sm <= b + 1e-6:
                return True
        return False

    for i in range(len(new_pts) - 1):
        a, b = new_pts[i], new_pts[i + 1]
        L = dist(a, b)
        if L < 1e-4:
            continue
        sm = 0.5 * (new_s[i] + new_s[i + 1])
        if in_dep(sm):
            layer, label = LAYER_DEPRESSED, "Depressed"
        elif window_is_curved(new_pts, new_s, i):
            layer, label = LAYER_CORNER, "Corner"
        else:
            layer, label = LAYER_STRAIGHT, "Straight"
        line = ms.AddLine(vpt(*a), vpt(*b))
        line.Layer = layer
        line.Color = 256
        totals[layer] += L
        counts[layer] += 1
        print(f"    {L:6.3f} {label}")

    return True, used_ramps


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
        r = ramp_flare_data(doc, h)
        ramps.append(r)
        print(f"ramp {h} {r['name']} flares={len(r['flares'])} area={r['area']:.2f}")

    totals = {k: 0.0 for k in COLORS}
    counts = {k: 0 for k in COLORS}
    claimed = set()

    for kind, h in SOURCES:
        try:
            e = doc.HandleToObject(h)
        except Exception:
            print(f"\n=== missing {h} ===")
            continue
        if kind == "3dpoly":
            pts = poly_pts(e)
        else:
            print(f"\n=== sample FL {h} ===")
            pts = sample_fl_vlax(doc, h)
            print(f"  n={len(pts)}")
            if len(pts) < 3:
                continue
        svals = stations(pts)
        print(f"\n=== {kind} {h} n={len(pts)} L={svals[-1]:.2f} ===")
        ok, used = classify_clean(ms, pts, ramps, totals, counts, h, claimed)
        if not ok:
            continue
        claimed.update(used)
        try:
            e.Delete()
            print("deleted source", h)
        except Exception as ex:
            print("del fail", ex)

    # Remove NW duplicate FLs (short faces / parallel) if still present
    for h in NW_DUP_FL:
        try:
            doc.HandleToObject(h).Delete()
            print("deleted dup FL", h)
        except Exception:
            pass

    doc.Regen(1)
    print("\n===== QUANTITIES (LF) =====")
    for lyr in (LAYER_STRAIGHT, LAYER_DEPRESSED, LAYER_CORNER):
        print(f"{lyr}\n  n={counts[lyr]}  LF={totals[lyr]:.2f}")
    print(f"TOTAL LF={sum(totals.values()):.2f}")
    print("claimed ramps", sorted(claimed))
    print("\n===== DWS SF =====")
    for r in ramps:
        print(f"  {r['name']} ({r['handle']}): {r['area']:.2f}")
    print(f"TOTAL DWS SF={sum(r['area'] for r in ramps):.2f}")
    print("DONE")


if __name__ == "__main__":
    main()
