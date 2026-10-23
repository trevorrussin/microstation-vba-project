"""Loc01 OTHER islands: NW rules, layers only, no flares.

WHERE EACH LAYER GOES (engineer / NW authority — always):
  Depressed — along the INNER curb from the outer edge of one yellow flare
              to the outer edge of the other flare on the SAME ped ramp
              (includes the flares and the DWS between them). Split verts
              at those stations so the layer change lands on the flare edge.
  Corner    — NOT inside a depressed span, AND the curb is a curve made of
              short tessellation chords (larger arc: path/chord + heading
              turn in a ~5 ft window of SHORT segs only).
  Straight  — NOT inside a depressed span, AND essentially colinear —
              including LONG elementary segments (>= ~2.5 ft). Long segs
              are never treated as arc chords (that is how NW gets 11–14 ft
              red runs).

NO flare drawing. Leave NW steel untouched. INNER only; hide sources.
"""
from __future__ import annotations

import math
import sys
import time
import pythoncom
from win32com.client import Dispatch

sys.path.insert(0, r"c:\repos\microstation-vba-project\scripts")
import _acad_reclassify_loc01_v7_keep as v7

SOURCES = [
    ("fl", "13D25E", "N"),
    ("fl", "13D7E8", "SW"),
    ("fl", "13CED2", "E"),
    ("fl", "13D2EE", "MidN"),
    ("fl", "13D7BF", "MidS"),
]
SIBLINGS = {
    "N": ["13D263", "13D265", "13D282"],
    "SW": ["13D7DB", "13D7E0"],
    "E": ["13CED4"],
    "MidN": ["13D425", "13D42D"],
    "MidS": ["13D7C4", "13D7CC"],
}
NW_CENTER = (1028470.0, 187478.0)
NW_KEEP_R = 28.0  # tighter — MidN was leaking into NW keep at 32

# Match NW/v5: long chords = Straight; Corner only from short arc chords
LONG_SEG_FT = 2.5
CURVE_WIN_FT = 5.0
PATH_CHORD_CURVE = 1.006
TURN_CURVE_DEG = 6.0
KINK_DEG = 4.0
RDP_EPS = 0.10  # ft — collapse colinear FL samples into long Straight runs


def clear_steel_except_nw(doc):
    ms = doc.ModelSpace
    to_del = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer not in v7.COLORS or e.ObjectName != "AcDbLine":
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if math.hypot(cx - NW_CENTER[0], cy - NW_CENTER[1]) <= NW_KEEP_R:
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


def unhide(doc, h):
    if h in v7.OUTER_HANDLES:
        return
    try:
        doc.HandleToObject(h).Visible = True
    except Exception:
        pass


def _perp_dist(p, a, b):
    ax, ay = a[0], a[1]
    bx, by = b[0], b[1]
    px, py = p[0], p[1]
    dx, dy = bx - ax, by - ay
    L2 = dx * dx + dy * dy
    if L2 < 1e-18:
        return math.hypot(px - ax, py - ay)
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / L2))
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


def rdp(pts, eps):
    """Ramer–Douglas–Peucker — restore long Straight runs like NW 3dpoly."""
    if len(pts) < 3:
        return list(pts)
    stack = [(0, len(pts) - 1)]
    keep = {0, len(pts) - 1}
    while stack:
        i0, i1 = stack.pop()
        a, b = pts[i0], pts[i1]
        best_i, best_d = None, 0.0
        for i in range(i0 + 1, i1):
            d = _perp_dist(pts[i], a, b)
            if d > best_d:
                best_d, best_i = d, i
        if best_i is not None and best_d > eps:
            keep.add(best_i)
            stack.append((i0, best_i))
            stack.append((best_i, i1))
    return [pts[i] for i in sorted(keep)]


def window_is_curved_nw(pts, svals, i_seg):
    """Same intent as NW/v5: long segs are Straight; Corner = short arc chords."""
    n = len(pts)
    if n < 3:
        return False
    seg_len = v7.dist(pts[i_seg], pts[i_seg + 1])
    if seg_len >= LONG_SEG_FT:
        return False  # long chord = Straight run (NW 11–14 ft red)
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
    short_idx = [k for k in range(lo, hi) if v7.dist(pts[k], pts[k + 1]) < LONG_SEG_FT]
    if len(short_idx) < 2:
        return False
    a0, a1 = short_idx[0], short_idx[-1] + 1
    path2 = svals[a1] - svals[a0]
    if path2 < 0.4:
        return False
    chord = v7.dist(pts[a0], pts[a1])
    ratio = path2 / max(chord, 1e-9)
    turn = 0.0
    max_kink = 0.0
    for k in range(a0 + 1, a1):
        kdeg = abs(
            math.degrees(
                v7.ang_diff(v7.heading(pts[k - 1], pts[k]), v7.heading(pts[k], pts[k + 1]))
            )
        )
        turn += kdeg
        max_kink = max(max_kink, kdeg)
    return (ratio >= PATH_CHORD_CURVE) or (turn >= TURN_CURVE_DEG) or (max_kink >= KINK_DEG)


def classify_nw_rules(ms, pts, ramps, totals, counts, tag, claimed):
    """Depressed flare-to-flare + split; then Corner vs Straight as NW."""
    svals = v7.stations(pts)
    total_len = svals[-1]
    if total_len < v7.MIN_SOURCE_LEN:
        print(f"  [{tag}] too short L={total_len:.2f} — skip")
        return False, []

    spans = []
    split_s = []
    used_ramps = []
    for ramp in ramps:
        if ramp["handle"] in claimed:
            continue
        span = v7.depressed_span(pts, svals, ramp)
        if not span:
            continue
        lo, hi, wrap = span
        slen = (total_len - lo) + hi if wrap else (hi - lo)
        if slen < 0.4:
            continue
        spans.append((lo, hi, wrap, ramp["handle"]))
        split_s.extend([lo, hi])
        used_ramps.append(ramp["handle"])
        print(
            f"  [{tag}] Depressed {ramp['name']} [{lo:.2f}..{hi:.2f}]"
            f"{' wrap' if wrap else ''} = {slen:.2f} ft  (flare-to-flare)"
        )

    if len(spans) == 1:
        a, b, wrap, _ = spans[0]
        slen = (total_len - a) + b if wrap else (b - a)
        if (not wrap) and slen / total_len >= v7.MAX_SINGLE_SPAN_FRAC:
            print(f"  [{tag}] reject face-FL")
            return False, []

    new_pts = v7.insert_stations(pts, svals, split_s) if split_s else list(pts)
    new_s = v7.stations(new_pts)
    print(f"  [{tag}] verts {len(pts)} -> {len(new_pts)} after flare-edge splits")

    def in_dep(sm):
        for a, b, wrap, _ in spans:
            if wrap:
                if sm >= a - 1e-6 or sm <= b + 1e-6:
                    return True
            elif a - 1e-6 <= sm <= b + 1e-6:
                return True
        return False

    n_s = n_d = n_c = 0
    for i in range(len(new_pts) - 1):
        a, b = new_pts[i], new_pts[i + 1]
        L = v7.dist(a, b)
        if L < 1e-4:
            continue
        sm = 0.5 * (new_s[i] + new_s[i + 1])
        if in_dep(sm):
            layer = v7.LAYER_DEPRESSED
            n_d += 1
        elif window_is_curved_nw(new_pts, new_s, i):
            layer = v7.LAYER_CORNER
            n_c += 1
        else:
            layer = v7.LAYER_STRAIGHT
            n_s += 1
        line = ms.AddLine(v7.vpt(*a), v7.vpt(*b))
        line.Layer = layer
        try:
            line.Color = 256
        except Exception:
            from win32com.client.dynamic import Dispatch as DynDispatch

            DynDispatch(line).Color = 256
        totals[layer] += L
        counts[layer] += 1

    print(f"  [{tag}] segs Straight={n_s} Depressed={n_d} Corner={n_c}")
    return True, used_ramps


def main():
    print(__doc__.split("NO flare")[0])
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

    v7.ensure_layers(doc)
    print("cleared non-NW steel", clear_steel_except_nw(doc))

    ramps = []
    for h in v7.RAMP_HANDLES:
        r = v7.ramp_flare_data(doc, h)
        ramps.append(r)
        print(f"  measure ramp {r['name']} flares={len(r['flares'])} (not drawing)")

    totals = {k: 0.0 for k in v7.COLORS}
    counts = {k: 0 for k in v7.COLORS}
    claimed = set()

    for kind, h, island in SOURCES:
        unhide(doc, h)
        try:
            e = doc.HandleToObject(h)
        except Exception:
            print(f"\n=== missing {island} {h} ===")
            continue
        print(f"\n=== {island} sample FL {h} dense then RDP eps={RDP_EPS} ===")
        dense = v7.sample_fl_vlax(doc, h, step=0.35)
        print(f"  dense n={len(dense)}")
        if len(dense) < 3:
            continue
        pts = rdp(dense, RDP_EPS)
        # ensure 3d tuples
        pts = [(p[0], p[1], p[2] if len(p) > 2 else 0.0) for p in pts]
        svals = v7.stations(pts)
        long_n = sum(
            1
            for i in range(len(pts) - 1)
            if v7.dist(pts[i], pts[i + 1]) >= LONG_SEG_FT
        )
        print(f"  after RDP n={len(pts)} L={svals[-1]:.2f} long_segs(>={LONG_SEG_FT})={long_n}")
        ok, used = classify_nw_rules(ms, pts, ramps, totals, counts, h, claimed)
        if not ok:
            continue
        claimed.update(used)
        v7.hide_entity(doc, h, f"{island} done")
        for sh in SIBLINGS.get(island, []):
            v7.hide_entity(doc, sh, f"{island} sibling")

    for oh in sorted(v7.OUTER_HANDLES):
        try:
            doc.HandleToObject(oh)
            print(f"OUTER OK {oh}")
        except Exception:
            print(f"OUTER MISSING {oh}")

    try:
        doc.Regen(1)
    except Exception:
        pass
    print("\n===== NON-NW TOTALS =====")
    for lyr in (v7.LAYER_STRAIGHT, v7.LAYER_DEPRESSED, v7.LAYER_CORNER):
        print(f"  {counts[lyr]:4d}  {totals[lyr]:8.2f} ft  {lyr}")
    print(f"  TOTAL   {sum(totals.values()):8.2f} ft")
    print("Running verify gate…")
    import subprocess

    rc = subprocess.call(
        [sys.executable, r"c:\repos\microstation-vba-project\scripts\_acad_verify_loc01_curb.py"]
    )
    if rc != 0:
        print("VERIFY FAILED — do not claim DONE; re-plan")
        raise SystemExit(rc)
    print("DONE — verify PASS")


if __name__ == "__main__":
    main()
