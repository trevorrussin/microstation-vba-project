"""Redo loc01 curb classification with fixed corner/depressed rules.

Depressed: flare-to-flare (includes flares).
Corner: between depressed spans on same curb, OR curved non-depressed runs.
Straight: outside all ramp spans and essentially linear.

Also attempts to explode nearby C-TOPO-FEAT feature lines for other islands.
"""
from __future__ import annotations

import math
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

RAMP_HANDLES = ["13CDD7", "13D60F", "13CDC0", "13D4BC", "13D72E"]

LAYER_STRAIGHT = "Straight Steel Faced Concrete Curb (18'' Deep)"
LAYER_DEPRESSED = "Depressed Steel Faced Concrete Curb (18'' Deep)"
LAYER_CORNER = "Corner Steel Faced Concrete Curb (18'' Deep)"
COLORS = {LAYER_STRAIGHT: 1, LAYER_DEPRESSED: 3, LAYER_CORNER: 5}
CURB_NEAR = 2.0
TURN_CORNER_DEG = 12.0
PATH_CHORD_RATIO = 1.015


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


def dedupe(pts, tol=1e-6):
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
            (cx + r * math.cos(a0 + (a1 - a0) * i / n), cy + r * math.sin(a0 + (a1 - a0) * i / n), cz)
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


def is_curved(pts):
    if len(pts) < 3:
        return False
    turn = 0.0
    for i in range(1, len(pts) - 1):
        turn += ang_diff(heading(pts[i - 1], pts[i]), heading(pts[i], pts[i + 1]))
    path = sum(dist(pts[i], pts[i + 1]) for i in range(len(pts) - 1))
    chord = dist(pts[0], pts[-1])
    if path < 0.3:
        return abs(math.degrees(turn)) > 5
    return (path / max(chord, 1e-6) > PATH_CHORD_RATIO) or (
        abs(math.degrees(turn)) >= TURN_CORNER_DEG
    )


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
            flares.append(uniq)
        if e.ObjectName == "AcDbHatch" and "WARN" in e.Layer.upper():
            area = float(e.Area)
            mn, mx = e.GetBoundingBox()
            if abs(float(mn[0])) < 50 and abs(float(mn[1])) < 50:
                dws_pts = [
                    xform(br, float(mn[0]), float(mn[1])),
                    xform(br, float(mx[0]), float(mn[1])),
                    xform(br, float(mx[0]), float(mx[1])),
                    xform(br, float(mn[0]), float(mx[1])),
                ]
            else:
                dws_pts = [
                    (float(mn[0]), float(mn[1])),
                    (float(mx[0]), float(mn[1])),
                    (float(mx[0]), float(mx[1])),
                    (float(mn[0]), float(mx[1])),
                ]
    return {"handle": handle, "name": br.Name, "flares": flares, "dws": dws_pts, "area": area}


def depressed_span(pts, svals, ramp):
    curb_s = []
    for flare in ramp["flares"]:
        for p in flare:
            s, d, _ = project_station(pts, svals, p[0], p[1])
            if d <= CURB_NEAR:
                curb_s.append(s)
    for p in ramp["dws"]:
        s, d, _ = project_station(pts, svals, p[0], p[1])
        if d <= CURB_NEAR + 0.5:
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


def classify_and_place(ms, pts, ramps, totals, counts, tag):
    svals = stations(pts)
    spans = []
    split_s = []
    for ramp in ramps:
        span = depressed_span(pts, svals, ramp)
        if span and span[1] - span[0] >= 0.5:
            spans.append(span)
            split_s.extend(span)
            print(f"  [{tag}] ramp depressed [{span[0]:.2f}..{span[1]:.2f}] = {span[1]-span[0]:.2f} ft")
    if not spans and tag != "force":
        print(f"  [{tag}] no depressed span — skip")
        return False

    new_pts = insert_stations(pts, svals, split_s) if split_s else pts
    new_s = stations(new_pts)
    s_env = (min(s[0] for s in spans), max(s[1] for s in spans)) if spans else (0, 0)

    def in_dep(sm):
        return any(a - 1e-6 <= sm <= b + 1e-6 for a, b in spans)

    # elementary segments
    elems = []
    for i in range(len(new_pts) - 1):
        a, b = new_pts[i], new_pts[i + 1]
        L = dist(a, b)
        if L < 1e-4:
            continue
        sm = 0.5 * (new_s[i] + new_s[i + 1])
        elems.append((a, b, L, sm))

    # classify with run-based corner/straight outside depressed
    i = 0
    while i < len(elems):
        a, b, L, sm = elems[i]
        if in_dep(sm):
            layer = LAYER_DEPRESSED
            ms.AddLine(vpt(*a), vpt(*b)).Layer = layer
            # set bylayer color via layer; force color 256
            # (AddLine returns entity — set props)
            # re-get last? easier create then set
            i += 1
            # fix: need handle to set color — redo properly below
            totals[layer] += L
            counts[layer] += 1
            continue
        # gather contiguous non-depressed run
        run = [elems[i]]
        j = i + 1
        while j < len(elems) and not in_dep(elems[j][3]):
            run.append(elems[j])
            j += 1
        run_pts = [run[0][0]] + [r[1] for r in run]
        smid = 0.5 * (run[0][3] + run[-1][3])
        # Between ramp envelope but not depressed => Corner (island corner)
        if spans and s_env[0] < smid < s_env[1]:
            layer = LAYER_CORNER
        elif is_curved(run_pts):
            layer = LAYER_CORNER
        else:
            layer = LAYER_STRAIGHT
        for a, b, L, sm in run:
            line = ms.AddLine(vpt(*a), vpt(*b))
            line.Layer = layer
            line.Color = 256
            totals[layer] += L
            counts[layer] += 1
            print(f"    {L:6.3f} {layer.split()[0]:9s}")
        i = j

    # Re-place depressed properly with color (first loop was incomplete)
    # Actually first loop added lines without storing — BUG.
    # Rewrite cleanly:
    return "NEED_REWRITE"


def place_segments(ms, elems_with_layer, totals, counts):
    for a, b, L, layer in elems_with_layer:
        line = ms.AddLine(vpt(*a), vpt(*b))
        line.Layer = layer
        line.Color = 256
        totals[layer] += L
        counts[layer] += 1
        print(f"    {L:6.3f} {layer.split()[0]:9s}")


def classify_clean(ms, pts, ramps, totals, counts, tag):
    svals = stations(pts)
    spans = []
    split_s = []
    for ramp in ramps:
        span = depressed_span(pts, svals, ramp)
        if span and span[1] - span[0] >= 0.5:
            spans.append(span)
            split_s.extend(span)
            print(f"  [{tag}] depressed [{span[0]:.2f}..{span[1]:.2f}] len={span[1]-span[0]:.2f}")
    if not spans:
        print(f"  [{tag}] no ramp span — skip")
        return False

    new_pts = insert_stations(pts, svals, split_s)
    new_s = stations(new_pts)
    s_env = (min(s[0] for s in spans), max(s[1] for s in spans))

    def in_dep(sm):
        return any(a - 1e-6 <= sm <= b + 1e-6 for a, b in spans)

    elems = []
    for i in range(len(new_pts) - 1):
        a, b = new_pts[i], new_pts[i + 1]
        L = dist(a, b)
        if L < 1e-4:
            continue
        elems.append((a, b, L, 0.5 * (new_s[i] + new_s[i + 1])))

    out = []
    i = 0
    while i < len(elems):
        a, b, L, sm = elems[i]
        if in_dep(sm):
            out.append((a, b, L, LAYER_DEPRESSED))
            print(f"    {L:6.3f} Depressed")
            i += 1
            continue
        run = [elems[i]]
        j = i + 1
        while j < len(elems) and not in_dep(elems[j][3]):
            run.append(elems[j])
            j += 1
        run_pts = [run[0][0]] + [r[1] for r in run]
        smid = 0.5 * (run[0][3] + run[-1][3])
        if s_env[0] < smid < s_env[1]:
            layer = LAYER_CORNER
        elif is_curved(run_pts):
            layer = LAYER_CORNER
        else:
            layer = LAYER_STRAIGHT
        for a, b, L, sm in run:
            out.append((a, b, L, layer))
            print(f"    {L:6.3f} {layer.split()[0]:9s}")
        i = j

    place_segments(ms, out, totals, counts)
    return True


def try_explode_fl(doc, handle):
    """Copy+explode feature line via LISP; return new 3dpoly handles."""
    before = set()
    ms = doc.ModelSpace
    for i in range(ms.Count):
        try:
            before.add(ms.Item(i).Handle)
        except Exception:
            pass
    lisp = (
        f'(progn (vl-load-com) (setq e (handent "{handle}")) '
        f'(if e (progn (command "_.COPY" e "" "0,0,0" "0,0,0") '
        f'(setq c (entlast)) (command "_.EXPLODE" c) (princ "OK")) (princ "NO")) (princ)) '
    )
    try:
        doc.SendCommand(lisp)
        time.sleep(1.2)
    except Exception as ex:
        print("SendCommand fail", ex)
        return []
    new = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Handle in before:
                continue
            if e.ObjectName in ("AcDb3dPolyline", "AcDbPolyline", "AcDbLine", "AcDbArc"):
                new.append(e)
        except Exception:
            continue
    return new


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    print("file", doc.FullName)
    # cancel any pending
    try:
        doc.SendCommand(chr(27))
    except Exception:
        pass
    ensure_layers(doc)
    print("cleared", clear_classified(ms, doc))

    ramps = [ramp_flare_data(doc, h) for h in RAMP_HANDLES]
    print("ramps", len(ramps))

    totals = {k: 0.0 for k in COLORS}
    counts = {k: 0 for k in COLORS}

    # 1) Top-left 3d poly if still present; else skip
    sources = []
    for h in ["13D986"]:
        try:
            sources.append(("topo", doc.HandleToObject(h)))
        except Exception:
            print("13D986 missing — will use existing classified redo from FL explode")

    # 2) For each ramp, find nearest feature line and try explode
    fl_candidates = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer == "C-TOPO-FEAT" and e.ObjectName == "AeccDbFeatureLine":
                try:
                    mn, mx = e.GetBoundingBox()
                except Exception:
                    continue
                fl_candidates.append(
                    (
                        e.Handle,
                        (float(mn[0]) + float(mx[0])) / 2,
                        (float(mn[1]) + float(mx[1])) / 2,
                        float(mx[0]) - float(mn[0]),
                        float(mx[1]) - float(mn[1]),
                    )
                )
        except Exception:
            continue

    exploded = []
    used_fl = set()
    for ramp in ramps:
        # ramp center from flares
        xs = [p[0] for fl in ramp["flares"] for p in fl]
        ys = [p[1] for fl in ramp["flares"] for p in fl]
        rcx, rcy = sum(xs) / len(xs), sum(ys) / len(ys)
        # nearest small-ish FL
        cands = sorted(
            fl_candidates,
            key=lambda t: math.hypot(t[1] - rcx, t[2] - rcy),
        )
        for h, cx, cy, w, hgt in cands[:4]:
            if h in used_fl:
                continue
            if math.hypot(cx - rcx, cy - rcy) > 25:
                continue
            # prefer FL roughly ramp-sized
            if max(w, hgt) > 40:
                continue
            print(f"explode FL {h} for ramp {ramp['handle']} dist={math.hypot(cx-rcx,cy-rcy):.1f}")
            newents = try_explode_fl(doc, h)
            print(f"  new ents={len(newents)}")
            for ne in newents:
                exploded.append(ne)
            used_fl.add(h)
            break

    # Process topo sources
    for kind, e in sources:
        pts = poly_pts(e)
        print(f"\n=== {kind} {e.Handle} n={len(pts)} ===")
        ok = classify_clean(ms, pts, ramps, totals, counts, e.Handle)
        if ok and e.ObjectName == "AcDb3dPolyline":
            h = e.Handle
            try:
                e.Delete()
                print("deleted", h)
            except Exception as ex:
                print("del", ex)

    # Process exploded FL geometry — merge connected pieces per cluster
    for ne in exploded:
        try:
            pts = poly_pts(ne)
        except Exception:
            continue
        if len(pts) < 2:
            continue
        print(f"\n=== exploded {ne.Handle} {ne.ObjectName} n={len(pts)} ===")
        ok = classify_clean(ms, pts, ramps, totals, counts, ne.Handle)
        if ok:
            try:
                ne.Delete()
            except Exception:
                pass

    # Cleanup top-left FLs if still there
    for h in ["13D822", "13D824"]:
        try:
            doc.HandleToObject(h).Delete()
            print("deleted FL", h)
        except Exception:
            pass

    doc.Regen(1)
    print("\n===== QUANTITIES (LF) =====")
    for lyr in (LAYER_STRAIGHT, LAYER_DEPRESSED, LAYER_CORNER):
        print(f"{lyr}\n  n={counts[lyr]}  LF={totals[lyr]:.2f}")
    print(f"TOTAL LF={sum(totals.values()):.2f}")
    print("\n===== DWS SF =====")
    for r in ramps:
        print(f"  {r['name']}: {r['area']:.2f}" if r["area"] else f"  {r['name']}: n/a")
    print(f"TOTAL DWS SF={sum(r['area'] for r in ramps if r['area']):.2f}")


if __name__ == "__main__":
    main()
