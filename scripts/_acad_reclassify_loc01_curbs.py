"""Reclassify Location 01 curb faces with corrected rules.

Rules:
  Depressed = continuous span from outer edge of one flare to outer edge of
              the other flare on the same ped ramp (includes flares + DWS).
  Corner    = non-depressed segments that are part of overall curved curb.
  Straight  = non-depressed segments that are essentially linear.

Splits insert vertices at projected flare-edge stations so layer changes
land on flare/DWS boundaries.

Sources: explodable C-TOPO-FEAT 3d polyline(s) and PRO-S-CURB-LINE
polylines/arcs near each PROP-D-ROAD-RAMPS block.
"""
from __future__ import annotations

import math
import pythoncom
from win32com.client import Dispatch, VARIANT

RAMP_HANDLES = ["13CDD7", "13D60F", "13CDC0", "13D4BC", "13D72E"]

LAYER_STRAIGHT = "Straight Steel Faced Concrete Curb (18'' Deep)"
LAYER_DEPRESSED = "Depressed Steel Faced Concrete Curb (18'' Deep)"
LAYER_CORNER = "Corner Steel Faced Concrete Curb (18'' Deep)"
COLORS = {LAYER_STRAIGHT: 1, LAYER_DEPRESSED: 3, LAYER_CORNER: 5}

CURB_NEAR = 2.0  # ft — vertex considered on curb if within this
TURN_CORNER_DEG = 12.0  # cumulative |turn| above this => corner
PATH_CHORD_RATIO = 1.015  # path/chord above this => corner


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


def heading(a, b):
    return math.atan2(b[1] - a[1], b[0] - a[0])


def ang_diff(a, b):
    d = b - a
    while d > math.pi:
        d -= 2 * math.pi
    while d < -math.pi:
        d += 2 * math.pi
    return d


def dist(a, b):
    return math.hypot(b[0] - a[0], b[1] - a[1])


def dedupe(pts, tol=1e-6):
    out = []
    for p in pts:
        if not out or dist(out[-1], p) > tol:
            out.append(p)
    return out


def poly_pts(e):
    if e.ObjectName == "AcDbArc":
        # densify arc to ~0.25 ft
        c = e.Center
        cx, cy, cz = float(c[0]), float(c[1]), float(c[2]) if len(list(c)) > 2 else 0.0
        r = float(e.Radius)
        a0 = float(e.StartAngle)
        a1 = float(e.EndAngle)
        if a1 < a0:
            a1 += 2 * math.pi
        arc_len = r * (a1 - a0)
        n = max(8, int(arc_len / 0.25) + 1)
        pts = []
        for i in range(n + 1):
            a = a0 + (a1 - a0) * i / n
            pts.append((cx + r * math.cos(a), cy + r * math.sin(a), cz))
        return pts
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
    """Closest point station along polyline."""
    best_d = 1e9
    best_s = 0.0
    best_pt = pts[0]
    for i in range(len(pts) - 1):
        ax, ay = pts[i][0], pts[i][1]
        bx, by = pts[i + 1][0], pts[i + 1][1]
        dx, dy = bx - ax, by - ay
        L2 = dx * dx + dy * dy
        if L2 < 1e-12:
            t = 0.0
        else:
            t = max(0.0, min(1.0, ((x - ax) * dx + (y - ay) * dy) / L2))
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
    """Return new pts list including interpolated points at new_s stations."""
    all_s = sorted(set([round(s, 5) for s in svals + list(new_s)]))
    # map back
    out = []
    for s in all_s:
        # prefer existing vertex if very close
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
        # single segment: treat as straight unless very short connector on arc
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
    flares = []
    dws_pts = []
    area = None
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
            # unique
            uniq = []
            for p in pts:
                if not any(math.hypot(p[0] - q[0], p[1] - q[1]) < 1e-5 for q in uniq):
                    uniq.append(p)
            flares.append(uniq)
        if e.ObjectName == "AcDbHatch" and "WARN" in e.Layer.upper():
            area = float(e.Area)
            try:
                mn, mx = e.GetBoundingBox()
                if abs(float(mn[0])) < 50 and abs(float(mn[1])) < 50:
                    corners = [
                        xform(br, float(mn[0]), float(mn[1])),
                        xform(br, float(mx[0]), float(mn[1])),
                        xform(br, float(mx[0]), float(mx[1])),
                        xform(br, float(mn[0]), float(mx[1])),
                    ]
                else:
                    corners = [
                        (float(mn[0]), float(mn[1])),
                        (float(mx[0]), float(mn[1])),
                        (float(mx[0]), float(mx[1])),
                        (float(mn[0]), float(mx[1])),
                    ]
                dws_pts = corners
            except Exception:
                pass
    return {"handle": handle, "name": br.Name, "flares": flares, "dws": dws_pts, "area": area}


def depressed_span_on_curb(pts, svals, ramp):
    """Return (s_lo, s_hi) for flare-to-flare depressed zone, or None."""
    curb_pts = []
    for flare in ramp["flares"]:
        for p in flare:
            s, d, _ = project_station(pts, svals, p[0], p[1])
            if d <= CURB_NEAR:
                curb_pts.append(s)
    for p in ramp["dws"]:
        s, d, _ = project_station(pts, svals, p[0], p[1])
        if d <= CURB_NEAR + 0.5:
            curb_pts.append(s)
    if len(curb_pts) < 2:
        return None
    return (min(curb_pts), max(curb_pts))


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


def clear_classified(doc, ms):
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


def find_curb_sources(doc, ms, ramps):
    """Return list of entities to classify: prefer C-TOPO-FEAT 3dpolys, else nearby curb lines/arcs."""
    sources = []
    # Always include known top-left 3d poly if present
    for h in ["13D986"]:
        try:
            e = doc.HandleToObject(h)
            sources.append(e)
        except Exception:
            pass

    # Near each ramp, collect PRO-S-CURB-LINE polys/arcs whose bbox intersects ramp union
    for ramp in ramps:
        xs, ys = [], []
        for fl in ramp["flares"]:
            for p in fl:
                xs.append(p[0])
                ys.append(p[1])
        if not xs:
            continue
        lx, hx, ly, hy = min(xs) - 8, max(xs) + 8, min(ys) - 8, max(ys) + 8
        for i in range(ms.Count):
            try:
                e = ms.Item(i)
                if e.Layer not in ("PRO-S-CURB-LINE", "C-TOPO-FEAT"):
                    continue
                if e.ObjectName not in (
                    "AcDbPolyline",
                    "AcDb3dPolyline",
                    "AcDbArc",
                    "AcDbLine",
                ):
                    continue
                mn, mx = e.GetBoundingBox()
                if float(mx[0]) < lx or float(mn[0]) > hx or float(mx[1]) < ly or float(mn[1]) > hy:
                    continue
                # skip tiny scratch
                w = float(mx[0]) - float(mn[0])
                h = float(mx[1]) - float(mn[1])
                if max(w, h) < 3:
                    continue
                if any(e.Handle == s.Handle for s in sources):
                    continue
                sources.append(e)
            except Exception:
                continue
    return sources


def classify_polyline(ms, pts, ramps, totals, counts, tag):
    if len(pts) < 2:
        return
    svals = stations(pts)
    # depressed spans that land on this curb
    spans = []
    split_s = []
    for ramp in ramps:
        span = depressed_span_on_curb(pts, svals, ramp)
        if span is None:
            continue
        # require meaningful length on this curb
        if span[1] - span[0] < 0.5:
            continue
        spans.append((span[0], span[1], ramp["handle"]))
        split_s.extend([span[0], span[1]])
        print(
            f"  [{tag}] ramp {ramp['handle']} depressed s=[{span[0]:.2f},{span[1]:.2f}] "
            f"len={span[1]-span[0]:.2f}"
        )

    # rebuild with split vertices
    new_pts = insert_stations(pts, svals, split_s) if split_s else pts
    new_s = stations(new_pts)

    def in_depressed(s_mid):
        for a, b, _ in spans:
            if a - 1e-6 <= s_mid <= b + 1e-6:
                return True
        return False

    # Build segment runs by class
    # First assign each elementary segment
    seg_classes = []
    for i in range(len(new_pts) - 1):
        a, b = new_pts[i], new_pts[i + 1]
        length = dist(a, b)
        if length < 1e-4:
            continue
        s_mid = 0.5 * (new_s[i] + new_s[i + 1])
        if in_depressed(s_mid):
            cls = LAYER_DEPRESSED
        else:
            cls = None  # decide curved/straight later in runs
        seg_classes.append((a, b, length, s_mid, cls, i))

    # Group consecutive non-depressed into runs; test curvature on each run
    i = 0
    classified_segs = []
    while i < len(seg_classes):
        a, b, length, s_mid, cls, idx = seg_classes[i]
        if cls == LAYER_DEPRESSED:
            classified_segs.append((a, b, length, LAYER_DEPRESSED))
            i += 1
            continue
        # gather run of non-depressed
        run = [seg_classes[i]]
        j = i + 1
        while j < len(seg_classes) and seg_classes[j][4] != LAYER_DEPRESSED:
            run.append(seg_classes[j])
            j += 1
        run_pts = [run[0][0]] + [r[1] for r in run]
        layer = LAYER_CORNER if is_curved(run_pts) else LAYER_STRAIGHT
        for a, b, length, s_mid, _, _ in run:
            classified_segs.append((a, b, length, layer))
        i = j

    for a, b, length, layer in classified_segs:
        line = ms.AddLine(vpt(a[0], a[1], a[2]), vpt(b[0], b[1], b[2]))
        line.Layer = layer
        line.Color = 256
        totals[layer] += length
        counts[layer] += 1
        print(f"    {length:6.3f}  {layer.split()[0]:9s}")

    return classified_segs


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    print("file", doc.FullName)
    ensure_layers(doc)
    nclr = clear_classified(doc, ms)
    print(f"cleared old classified: {nclr}")

    ramps = [ramp_flare_data(doc, h) for h in RAMP_HANDLES]
    print(f"ramps={len(ramps)}")
    for r in ramps:
        print(
            f"  {r['handle']} {r['name']} flares={len(r['flares'])} "
            f"dws_area={r['area']}"
        )

    sources = find_curb_sources(doc, ms, ramps)
    print(f"curb sources={len(sources)}")
    for e in sources:
        print(f"  {e.Handle} {e.ObjectName} lyr={e.Layer}")

    totals = {k: 0.0 for k in COLORS}
    counts = {k: 0 for k in COLORS}
    used_handles = set()

    for e in sources:
        if e.Handle in used_handles:
            continue
        pts = poly_pts(e)
        if len(pts) < 2:
            continue
        # Only process if at least one ramp projects onto it
        svals = stations(pts)
        hits = 0
        for ramp in ramps:
            span = depressed_span_on_curb(pts, svals, ramp)
            if span and span[1] - span[0] >= 0.5:
                hits += 1
        if hits == 0 and e.Layer != "C-TOPO-FEAT":
            # skip curb pieces with no ramp attachment (unless main topo)
            continue
        print(f"\n=== source {e.Handle} hits={hits} npts={len(pts)} ===")
        classify_polyline(ms, pts, ramps, totals, counts, e.Handle)
        used_handles.add(e.Handle)
        # delete source if it was the topo 3d poly we exploded
        if e.ObjectName == "AcDb3dPolyline" and e.Layer == "C-TOPO-FEAT":
            try:
                e.Delete()
                print(f"  deleted source 3dpoly {e.Handle}")
            except Exception as ex:
                print(f"  del fail {ex}")

    # Also delete matching feature lines on top-left that duplicate 13D986 path
    for h in ["13D822", "13D824"]:
        try:
            doc.HandleToObject(h).Delete()
            print("deleted FL", h)
        except Exception:
            pass

    doc.Regen(1)
    print("\n===== QUANTITIES (LF) =====")
    for lyr in (LAYER_STRAIGHT, LAYER_DEPRESSED, LAYER_CORNER):
        print(f"{lyr}")
        print(f"  n={counts[lyr]}  LF={totals[lyr]:.2f}")
    print(f"TOTAL LF={sum(totals.values()):.2f}")

    print("\n===== DWS AREAS (SF) =====")
    for r in ramps:
        print(f"  {r['handle']} {r['name']}: {r['area']:.2f} SF" if r["area"] else f"  {r['handle']}: n/a")
    areas = [r["area"] for r in ramps if r["area"]]
    print(f"TOTAL DWS SF={sum(areas):.2f}")


if __name__ == "__main__":
    main()
