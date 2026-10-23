"""Deep-compare NW correct flares vs MidS pocket geometry (outer/inner/PRO-S)."""
from __future__ import annotations

import math
import os
import time
import pythoncom
from win32com.client import Dispatch

PTS = r"c:\repos\microstation-vba-project\scripts\_fl_curve_pts.txt"


def sample_fl(doc, handle, step=0.25):
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
        time.sleep(0.15)
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
        if not out or math.hypot(p[0] - out[-1][0], p[1] - out[-1][1]) > 1e-4:
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


def dpoly(px, py, poly):
    best, hit = 1e9, None
    for i in range(len(poly) - 1):
        ax, ay = poly[i]
        bx, by = poly[i + 1]
        dx, dy = bx - ax, by - ay
        L2 = dx * dx + dy * dy
        t = 0.0 if L2 < 1e-12 else max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / L2))
        qx, qy = ax + t * dx, ay + t * dy
        d = math.hypot(px - qx, py - qy)
        if d < best:
            best, hit = d, (qx, qy)
    return best, hit


def solids_near(doc, cx0, cy0, r=30):
    ms = doc.ModelSpace
    out = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbSolid" or e.Layer != "PROP-D-ROAD-RAMPS":
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if math.hypot(cx - cx0, cy - cy0) > r:
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
            out.append((e.Handle, cx, cy, u))
        except Exception:
            continue
    return out


def steel_chain(doc, cx0, cy0, r=35, layer_prefix="Depressed"):
    ms = doc.ModelSpace
    segs = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbLine":
                continue
            if not e.Layer.startswith(layer_prefix) and "Steel Faced" not in e.Layer:
                continue
            s, en = e.StartPoint, e.EndPoint
            mx = 0.5 * (float(s[0]) + float(en[0]))
            my = 0.5 * (float(s[1]) + float(en[1]))
            if math.hypot(mx - cx0, my - cy0) > r:
                continue
            segs.append(
                (
                    e.Layer.split()[0],
                    (float(s[0]), float(s[1])),
                    (float(en[0]), float(en[1])),
                )
            )
        except Exception:
            continue
    return segs


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    try:
        doc.SendCommand(chr(27) + chr(27))
    except Exception:
        pass
    time.sleep(0.2)

    # --- NW reference ---
    print("======== NW REFERENCE ========")
    outer_nw = sample_fl(doc, "13D80A")
    print("outer", len(outer_nw))
    # PRO-S near NW
    pro_nw = []
    for h in ["138A44", "13A7EA", "13C903"]:
        try:
            pro_nw.extend(entity_pts(doc, h))
        except Exception:
            pass
    print("pro_nw", len(pro_nw))
    # all steel near NW as inner proxy
    steel = steel_chain(doc, 1028470, 187478, 40)
    inner_pts = []
    for lyr, a, b in steel:
        inner_pts.append(a)
        inner_pts.append(b)
    print("steel segs", len(steel))

    for h, cx, cy, u in solids_near(doc, 1028470, 187478, 25):
        print(f"\nNW flare {h} c=({cx:.2f},{cy:.2f})")
        for j, p in enumerate(u):
            do, _ = dpoly(p[0], p[1], outer_nw)
            dp, _ = dpoly(p[0], p[1], pro_nw) if pro_nw else (None, None)
            # dist to any steel
            best_s, best_lyr = 1e9, None
            for lyr, a, b in steel:
                d, _ = dpoly(p[0], p[1], [a, b])
                if d < best_s:
                    best_s, best_lyr = d, lyr
            print(
                f"  v{j} ({p[0]:.3f},{p[1]:.3f}) "
                f"d_outer={do:.3f} d_pro={dp} d_steel={best_s:.3f}({best_lyr})"
            )
        # edge lengths
        edges = [
            math.hypot(u[(j + 1) % 3][0] - u[j][0], u[(j + 1) % 3][1] - u[j][1])
            for j in range(3)
        ]
        print("  edges", [round(e, 3) for e in edges])

    # --- MidS pocket ---
    print("\n======== MidS POCKET ========")
    outer_s = sample_fl(doc, "13D7BC")
    # also inner-ish FL still present?
    for h in ["13D7BF", "13D7C4", "13D7CC", "13D7E8"]:
        try:
            e = doc.HandleToObject(h)
            print(f"still have {h} {e.ObjectName}")
        except Exception:
            print(f"missing {h}")

    pro_s = []
    for h in ["13C226", "138B86", "138B6B", "13C115", "13C116", "138B64"]:
        try:
            pts = entity_pts(doc, h)
            print(f"PRO {h} n={len(pts)} first={pts[0]} last={pts[-1]}")
            pro_s.extend(pts)
        except Exception as ex:
            print(f"PRO {h} fail {ex}")

    steel_s = steel_chain(doc, 1028525, 187428, 40)
    print("MidS steel segs", len(steel_s))
    for lyr, a, b in steel_s[:8]:
        print(f"  {lyr} ({a[0]:.2f},{a[1]:.2f})-({b[0]:.2f},{b[1]:.2f})")

    for h, cx, cy, u in solids_near(doc, 1028525, 187428, 30):
        print(f"\nMidS flare {h} c=({cx:.2f},{cy:.2f})")
        for j, p in enumerate(u):
            do, _ = dpoly(p[0], p[1], outer_s)
            dp, hp = dpoly(p[0], p[1], pro_s)
            best_s, best_lyr = 1e9, None
            for lyr, a, b in steel_s:
                d, _ = dpoly(p[0], p[1], [a, b])
                if d < best_s:
                    best_s, best_lyr = d, lyr
            print(
                f"  v{j} ({p[0]:.3f},{p[1]:.3f}) "
                f"d_outer={do:.3f} d_pro={dp:.3f} d_steel={best_s:.3f}({best_lyr})"
            )

    # Find candidate triangle pockets at MidS: for points along outer near DWS,
    # project to inner steel and to PRO-S — that forms the correct 3-line triangle
    print("\n======== MidS pocket candidates along outer ========")
    # DWS hatch
    try:
        e = doc.HandleToObject("138B88")
        mn, mx = e.GetBoundingBox()
        print(
            f"DWS bbox ({float(mn[0]):.2f},{float(mn[1]):.2f})-"
            f"({float(mx[0]):.2f},{float(mx[1]):.2f})"
        )
    except Exception as ex:
        print("DWS", ex)

    # Build inner poly from depressed+corner+straight steel near MidS
    inner_poly = []
    for lyr, a, b in steel_s:
        inner_poly.append(a)
        inner_poly.append(b)

    # Sample outer near DWS y range 187419-187437
    for i, p in enumerate(outer_s):
        if not (187418 < p[1] < 187438 and 1028515 < p[0] < 1028535):
            continue
        if i % 3 != 0:
            continue
        di, hi = dpoly(p[0], p[1], inner_poly) if inner_poly else (None, None)
        dp, hp = dpoly(p[0], p[1], pro_s)
        print(
            f"  outer({p[0]:.2f},{p[1]:.2f}) d_inner={di} d_pro={dp:.2f} "
            f"pro_hit={hp}"
        )


if __name__ == "__main__":
    main()
