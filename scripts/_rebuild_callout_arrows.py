"""Rebuild 7 callout note-dims with arrow tips (stock look).

Recipe (pre-association-experiment working look):
  AssociationLock ON only for PLACE NOTE tip snap
  PLACE NOTE tip -> stub end
  PrimaryText='*', level/color like stock
  keep existing TextNode + stub; shared GraphicGroup
  AssociationLock OFF after; GraphicGroupLock ON
"""
from __future__ import annotations

import math
import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

import pythoncom
import ms_connect
from win32com.client import Dispatch

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
mr = app.ActiveModelReference
q = app.CadInputQueue
as_ = app.ActiveSettings
adf = app.ActiveDesignFile

# (gg, dim_id, tn_id, stub_id)
CALLOUTS = [
    (6114, 6107, 6117, 6138),
    (6115, 6108, 6120, 6139),
    (6116, 6109, 6123, 6140),
    (6117, 6110, 6126, 6141),
    (6118, 6111, 6129, 6142),
    (6119, 6112, 6132, 6143),
    (6120, 6113, 6135, 6144),
]

KNOWN_TARGETS = [4889, 5840, 3997, 3399, 5562, 5844, 4885]


def pt3(x, y, z=0.0):
    return app.Point3dFromXYZ(float(x), float(y), float(z))


def dist2(a, b):
    return (a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2


def line_ends(e):
    le = e.AsLineElement
    s, t = le.StartPoint, le.EndPoint
    return (float(s.X), float(s.Y)), (float(t.X), float(t.Y))


def find_target_near(tip_xy, prefer_ids=None):
    prefer_ids = prefer_ids or []
    best = None
    best_d = 1e18
    # Prefer known targets first
    for tid in prefer_ids:
        try:
            e = mr.GetElementByID64(tid)
        except Exception:
            continue
        r = e.Range
        cx = 0.5 * (float(r.Low.X) + float(r.High.X))
        cy = 0.5 * (float(r.Low.Y) + float(r.High.Y))
        d = dist2(tip_xy, (cx, cy))
        # also compare to range corners proximity via tip inside expanded range
        d2 = min(
            dist2(tip_xy, (float(r.Low.X), float(r.Low.Y))),
            dist2(tip_xy, (float(r.High.X), float(r.High.Y))),
            dist2(tip_xy, (cx, cy)),
        )
        d = min(d, d2)
        if d < best_d:
            best_d = d
            best = e
    if best is not None and best_d < 25.0**2:
        return best, math.sqrt(best_d)

    # Fallback: scan nearby graphical elements
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    while ee.MoveNext():
        e = ee.Current
        try:
            eid = int(e.ID64)
            if eid in {c[1] for c in CALLOUTS} | {c[2] for c in CALLOUTS} | {c[3] for c in CALLOUTS}:
                continue
            if int(e.Type) in (7, 17, 33):  # text / textnode / dim
                continue
            r = e.Range
            cx = 0.5 * (float(r.Low.X) + float(r.High.X))
            cy = 0.5 * (float(r.Low.Y) + float(r.High.Y))
            d = dist2(tip_xy, (cx, cy))
            if d < best_d and d < 40.0**2:
                best_d = d
                best = e
        except Exception:
            continue
    return best, (math.sqrt(best_d) if best is not None else None)


def newest_dim_id(before_ids: set[int]) -> int | None:
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    best = None
    while ee.MoveNext():
        e = ee.Current
        try:
            if int(e.Type) != 33:
                continue
            eid = int(e.ID64)
            if eid in before_ids:
                continue
            if best is None or eid > best:
                best = eid
        except Exception:
            pass
    return best


def all_dim_ids() -> set[int]:
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    out = set()
    while ee.MoveNext():
        e = ee.Current
        try:
            if int(e.Type) == 33:
                out.add(int(e.ID64))
        except Exception:
            pass
    return out


level = adf.Levels("TCPP_P_T")
style = adf.DimensionStyles("ny_Plan")

# Collect geometry before mutating
plans = []
for gg, dim_id, tn_id, stub_id in CALLOUTS:
    dim = mr.GetElementByID64(dim_id)
    d = dim.AsDimensionElement
    tip = (float(d.Point(1).X), float(d.Point(1).Y))
    text_end = (float(d.Point(2).X), float(d.Point(2).Y))
    stub = mr.GetElementByID64(stub_id)
    s0, s1 = line_ends(stub)
    # stub end should match text_end; pick stub endpoint closer to tip as tip-side? 
    # Point2 is stub end (text side). Confirm against stub.
    if dist2(text_end, s0) < dist2(text_end, s1):
        stub_end = s0
    else:
        stub_end = s1
    tgt, td = find_target_near(tip, KNOWN_TARGETS)
    plans.append({
        "gg": gg,
        "dim_id": dim_id,
        "tn_id": tn_id,
        "stub_id": stub_id,
        "tip": tip,
        "stub_end": stub_end,
        "target": tgt,
        "target_id": int(tgt.ID64) if tgt is not None else None,
        "target_dist": td,
    })
    print(
        f"plan GG{gg}: tip={tip} stub_end={stub_end} "
        f"target={int(tgt.ID64) if tgt else None} dist={td}"
    )

# Rebuild each note dim
as_.AssociationLockEnabled = True
as_.GraphicGroupLockEnabled = False  # avoid GG side-effects during recreate
print("Assoc ON for PLACE NOTE")

new_dims = []
for p in plans:
    before = all_dim_ids()
    # Delete old dim
    try:
        mr.RemoveElement(mr.GetElementByID64(p["dim_id"]))
        print("deleted dim", p["dim_id"])
    except Exception as ex:
        print("delete fail", p["dim_id"], ex)

    tip_pt = pt3(*p["tip"])
    end_pt = pt3(*p["stub_end"])
    tgt = p["target"]

    q.SendKeyin("NULL")
    time.sleep(0.05)
    q.SendKeyin("PLACE NOTE")
    time.sleep(0.1)
    if tgt is not None:
        q.SendTentativePoint(tip_pt, 1)
        time.sleep(0.05)
        q.SendDataPointForLocate(tgt, tip_pt)
        time.sleep(0.08)
    else:
        q.SendDataPoint(tip_pt, 1)
        time.sleep(0.05)
    q.SendDataPoint(end_pt, 1)
    time.sleep(0.25)
    q.SendReset()
    time.sleep(0.15)

    nid = newest_dim_id(before)
    if nid is None:
        print("FAIL no new dim for GG", p["gg"])
        new_dims.append(None)
        continue
    print("new dim", nid, "for GG", p["gg"])

    el = mr.GetElementByID64(nid)
    d = el.AsDimensionElement
    try:
        d.PrimaryText = "*"
    except Exception as ex:
        print(" PrimaryText", ex)
    try:
        d.DimensionStyle = style
    except Exception as ex:
        print(" style", ex)
    try:
        el.Level = level
    except Exception as ex:
        print(" level", ex)
    try:
        el.Color = 3
    except Exception:
        pass
    try:
        el.IsHidden = False
    except Exception:
        pass
    try:
        el.GraphicGroup = p["gg"]
    except Exception as ex:
        print(" gg dim", ex)
    el.Rewrite()

    # Re-assert TN + stub GG
    for eid in (p["tn_id"], p["stub_id"]):
        try:
            e = mr.GetElementByID64(eid)
            e.GraphicGroup = p["gg"]
            e.Rewrite()
        except Exception as ex:
            print(" gg member", eid, ex)

    new_dims.append(nid)
    # verify points
    d2 = mr.GetElementByID64(nid).AsDimensionElement
    print(
        "  pts",
        (round(float(d2.Point(1).X), 3), round(float(d2.Point(1).Y), 3)),
        (round(float(d2.Point(2).X), 3), round(float(d2.Point(2).Y), 3)),
        "prim",
        repr(d2.PrimaryText),
        "hidden",
        mr.GetElementByID64(nid).IsHidden,
    )

# Restore locks: stock move-as-unit, no association-lock experiments
as_.AssociationLockEnabled = False
as_.GraphicGroupLockEnabled = True
q.SendKeyin("NULL")
q.SendKeyin("SAVE DESIGN")
time.sleep(0.3)

print("=== result ===")
print("new_dims", new_dims)
print("Assoc", as_.AssociationLockEnabled, "GG", as_.GraphicGroupLockEnabled)
print("done")
