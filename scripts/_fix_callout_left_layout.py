"""Reposition NEW conduit callouts to left-of-tip layout like stock/left side.

Stock / left-side pattern:
  text (RJ) WEST of tip → 5' stub EAST toward tip → note-dim tip
  tip must sit clear of text High.X (not under the text block)
"""
from __future__ import annotations

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

assert "fea_tsp" in adf.FullName.lower()

LEADER_GAP = 25.0  # ft between text right edge and tip
STUB_LEN = 5.0
STUB_DX0 = 1.054
STUB_DY0 = -1.498


def pt3(x, y, z=0.0):
    return app.Point3dFromXYZ(float(x), float(y), float(z))


def gg_members(gg: int):
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    tn = stub = dim = None
    while ee.MoveNext():
        e = ee.Current
        try:
            if int(e.GraphicGroup) != gg:
                continue
            if e.IsTextNodeElement:
                tn = e
            elif int(e.Type) == 3:
                stub = e
            elif int(e.Type) == 33:
                dim = e
        except Exception:
            pass
    return tn, stub, dim


def nearest_tci(xy):
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    best = None
    best_d = 1e18
    for lv_ok in (("TCI", "TCI_P"), ("UUO",)):
        ee = mr.Scan(o)
        while ee.MoveNext():
            e = ee.Current
            try:
                lv = e.Level.Name if e.Level else ""
                if lv not in lv_ok:
                    continue
                r = e.Range
                cx = 0.5 * (float(r.Low.X) + float(r.High.X))
                cy = 0.5 * (float(r.Low.Y) + float(r.High.Y))
                d = (cx - xy[0]) ** 2 + (cy - xy[1]) ** 2
                if d < best_d:
                    best_d = d
                    best = (cx, cy)
            except Exception:
                continue
        if best is not None and best_d < 80**2:
            return best
    return best or xy


level = adf.Levels("TCPP_P_T")
style = adf.DimensionStyles("ny_Plan")
as_.GraphicGroupLockEnabled = False
as_.AssociationLockEnabled = False

# Delete engineer/orphan copy GG 21 if present
print("=== cleanup GG 21 ===")
tn21, stub21, dim21 = gg_members(21)
for e in (tn21, stub21, dim21):
    if e is None:
        continue
    try:
        print("  del", int(e.ID64))
        mr.RemoveElement(e)
    except Exception as ex:
        print("  fail", ex)

print("=== reposition ===")
for gg in range(7201, 7213):
    tn, stub, dim = gg_members(gg)
    if not (tn and stub and dim):
        print(f"GG{gg} incomplete, skip")
        continue

    tn_id, stub_id, dim_id = int(tn.ID64), int(stub.ID64), int(dim.ID64)
    r = tn.Range
    text_w = float(r.High.X) - float(r.Low.X)
    text_h = float(r.High.Y) - float(r.Low.Y)
    de = dim.AsDimensionElement
    tip = (float(de.Point(1).X), float(de.Point(1).Y))

    # Prefer tip on conduit near current tip
    tip = nearest_tci(tip)

    # Need tip.X >= text High.X + small gap  →  origin (RJ right) = tip.X - LEADER_GAP
    # and text extends left by text_w, so Low.X = origin - text_w
    target_origin_x = tip[0] - LEADER_GAP
    # Keep Y near current text center (multi-line block sits above stub)
    cur_oy = float(tn.AsTextNodeElement.Origin.Y)
    cur_ox = float(tn.AsTextNodeElement.Origin.X)
    target_origin_y = cur_oy

    # If already clear (tip east of text by >= 15), skip move
    clear = tip[0] - float(r.High.X) >= 15.0
    if clear and abs(target_origin_x - cur_ox) < 5.0:
        print(f"GG{gg} already clear tip-HighX={tip[0]-float(r.High.X):.1f}")
        continue

    dx = target_origin_x - cur_ox
    dy = target_origin_y - cur_oy
    print(
        f"GG{gg} move ({dx:.1f},{dy:.1f}) origin->({target_origin_x:.1f},{target_origin_y:.1f}) "
        f"tip=({tip[0]:.1f},{tip[1]:.1f}) text_w={text_w:.1f}"
    )

    mv = pt3(dx, dy)
    for eid in (tn_id, stub_id, dim_id):
        e = mr.GetElementByID64(eid)
        e.Move(mv)
        e.Rewrite()

    # Rebuild stub: always tip_dir = +1 (east toward tip)
    ox = target_origin_x
    oy = target_origin_y
    stub_start = (ox + STUB_DX0, oy + STUB_DY0)
    stub_end = (stub_start[0] + STUB_LEN, stub_start[1])

    # Replace stub line
    try:
        mr.RemoveElement(mr.GetElementByID64(stub_id))
    except Exception:
        pass
    stub_el = app.CreateLineElement1(None, (pt3(*stub_start), pt3(*stub_end)))
    if isinstance(stub_el, tuple):
        stub_el = stub_el[0]
    stub_el.Level = level
    stub_el.Color = 3
    stub_el.LineWeight = 0
    stub_el.GraphicGroup = gg
    mr.AddElement(stub_el)
    stub_el.Rewrite()
    stub_id = int(stub_el.ID64)

    # Replace dim via PLACE NOTE
    try:
        mr.RemoveElement(mr.GetElementByID64(dim_id))
    except Exception:
        pass

    before = set()
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    while ee.MoveNext():
        try:
            if int(ee.Current.Type) == 33:
                before.add(int(ee.Current.ID64))
        except Exception:
            pass

    q.SendKeyin("NULL")
    time.sleep(0.04)
    q.SendKeyin("PLACE NOTE")
    time.sleep(0.08)
    q.SendDataPoint(pt3(*tip), 1)
    time.sleep(0.04)
    q.SendDataPoint(pt3(*stub_end), 1)
    time.sleep(0.2)
    q.SendReset()
    time.sleep(0.1)

    new_dim = None
    ee = mr.Scan(o)
    while ee.MoveNext():
        e = ee.Current
        try:
            if int(e.Type) != 33:
                continue
            eid = int(e.ID64)
            if eid not in before and (new_dim is None or eid > new_dim):
                new_dim = eid
        except Exception:
            pass
    if new_dim is None:
        print(f"  FAIL no dim GG{gg}")
        continue

    dim_el = mr.GetElementByID64(new_dim)
    de = dim_el.AsDimensionElement
    try:
        de.PrimaryText = "*"
    except Exception:
        pass
    try:
        de.DimensionStyle = style
    except Exception:
        pass
    try:
        dim_el.Level = level
        dim_el.Color = 3
        dim_el.IsHidden = False
    except Exception:
        pass
    dim_el.GraphicGroup = gg
    dim_el.Rewrite()

    # ensure pt order tip then stub_end
    de = mr.GetElementByID64(new_dim).AsDimensionElement
    p1 = (float(de.Point(1).X), float(de.Point(1).Y))
    p2 = (float(de.Point(2).X), float(de.Point(2).Y))
    d1 = (p1[0] - tip[0]) ** 2 + (p1[1] - tip[1]) ** 2
    d2 = (p2[0] - tip[0]) ** 2 + (p2[1] - tip[1]) ** 2
    if d2 < d1:
        while de.ReferencePointsCount > 0:
            de.DeletePoint(1)
        de.AddReferencePoint(mr, pt3(*tip))
        de.AddReferencePoint(mr, pt3(*stub_end))
        try:
            de.PrimaryText = "*"
        except Exception:
            pass
        mr.GetElementByID64(new_dim).Rewrite()

    print(f"  new stub={stub_id} dim={new_dim}")

as_.GraphicGroupLockEnabled = True
q.SendKeyin("SAVE DESIGN")
time.sleep(0.4)

print("=== clearance check ===")
for gg in range(7201, 7213):
    tn, stub, dim = gg_members(gg)
    if not (tn and dim):
        print(gg, "missing")
        continue
    r = tn.Range
    tip = (
        float(dim.AsDimensionElement.Point(1).X),
        float(dim.AsDimensionElement.Point(1).Y),
    )
    gap = tip[0] - float(r.High.X)
    ok = gap >= 10.0
    print(
        f"GG{gg} tip-HighX={gap:.1f} {'OK' if ok else 'BAD'} "
        f"text=[{float(r.Low.X):.1f},{float(r.High.X):.1f}] tip=({tip[0]:.1f},{tip[1]:.1f})"
    )

# overlap gate
o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
texts = []
while ee.MoveNext():
    e = ee.Current
    try:
        if not (e.IsTextElement or e.IsTextNodeElement):
            continue
        if e.Level and e.Level.Name not in ("TCPP_P_T", "TCSH_P_T"):
            continue
        r = e.Range
        texts.append(
            (
                int(e.ID64),
                (float(r.Low.X), float(r.Low.Y)),
                (float(r.High.X), float(r.High.Y)),
            )
        )
    except Exception:
        pass
clashes = []
PAD = 0.25
for i in range(len(texts)):
    for j in range(i + 1, len(texts)):
        a, b = texts[i], texts[j]
        if (
            a[1][0] - PAD <= b[2][0]
            and b[1][0] - PAD <= a[2][0]
            and a[1][1] - PAD <= b[2][1]
            and b[1][1] - PAD <= a[2][1]
        ):
            clashes.append((a[0], b[0]))
print(f"overlap {'PASS' if not clashes else 'FAIL'} clashes={clashes}")
print("DONE")
