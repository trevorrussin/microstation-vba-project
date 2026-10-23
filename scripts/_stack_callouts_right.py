"""Place NEW conduit callouts on the RIGHT (opposite of left APS column).

Mirror of stock left layout:
  tip on conduit (west) → note-dim → stub end ← 5' stub ← text (RJ, east)
  tip.X < text Low.X with clear gap (leader never under text)
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

# Right column: RJ origin (text right edge). Poles/TCI ~1029460–1029510.
COL_ORIGIN_X = 1029620.0
Y_TOP = 239960.0
Y_GAP = 4.0
STUB_LEN = 5.0
STUB_DX0 = 1.054  # stub offset from text edge toward tip
STUB_DY0 = -1.498
LEADER_GAP = 25.0  # min tip → text Low.X


def pt3(x, y, z=0.0):
    return app.Point3dFromXYZ(float(x), float(y), float(z))


def nearest_tci_west_of(x_max: float, y_prefer: float):
    """Pick TCI/TCI_P/UUO center west of x_max, nearest to y_prefer."""
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    best = None
    best_d = 1e18
    ee = mr.Scan(o)
    while ee.MoveNext():
        e = ee.Current
        try:
            lv = e.Level.Name if e.Level else ""
            if lv not in ("TCI", "TCI_P", "UUO"):
                continue
            r = e.Range
            cx = 0.5 * (float(r.Low.X) + float(r.High.X))
            cy = 0.5 * (float(r.Low.Y) + float(r.High.Y))
            if cx > x_max - 5.0:
                continue
            d = (cy - y_prefer) ** 2 + 0.15 * (cx - x_max) ** 2
            if d < best_d:
                best_d = d
                best = (cx, cy)
        except Exception:
            continue
    return best


def gg_parts(gg):
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


level = adf.Levels("TCPP_P_T")
style = adf.DimensionStyles("ny_Plan")
as_.GraphicGroupLockEnabled = False
as_.AssociationLockEnabled = False
q.SendKeyin("NULL")
time.sleep(0.05)

# Top→bottom by current Y
items = []
for gg in range(7201, 7213):
    tn, stub, dim = gg_parts(gg)
    if tn is None:
        print("missing", gg)
        continue
    oy = float(tn.AsTextNodeElement.Origin.Y)
    h = float(tn.Range.High.Y) - float(tn.Range.Low.Y)
    w = float(tn.Range.High.X) - float(tn.Range.Low.X)
    items.append((oy, gg, int(tn.ID64), h, w))
items.sort(key=lambda x: -x[0])

print("order", [(gg, round(h, 1), round(w, 1)) for _, gg, _, h, w in items])

cursor_y = Y_TOP
for _, gg, tn_id, text_h, text_w in items:
    tn, stub, dim = gg_parts(gg)
    if tn is None:
        continue

    ox = COL_ORIGIN_X
    oy = cursor_y
    cur = tn.AsTextNodeElement.Origin
    dx = ox - float(cur.X)
    dy = oy - float(cur.Y)
    print(f"GG{gg} -> ({ox:.1f},{oy:.1f}) d=({dx:.1f},{dy:.1f}) w={text_w:.1f}")

    mv = pt3(dx, dy)
    e = mr.GetElementByID64(tn_id)
    e.Move(mv)
    e.Rewrite()

    # refresh size after move
    tn2 = mr.GetElementByID64(tn_id)
    r = tn2.Range
    text_w = float(r.High.X) - float(r.Low.X)
    low_x = float(r.Low.X)  # should be ~ ox - text_w

    # tip west of text left edge
    tip = nearest_tci_west_of(low_x - LEADER_GAP, oy)
    if tip is None:
        tip = (low_x - LEADER_GAP - 40.0, oy)
    # force clearance
    if tip[0] > low_x - 15.0:
        tip = (low_x - LEADER_GAP, tip[1])

    # stub on LEFT edge of text, pointing WEST toward tip
    stub_start = (low_x - STUB_DX0, oy + STUB_DY0)
    stub_end = (stub_start[0] - STUB_LEN, stub_start[1])

    if stub is not None:
        try:
            mr.RemoveElement(stub)
        except Exception:
            pass
    if dim is not None:
        try:
            mr.RemoveElement(dim)
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
    time.sleep(0.03)
    q.SendKeyin("PLACE NOTE")
    time.sleep(0.06)
    q.SendDataPoint(pt3(*tip), 1)
    time.sleep(0.03)
    q.SendDataPoint(pt3(*stub_end), 1)
    time.sleep(0.18)
    q.SendReset()
    time.sleep(0.08)

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
        print("  FAIL dim")
        cursor_y = float(tn2.Range.Low.Y) - Y_GAP
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
    dim_el.Level = level
    dim_el.Color = 3
    dim_el.IsHidden = False
    dim_el.GraphicGroup = gg
    dim_el.Rewrite()

    # ensure tip is pt1
    de = mr.GetElementByID64(new_dim).AsDimensionElement
    p1 = (float(de.Point(1).X), float(de.Point(1).Y))
    p2 = (float(de.Point(2).X), float(de.Point(2).Y))
    if (p2[0] - tip[0]) ** 2 + (p2[1] - tip[1]) ** 2 < (p1[0] - tip[0]) ** 2 + (
        p1[1] - tip[1]
    ) ** 2:
        while de.ReferencePointsCount > 0:
            de.DeletePoint(1)
        de.AddReferencePoint(mr, pt3(*tip))
        de.AddReferencePoint(mr, pt3(*stub_end))
        try:
            de.PrimaryText = "*"
        except Exception:
            pass
        mr.GetElementByID64(new_dim).Rewrite()

    tn3 = mr.GetElementByID64(tn_id)
    gap = float(tn3.Range.Low.X) - tip[0]
    print(
        f"  stub={stub_id} dim={new_dim} tip=({tip[0]:.1f},{tip[1]:.1f}) "
        f"LowX-tip={gap:.1f} stub {stub_start[0]:.1f}->{stub_end[0]:.1f}"
    )
    cursor_y = float(tn3.Range.Low.Y) - Y_GAP

as_.GraphicGroupLockEnabled = True
q.SendKeyin("NULL")
q.SendKeyin("SAVE DESIGN")
time.sleep(0.4)

print("=== verify ===")
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

for gg in range(7201, 7213):
    tn, stub, dim = gg_parts(gg)
    if not (tn and stub and dim):
        print(gg, "INCOMPLETE")
        continue
    r = tn.Range
    tip = (
        float(dim.AsDimensionElement.Point(1).X),
        float(dim.AsDimensionElement.Point(1).Y),
    )
    le = stub.AsLineElement
    stub_dx = float(le.EndPoint.X) - float(le.StartPoint.X)
    gap = float(r.Low.X) - tip[0]
    ok = gap >= 10.0 and stub_dx < 0
    print(
        f"GG{gg} LowX={float(r.Low.X):.1f} tipX={tip[0]:.1f} gap={gap:.1f} "
        f"stub_dx={stub_dx:.1f} {'OK' if ok else 'BAD'}"
    )

PAD = 0.5
clashes = []
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
print(f"overlap {'PASS' if not clashes else 'FAIL'} {clashes}")
print("DONE")
