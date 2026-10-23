"""Stack NEW conduit callouts on the LEFT like stock left-side layout.

Layout (matches left APS / left conduit look):
  Text column on the west (RJ), 5' stub east, note-dim to tip on conduit.
  Long clear leader — tip never under text.
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

# Left column: right edge of RJ text (west of stock APS ~1029333)
COL_RIGHT_X = 1029315.0
Y_TOP = 239960.0
Y_GAP = 4.0  # gap between callout blocks
STUB_LEN = 5.0
STUB_DX0 = 1.054
STUB_DY0 = -1.498


def pt3(x, y, z=0.0):
    return app.Point3dFromXYZ(float(x), float(y), float(z))


def nearest_tci(xy):
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
            # prefer tips east of the text column
            if cx < COL_RIGHT_X + 10:
                continue
            d = (cx - xy[0]) ** 2 + (cy - xy[1]) ** 2
            if d < best_d:
                best_d = d
                best = (cx, cy)
        except Exception:
            continue
    return best or (xy[0] + 80.0, xy[1])


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

# Order top→bottom by current Y (keep relative reading order)
items = []
for gg in range(7201, 7213):
    tn, stub, dim = gg_parts(gg)
    if tn is None:
        print("missing", gg)
        continue
    oy = float(tn.AsTextNodeElement.Origin.Y)
    h = float(tn.Range.High.Y) - float(tn.Range.Low.Y)
    items.append((oy, gg, int(tn.ID64), h))
items.sort(key=lambda x: -x[0])  # top first

print("stack order", [(gg, round(oy, 1), round(h, 1)) for oy, gg, _, h in items])

cursor_y = Y_TOP
placed = []

for _, gg, tn_id, text_h in items:
    tn, stub, dim = gg_parts(gg)
    if not (tn and dim):
        continue
    # Target: origin at COL_RIGHT_X; top of text near cursor_y
    # TN origin is roughly at top-right for just=12; use origin Y = cursor_y
    ox = COL_RIGHT_X
    oy = cursor_y
    cur = tn.AsTextNodeElement.Origin
    dx = ox - float(cur.X)
    dy = oy - float(cur.Y)
    print(f"GG{gg} -> ({ox:.1f},{oy:.1f}) d=({dx:.1f},{dy:.1f})")

    mv = pt3(dx, dy)
    # move TN only first
    e = mr.GetElementByID64(tn_id)
    e.Move(mv)
    e.Rewrite()

    # remove old stub/dim
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

    stub_start = (ox + STUB_DX0, oy + STUB_DY0)
    stub_end = (stub_start[0] + STUB_LEN, stub_start[1])
    tip = nearest_tci((ox + 60.0, oy))

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

    # refresh height after move
    tn2 = mr.GetElementByID64(tn_id)
    h = float(tn2.Range.High.Y) - float(tn2.Range.Low.Y)
    cursor_y = float(tn2.Range.Low.Y) - Y_GAP
    gap = tip[0] - float(tn2.Range.High.X)
    print(f"  stub={stub_id} dim={new_dim} tip=({tip[0]:.1f},{tip[1]:.1f}) tip-HighX={gap:.1f} h={h:.1f}")
    placed.append(gg)

as_.GraphicGroupLockEnabled = True
q.SendKeyin("SAVE DESIGN")
time.sleep(0.4)

# verify
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
                int(e.GraphicGroup),
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
    stub_dir = float(le.EndPoint.X) - float(le.StartPoint.X)
    print(
        f"GG{gg} textHigh={float(r.High.X):.1f} tipX={tip[0]:.1f} "
        f"gap={tip[0]-float(r.High.X):.1f} stub_dx={stub_dir:.1f}"
    )

PAD = 0.5
clashes = []
for i in range(len(texts)):
    for j in range(i + 1, len(texts)):
        a, b = texts[i], texts[j]
        if (
            a[2][0] - PAD <= b[3][0]
            and b[2][0] - PAD <= a[3][0]
            and a[2][1] - PAD <= b[3][1]
            and b[2][1] - PAD <= a[3][1]
        ):
            clashes.append((a[0], b[0]))
print(f"overlap {'PASS' if not clashes else 'FAIL'} {clashes}")
print("DONE placed", placed)
