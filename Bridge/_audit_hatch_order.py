"""Audit south corridor: W04 face, hatch, signs vs order-table tips."""
from __future__ import annotations

import pythoncom
import win32com.client
from win32com.client import Dispatch

pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("MicroStationDGN.Application")
mr = app.ActiveModelReference

# --- W04-02R face detail ---
oScan = Dispatch("MicroStationDGN.ElementScanCriteria")
oScan.ExcludeNonGraphical()
ee = mr.Scan(oScan)
print("=== W04-02R / yellow faces ===")
while ee.MoveNext():
    el = ee.Current
    if not el.IsCellElement:
        continue
    c = el.AsCellElement
    if "W04" not in c.Name.upper() and "W4" not in c.Name.upper():
        continue
    oy = float(c.Origin.Y)
    if oy < 216850 or oy > 217100:
        continue
    print(f"CELL {c.Name} id={int(el.ID.Low)} origin=({c.Origin.X:.2f},{c.Origin.Y:.2f}) color={el.Color} wt={el.LineWeight}")
    try:
        ce = c.GetSubElements()
        i = 0
        while ce.MoveNext() and i < 20:
            sub = ce.Current
            try:
                sc = int(sub.Color)
            except Exception:
                sc = -1
            try:
                sw = int(sub.LineWeight)
            except Exception:
                sw = -1
            try:
                sl = sub.Level.Name
            except Exception:
                sl = "?"
            # fill?
            filled = ""
            try:
                if hasattr(sub, "FillMode") or sub.IsShapeElement:
                    filled = f" IsShape={sub.IsShapeElement}"
            except Exception:
                pass
            sr = sub.Range
            print(f"  sub[{i}] type={sub.Type} color={sc} wt={sw} lvl={sl}{filled} "
                  f"Y=[{sr.Low.Y:.1f},{sr.High.Y:.1f}]")
            i += 1
    except Exception as ex:
        print("  sub fail", ex)

# --- workspace shapes / hatch ---
print("\n=== shapes near south workspace ===")
oScan2 = Dispatch("MicroStationDGN.ElementScanCriteria")
oScan2.ExcludeNonGraphical()
ee2 = mr.Scan(oScan2)
while ee2.MoveNext():
    el = ee2.Current
    try:
        rng = el.Range
        eid = int(el.ID.Low)
    except Exception:
        continue
    cx = 0.5 * (float(rng.Low.X) + float(rng.High.X))
    cy = 0.5 * (float(rng.Low.Y) + float(rng.High.Y))
    if not (1022200 <= cx <= 1022400 and 217020 <= cy <= 217050):
        continue
    if el.IsShapeElement:
        ce = el.AsClosedElement
        try:
            hp = ce.HasPattern()
        except Exception:
            hp = getattr(ce, "HasPattern", None)
        try:
            lvl = el.Level.Name
        except Exception:
            lvl = "?"
        print(f"SHAPE id={eid} color={el.Color} wt={el.LineWeight} lvl={lvl} HasPattern={hp} "
              f"X=[{rng.Low.X:.1f},{rng.High.X:.1f}] Y=[{rng.Low.Y:.1f},{rng.High.Y:.1f}]")
        # try get pattern params
        try:
            # Pattern display
            print(f"  FillMode? ", end="")
            print(getattr(el, "FillMode", "n/a"), getattr(el, "FillColor", "n/a"))
        except Exception as ex:
            print("  fill", ex)

# --- perps vs signs on south align ---
print("\n=== south perps (80ft) and signs at tips ===")
oScan3 = Dispatch("MicroStationDGN.ElementScanCriteria")
oScan3.ExcludeNonGraphical()
ee3 = mr.Scan(oScan3)
perps = []
faces = []
posts = []
while ee3.MoveNext():
    el = ee3.Current
    if el.IsLineElement:
        try:
            s = el.AsLineElement.StartPoint
            e = el.AsLineElement.EndPoint
            length = ((float(e.X) - float(s.X)) ** 2 + (float(e.Y) - float(s.Y)) ** 2) ** 0.5
        except Exception:
            continue
        if abs(length - 80) > 0.5:
            continue
        if abs(float(s.X) - float(e.X)) > 1:
            continue
        tip = s if float(s.Y) < float(e.Y) else e
        if abs(float(tip.Y) - 217000.6) > 2:
            continue
        if not (1019700 < float(tip.X) < 1022800):
            continue
        perps.append((float(tip.X), float(tip.Y), int(el.ID.Low)))
    elif el.IsCellElement:
        c = el.AsCellElement
        ox, oy = float(c.Origin.X), float(c.Origin.Y)
        if oy < 216850 or oy > 217050:
            continue
        if 1019700 < ox < 1022800:
            if c.Name == "TWZSGN_P":
                posts.append((ox, oy, int(el.ID.Low)))
            elif c.Name in ("W20-01RA", "W20-05RA", "W04-02R", "G20-02") or c.Name.startswith("W") or c.Name.startswith("G"):
                faces.append((c.Name, ox, oy, int(el.ID.Low)))

perps.sort()
faces.sort(key=lambda t: t[1])
print(f"perps={len(perps)} faces={len(faces)} posts={len(posts)}")
for tipx, tipy, pid in perps:
    # find face near this tip X
    match = None
    for name, ox, oy, fid in faces:
        if abs(ox - tipx) < 2:
            match = (name, ox, oy, fid)
            break
    post = None
    for ox, oy, pid2 in posts:
        if abs(ox - tipx) < 5:
            post = pid2
            break
    status = f"SIGN {match[0]}" if match else "BLANK"
    print(f"  tip X={tipx:.1f} perp={pid} -> {status}")
