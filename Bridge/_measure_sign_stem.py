"""Measure user-ref vs probe sign assemblies for stem/post geometry."""
from __future__ import annotations

import pythoncom
import win32com.client
from win32com.client import Dispatch

pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("MicroStationDGN.Application")
mr = app.ActiveModelReference


def scan():
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    return mr.Scan(o)


def dump_cluster(label, x0, y0, xspan=120, yspan=150):
    print(f"\n==== {label} around ({x0},{y0}) ====")
    ee = scan()
    rows = []
    while ee.MoveNext():
        el = ee.Current
        rng = el.Range
        if el.IsCellElement:
            c = el.AsCellElement
            ox, oy = float(c.Origin.X), float(c.Origin.Y)
            if abs(ox - x0) > xspan or abs(oy - y0) > yspan:
                continue
            rows.append(
                (
                    "CELL",
                    c.Name,
                    ox,
                    oy,
                    float(c.Scale.X),
                    float(rng.High.X - rng.Low.X),
                    float(rng.High.Y - rng.Low.Y),
                    float(rng.Low.X),
                    float(rng.Low.Y),
                    float(rng.High.X),
                    float(rng.High.Y),
                )
            )
        elif el.IsLineElement:
            try:
                s = el.AsLineElement.StartPoint
                e = el.AsLineElement.EndPoint
            except Exception:
                continue
            sx, sy, ex, ey = float(s.X), float(s.Y), float(e.X), float(e.Y)
            mx, my = (sx + ex) / 2, (sy + ey) / 2
            if abs(mx - x0) > xspan or abs(my - y0) > yspan:
                continue
            length = ((ex - sx) ** 2 + (ey - sy) ** 2) ** 0.5
            rows.append(("LINE", length, sx, sy, ex, ey, None, None, None, None, None))
    for r in sorted(rows, key=lambda t: (t[0], t[2] if t[0] == "CELL" else t[3])):
        print(r)


# User ref (south of agent corridor)
dump_cluster("USER_REF", 1022111, 217270, 80, 120)
# Latest probe (dir south)
dump_cluster("PROBE", 1022600, 217020, 100, 120)
# Old agent (east of alignment)
dump_cluster("OLD_AGENT", 1022120, 217353, 80, 60)
