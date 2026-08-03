"""Prove visual dimension (extension lines + dim line + distance text)."""
from __future__ import annotations

import math
import sys

import pythoncom
import win32com.client

pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("MicroStationDGN.Application")
mr = app.ActiveModelReference

x1, y1, x2, y2 = 35600.0, 35000.0, 35700.0, 35000.0
ox, oy = 35650.0, 35030.0
dx, dy = x2 - x1, y2 - y1
L = math.hypot(dx, dy) or 1.0
ux, uy = dx / L, dy / L
mx, my = 0.5 * (x1 + x2), 0.5 * (y1 + y2)
oxv, oyv = ox - mx, oy - my
px, py = -uy, ux
side = 1.0 if (oxv * px + oyv * py) >= 0 else -1.0
off = abs(oxv * px + oyv * py) or 30.0
d1x, d1y = x1 + side * off * px, y1 + side * off * py
d2x, d2y = x2 + side * off * px, y2 + side * off * py


def add_line(a, b, col=2, wt=0):
    ln = app.CreateLineElement2(
        None,
        app.Point3dFromXYZ(a[0], a[1], 0),
        app.Point3dFromXYZ(b[0], b[1], 0),
    )
    if isinstance(ln, tuple):
        ln = ln[0]
    ln.Color = col
    ln.LineWeight = wt
    mr.AddElement(ln)
    ln.Rewrite()
    return int(ln.ID.Low)


ids = [
    add_line((x1, y1), (d1x, d1y)),
    add_line((x2, y2), (d2x, d2y)),
    add_line((d1x, d1y), (d2x, d2y)),
]
print("lines", ids, "len", L)

ci = app.CadInputQueue
ci.SendCommand("TEXTEDITOR PLACE")
txt = f"{int(round(L))}'"
ci.SendKeyin(f'TEXTEDITOR PLAYCOMMAND INSERT_TEXT "{txt}"')
ci.SendDataPoint(app.Point3dFromXYZ(ox, oy, 0), 1)
ci.SendReset()
print("OK visual dim")
sys.exit(0)
