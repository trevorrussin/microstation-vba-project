"""Inspect last 3 estimate-table rows in detail."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch
from view_capture import capture_window

ITEM_X, DESC_X, UNIT_X, QTY_X = 1030191.0, 1030272.88, 1030347.37, 1030367.63

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
mr = None
for m in app.ActiveDesignFile.Models:
    if m.Name == "XS Stacked Views":
        mr = m
        break

o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
rows = {"item": [], "desc": [], "unit": [], "qty": [], "band": []}
while ee.MoveNext():
    el = ee.Current
    try:
        rng = el.Range
        lvl = el.Level.Name
    except Exception:
        continue
    cy = (float(rng.High.Y) + float(rng.Low.Y)) / 2
    cx = (float(rng.High.X) + float(rng.Low.X)) / 2
    if cy < 239535 or cy > 239590:
        continue
    if float(rng.Low.X) < 1030150 or float(rng.High.X) > 1030400:
        continue
    h = float(rng.High.Y) - float(rng.Low.Y)
    w = float(rng.High.X) - float(rng.Low.X)
    if lvl == "O_Details_Table Border_P" and int(el.Type) == 6 and w > 150 and 5 < h < 14:
        rows["band"].append((cy, h, float(rng.Low.Y), float(rng.High.Y)))
        continue
    if lvl != "TCSH_P_T":
        continue
    if el.IsTextNodeElement:
        tn = el.AsTextNodeElement
        t = " ".join(tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)).strip()
        if abs(cx - ITEM_X) < 8:
            rows["item"].append((cy, t))
        elif abs(cx - UNIT_X) < 8:
            rows["unit"].append((cy, t))
        elif abs(cx - QTY_X) < 8:
            rows["qty"].append((cy, t))
    elif el.IsTextElement and abs(cx - DESC_X) < 60:
        te = el.AsTextElement
        rows["desc"].append(
            (
                cy,
                te.Text.strip(),
                float(rng.Low.X),
                float(rng.High.X),
                float(rng.Low.Y),
                float(rng.High.Y),
                w,
            )
        )

for k in rows:
    rows[k].sort(key=lambda x: -x[0])

print("=== LAST ROWS DETAIL ===")
for it in rows["item"]:
    cy, item = it
    band = min(rows["band"], key=lambda b: abs(b[0] - cy))
    descs = [d for d in rows["desc"] if abs(d[0] - cy) < 8 or abs(d[0] - band[0]) < 8]
    # better: descs whose cy within band y range
    descs = [
        d
        for d in rows["desc"]
        if d[4] >= band[2] - 0.5 and d[5] <= band[3] + 0.5
    ]
    units = [u for u in rows["unit"] if abs(u[0] - cy) < 1]
    qtys = [q for q in rows["qty"] if abs(q[0] - cy) < 1]
    print(f"\nITEM {item} cy={cy:.3f} band_h={band[1]:.2f} band_y=[{band[2]:.3f},{band[3]:.3f}]")
    print(f"  unit={units} qty={qtys}")
    for d in descs:
        top = band[3] - d[5]
        bot = d[4] - band[2]
        print(
            f"  DESC cy={d[0]:.3f} w={d[6]:.1f} pad_top={top:.2f} pad_bot={bot:.2f} {d[1]!r}"
        )
        if len(descs) == 2:
            pass
    if len(descs) == 2:
        gap = descs[0][0] - descs[1][0]
        clear = descs[0][4] - descs[1][5]
        print(f"  2line pitch={gap:.2f} box_clear={clear:.2f}")

# navigate view tightly
v = app.ActiveDesignFile.Views(1)
v.Center = app.Point3dFromXYZ(1030275, 239560, 0)
v.Extents = app.Point3dFromXYZ(200, 70, 0)
app.CadInputQueue.SendKeyin("UPDATE VIEW EXTENDED 1")
time.sleep(2.2)
p = capture_window(
    "MicroStation",
    out_path=r"c:\repos\microstation-vba-project\scripts\_bts_fmt\table_last3_tight.png",
)
print("\ncapture", p)
