"""Dump estimate table last rows + capture via PrintWindow on tsp_02."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch
from view_capture import capture_window, navigate_view

DESC_X = 1030272.88
ITEM_X = 1030191.0

pythoncom.CoInitialize()
# Prefer any MicroStation with the sheet open
apps = list(ms_connect._candidate_apps())
print("instances", len(apps))
for i, (pid, app) in enumerate(apps):
    try:
        print(i, pid, app.ActiveDesignFile.FullName, app.ActiveModelReference.Name)
    except Exception as e:
        print(i, pid, e)

# Use first that has XS Stacked Views
app = None
mr = None
for _, a in apps:
    try:
        for m in a.ActiveDesignFile.Models:
            if m.Name == "XS Stacked Views":
                app = a
                mr = m
                break
    except Exception:
        continue
    if mr:
        break

if mr is None:
    # try active model
    _, app = apps[0]
    mr = app.ActiveModelReference
    print("fallback model", mr.Name)
else:
    print("using", app.ActiveDesignFile.FullName, mr.Name)

o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
items = []
descs = []
bands = []
while ee.MoveNext():
    el = ee.Current
    try:
        rng = el.Range
        lvl = el.Level.Name
    except Exception:
        continue
    cy = (float(rng.High.Y) + float(rng.Low.Y)) / 2
    cx = (float(rng.High.X) + float(rng.Low.X)) / 2
    if cy < 239540 or cy > 239720:
        continue
    if float(rng.Low.X) < 1030150 or float(rng.High.X) > 1030400:
        continue
    h = float(rng.High.Y) - float(rng.Low.Y)
    w = float(rng.High.X) - float(rng.Low.X)
    if lvl == "O_Details_Table Border_P" and int(el.Type) == 6 and w > 150 and 5 < h < 14:
        bands.append((float(rng.Low.Y), float(rng.High.Y), cy, h))
    if lvl != "TCSH_P_T":
        continue
    if el.IsTextNodeElement and abs(cx - ITEM_X) < 8:
        tn = el.AsTextNodeElement
        t = " ".join(tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)).strip()
        items.append((cy, t, float(rng.Low.Y), float(rng.High.Y)))
    if el.IsTextElement and abs(cx - DESC_X) < 60:
        descs.append(
            (
                cy,
                el.AsTextElement.Text.strip(),
                float(rng.Low.X),
                float(rng.High.X),
                float(rng.Low.Y),
                float(rng.High.Y),
                h,
                w,
            )
        )

items = sorted(items, key=lambda x: -x[0])
descs = sorted(descs, key=lambda x: -x[0])
bands = sorted(bands, key=lambda x: -x[2])

print("\n=== ITEMS (bottom 6) ===")
for row in items[-6:]:
    print(f"cy={row[0]:.3f} {row[1]!r}")

print("\n=== ALL DESCS from bottom third ===")
for d in descs:
    if d[0] < 239600:
        print(
            f"cy={d[0]:.3f} w={d[7]:.1f} y=[{d[4]:.2f},{d[5]:.2f}] {d[1]!r}"
        )

print("\n=== LAST 5 BANDS ===")
for b in bands[-5:]:
    print(f"cy={b[2]:.3f} h={b[3]:.2f} y=[{b[0]:.3f},{b[1]:.3f}]")

# check text overflow past unit column (~1030340)
print("\n=== DESC overflow past x=1030335 ===")
for d in descs:
    if d[3] > 1030335:
        print(f"OVERFLOW cy={d[0]:.3f} x1={d[3]:.1f} {d[1]!r}")

# Capture bottom of table via COM navigate on this app
try:
    # monkey: set ms_connect cache? navigate_view uses get_microstation_app
    # Just use PrintWindow of MicroStation after keyin zoom
    app.CadInputQueue.SendKeyin("VIEW 1")
    v = None
    # Views collection
    try:
        v = app.ActiveDesignFile.Views(1)
        v.Center = app.Point3dFromXYZ(1030272, 239575, 0)
        v.Extents = app.Point3dFromXYZ(220, 100, 0)
        app.CadInputQueue.SendKeyin("UPDATE VIEW EXTENDED 1")
        time.sleep(2.0)
    except Exception as e:
        print("view nav", e)
    p = capture_window("MicroStation", out_path=r"c:\repos\microstation-vba-project\scripts\_bts_fmt\table_last3_now.png")
    print("capture", p)
except Exception as e:
    print("capture fail", e)
