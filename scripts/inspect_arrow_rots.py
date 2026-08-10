"""Inspect turn-arrow cell origins/rotations near the smoke junction."""
from __future__ import annotations

import math
import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

import pythoncom

pythoncom.CoInitialize()

import ms_connect
from bridge_client import chat_bridge
import wztc_ops

wztc_ops.set_bridge(chat_bridge)
app = ms_connect.get_microstation_app()


def dump(eid: int) -> None:
    try:
        el = app.ActiveModelReference.GetElementByID2(eid)
    except Exception as e:
        print(f"id={eid} getfail={e}")
        return
    try:
        nm = str(el.Name)
    except Exception:
        nm = ""
    try:
        o = el.Origin
        ox, oy = float(o.X), float(o.Y)
    except Exception:
        ox = oy = None
    try:
        m = el.Rotation
        rx = float(m.RowX.X)
        ry = float(m.RowX.Y)
        ang = math.degrees(math.atan2(ry, rx))
        print(f"id={eid} name={nm!r} origin=({ox:.2f},{oy:.2f}) RowX=({rx:.4f},{ry:.4f}) ang={ang:.1f}")
    except Exception as e:
        print(f"id={eid} name={nm!r} origin=({ox},{oy}) rot_fail={e}")


ids = [
    139362,
    139377,
    139397,
    139410,
    139430,
    139443,
    139463,
    139476,
    139496,
    139515,
]
print("--- known agent arrows ---")
for i in ids:
    dump(i)

print("--- all CELL near junction ---")
rows = wztc_ops.find_elements_near(23800, 292500, 400, type_filter="CELL")
seen = set()
for e in rows or []:
    eid = int(float(e.get("elementId")))
    if eid in seen:
        continue
    seen.add(eid)
    dump(eid)

print("--- CELL south of junction (side road example?) ---")
rows2 = wztc_ops.find_elements_near(23800, 292200, 500, type_filter="CELL")
for e in rows2 or []:
    eid = int(float(e.get("elementId")))
    if eid in seen:
        continue
    seen.add(eid)
    dump(eid)

print("--- CELL east of junction ---")
rows3 = wztc_ops.find_elements_near(24100, 292500, 500, type_filter="CELL")
for e in rows3 or []:
    eid = int(float(e.get("elementId")))
    if eid in seen:
        continue
    seen.add(eid)
    dump(eid)
