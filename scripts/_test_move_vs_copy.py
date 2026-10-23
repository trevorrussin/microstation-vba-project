"""Live test: does MOVE duplicate a GG callout? Force nocopy settings first."""
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

as_.AssociationLockEnabled = False
as_.GraphicGroupLockEnabled = True
print("locks Assoc", as_.AssociationLockEnabled, "GG", as_.GraphicGroupLockEnabled)

# Force settings while Move active
q.SendKeyin("NULL")
time.sleep(0.1)
q.SendKeyin("MOVE ELEMENT")
time.sleep(0.2)
for cmd in [
    "SET ITEM TOOLSETTINGS MAKECOPY=0",
    "SET ITEM TOOLSETTINGS COPIES=0",
    "SET ITEM TOOLSETTINGS COPYTOGGLE=0",
    "SET ITEM TOOLSETTINGS WANTMAKECOPY=0",
]:
    q.SendKeyin(cmd)
q.SendReset()
time.sleep(0.2)

TN_ID = 6117  # TextNode in GG 6114
GG = 6114

def elems_in_gg(gg):
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    out = []
    while ee.MoveNext():
        e = ee.Current
        try:
            if int(e.GraphicGroup) == gg:
                out.append(int(e.ID64))
        except Exception:
            pass
    return sorted(out)

def get_tn_origin(eid):
    e = mr.GetElementByID64(eid)
    tn = e.AsTextNodeElement
    o = tn.Origin
    return (float(o.X), float(o.Y), float(o.Z))

before_ids = elems_in_gg(GG)
print("before GG", GG, before_ids)
ox, oy, oz = get_tn_origin(TN_ID)
print("origin before", ox, oy)

dx, dy = 5.0, 5.0  # small nudge
from_pt = app.Point3dFromXYZ(ox, oy, oz)
to_pt = app.Point3dFromXYZ(ox + dx, oy + dy, oz)

# Programmatic move of whole GG via COM (bypasses Make Copy UI) — first
# prove whether USER tool copies by using CadInputQueue MOVE.
q.SendKeyin("NULL")
time.sleep(0.1)
q.SendKeyin("MOVE ELEMENT")
time.sleep(0.15)
q.SendKeyin("SET ITEM TOOLSETTINGS MAKECOPY=0")
time.sleep(0.05)

# Select the text node
e = mr.GetElementByID64(TN_ID)
# Use locate datapoints
q.SendDataPointForLocate(e, from_pt)
time.sleep(0.15)
q.SendDataPoint(from_pt, 1)
time.sleep(0.15)
q.SendDataPoint(to_pt, 1)
time.sleep(0.4)
q.SendReset()
time.sleep(0.2)

after_ids = elems_in_gg(GG)
print("after GG ids", after_ids)
try:
    nx, ny, nz = get_tn_origin(TN_ID)
    print("origin after same id", nx, ny, "delta", nx - ox, ny - oy)
except Exception as ex:
    print("origin after FAIL", ex)

# Count textnodes near old and new origin
o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
near_old = near_new = 0
while ee.MoveNext():
    e = ee.Current
    try:
        if not e.IsTextNodeElement:
            continue
        p = e.AsTextNodeElement.Origin
        if abs(p.X - ox) < 0.5 and abs(p.Y - oy) < 0.5:
            near_old += 1
        if abs(p.X - (ox + dx)) < 0.5 and abs(p.Y - (oy + dy)) < 0.5:
            near_new += 1
    except Exception:
        pass
print("textnodes near old", near_old, "near new", near_new)
print("id count before", len(before_ids), "after", len(after_ids))

# If copy happened, undo
if len(after_ids) > len(before_ids) or near_old >= 1 and near_new >= 1:
    print("COPY DETECTED — undoing")
    q.SendKeyin("UNDO")
    time.sleep(0.3)
else:
    print("MOVE looks clean (or locate failed)")
    # undo nudge anyway so drawing stays put
    q.SendKeyin("UNDO")
    time.sleep(0.2)

print("final origin", get_tn_origin(TN_ID))
print("done")
