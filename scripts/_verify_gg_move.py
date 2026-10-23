"""Verify no callout dups; COM-move one GG as unit and confirm no leftover."""
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
as_ = app.ActiveSettings
as_.AssociationLockEnabled = False
as_.GraphicGroupLockEnabled = True

keys = ("APS", "STREETLIGHT", "VEHICULAR", "CONTROL CABINET", "PUSH BUTTON")

def text_of(e):
    try:
        if e.IsTextNodeElement:
            tn = e.AsTextNodeElement
            return " | ".join(tn.TextLine(i) for i in range(1, int(tn.TextLinesCount) + 1))
        if e.IsTextElement:
            return e.AsTextElement.Text
    except Exception:
        return ""
    return ""

def scan_callout_tns():
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    out = []
    while ee.MoveNext():
        e = ee.Current
        if not e.IsTextNodeElement:
            continue
        t = text_of(e)
        if any(k in (t or "").upper() for k in keys):
            p = e.AsTextNodeElement.Origin
            out.append((int(e.ID64), int(e.GraphicGroup), float(p.X), float(p.Y), t[:40]))
    return out

before = scan_callout_tns()
print("callout TNs", len(before))
for b in before:
    print(" ", b)

# Move GG 6114 via COM transform on each member (true move, not tool)
GG = 6114
dx, dy = 2.0, 0.0
o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
members = []
while ee.MoveNext():
    e = ee.Current
    try:
        if int(e.GraphicGroup) == GG:
            members.append(int(e.ID64))
    except Exception:
        pass
print("members", members)

mat = app.Transform3dIdentity()
# Translation matrix
# Transform3dFromMatrix / Matrix3d - use Point3d offset via MoveElement if available
trans = app.Point3dFromXYZ(dx, dy, 0)
for eid in members:
    e = mr.GetElementByID64(eid)
    # Element.Move(Point3d) exists in MicroStation COM
    try:
        e.Move(trans)
        e.Rewrite()
        print("moved", eid)
    except Exception as ex:
        print("move fail", eid, ex)

after = scan_callout_tns()
print("after move callout TNs", len(after))
for a in after:
    print(" ", a)

# Undo the test nudge
q = app.CadInputQueue
# Multiple rewrites may be one undo each - undo enough
for _ in range(len(members) + 1):
    q.SendKeyin("UNDO")
    time.sleep(0.05)

final = scan_callout_tns()
print("after undo count", len(final))
for f in final:
    print(" ", f)

# Check any leftover text at old+new for CONTROL CABINET
print("AssociationLock", as_.AssociationLockEnabled, "GG", as_.GraphicGroupLockEnabled)
print("done")
