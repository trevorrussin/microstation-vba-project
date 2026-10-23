"""Delete orphan callout leftovers; force Move nocopy settings; verify."""
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

KEEP_IDS = {
    6107, 6108, 6109, 6110, 6111, 6112, 6113,  # dims
    6117, 6120, 6123, 6126, 6129, 6132, 6135,  # textnodes
    6138, 6139, 6140, 6141, 6142, 6143, 6144,  # stubs
}
KEEP_GG = set(range(6114, 6121))

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

o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
orphans = []
while ee.MoveNext():
    e = ee.Current
    eid = int(e.ID64)
    if eid in KEEP_IDS:
        continue
    try:
        gg = int(e.GraphicGroup)
    except Exception:
        gg = 0
    # Orphan callout-like text not in keep set
    if e.IsTextNodeElement or e.IsTextElement:
        t = (text_of(e) or "").upper()
        if any(k in t for k in keys):
            orphans.append((eid, int(e.Type), gg, text_of(e)[:60]))
            continue
    # Elements that somehow sit in our GG numbers but aren't keep ids
    if gg in KEEP_GG and eid not in KEEP_IDS:
        orphans.append((eid, int(e.Type), gg, text_of(e)[:60]))

print("orphans to delete", orphans)
for eid, typ, gg, txt in orphans:
    try:
        el = mr.GetElementByID64(eid)
        el.RemoveFromModel()
        print("deleted", eid, "type", typ, "gg", gg, repr(txt))
    except Exception as ex:
        print("delete fail", eid, ex)

# Force tool settings again
q.SendKeyin("NULL")
time.sleep(0.1)
q.SendKeyin("MOVE ELEMENT")
time.sleep(0.2)
for cmd in [
    "SET ITEM TOOLSETTINGS MAKECOPY=0",
    "SET ITEM TOOLSETTINGS COPIES=0",
    "SET ITEM TOOLSETTINGS COPYTOGGLE=0",
]:
    q.SendKeyin(cmd)
q.SendReset()

# Verify no duplicate callout text clusters
o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
callouts = []
while ee.MoveNext():
    e = ee.Current
    if not e.IsTextNodeElement:
        continue
    t = text_of(e)
    if any(k in (t or "").upper() for k in keys):
        p = e.AsTextNodeElement.Origin
        callouts.append((int(e.ID64), int(e.GraphicGroup), (round(float(p.X), 3), round(float(p.Y), 3)), t[:50]))

print("remaining callout TNs", len(callouts))
for c in callouts:
    print(" ", c)

from collections import Counter
xy_counts = Counter(c[2] for c in callouts)
dups = [xy for xy, n in xy_counts.items() if n > 1]
print("dup xy remaining", dups)
print("AssociationLock", as_.AssociationLockEnabled, "GG", as_.GraphicGroupLockEnabled)
print("done")
