"""Click Move mode label in tool settings; dump all Controls mentioning Move/Copy."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

import pythoncom
import ms_connect
import uiautomation as auto

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
q = app.CadInputQueue
as_ = app.ActiveSettings
as_.AssociationLockEnabled = False
as_.GraphicGroupLockEnabled = True

q.SendKeyin("NULL")
time.sleep(0.1)
q.SendKeyin("MOVE ELEMENT")
time.sleep(0.5)

root = auto.GetRootControl()
ms = None
for w in root.GetChildren():
    if "MicroStation" in (w.Name or ""):
        ms = w
        break
assert ms

hits = []

def walk(ctrl, depth=0, trail=None):
    if trail is None:
        trail = []
    if depth > 16:
        return
    try:
        name = ctrl.Name or ""
        ctype = ctrl.ControlTypeName
        aid = ""
        try:
            aid = ctrl.AutomationId or ""
        except Exception:
            pass
        low = name.lower()
        if low in ("move", "copy", "make copy") or "make copy" in low:
            hits.append((depth, ctype, name, aid, " / ".join(trail[-6:])))
        trail2 = trail + [f"{ctype}:{name[:40]}"]
    except Exception:
        return
    try:
        for c in ctrl.GetChildren():
            walk(c, depth + 1, trail2)
    except Exception:
        pass

walk(ms)
print("hits", len(hits))
for h in hits:
    print(h)

# Click every exact 'Move' that is Text/Radio/ListItem under a Move Element trail
for depth, ctype, name, aid, trail in hits:
    if name != "Move":
        continue
    if "Copy Element" in trail:
        continue
    print("TRY CLICK", ctype, aid, trail)
    # re-find by walking again is hard; use ControlFromPoint after locating BoundingRectangle via search
# Use auto search
for ctrl in [
    ms.TextControl(searchDepth=20, Name="Move"),
    ms.RadioButtonControl(searchDepth=20, Name="Move"),
    ms.ListItemControl(searchDepth=20, Name="Move"),
]:
    try:
        if ctrl.Exists(0, 0):
            print("found", ctrl.ControlTypeName, ctrl.Name, ctrl.BoundingRectangle)
            # Check parent names
            p = ctrl.GetParentControl()
            print(" parent", p.ControlTypeName if p else None, p.Name if p else None)
            ctrl.Click()
            print(" clicked")
            time.sleep(0.2)
    except Exception as e:
        print("err", e)

# Try keyins that match Mode option (Move=0 Copy=1) common pattern
for name in (
    "METHOD",
    "MODE",
    "OPERATIONCHOICE",
    "METHODCHOICE",
    "MOVECOPYCHOICE",
    "ACTIONCHOICE",
    "TASKCHOICE",
    "TYPECHOICE",
    "OPTIONCHOICE",
    "CHOICE",
    "USEMETHOD",
    "ACTIVEOPTION",
):
    for val in ("0", "1"):
        cmd = f"SET ITEM TOOLSETTINGS {name}={val}"
        try:
            q.SendKeyin(cmd)
        except Exception:
            pass
    print("tried", name)

# Explicit: start MOVE then set make copy false variants
q.SendKeyin("MOVE ELEMENT;SET ITEM TOOLSETTINGS MAKECOPY=0")
print("AssociationLock", as_.AssociationLockEnabled, "GG", as_.GraphicGroupLockEnabled)
print("done")
