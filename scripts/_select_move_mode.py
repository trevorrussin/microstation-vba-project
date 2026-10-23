"""Select Move mode (not Copy) in Move Element tool settings."""
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
time.sleep(0.15)
q.SendKeyin("MOVE ELEMENT")
time.sleep(0.5)
q.SendKeyin("DIALOG TOOLSETTINGS OPEN")
time.sleep(0.4)

root = auto.GetRootControl()
ms = None
for w in root.GetChildren():
    if "MicroStation" in (w.Name or ""):
        ms = w
        break
assert ms, "no MicroStation window"

# Dump full tool settings tree under Move Element HostForm
host = ms.WindowControl(searchDepth=3, Name="Move Element")
print("host found", bool(host), host.Name if host else None)

def dump(ctrl, depth=0, max_depth=14):
    if depth > max_depth:
        return
    try:
        name = ctrl.Name or ""
        ctype = ctrl.ControlTypeName
        aid = ""
        try:
            aid = ctrl.AutomationId or ""
        except Exception:
            pass
        print(f"{'  '*depth}{ctype} {name!r} id={aid!r}")
    except Exception as e:
        print(f"{'  '*depth}? {e}")
        return
    try:
        for c in ctrl.GetChildren():
            dump(c, depth + 1, max_depth)
    except Exception:
        pass

if host:
    dump(host)

# Click the 'Move' text/radio if present; avoid clicking 'Copy'
clicked = False
def find_click_move(ctrl, depth=0):
    global clicked
    if clicked or depth > 14:
        return
    try:
        name = (ctrl.Name or "").strip()
        ctype = ctrl.ControlTypeName
        if name == "Move" and ctype in (
            "RadioButtonControl",
            "TextControl",
            "ListItemControl",
            "ButtonControl",
            "CheckBoxControl",
            "HyperlinkControl",
        ):
            print("clicking", ctype, repr(name), "id", getattr(ctrl, "AutomationId", None))
            # Prefer SelectionItem / Toggle
            try:
                sp = ctrl.GetSelectionItemPattern()
                if sp:
                    sp.Select()
                    print("  SelectionItem.Select")
                    clicked = True
                    return
            except Exception as e:
                print("  sel", e)
            try:
                ctrl.Click()
                print("  Click()")
                clicked = True
                return
            except Exception as e:
                print("  click", e)
    except Exception:
        pass
    try:
        for c in ctrl.GetChildren():
            find_click_move(c, depth + 1)
    except Exception:
        pass

if host:
    find_click_move(host)

# Try mode/method keyins
for cmd in [
    "SET ITEM TOOLSETTINGS METHOD=0",
    "SET ITEM TOOLSETTINGS MODE=0",
    "SET ITEM TOOLSETTINGS OPERATION=0",
    "SET ITEM TOOLSETTINGS ACTION=0",
    "SET ITEM TOOLSETTINGS TASK=0",
    "SET ITEM TOOLSETTINGS TYPE=0",
    "SET ITEM TOOLSETTINGS CHOICE=0",
    "SET ITEM TOOLSETTINGS OPTION=0",
    "SET ITEM TOOLSETTINGS MOVECOPY=0",
    "SET ITEM TOOLSETTINGS MOVEORCOPY=0",
    "SET ITEM TOOLSETTINGS USECOPY=0",
    "SET ITEM TOOLSETTINGS COPYMODE=0",
    "SET ITEM TOOLSETTINGS MAKECOPY=0",
]:
    try:
        q.SendKeyin(cmd)
        print("sent", cmd)
    except Exception as e:
        print("fail", cmd, e)

print("clicked_move", clicked)
print("AssociationLock", as_.AssociationLockEnabled, "GG", as_.GraphicGroupLockEnabled)
print("done")
