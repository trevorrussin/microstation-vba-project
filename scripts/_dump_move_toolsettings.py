"""Deep-dump Move Element WinForms tool settings; toggle any Copy-related control."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

import pythoncom
import ms_connect
import uiautomation as auto
import win32gui
import win32con

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
q = app.CadInputQueue
as_ = app.ActiveSettings
as_.AssociationLockEnabled = False
as_.GraphicGroupLockEnabled = True

q.SendKeyin("NULL")
time.sleep(0.1)
q.SendKeyin("MOVE ELEMENT")
time.sleep(0.4)
q.SendKeyin("DIALOG TOOLSETTINGS OPEN")
time.sleep(0.6)

# Find top-level Move Element windows
root = auto.GetRootControl()
move_wins = []
for w in root.GetChildren():
    if (w.Name or "") == "Move Element":
        move_wins.append(w)
        print("found Move Element top", w.ControlTypeName, w.ClassName)

def dump(ctrl, depth=0, max_depth=20):
    if depth > max_depth:
        return
    try:
        name = ctrl.Name or ""
        ctype = ctrl.ControlTypeName
        cls = ""
        try:
            cls = ctrl.ClassName or ""
        except Exception:
            pass
        aid = ""
        try:
            aid = ctrl.AutomationId or ""
        except Exception:
            pass
        rect = None
        try:
            rect = ctrl.BoundingRectangle
        except Exception:
            pass
        # print almost everything with size
        if rect and rect.width() > 0 and rect.height() > 0:
            print(f"{'  '*depth}{ctype}/{cls} name={name!r} id={aid!r} {rect}")
        elif name or aid:
            print(f"{'  '*depth}{ctype}/{cls} name={name!r} id={aid!r}")
    except Exception as e:
        print(f"{'  '*depth}ERR {e}")
        return
    try:
        for c in ctrl.GetChildren():
            dump(c, depth + 1, max_depth)
    except Exception:
        pass

for w in move_wins:
    print("==== DUMP", w.Name)
    dump(w)

# Probe C expressions aggressively
print("==== CEXPR")
candidates = []
prefixes = [
    "tcb->msToolSettings",
    "tcb->toolSettings",
    "tcb->msToolSettings.manip",
    "tcb->msToolSettings.move",
    "tcb->msToolSettings.copy",
]
fields = [
    "copy", "makeCopy", "MakeCopy", "wantMakeCopy", "WantMakeCopy",
    "copies", "useCopy", "copyMode", "mode", "method", "action",
    "moveCopy", "flags", "options", "iCopy", "bCopy",
]
for p in prefixes:
    for f in fields:
        candidates.append(f"{p}.{f}")
# also flat
candidates += [
    "tcb->msToolSettings.copy",
    "session.msToolSettings.copy",
]

for expr in candidates:
    try:
        v = app.GetCExpressionValue(expr)
        print("GET", expr, "->", v)
        try:
            app.SetCExpressionValue(expr, 0)
            print("  SET0 ->", app.GetCExpressionValue(expr))
        except Exception as e:
            print("  setfail", e)
    except Exception:
        pass

# Try mdl key-ins
for cmd in [
    "MDL COMMAND DIALOG TOOLSETTINGS",
    "EXPAND SET ITEM TOOLSETTINGS",
    "SET ITEM TOOLSETTINGS ?",
]:
    try:
        q.SendKeyin(cmd)
        print("keyin", cmd)
    except Exception as e:
        print("keyin fail", cmd, e)

print("AssociationLock", as_.AssociationLockEnabled, "GG", as_.GraphicGroupLockEnabled)
print("done")
