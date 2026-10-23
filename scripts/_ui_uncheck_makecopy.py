"""Find and uncheck 'Make Copy' on Move Element tool settings only."""
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

print("COM file:", app.ActiveDesignFile.FullName)
print("COM model:", app.ActiveModelReference.Name)

as_.AssociationLockEnabled = False
as_.GraphicGroupLockEnabled = True
print("AssociationLock", as_.AssociationLockEnabled, "GG", as_.GraphicGroupLockEnabled)

q.SendKeyin("NULL")
time.sleep(0.2)
# Prefer MOVE ELEMENT over ribbon Copy
q.SendKeyin("MOVE ELEMENT")
time.sleep(0.6)
q.SendKeyin("DIALOG TOOLSETTINGS OPEN")
time.sleep(0.5)

root = auto.GetRootControl()
ms_windows = []
for w in root.GetChildren():
    name = w.Name or ""
    if "MicroStation" in name:
        ms_windows.append(w)
        print("MS WIN:", name[:120])

def dump_and_fix(ctrl, depth=0, max_depth=12, path=""):
    found = []
    try:
        name = (ctrl.Name or "").strip()
        ctype = ctrl.ControlTypeName
        aid = ""
        try:
            aid = ctrl.AutomationId or ""
        except Exception:
            pass
        interesting = False
        low = name.lower()
        if any(k in low for k in ("make copy", "makecopy", "copy", "move", "tool setting")):
            interesting = True
        if ctype in ("CheckBoxControl", "ToggleButtonControl") and name:
            interesting = True
        if interesting and depth <= max_depth:
            print(f"{'  '*depth}{ctype} name={name!r} id={aid!r}")
        if low == "make copy" or "make copy" in low:
            found.append(ctrl)
        if ctype == "CheckBoxControl" and "copy" in low:
            found.append(ctrl)
    except Exception:
        pass
    if depth >= max_depth:
        return found
    try:
        for c in ctrl.GetChildren():
            found.extend(dump_and_fix(c, depth + 1, max_depth, path))
    except Exception:
        pass
    return found

targets = []
for w in ms_windows:
    # Prefer Tool Settings / Move Element panes
    targets.extend(dump_and_fix(w, 0, 10))

print("candidates", len(targets))
for t in targets:
    name = t.Name or ""
    print("TARGET", t.ControlTypeName, repr(name))
    try:
        tp = t.GetTogglePattern()
        if tp:
            print("  state", tp.ToggleState)
            # ToggleState: 0 Off, 1 On, 2 Indeterminate
            if int(tp.ToggleState) == 1:
                tp.Toggle()
                print("  toggled to", tp.ToggleState)
            else:
                print("  already not On")
            continue
    except Exception as e:
        print("  GetTogglePattern", e)
    try:
        # Legacy TogglePattern via PatternId
        pat = t.GetPattern(auto.PatternId.TogglePattern)
        if pat:
            print("  ToggleState", pat.ToggleState)
            if pat.ToggleState == auto.ToggleState.On:
                pat.Toggle()
                print("  toggled OFF ->", pat.ToggleState)
            continue
    except Exception as e:
        print("  PatternId toggle", e)
    # last resort: click only if name is exactly Make Copy
    if (name or "").strip().lower() == "make copy":
        try:
            t.Click()
            print("  clicked Make Copy")
        except Exception as e:
            print("  click", e)

# Also try common toolsetting names after Move is active
for name in ("MAKECOPY", "WANTMAKECOPY", "COPYTOGGLE", "COPIES"):
    cmd = f"SET ITEM TOOLSETTINGS {name}=0"
    try:
        q.SendKeyin(cmd)
        print("sent", cmd)
    except Exception as e:
        print("fail", cmd, e)

print("final AssociationLock", as_.AssociationLockEnabled)
print("final GG", as_.GraphicGroupLockEnabled)
print("done")
