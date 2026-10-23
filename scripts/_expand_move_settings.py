"""Expand Move Element tool settings; find Copies/Make Copy; set to 0/off."""
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
time.sleep(0.3)
q.SendKeyin("DIALOG TOOLSETTINGS OPEN")
time.sleep(0.4)

# Resize Move Element window taller so Copies/Make Copy is visible
def find_hwnd():
    found = []
    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) == "Move Element":
            found.append(hwnd)
        return True
    win32gui.EnumWindows(cb, None)
    return found

for hwnd in find_hwnd():
    # GetWindowRect and enlarge height
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    print("hwnd", hex(hwnd), "rect", l, t, r, b)
    win32gui.SetWindowPos(hwnd, None, l, t, max(r - l, 320), 420, win32con.SWP_NOZORDER)
    print("resized")

time.sleep(0.3)

root = auto.GetRootControl()
for w in root.GetChildren():
    if (w.Name or "") != "Move Element":
        continue
    print("=== Move Element tree ===")
    stack = [(w, 0)]
    while stack:
        c, d = stack.pop(0)
        try:
            name = c.Name or ""
            ctype = c.ControlTypeName
            aid = ""
            try:
                aid = c.AutomationId or ""
            except Exception:
                pass
            if d < 12:
                print(f"{'  '*d}{ctype} {name!r} id={aid!r}")
            # If Copies edit / spinner
            if "copy" in name.lower() or aid.lower().find("copy") >= 0:
                print("  ** copy-related **")
                try:
                    vp = c.GetValuePattern()
                    if vp:
                        print("  value", vp.Value)
                        vp.SetValue("0")
                        print("  set 0 ->", vp.Value)
                except Exception as e:
                    print("  value err", e)
                try:
                    tp = c.GetTogglePattern()
                    if tp:
                        print("  toggle", tp.ToggleState)
                        if int(tp.ToggleState) == 1:
                            tp.Toggle()
                            print("  toggled off")
                except Exception as e:
                    print("  toggle err", e)
            for ch in c.GetChildren():
                stack.append((ch, d + 1))
        except Exception:
            pass

# Keyins for copies count
for cmd in [
    "SET ITEM TOOLSETTINGS COPIES=0",
    "SET ITEM TOOLSETTINGS NUMCOPIES=0",
    "SET ITEM TOOLSETTINGS COPYCOUNT=0",
    "SET ITEM TOOLSETTINGS MAKECOPY=0",
    "SET ITEM TOOLSETTINGS COPYTOGGLE=0",
]:
    q.SendKeyin(cmd)
    print("sent", cmd)

print("AssociationLock", as_.AssociationLockEnabled, "GG", as_.GraphicGroupLockEnabled)
print("done")
