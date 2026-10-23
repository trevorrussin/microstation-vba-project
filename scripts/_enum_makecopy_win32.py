"""Enum all Win32 child windows; find Make Copy / tool settings checkboxes."""
from __future__ import annotations

import sys
import time
import ctypes
from ctypes import wintypes

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

import pythoncom
import ms_connect
import win32gui
import win32con
import win32api

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
time.sleep(0.5)

user32 = ctypes.windll.user32

def get_text(hwnd):
    n = win32gui.GetWindowTextLength(hwnd)
    if n <= 0:
        return ""
    return win32gui.GetWindowText(hwnd)

def get_class(hwnd):
    return win32gui.GetClassName(hwnd)

ms_hwnds = []

def enum_top(hwnd, _):
    title = get_text(hwnd)
    if "MicroStation" in title:
        ms_hwnds.append(hwnd)
        print("TOP", hex(hwnd), title[:100])
    if "Move Element" in title or "Tool Settings" in title or "Key-in" in title:
        print("TOPOTHER", hex(hwnd), title[:100], get_class(hwnd))
    return True

win32gui.EnumWindows(enum_top, None)

interesting = []

def enum_child(hwnd, results):
    try:
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = get_text(hwnd)
        cls = get_class(hwnd)
        low = title.lower()
        if title and (
            "copy" in low
            or "make" in low
            or "move" in low
            or "tool" in low
            or cls.lower() in ("button", "button", "bst_checkbox")
            or "checkbox" in cls.lower()
        ):
            results.append((hwnd, cls, title))
        # also collect all buttons with empty? no
    except Exception:
        pass
    return True

all_hits = []
for top in ms_hwnds:
    win32gui.EnumChildWindows(top, enum_child, all_hits)

# Also enum children of any top-level Move Element
def enum_top2(hwnd, _):
    title = get_text(hwnd)
    if title in ("Move Element", "Tool Settings", "Key-in"):
        win32gui.EnumChildWindows(hwnd, enum_child, all_hits)
    return True
win32gui.EnumWindows(enum_top2, None)

print("hits", len(all_hits))
for hwnd, cls, title in all_hits:
    print(hex(hwnd), cls, repr(title))

# Click any checkbox whose text contains Make Copy
for hwnd, cls, title in all_hits:
    if "make copy" in title.lower():
        # BM_GETCHECK
        state = win32gui.SendMessage(hwnd, win32con.BM_GETCHECK, 0, 0)
        print("Make Copy checkbox state", state)
        if state == win32con.BST_CHECKED:
            win32gui.SendMessage(hwnd, win32con.BM_CLICK, 0, 0)
            print("clicked to uncheck")
            state2 = win32gui.SendMessage(hwnd, win32con.BM_GETCHECK, 0, 0)
            print("new state", state2)

# WPF checkboxes may not be classic BUTTON - try UI Automation Toggle from all checkboxes again with deeper search
import uiautomation as auto
root = auto.GetRootControl()
for w in root.GetChildren():
    if "MicroStation" not in (w.Name or "") and (w.Name or "") != "Move Element":
        continue
    cbs = w.GetChildren()
    # recursive checkbox hunt
    stack = [w]
    while stack:
        c = stack.pop()
        try:
            ctype = c.ControlTypeName
            name = c.Name or ""
            if ctype == "CheckBoxControl":
                print("CB", repr(name), c.BoundingRectangle)
                try:
                    tp = c.GetTogglePattern()
                    print("  toggle", tp.ToggleState if tp else None)
                    if tp and int(tp.ToggleState) == 1 and "copy" in name.lower():
                        tp.Toggle()
                        print("  unchecked", name)
                except Exception as e:
                    print("  err", e)
            for ch in c.GetChildren():
                stack.append(ch)
        except Exception:
            pass

print("AssociationLock", as_.AssociationLockEnabled, "GG", as_.GraphicGroupLockEnabled)
print("done")
