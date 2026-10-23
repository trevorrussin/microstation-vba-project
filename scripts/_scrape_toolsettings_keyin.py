"""Scrape Key-in browser 4th column for SET ITEM TOOLSETTINGS after MOVE."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

import pythoncom
import ms_connect
import uiautomation as auto
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
q.SendKeyin("DIALOG KEYIN OPEN")
time.sleep(0.8)

root = auto.GetRootControl()
ms = None
for w in root.GetChildren():
    if "MicroStation" in (w.Name or ""):
        ms = w
        break

# Find Key-in window
keyin = None
for w in root.GetChildren():
    name = (w.Name or "").lower()
    if "key-in" in name or "keyin" in name:
        keyin = w
        print("top keyin", w.Name)
        break

if ms and keyin is None:
    # nested
    def find_keyin(ctrl, depth=0):
        if depth > 6:
            return None
        try:
            n = (ctrl.Name or "").lower()
            if "key-in" in n:
                return ctrl
        except Exception:
            return None
        try:
            for c in ctrl.GetChildren():
                r = find_keyin(c, depth + 1)
                if r:
                    return r
        except Exception:
            pass
        return None
    keyin = find_keyin(ms)

print("keyin", keyin.Name if keyin else None)

def dump(ctrl, depth=0, max_depth=8):
    if depth > max_depth:
        return
    try:
        print(f"{'  '*depth}{ctrl.ControlTypeName} {(ctrl.Name or '')!r}")
    except Exception:
        return
    try:
        for c in ctrl.GetChildren():
            dump(c, depth + 1, max_depth)
    except Exception:
        pass

if keyin:
    dump(keyin)

# Type into key-in without Enter: use SendKeys to focused edit
# MicroStation key-in often accepts via CadInputQueue partial? Try SendKeyin without completing.
# Instead: find Edit control and set value.

edit = None

def find_edits(ctrl, depth=0, acc=None):
    if acc is None:
        acc = []
    if depth > 10:
        return acc
    try:
        if ctrl.ControlTypeName in ("EditControl", "DocumentControl"):
            acc.append(ctrl)
    except Exception:
        return acc
    try:
        for c in ctrl.GetChildren():
            find_edits(c, depth + 1, acc)
    except Exception:
        pass
    return acc

if keyin:
    edits = find_edits(keyin)
    print("edits", len(edits))
    for e in edits:
        print(" edit", repr(e.Name), "val", getattr(e, "GetValuePattern", lambda: None))
        try:
            vp = e.GetValuePattern()
            if vp:
                print("  current", vp.Value)
                vp.SetValue("SET ITEM TOOLSETTINGS")
                print("  set to", vp.Value)
                edit = e
        except Exception as ex:
            print("  err", ex)

time.sleep(0.5)
# After typing without enter, list boxes should populate - dump again
if keyin:
    print("--- after set ---")
    dump(keyin, 0, 10)
    # collect list items
    items = []
    def collect_items(ctrl, depth=0):
        if depth > 12:
            return
        try:
            if ctrl.ControlTypeName in ("ListItemControl", "TextControl", "DataItemControl"):
                n = ctrl.Name or ""
                if n:
                    items.append((ctrl.ControlTypeName, n))
        except Exception:
            pass
        try:
            for c in ctrl.GetChildren():
                collect_items(c, depth + 1)
        except Exception:
            pass
    collect_items(keyin)
    print("items count", len(items))
    for it in items[:200]:
        print(" ", it)

print("done")
