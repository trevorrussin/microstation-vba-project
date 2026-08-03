"""Dismiss VBA error dialog if present, Reset, then hot-reload modules."""
from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client
import win32con
import win32gui

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import hot_reload
import ms_connect

pythoncom.CoInitialize()


def find_vba_error_dialog():
    hits = []

    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title == "Microsoft Visual Basic for Applications":
                hits.append(hwnd)
        return True

    win32gui.EnumWindows(cb, None)
    return hits


def click_ok(hwnd):
    kids = []

    def cb(ch, _):
        cls = win32gui.GetClassName(ch)
        txt = win32gui.GetWindowText(ch)
        if cls == "Button" and txt.strip() == "OK":
            kids.append(ch)
        return True

    win32gui.EnumChildWindows(hwnd, cb, None)
    if not kids:
        raise RuntimeError("no OK button")
    win32gui.SendMessage(kids[0], win32con.BM_CLICK, 0, 0)


dialogs = find_vba_error_dialog()
print("error dialogs", len(dialogs))
for h in dialogs:
    try:
        click_ok(h)
        print("clicked OK on", h)
    except Exception as e:
        print("click fail", e)

app = ms_connect.get_microstation_app("Test")
vbe = app.VBE

# Active pane selection for diagnosis
try:
    cp = vbe.ActiveCodePane
    if cp is not None:
        sel = cp.GetSelection()
        print("selection", sel)
        line = sel[0]
        print("line text", repr(cp.CodeModule.Lines(line, 1)))
except Exception as e:
    print("pane", e)

# Reset
bar = vbe.CommandBars("Standard")
for i in range(1, bar.Controls.Count + 1):
    c = bar.Controls(i)
    try:
        if c.Id == 228:
            c.Execute()
            print("Reset executed")
            break
    except Exception:
        pass

# Compile Test
try:
    menu = vbe.CommandBars("Menu Bar")
    for i in range(1, menu.Controls.Count + 1):
        c = menu.Controls(i)
        try:
            cap = c.Caption
        except Exception:
            continue
        if "&Debug" in cap or cap.replace("&", "") == "Debug":
            for j in range(1, c.Controls.Count + 1):
                sub = c.Controls(j)
                try:
                    if sub.Id == 578:
                        print("compiling via", sub.Caption)
                        sub.Execute()
                        break
                except Exception as e:
                    print("compile ctrl", e)
            break
except Exception as e:
    print("compile menu", e)

dialogs = find_vba_error_dialog()
print("error dialogs after compile", len(dialogs))
if dialogs:
    app2 = ms_connect.get_microstation_app("Test")
    try:
        cp = app2.VBE.ActiveCodePane
        sel = cp.GetSelection()
        print("ERR line", sel[0], repr(cp.CodeModule.Lines(sel[0], 3)))
    except Exception as e:
        print("read err", e)
    for h in dialogs:
        click_ok(h)
    # reset again
    for i in range(1, vbe.CommandBars("Standard").Controls.Count + 1):
        c = vbe.CommandBars("Standard").Controls(i)
        try:
            if c.Id == 228:
                c.Execute()
                print("Reset again")
        except Exception:
            pass

# Now hot reload
project = vbe.VBProjects("Test")
for rel in ("DrawSign.bas", "SignLibrary.bas", "WZTCExec.bas"):
    path = Path(r"c:\repos\microstation-vba-project\Modules") / rel
    try:
        hot_reload.reload_file(project, path)
        print("OK", rel, project.VBComponents(path.stem).CodeModule.CountOfLines)
    except Exception as e:
        print("FAIL", rel, e)
