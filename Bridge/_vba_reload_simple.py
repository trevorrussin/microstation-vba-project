"""Minimal VBA recover + reload -- no Compile (that hung last time)."""
from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
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
            if win32gui.GetWindowText(hwnd) == "Microsoft Visual Basic for Applications":
                hits.append(hwnd)
        return True

    win32gui.EnumWindows(cb, None)
    return hits


def click_ok(hwnd):
    kids = []

    def cb(ch, _):
        if win32gui.GetClassName(ch) == "Button" and win32gui.GetWindowText(ch).strip() == "OK":
            kids.append(ch)
        return True

    win32gui.EnumChildWindows(hwnd, cb, None)
    if kids:
        win32gui.SendMessage(kids[0], win32con.BM_CLICK, 0, 0)


for h in find_vba_error_dialog():
    print("dismiss", h)
    click_ok(h)

app = ms_connect.get_microstation_app("Test")
print("app ok", app.ActiveDesignFile.Name)
vbe = app.VBE

# Reset if possible
try:
    bar = vbe.CommandBars("Standard")
    for i in range(1, bar.Controls.Count + 1):
        c = bar.Controls(i)
        if getattr(c, "Id", None) == 228:
            c.Execute()
            print("Reset ok")
            break
except Exception as e:
    print("Reset skip", e)

project = vbe.VBProjects("Test")
for rel in ("DrawSign.bas", "SignLibrary.bas", "WZTCExec.bas"):
    path = Path(r"c:\repos\microstation-vba-project\Modules") / rel
    try:
        hot_reload.reload_file(project, path)
        n = project.VBComponents(path.stem).CodeModule.CountOfLines
        print(f"OK {rel} lines={n}")
    except Exception as e:
        print(f"FAIL {rel}: {e}")

print("done")
