"""Dump ROT monikers; try activate ORD hwnd with tsp_02 then re-read ActiveDesignFile."""
from __future__ import annotations

import sys
import time

import pythoncom
import win32gui
import win32process
import win32com.client

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import ms_connect
import com_cache

pythoncom.CoInitialize()

print("=== ALL ROT monikers (bentley/micro/dgn) ===")
rot = pythoncom.GetRunningObjectTable()
ctx = pythoncom.CreateBindCtx(0)
enum = rot.EnumRunning()
while True:
    mons = enum.Next(1)
    if not mons:
        break
    try:
        name = mons[0].GetDisplayName(ctx, None)
    except Exception:
        continue
    low = name.lower()
    if any(k in low for k in ("bentley", "microstation", "dgn", "openroad", "6ba41")):
        print(name[:200])

print("\n=== HWND for OpenRoads with tsp_02 ===")
targets = []


def enum_cb(hwnd, _):
    if not win32gui.IsWindowVisible(hwnd):
        return
    title = win32gui.GetWindowText(hwnd)
    if "OpenRoads" in title and "tsp_02" in title:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        targets.append((hwnd, pid, title))


win32gui.EnumWindows(enum_cb, None)
for hwnd, pid, title in targets:
    print(f"hwnd={hwnd} pid={pid} title={title[:100]}")
    win32gui.SetForegroundWindow(hwnd)
    time.sleep(1.0)

print("\n=== ActiveDesignFile after foreground ===")
for i, (desc, app) in enumerate(ms_connect._candidate_apps()):
    print(i, "HasActive", app.HasActiveDesignFile, "Caption", app.Caption)
    try:
        print("  file", app.ActiveDesignFile.FullName)
        print("  model", app.ActiveModelReference.Name)
    except Exception as e:
        print("  ERR", e)
