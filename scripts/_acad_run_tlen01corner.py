"""Unhang AutoCAD + type TLEN01CORNER so the CLI type-in prompt appears."""
from __future__ import annotations

import time
import ctypes
import win32api
import win32clipboard
import win32con
import win32gui
import win32process

CMD = '(progn (load "c:/repos/microstation-vba-project/scripts/TLEN01CORNER.lsp") (C:TLEN01CORNER))\n'


def find_acad_hwnd():
    found = []

    def cb(h, _):
        if win32gui.IsWindowVisible(h):
            t = win32gui.GetWindowText(h)
            c = win32gui.GetClassName(h)
            if c == "AfxMDIFrame140u" and ("AutoCAD" in t or "Civil" in t):
                found.append((h, t))
        return True

    win32gui.EnumWindows(cb, None)
    return found[0] if found else (None, None)


def focus(hwnd) -> bool:
    if not hwnd:
        return False
    user32 = ctypes.windll.user32
    user32.AllowSetForegroundWindow(-1)
    try:
        fg = win32gui.GetForegroundWindow()
        tid_fg = win32process.GetWindowThreadProcessId(fg)[0]
        tid = win32process.GetWindowThreadProcessId(hwnd)[0]
        user32.AttachThreadInput(tid_fg, tid, True)
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.BringWindowToTop(hwnd)
        win32gui.SetForegroundWindow(hwnd)
        user32.AttachThreadInput(tid_fg, tid, False)
        time.sleep(0.35)
        return win32gui.GetForegroundWindow() == hwnd
    except Exception as ex:
        print("focus_err", ex)
        return False


def tap(vk, times=1, gap=0.06):
    for _ in range(times):
        win32api.keybd_event(vk, 0, 0, 0)
        win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(gap)


def clipboard_set(text: str) -> None:
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


def paste_enter():
    # Ctrl+V
    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(ord("V"), 0, 0, 0)
    win32api.keybd_event(ord("V"), 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.15)
    tap(win32con.VK_RETURN, 1)


def main() -> None:
    hwnd, title = find_acad_hwnd()
    print("hwnd", hwnd, title, flush=True)
    if not hwnd:
        raise SystemExit("AutoCAD window not found")

    ok = focus(hwnd)
    print("focused", ok, flush=True)

    # Cancel hung getkword / GetString / dialog
    print("sending Esc…", flush=True)
    tap(win32con.VK_ESCAPE, 10, 0.08)
    time.sleep(0.4)

    # Re-focus, paste LISP load+run (no COM)
    focus(hwnd)
    clipboard_set(CMD)
    time.sleep(0.1)
    print("pasting TLEN01CORNER…", flush=True)
    paste_enter()
    time.sleep(0.5)

    print("PROMPT READY", flush=True)
    print("On AutoCAD command line type: NW   (or N / SW / E / MidN / MidS)", flush=True)
    print("If nothing happened: click AutoCAD, press Esc, paste/run TLEN01CORNER", flush=True)


if __name__ == "__main__":
    main()
