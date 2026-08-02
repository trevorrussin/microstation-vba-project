"""
OS-level screenshot of a MicroStation-related window -- the working
alternative to driving MicroStation's own raster-export dialogs (CAPTURE
VIEW / DIALOG SAVEIMAGE) headlessly, which turned out not to be viable: live
testing found this install's CAPTURE VIEW never opens a dialog at all, and
DIALOG SAVEIMAGE opens a real "Save Image" dialog whose own "Save" control
opens a second, unknown-command "Save Image As" dialog -- guessing that
command already hung the VBA thread once (see Modules/WZTCViewCapture.bas's
header). This module never touches CadInputQueue, so it can't hang
MicroStation.

Uses Win32 PrintWindow (PW_RENDERFULLCONTENT) rather than a screen-region
grab -- confirmed live that an ImageGrab-based screen-rect approach requires
the window to actually be on top, and SetForegroundWindow silently no-ops
when called from a background process while a different window (e.g. this
editor) is the one the user is actively using, a well-known Windows
foreground-lock restriction. It captured the wrong window as a result.
PrintWindow renders the target window's contents directly into a supplied
device context regardless of stacking order or occlusion, and
PW_RENDERFULLCONTENT (added for exactly this) also handles GPU/DWM-composited
surfaces that older PrintWindow calls render as black.

MicroStation's main frame and the WZTCChatPanel form are separate top-level
windows (the panel is a modeless UserForm, its own OS window, not nested
inside the main frame) -- confirmed live when a capture of the main frame's
title didn't include the panel at all. capture_window() targets any visible
window by a title substring; capture_microstation() is the main-frame-only
convenience wrapper most callers want.
"""
from __future__ import annotations

from ctypes import windll
from pathlib import Path

import win32gui
import win32ui
from PIL import Image

CAPTURES_DIR = Path(r"c:\repos\microstation-vba-project\Bridge\captures")
PW_RENDERFULLCONTENT = 2

# Anthropic resizes any image down to this long-edge before it's processed by
# a vision-capable model, so a raw capture wider than this (the full
# MicroStation window is ~1936px) buys nothing when the image is eventually
# handed to Claude -- it just costs extra bytes/upload time. Resizing here is
# free: same effective resolution Claude would see either way, smaller file.
MAX_LONG_EDGE = 1568


def _find_window(title_predicate) -> int:
    """Returns the hwnd of the first visible window whose title satisfies
    title_predicate(title) -> bool. Raises if none match."""
    matches: list[int] = []

    def _callback(hwnd: int, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and title_predicate(title):
                matches.append(hwnd)
        return True

    win32gui.EnumWindows(_callback, None)
    if not matches:
        raise RuntimeError("no visible window matched")
    return matches[0]


def _capture_hwnd(hwnd: int, out_path: str | Path | None, default_name: str) -> Path:
    """Shared PrintWindow -> resize -> save path for any window handle."""
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()
    save_bitmap = win32ui.CreateBitmap()
    save_bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
    save_dc.SelectObject(save_bitmap)

    try:
        result = windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), PW_RENDERFULLCONTENT)
        if result != 1:
            raise RuntimeError(f"PrintWindow failed (return value {result})")

        bmp_info = save_bitmap.GetInfo()
        bmp_bits = save_bitmap.GetBitmapBits(True)
        img = Image.frombuffer(
            "RGB", (bmp_info["bmWidth"], bmp_info["bmHeight"]), bmp_bits, "raw", "BGRX", 0, 1)

        long_edge = max(img.size)
        if long_edge > MAX_LONG_EDGE:
            scale = MAX_LONG_EDGE / long_edge
            new_size = (round(img.size[0] * scale), round(img.size[1] * scale))
            img = img.resize(new_size, Image.LANCZOS)
    finally:
        win32gui.DeleteObject(save_bitmap.GetHandle())
        save_dc.DeleteDC()
        mfc_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)

    CAPTURES_DIR.mkdir(parents=True, exist_ok=True)
    dest = Path(out_path) if out_path else CAPTURES_DIR / default_name
    img.save(dest, format="PNG")
    return dest


def navigate_view(x: float, y: float, width: float, height: float,
                   z: float = 0.0, view_num: int = 1,
                   settle_seconds: float = 2.0) -> None:
    """Point a MicroStation view at a specific model-space location before
    calling capture_microstation() -- MicroStation's own interactive
    fit/zoom keyins (VIEW_FIT, ZOOM_OUT, etc., via run_registry_command)
    can't complete headlessly, they end by prompting for a datapoint click
    that never arrives. Setting View.Center/Extents directly via COM works
    instead, confirmed live 2026-08-02 (see the "sign face cell oversized
    bbox" investigation), but two things matter:
      1. Extents.Z must be 0 for a 2D model -- a nonzero Z produced a
         completely blank render on this install's 2D DGN. This checks
         ActiveModelReference.Is3D and branches automatically.
      2. The repaint is NOT synchronous with the property write -- a
         screenshot taken immediately after setting Center/Extents came
         back blank in testing. settle_seconds (default 2.0) sleeps before
         returning so a capture_microstation() call right after this one
         shows the real content; the minimum safe delay wasn't narrowed
         down further, 2.0s just confirmed to work.
    3D branch is UNTESTED -- no 3D model exists in this project yet. For
    3D, Extents.Z is set to max(width, height) as a reasonable depth guess
    and camera/perspective settings are left untouched; verify this works
    before relying on it once a 3D file is available.
    """
    import time
    from win32com.client import gencache

    app = gencache.EnsureDispatch("MicroStationDGN.Application")
    is_3d = app.ActiveModelReference.Is3D
    v = app.ActiveDesignFile.Views(view_num)

    z_extent = 0.0 if not is_3d else max(width, height, 1.0)
    v.Extents = app.Point3dFromXYZ(width, height, z_extent)
    v.Center = app.Point3dFromXYZ(x, y, z)
    v.Redraw()
    time.sleep(settle_seconds)


def capture_microstation(out_path: str | Path | None = None) -> Path:
    """Screenshots MicroStation's main frame window -- the one whose title
    ends in "- MicroStation" (confirmed live: the design file path/name is
    the rest of the title, e.g. "...DELETE.dgn [2D - V8 DGN] -
    MicroStation"). Does NOT include separate top-level windows like the
    WZTC chat panel -- use capture_window() for those. Returns the path
    written; defaults to Bridge/captures/capture_live.png (fixed name,
    overwritten each call)."""
    hwnd = _find_window(lambda t: t.endswith("- MicroStation"))
    return _capture_hwnd(hwnd, out_path, "capture_live.png")


def capture_window(title_substring: str, out_path: str | Path | None = None) -> Path:
    """Screenshots the first visible window whose title contains
    title_substring (case-sensitive substring match) -- e.g. "WZTC Agent
    Chat" for the chat panel. Returns the path written; defaults to
    Bridge/captures/capture_window.png (fixed name, overwritten each call)."""
    hwnd = _find_window(lambda t: title_substring in t)
    return _capture_hwnd(hwnd, out_path, "capture_window.png")


if __name__ == "__main__":
    path = capture_microstation()
    print(f"OK: saved {path}")
