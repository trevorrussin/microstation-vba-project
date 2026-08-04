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

# Default settle after View.Center/Extents write. Live 2026-08-03: captures
# taken too soon after navigate looked blank (repaint not finished); 2.0s
# was the prior floor, 2.5s is safer under load.
DEFAULT_SETTLE_SECONDS = 2.5


def _find_window(title_predicate) -> int:
    """Returns the hwnd of the visible window whose title satisfies
    title_predicate(title) -> bool. Raises if none match -- and, per the
    2026-08-02 crash incident (see ms_connect.py's docstring), also raises
    if MORE than one matches rather than silently capturing whichever
    happened to enumerate first. Two MicroStation windows (or two chat
    panels) open at once is exactly the ambiguous state that led there;
    surfacing it here is cheap insurance even though this function itself
    doesn't touch COM."""
    matches: list[tuple[int, str]] = []

    def _callback(hwnd: int, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title and title_predicate(title):
                matches.append((hwnd, title))
        return True

    win32gui.EnumWindows(_callback, None)
    if not matches:
        raise RuntimeError("no visible window matched")
    if len(matches) > 1:
        titles = "\n  ".join(t for _, t in matches)
        raise RuntimeError(
            f"{len(matches)} visible windows matched -- can't tell which one you "
            f"mean. Close the extra one(s) first. Matches:\n  {titles}"
        )
    return matches[0][0]


def _find_view_child(main_hwnd: int, view_num: int = 1) -> int | None:
    """Locates the MStnChild window for a specific view (titled e.g. "View 1,
    left lane closure") inside MicroStation's main frame. Its top edge is
    exactly where the per-view mini-toolbar (fit/zoom/pan icons + the view
    title itself) begins -- i.e. exactly where the app-level ribbon/toolbar
    rows above it end. Confirmed live via EnumChildWindows (2026-08-02):
    main frame top toolbar row ("TBxBg2(0)") ends and "View 1, ..." begins
    at the same y. Returns None if not found (e.g. that view isn't open),
    so callers can fall back to an uncropped capture rather than erroring."""
    prefix = f"View {view_num},"
    matches: list[int] = []

    def _callback(hwnd: int, _):
        if win32gui.GetWindowText(hwnd).startswith(prefix):
            matches.append(hwnd)
        return True

    win32gui.EnumChildWindows(main_hwnd, _callback, None)
    return matches[0] if matches else None


def _capture_hwnd(hwnd: int, out_path: str | Path | None, default_name: str,
                   crop_top_px: int = 0) -> Path:
    """Shared PrintWindow -> crop -> resize -> save path for any window
    handle. crop_top_px removes that many pixels off the top of the raw
    capture (before the MAX_LONG_EDGE resize, so the pixel count is measured
    against the real window, not a scaled-down copy) -- see
    _find_view_child for why this is how the top toolbar/ribbon gets
    excluded from a MicroStation capture."""
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

        if crop_top_px > 0:
            img = img.crop((0, min(crop_top_px, img.height - 1), img.width, img.height))

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


def get_view_state(view_num: int = 1) -> dict:
    """Reads the current view's center and extents via COM -- the read
    counterpart to navigate_view() below, letting a caller compute a
    RELATIVE zoom/pan (e.g. "zoom out 40%") instead of only being able to
    set an absolute center/width/height. Confirmed live 2026-08-02:
    View.Extents.X/Y are the current width/height in design units,
    View.Center.X/Y/Z is the current center point -- exactly the shape
    navigate_view() already writes, just read instead of written."""
    import ms_connect
    import pythoncom

    # MCP dispatches each tool call on a possibly-different worker thread
    # (see bridge_client.py's call() for the full explanation) -- COM must
    # be initialized on whichever thread is actually running this.
    pythoncom.CoInitialize()
    app = ms_connect.get_microstation_app()
    v = app.ActiveDesignFile.Views(view_num)
    ext = v.Extents
    ctr = v.Center
    return {
        "centerX": ctr.X, "centerY": ctr.Y, "centerZ": ctr.Z,
        "width": ext.X, "height": ext.Y,
    }


def _view_window_aspect(view_num: int = 1) -> float:
    """Pixel aspect (width/height) of the View N child window. MicroStation
    forces View.Extents to this aspect — a mismatched 200x160 request
    became ~417x160 (live 2026-08-03)."""
    try:
        main = _find_window(lambda t: t.endswith("- MicroStation"))
        child = _find_view_child(main, view_num)
        if child is None:
            return 2.4
        left, top, right, bottom = win32gui.GetWindowRect(child)
        pw, ph = max(right - left, 1), max(bottom - top, 1)
        return pw / ph
    except Exception:
        return 2.4


def _fit_extents_to_aspect(width: float, height: float, aspect: float) -> tuple[float, float]:
    """Expand width or height so the requested model rectangle stays fully
    visible under the view's pixel aspect."""
    width = max(float(width), 1.0)
    height = max(float(height), 1.0)
    aspect = max(float(aspect), 0.05)
    if (width / height) > aspect:
        height = width / aspect
    else:
        width = height * aspect
    return width, height


def _drawing_looks_empty(img: Image.Image) -> bool:
    """True when the drawing band is nearly uniform dark-grey (repaint race
    or nothing visible). Ignores chrome via a central crop. Over-zoomed
    overviews with tiny signs can also trip this — frame tighter (~150–400
    ft for a sign closeup)."""
    w, h = img.size
    if w < 20 or h < 20:
        return True
    region = img.crop((w // 20, h // 10, w - w // 20, max(h // 10 + 1, h - h // 6)))
    small = region.resize((64, 48), Image.BILINEAR)
    pixels = list(small.getdata())
    n = len(pixels)
    if n == 0:
        return True
    bg = 0
    signal = 0
    for pix in pixels:
        r, g, b = pix[0], pix[1], pix[2]
        if abs(r - g) < 14 and abs(g - b) < 14 and 35 <= r <= 120:
            bg += 1
        if r > 170 or g > 170 or abs(r - g) > 30 or abs(g - b) > 30:
            signal += 1
    return signal < 10 and bg >= int(0.80 * n)


def _force_view_redraw(app, view_num: int = 1) -> None:
    """Best-effort repaint after Center/Extents write."""
    v = app.ActiveDesignFile.Views(view_num)
    try:
        v.Redraw()
    except Exception:
        pass
    try:
        app.RedrawAllViews()
    except Exception:
        pass
    try:
        v.Update()
    except Exception:
        pass


def navigate_view(x: float, y: float, width: float, height: float,
                   z: float = 0.0, view_num: int = 1,
                   settle_seconds: float | None = None) -> dict:
    """Point a MicroStation view at a model-space location before capture.

    MicroStation fit/zoom keyins can't complete headlessly (datapoint wait).
    COM Center/Extents works (live 2026-08-02), with these rules:
      1. Write Extents.Z=0 for 2D (install may round to ~0.02 — OK).
      2. Repaint is async — sleep settle_seconds (default 2.5) before capture.
      3. Extents are expanded to the view window's pixel aspect so the
         requested rectangle stays fully visible; return value reports what
         was applied.
      4. Framing: a sign face is ~4–50 ft. A 2000-ft overview makes each
         sign a few pixels — captures look 'blank' to vision models even
         when geometry exists. Use ~150–400 ft width for sign closeups.

    Returns the applied view state dict (center/width/height).
    """
    import time

    import ms_connect
    import pythoncom

    if settle_seconds is None:
        settle_seconds = DEFAULT_SETTLE_SECONDS

    # MCP dispatches each tool call on a possibly-different worker thread --
    # see bridge_client.py's call() for the full explanation.
    pythoncom.CoInitialize()
    app = ms_connect.get_microstation_app()
    is_3d = app.ActiveModelReference.Is3D
    v = app.ActiveDesignFile.Views(view_num)

    aspect = _view_window_aspect(view_num)
    width, height = _fit_extents_to_aspect(width, height, aspect)

    z_extent = 0.0 if not is_3d else max(width, height, 1.0)
    # Center first, then Extents — more reliable when jumping a long way
    # (live 2026-08-03 blank-capture chase).
    v.Center = app.Point3dFromXYZ(x, y, z)
    v.Extents = app.Point3dFromXYZ(width, height, z_extent)
    _force_view_redraw(app, view_num)
    time.sleep(settle_seconds)

    state = get_view_state(view_num=view_num)
    if abs(state["centerX"] - x) > 1.0 or abs(state["centerY"] - y) > 1.0:
        v.Center = app.Point3dFromXYZ(x, y, z)
        v.Extents = app.Point3dFromXYZ(width, height, z_extent)
        _force_view_redraw(app, view_num)
        time.sleep(settle_seconds)
        state = get_view_state(view_num=view_num)
    return state


def capture_microstation(out_path: str | Path | None = None, view_num: int = 1,
                          crop_toolbar: bool = True,
                          retry_if_empty: bool = True) -> Path:
    """Screenshots MicroStation's main frame window -- the one whose title
    ends in "- MicroStation" (confirmed live: the design file path/name is
    the rest of the title, e.g. "...DELETE.dgn [2D - V8 DGN] -
    MicroStation"). Does NOT include separate top-level windows like the
    WZTC chat panel -- use capture_window() for those. Returns the path
    written; defaults to Bridge/captures/capture_live.png (fixed name,
    overwritten each call).

    crop_toolbar (default True, per 2026-08-02 feedback) crops off the
    app-level title bar/ribbon/toolbar rows above view_num's view, keeping
    everything from that view's own mini-toolbar down through the bottom
    toolbar/status area. Falls back to uncropped if that view isn't open.

    retry_if_empty (default True): if the drawing band looks like uniform
    dark grey (repaint race), force a redraw, wait briefly, and capture
    once more. Does not fix over-zoomed framing — see navigate_view.
    """
    import time

    import ms_connect
    import pythoncom

    hwnd = _find_window(lambda t: t.endswith("- MicroStation"))
    crop_top_px = 0
    if crop_toolbar:
        main_top = win32gui.GetWindowRect(hwnd)[1]
        view_hwnd = _find_view_child(hwnd, view_num)
        if view_hwnd is not None:
            crop_top_px = max(0, win32gui.GetWindowRect(view_hwnd)[1] - main_top)
    dest = _capture_hwnd(hwnd, out_path, "capture_live.png", crop_top_px=crop_top_px)

    if retry_if_empty:
        try:
            if _drawing_looks_empty(Image.open(dest)):
                pythoncom.CoInitialize()
                app = ms_connect.get_microstation_app()
                _force_view_redraw(app, view_num)
                time.sleep(1.5)
                dest = _capture_hwnd(hwnd, out_path, "capture_live.png",
                                     crop_top_px=crop_top_px)
        except Exception:
            pass
    return dest


def navigate_and_capture(x: float, y: float, width: float, height: float,
                          out_path: str | Path | None = None,
                          view_num: int = 1,
                          settle_seconds: float | None = None) -> dict:
    """navigate_view + capture_microstation. Returns
    {"path", "view", "retriedEmpty"} for QA scripts.
    """
    state = navigate_view(x, y, width, height, view_num=view_num,
                          settle_seconds=settle_seconds)
    path = capture_microstation(out_path=out_path, view_num=view_num,
                                retry_if_empty=False)
    retried = False
    try:
        if _drawing_looks_empty(Image.open(path)):
            retried = True
            path = capture_microstation(out_path=out_path, view_num=view_num,
                                        retry_if_empty=True)
    except Exception:
        path = capture_microstation(out_path=out_path, view_num=view_num,
                                    retry_if_empty=True)
        retried = True
    return {"path": str(path), "view": state, "retriedEmpty": retried}


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
