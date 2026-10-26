"""Drive Place Named Boundary Civil Profile (N5) and Create Drawing (N6/N7) in OpenRoads Designer.

UIA reaches only the ribbon item (InvokePattern). Both dialogs are single MStnChild panes with no named
children, so fields are set by posting WM_LBUTTONDOWN/UP, WM_KEYDOWN and WM_CHAR to the dialog window
(no cursor movement). The two clicks in the profile view are real mouse clicks with ORD in the foreground.
Verified live 2026-10-24 on WZTC_TEST_ALG2 (scratch 727 file): boundary WZTC_TEST_NB_N5, Drawing model id 19,
Sheet model id 20.

Procedure that worked (screen coordinates are for the 1920x1080 layout used then; pass --view-xy to change):
  1. Invoke ribbon item "Place Named Boundary" -> dialog "Place Named Boundary Civil Profile".
  2. Set Drawing Seed (Down+Enter in the combo), names, clearances, untick Use Terrains, Profile Shifts.
     Edits only take effect while the prompt is "Identify Profile View"; mid-placement edits are ignored,
     except Length, which can still be changed at the end-point prompt.
  3. Click the profile in the profile view (Identify Profile View).
  4. Click the Start and Stop arrows next to the Start/Stop boxes: they fill the profile's real station range
     and tick both lock boxes. Do not click a start/stop by hand: a hand click gives a boundary that "will not fit".
  5. Length = the length of ONE sheet, not the whole profile: ORD cuts the span into a group of boundaries (name, name-1, ...), one Drawing and
     one Sheet model each. 17" sheet: 1375 ft is the seed's value at 1"=100' (13.75"), the engineer's CL1 sheet uses 1400 ft (14.0");
     at 1"=40' it is 550 ft. A Length longer than the sheet makes the profile run off the border (my 2500 ft mistake).
  6. Click start, click end, click once more to accept ("Accept/Reject. Data point in Profile View").
  7. Create Drawing opens (Create Drawing + Show Dialog ticked). Its Drawing Model annotation scale ALWAYS opens at 1" = 40' whatever the boundary
     scale (the Sheet Detail Scale follows the boundary): set it to the same scale or the text is drawn at the wrong size. Untick Open Model, press OK.

SCALE (argv[6], default 40): Detail Scale list position from Full Size (SCALE_INDEX): 40 -> 1"=40', 100 -> 1"=100'. Choosing 1"=100' makes the seed
refill Length 1375, VE 5, Available Profile Height 125, clearances 5. Proof that 1" = 100' on the sheet: boundary scale=1200, Drawing annotation
scale 1200 (describe_drawing_state), sheet attachment scale 1/1200 = 0.000833, so 100 ft of drawing is 0.0833 ft = 1.000" on the sheet, and
profile_sheet_verify(detail_scale="1\" = 100'") PASSes. Text at 1:1200 is 6.252 ft (HDM Exhibit 22-27): measured 6.3-6.4 ft on axis and annotation text.
Scale bar (not done by this script): HDM 21.3.9.2.Q needs a graphic scale bar from barscale.cel; the cell for 1"=100' is 1200_Vert (250.47 x 98.45 ft = 2.5" x
0.82"). Placed in the Drawing model with place_cell at origin (grid_right - 208.2, grid_bottom_y - 110.9), grid_right = grid_left + 1400, i.e. in the blank
band between the grid and the title block, right aligned with the grid.
Needs: ribbon workflow NYSDOT Workflow, tab Plan Sheets; view 2 open on the profile (key-in VIEW ON 2, then cif_exec set_profile_view) with Default active in view 1.
WHAT BROKE EARLIER RUNS (2026-10-24) and how main() now handles it:
  - picking a seed reloads every field; edits sent before it finishes are lost -> 2.5 s wait, then set the fields, then check Use Terrains by pixel;
  - Name resets to the next free number when the seed changes, when the profile is identified and when Length changes -> Name is set LAST;
  - the Start/Stop arrow click is sometimes ignored -> fill_limits() compares crops of the two boxes and retries, and refuses to place if they match;
  - Open Model / Use Terrains keep their last state between runs -> read the box colour before clicking, never toggle blind;
  - view 2 must show the profile at a sane zoom: pass "cx,cy,ex,ey" as the 5th argument (view change only).
Verified: dry run (WZTC_NB_DRY=1) showed every field right, and two real runs (N5F in two parts, N5G in one pass) gave PASS from
profile_sheet_verify with the active model unchanged. Always read back with profile_sheet_verify afterwards; a stale Name or clearance is only visible there.
Read back with cif_read profile_boundaries / list_models. Needs: pip install uiautomation pywin32 pillow.
Capture with PIL ImageGrab: PrintWindow returns a blank image for these dialogs.
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

import win32api
import win32con
import win32gui
from PIL import ImageGrab

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "mcp-server"))
import ord_uia  # noqa: E402

NB_TITLE = "Place Named Boundary Civil Profile"
CD_TITLE = "Create Drawing"

# dialog-relative coordinates (screenshot of the dialog window, origin at its top-left)
SCALE_INDEX = {"40": 18, "50": 19, "100": 20}          # position in the Detail Scale list counted from Full Size 1=1
SCALE_FIELDS = {"40": {"height": "50", "clr": "1"},      # NYSDOT 1"=40' tutorial values (PDF page 2)
                "100": {"height": "125", "clr": "5"}}    # seed values for 1"=100' = the engineer's CL1 boundary (height 125, clearance 5)
NB = {"start": (280, 274), "stop": (280, 297), "detail_scale": (280, 112), "seed": (280, 89), "name": (280, 135), "group": (280, 227), "start_arrow": (417, 274),
      "stop_arrow": (417, 297), "length": (280, 322), "ve": (280, 346), "height": (280, 369),
      "top_clr": (280, 392), "bot_clr": (280, 415), "shifts": (280, 484), "use_terrains": (170, 506)}


def _find(title: str):
    out: list[int] = []
    win32gui.EnumWindows(lambda h, _: out.append(h) if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h) == title else None, None)
    return out[0] if out else None


def _child(top: int) -> int:
    out: list[int] = []
    win32gui.EnumChildWindows(top, lambda c, _: out.append(c) if win32gui.GetClassName(c) == "MStnChild" else None, None)
    return out[0] if out else top


def post_click(hwnd: int, cx: int, cy: int) -> None:
    lp = win32api.MAKELONG(cx, cy)
    win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lp)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
    time.sleep(0.05)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lp)
    time.sleep(0.2)


def post_key(hwnd: int, vk: int) -> None:
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, 0)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk, 0xC0000001)
    time.sleep(0.03)


def post_text(hwnd: int, s: str) -> None:
    for ch in s:
        win32gui.PostMessage(hwnd, win32con.WM_CHAR, ord(ch), 0)
        time.sleep(0.02)


def _client(top: int, child: int, x: int, y: int) -> tuple[int, int]:
    wl, wt, _, _ = win32gui.GetWindowRect(top)
    cl, ct, _, _ = win32gui.GetWindowRect(child)
    return x - (cl - wl), y - (ct - wt)


def nb_click(x: int, y: int) -> None:
    top = _find(NB_TITLE)
    ch = _child(top)
    post_click(ch, *_client(top, ch, x, y))


def nb_set(field: str, text: str) -> None:
    x, y = NB[field]
    top = _find(NB_TITLE)
    ch = _child(top)
    for _ in range(2):                       # the first edit after the dialog settles is sometimes dropped (run 4: name and top clearance); the edit is idempotent
        post_click(ch, *_client(top, ch, x, y))
        time.sleep(0.3)
        post_key(ch, win32con.VK_END)
        for _ in range(40):
            post_key(ch, win32con.VK_BACK)
        post_text(ch, text)
        post_key(ch, win32con.VK_TAB)
        time.sleep(0.3)


def checked(x: int, y: int, title: str = NB_TITLE) -> bool:
    """A ticked box is filled blue; read the pixel at dialog-relative (x, y) from the screen."""
    l, t, _r, _b = win32gui.GetWindowRect(_find(title))
    r, g, b = ImageGrab.grab(bbox=(l + x, t + y, l + x + 1, t + y + 1), all_screens=True).getpixel((0, 0))[:3]
    return b > 140 and r < 80


def view_rect(n: int = 2) -> tuple[int, int, int, int]:
    """Screen rect of the drawing area of view n (the MStnViewport child 'VpWindow n')."""
    out: list[int] = []

    def cb(h, _):
        if win32gui.GetClassName(h) == "MStnViewport" and win32gui.GetWindowText(h) == f"VpWindow {n}":
            out.append(h)
        return True

    def top(h, _):
        win32gui.EnumChildWindows(h, cb, None)
        return True

    win32gui.EnumWindows(top, None)
    if not out:
        raise RuntimeError(f"view {n} window not found")
    return win32gui.GetWindowRect(out[0])


def find_red_on_screen(rect: tuple[int, int, int, int]) -> tuple[int, int]:
    """A pixel on the red profile line in the view: the profile is drawn in pure red (255,0,0)."""
    l, t, r, b = rect
    img = ImageGrab.grab(bbox=(l, t, r, b), all_screens=True).convert("RGB")
    w, h = img.size
    pts = [(x, y) for y in range(0, h, 2) for x in range(0, w, 2) if (lambda p: p[0] > 200 and p[1] < 60 and p[2] < 60)(img.getpixel((x, y)))]
    if not pts:
        raise RuntimeError("no red profile line visible in the view: pan/zoom view 2 over the profile first")
    pts.sort()
    x, y = pts[len(pts) // 2]
    return l + x, t + y


def real_click(x: int, y: int) -> None:
    win32api.SetCursorPos((x, y))
    time.sleep(0.6)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.08)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(1.2)


def _find_ctrl(main, name: str, ctype: str):
    hit: list = []

    def find(c, d=0):
        if hit or d > 14:
            return
        try:
            if c.Name == name and c.ControlTypeName == ctype:
                hit.append(c)
                return
            for k in c.GetChildren():
                find(k, d + 1)
        except Exception:
            pass

    find(main)
    return hit[0] if hit else None


def open_nb_dialog() -> int | None:
    ord_uia.activate_ord()
    main = ord_uia.find_ord_main()
    item = _find_ctrl(main, "Place Named Boundary", "MenuItemControl")
    if item is None:                         # the dropdown's items only exist after it has been opened once (ribbon must be NYSDOT Workflow > Plan Sheets)
        grp = _find_ctrl(main, "Named Boundaries", "GroupControl")      # the big button has no UIA name; its group does
        if grp is None:
            raise RuntimeError("ribbon group 'Named Boundaries' not found: switch the workflow to NYSDOT Workflow, tab Plan Sheets")
        r = grp.BoundingRectangle
        real_click((r.left + r.right) // 2, r.top + 52)                 # lower half of the big button = its dropdown arrow
        time.sleep(1.2)
        win32api.keybd_event(0x1B, 0, 0, 0); win32api.keybd_event(0x1B, 0, 2, 0); time.sleep(0.5)
        item = _find_ctrl(main, "Place Named Boundary", "MenuItemControl")
        if item is None:
            raise RuntimeError("menu item 'Place Named Boundary' not found")
    item.GetInvokePattern().Invoke()
    for _ in range(20):
        time.sleep(0.4)
        if _find(NB_TITLE):
            return _find(NB_TITLE)
    return None


def close_nb_dialog() -> None:
    import uiautomation as auto
    h = _find(NB_TITLE)
    if h:
        b = auto.ControlFromHandle(h).ButtonControl(searchDepth=3, Name="Close")
        if b.Exists(2):
            b.GetInvokePattern().Invoke()


def shot(path: str, title: str = NB_TITLE) -> None:
    ImageGrab.grab(bbox=win32gui.GetWindowRect(_find(title)), all_screens=True).save(path)


def cd_click_screen(sx: int, sy: int) -> None:
    h = _find(CD_TITLE)
    cx, cy = win32gui.ScreenToClient(h, (sx, sy))
    post_click(h, cx, cy)


def _crop(x0: int, y0: int, x1: int, y1: int, title: str = NB_TITLE) -> bytes:
    l, t, _r, _b = win32gui.GetWindowRect(_find(title))
    return ImageGrab.grab(bbox=(l + x0, t + y0, l + x1, t + y1), all_screens=True).tobytes()


START_BOX, STOP_BOX = (165, 264, 395, 284), (165, 288, 395, 308)     # dialog-relative text boxes next to the arrows


def fill_limits(tries: int = 3) -> bool:
    """Click the Start then Stop arrow until both boxes changed and show different text (the arrow click is sometimes ignored)."""
    for arrow, box in (("start_arrow", START_BOX), ("stop_arrow", STOP_BOX)):
        for _ in range(tries):
            before = _crop(*box)
            nb_click(*NB[arrow]); time.sleep(1.5)
            if _crop(*box) != before:
                break
        else:
            return False
    return _crop(*START_BOX) != _crop(*STOP_BOX)


def set_view2(cx: float, cy: float, ex: float, ey: float) -> None:
    """View 2 center and extents (view change only; nothing in the model changes)."""
    import ms_connect
    v = ms_connect.get_microstation_app().ActiveDesignFile.Views(2)
    e = v.Extents; e.X = ex; e.Y = ey; e.Z = 0; v.Extents = e
    c = v.Center; c.X = cx; c.Y = cy; v.Center = c
    v.Redraw()
    time.sleep(1.5)


def main(argv: list[str]) -> int:
    """python scripts/ord_named_boundary.py <name> <profile_xy> <accept_xy> <length>   e.g. WZTC_TEST_NB_N5 1390,310 1600,450 2500
    Needs an open QA task on a scratch file. Runs N5 then Create Drawing (Open Model off, OK). Returns 0 when Create Drawing closed."""
    name, prof, acc, length = argv[1], argv[2], argv[3], argv[4]
    if len(argv) > 5 and argv[5] != "-":     # optional "cx,cy,ex,ey": put view 2 at a sane zoom over the profile first (view change only)
        set_view2(*map(float, argv[5].split(",")))
    scale = argv[6] if len(argv) > 6 else "40"          # horizontal scale: 1" = <scale>'
    os.environ["WZTC_NB_SCALE"] = scale
    ord_uia.activate_ord(); time.sleep(0.5)
    vr = view_rect(2)
    px, py = find_red_on_screen(vr) if prof == "auto" else map(int, prof.split(","))      # a pixel on the red profile line
    if acc == "auto":                        # start, end and accept clicks in the empty lower part of the view
        ax, ay = vr[0] + (vr[2] - vr[0]) // 5, vr[1] + (vr[3] - vr[1]) * 3 // 4
    else:
        ax, ay = map(int, acc.split(","))
    open_nb_dialog()
    time.sleep(1.5)
    top = _find(NB_TITLE)
    ch = _child(top)
    nb_click(*NB["seed"]); time.sleep(0.8)   # settings persist between runs: Home -> (none), Down -> ny_Profile 40 Scale
    post_key(ch, win32con.VK_HOME); post_key(ch, win32con.VK_DOWN); post_key(ch, win32con.VK_RETURN)
    time.sleep(2.5)                          # picking a seed reloads every field from the seed; edits made before it finishes are lost
    nb_click(*NB["detail_scale"]); time.sleep(0.8)                      # Detail Scale list: Full Size, 3/4", ... 1"=1' ... 1"=40', 1"=50', 1"=100'
    post_key(ch, win32con.VK_HOME)
    for _ in range(SCALE_INDEX[scale]):
        post_key(ch, win32con.VK_DOWN)
    post_key(ch, win32con.VK_RETURN); time.sleep(2.5)                   # the seed refills Length, VE, height and clearances for that scale
    p = SCALE_FIELDS[scale]
    nb_set("top_clr", p["clr"]); nb_set("bot_clr", p["clr"]); nb_set("ve", "5"); nb_set("height", p["height"])
    for _ in range(3):                       # Use Terrains must be off
        if not checked(*NB["use_terrains"]):
            break
        nb_click(*NB["use_terrains"]); time.sleep(0.6)
    else:
        raise RuntimeError("Use Terrains is still ticked")
    if os.environ.get("WZTC_NB_SHIFTS", "datum") == "datum":      # WZTC_NB_SHIFTS=needed keeps the seed's "Where Needed"
        for _ in range(2):                   # first entry of the list = Datum Stations
            nb_click(*NB["shifts"]); time.sleep(1.0)
            post_key(ch, win32con.VK_HOME); post_key(ch, win32con.VK_RETURN); time.sleep(1.0)
    ord_uia.activate_ord(); time.sleep(0.5)
    real_click(px, py)                       # Identify Profile View
    if not fill_limits():
        raise RuntimeError("Start/Stop arrows did not fill two different locations; not placing")
    if os.environ.get("WZTC_NB_START"):      # type a round station instead of the profile's own start, e.g. 66+00
        nb_set("start", os.environ["WZTC_NB_START"])
    if os.environ.get("WZTC_NB_STOP"):
        nb_set("stop", os.environ["WZTC_NB_STOP"])
    nb_set("length", length)
    nb_set("name", name)                     # Name is reset by the seed, by identifying the profile and by changing Length: set it last
    if os.environ.get("WZTC_NB_DRY"):        # stop before placing; the dialog is left open for a screenshot
        return 3
    return place_and_create(ax, ay)


def place_and_create(ax: int, ay: int) -> int:
    """With the dialog filled and Start/Stop set: three clicks in the profile view, then Create Drawing (Open Model off, OK)."""
    real_click(ax, ay)                       # start
    real_click(ax + 300, ay)                 # end
    real_click(ax, ay)                       # accept
    for _ in range(20):                      # Create Drawing opens by itself (Create Drawing + Show Dialog ticked)
        if _find(CD_TITLE):
            break
        time.sleep(0.5)
    else:
        return 1
    time.sleep(1.0)
    l, t, _r, _b = win32gui.GetWindowRect(_find(CD_TITLE))
    scale = os.environ.get("WZTC_NB_SCALE", "40")
    if scale != "40":                        # the Drawing Model annotation scale always opens at 1" = 40' whatever the boundary scale: set it, or the text is drawn at the wrong size
        h = _find(CD_TITLE)
        cd_click_screen(l + 264, t + 327); time.sleep(1.0)
        post_key(h, win32con.VK_HOME)
        for _ in range(SCALE_INDEX[scale]):
            post_key(h, win32con.VK_DOWN)
        post_key(h, win32con.VK_RETURN); time.sleep(1.5)
    if os.environ.get("WZTC_CD_DRY"):        # stop with Create Drawing open and save a picture of it
        ImageGrab.grab(bbox=(l, t, _r, _b), all_screens=True).save(os.environ["WZTC_CD_DRY"])
        return 4
    if checked(156, 613, CD_TITLE):          # Open Model must be OFF: ticked, Create Drawing switches the active model. State persists between runs
        cd_click_screen(l + 156, t + 613)
        time.sleep(0.6)
        if checked(156, 613, CD_TITLE):
            raise RuntimeError("Open Model is still ticked; not pressing OK")
    cd_click_screen(l + 288, t + 635)        # OK
    for _ in range(40):
        if not _find(CD_TITLE):
            return 0
        time.sleep(0.5)
    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
