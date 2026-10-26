"""N4: Element Annotation of a profile with the Profile Annotation group, driven through the Annotate Element tool settings.

The managed route (PersistenceAnnotationProcessor.Annotate via the add-in's annotate_profile op) returns status 0 but draws nothing.
The tool itself works: arm it with the key-in, set its tool settings, locate the profile, Reset. Verified live 2026-10-24 on WZTC_TEST_ALG2
(PVC/PVI/PVT label blocks appeared at both vertical curves).

Steps (the settings window is a WinForms host with real BUTTON/STATIC children, so BM_CLICK / BM_GETCHECK work; no UIA fields):
  1. Key-in GEOMETRY ANNOTATE ELEMENT (CadInputQueue.SendKeyin) -> tool "Annotate Element" with a settings window.
  2. Untick "All Elements in Model", tick "Override Annotation Group".
  3. The "..." button opens a modal tree picker (PostMessage, SendMessage would block): Profile > Linear > Profile Annotation, Enter.
  4. Click the profile in the profile view (real click, ORD in front), then Reset (right click) to apply.
Needs an open QA task on a scratch file and view 2 showing the profile (cif_exec set_profile_view, then ord_named_boundary.set_view2).
Usage: python scripts/ord_annotate_profile.py <profile_click_x,y>        e.g. 1390,311
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

import win32api
import win32con
import win32gui

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "mcp-server"))
import ms_connect  # noqa: E402
import ord_named_boundary as nb  # noqa: E402

GROUP = ("Profile", "Linear", "Profile Annotation")


def settings_hwnd() -> int | None:
    out: list[int] = []
    win32gui.EnumWindows(lambda h, _: out.append(h) if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h).startswith("Annotate") else None, None)
    return out[0] if out else None


def _children(top: int) -> list[tuple[int, str, str, tuple]]:
    rows: list[tuple[int, str, str, tuple]] = []
    win32gui.EnumChildWindows(top, lambda c, _: rows.append((c, win32gui.GetClassName(c), win32gui.GetWindowText(c), win32gui.GetWindowRect(c))) if win32gui.IsWindowVisible(c) else None, None)
    return rows


def button_for(top: int, label: str) -> int:
    """The BUTTON (checkbox) on the same row, to the right of the STATIC with this text."""
    rows = _children(top)
    st = next(r for r in rows if "STATIC" in r[1] and r[2] == label)
    cands = [r for r in rows if "BUTTON" in r[1] and abs(r[3][1] - st[3][1]) < 8 and r[3][0] > st[3][0]]
    return min(cands, key=lambda r: r[3][0])[0]


def is_checked(btn: int) -> bool:
    """These checkboxes are owner-drawn: BM_GETCHECK always returns 0. A ticked box is filled blue at its centre."""
    from PIL import ImageGrab
    l, t, r, b = win32gui.GetWindowRect(btn)
    cx, cy = (l + r) // 2, (t + b) // 2
    img = ImageGrab.grab(bbox=(cx - 5, cy - 5, cx + 6, cy + 6), all_screens=True)    # the white tick sits at the centre: sample the corners of the glyph
    blue = [img.getpixel(p)[:3] for p in ((1, 1), (9, 1), (1, 9), (9, 9))]
    return sum(1 for px in blue if px[2] > 140 and px[0] < 80) >= 2


def set_check(top: int, label: str, want: bool) -> None:
    b = button_for(top, label)
    if is_checked(b) != want:
        win32gui.SendMessage(b, win32con.BM_CLICK, 0, 0)
        for _ in range(8):                                     # the window repaints and rebuilds its rows after a toggle
            time.sleep(0.6)
            try:
                if is_checked(button_for(settings_hwnd(), label)) == want:
                    return
            except StopIteration:
                pass
        raise RuntimeError(f"could not set '{label}' to {want}")


def pick_group() -> None:
    top = settings_hwnd()
    rows = _children(top)
    st = next(r for r in rows if "STATIC" in r[1] and r[2] == "Annotation Group")
    btn = max((r for r in rows if "BUTTON" in r[1] and abs(r[3][1] - st[3][1]) < 8 and r[3][0] > st[3][0]), key=lambda r: r[3][0])[0]
    win32gui.PostMessage(btn, win32con.BM_CLICK, 0, 0)       # modal picker: SendMessage would block until it closes
    time.sleep(2.0)
    import uiautomation as auto
    for part in GROUP[:-1]:                                    # expand each folder with the keyboard
        auto.SendKeys("{DOWN}{RIGHT}", waitTime=0.3)
        time.sleep(0.4)
    auto.SendKeys("{DOWN}{ENTER}", waitTime=0.3)
    time.sleep(1.5)


def main(argv: list[str]) -> int:
    px, py = map(int, argv[1].split(","))
    app = ms_connect.get_microstation_app()
    nb.ord_uia.activate_ord()
    time.sleep(0.5)
    app.CadInputQueue.SendReset()                              # a tool left mid-locate ignores its settings window
    time.sleep(1.0)
    app.CadInputQueue.SendKeyin("GEOMETRY ANNOTATE ELEMENT")
    for _ in range(20):
        if settings_hwnd():
            break
        time.sleep(0.4)
    top = settings_hwnd()
    if not top:
        return 1
    time.sleep(1.0)
    set_check(top, "All Elements in Model", False)
    set_check(top, "Override Annotation Group", True)
    pick_group()
    nb.ord_uia.activate_ord()
    time.sleep(0.5)
    nb.real_click(px, py)                                      # locate the profile
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0)  # Reset applies the annotation
    time.sleep(0.08)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0)
    time.sleep(3)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
