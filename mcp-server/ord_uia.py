"""Windows UI Automation helpers for OpenRoads Designer / MicroStation.

Surface probe + thin helpers that reuse proven patterns:
  - find / foreground ORD (AttachThreadInput + SetForegroundWindow)
  - open Key-in (CadInputQueue DIALOG KEYIN OPEN, else Ctrl+D → Key-in)
  - dismiss classic VBA OK dialogs (BM_CLICK)

Prefer key-ins for activation; do not SetValue on custom Bentley fields
(see Data/ord-ui-catalog.yaml and agent-log 2026-10-14).

Proven elsewhere (reuse, do not reimplement here):
  - mcp-server/view_capture.py — PrintWindow / MStnChild view crop
  - scripts/dump_keyin_browser.py — Key-in OCR (no UIA children)

Non-goals (dead ends — see catalog dead_ends):
  - UIA SetValue on Create Drawing / civil forms
  - UIA scrape of Key-in Browser list (MStnChild_Thin)
  - Make Copy checkbox UIA as the default move fix
  - Clicking inside Level Display level list (chrome_only)
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from ctypes import windll
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Iterable, Optional

try:
    import uiautomation as auto
except ImportError as e:  # pragma: no cover
    raise SystemExit("pip install uiautomation") from e

try:
    import win32con
    import win32gui
    import win32process
except ImportError as e:  # pragma: no cover
    raise SystemExit("pip install pywin32") from e

REPO_ROOT = Path(__file__).resolve().parents[1]
CATALOG_PATH = REPO_ROOT / "Data" / "ord-ui-catalog.yaml"
DEFAULT_REPORT = REPO_ROOT / "outputs" / "uia" / "ord_uia_surface.json"

ORD_PROCESS_HINTS = ("OpenRoads", "MicroStation", "ustation")
ORD_CLASS = "RadRibbonWindow"
VBA_ERROR_TITLES = (
    "Microsoft Visual Basic for Applications",
    "Microsoft Visual Basic",
)


@dataclass
class ControlHit:
    depth: int
    control_type: str
    name: str
    class_name: str = ""
    automation_id: str = ""


@dataclass
class DialogReport:
    name: str
    class_name: str
    node_count: int
    named_controls: list[ControlHit] = field(default_factory=list)
    verdict: str = "empty"  # useful | chrome_only | empty


# --- window discovery ---------------------------------------------------------


def find_ord_main() -> Optional[auto.Control]:
    """Return the main ORD/MicroStation RadRibbonWindow, if running."""
    for w in auto.GetRootControl().GetChildren():
        try:
            cls = w.ClassName or ""
            name = w.Name or ""
        except Exception:
            continue
        if cls == ORD_CLASS:
            return w
        low = name.lower()
        if any(h.lower() in low for h in ORD_PROCESS_HINTS) or "fea_rwy" in low or ".dgn" in low:
            if cls == ORD_CLASS or "Ribbon" in cls:
                return w
    return None


def find_ord_hwnd() -> Optional[int]:
    """Win32 hwnd for the ORD drawing frame (RadRibbonWindow / OpenRoads title)."""
    found: list[tuple[int, str]] = []

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        try:
            title = win32gui.GetWindowText(hwnd) or ""
            cls = win32gui.GetClassName(hwnd) or ""
        except Exception:
            return True
        if cls == ORD_CLASS or (
            "OpenRoads" in title and "Visual Basic" not in title
        ) or (
            "MicroStation" in title and "Visual Basic" not in title
        ):
            found.append((hwnd, title))
        return True

    win32gui.EnumWindows(cb, None)
    if not found:
        return None
    # Prefer RadRibbonWindow / pw: ProjectWise titles over bare helpers
    for hwnd, title in found:
        cls = win32gui.GetClassName(hwnd)
        if cls == ORD_CLASS:
            return hwnd
    return found[0][0]


def activate_ord(hwnd: Optional[int] = None, settle_s: float = 0.35) -> dict[str, Any]:
    """Bring ORD to the foreground. Uses AttachThreadInput (plain SetForeground often fails)."""
    hwnd = hwnd or find_ord_hwnd()
    if not hwnd:
        return {"ok": False, "error": "ORD hwnd not found"}
    user32 = windll.user32
    try:
        user32.AllowSetForegroundWindow(-1)
    except Exception:
        pass
    fg = win32gui.GetForegroundWindow()
    tid_fg = win32process.GetWindowThreadProcessId(fg)[0] if fg else 0
    tid = win32process.GetWindowThreadProcessId(hwnd)[0]
    attached = False
    try:
        if tid_fg and tid_fg != tid:
            attached = bool(user32.AttachThreadInput(tid_fg, tid, True))
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        ok = bool(win32gui.SetForegroundWindow(hwnd))
        try:
            ctrl = auto.ControlFromHandle(hwnd)
            ctrl.SetActive()
        except Exception:
            pass
    finally:
        if attached:
            user32.AttachThreadInput(tid_fg, tid, False)
    time.sleep(settle_s)
    title = ""
    try:
        title = win32gui.GetWindowText(hwnd)
    except Exception:
        pass
    return {
        "ok": True,
        "hwnd": int(hwnd),
        "set_foreground": ok,
        "title": title[:160],
        "foreground_now": int(win32gui.GetForegroundWindow() or 0) == int(hwnd),
    }


def iter_owned_windows() -> Iterable[auto.Control]:
    """ORD main window plus owned dialog windows (often children of RadRibbonWindow)."""
    main = find_ord_main()
    if main is None:
        return
    yield main
    try:
        kids = main.GetChildren()
    except Exception:
        kids = []
    for w in kids:
        try:
            ctype = w.ControlTypeName
            name = (w.Name or "").strip()
            cls = w.ClassName or ""
        except Exception:
            continue
        if ctype != "WindowControl":
            continue
        if name or cls.startswith("WindowsForms") or cls.startswith("Thunder"):
            yield w
    for w in auto.GetRootControl().GetChildren():
        try:
            name = (w.Name or "").strip()
            cls = w.ClassName or ""
        except Exception:
            continue
        if w is main:
            continue
        if name and (
            cls.startswith("WindowsForms10")
            or cls.startswith("Thunder")
            or name == "Key-in"
            or name.startswith("Key-in")
            or "Barmenu" in name
        ):
            yield w


# --- surface probe ------------------------------------------------------------


def _walk_named(
    ctrl: auto.Control,
    depth: int,
    max_depth: int,
    out: list[ControlHit],
    max_nodes: int,
    node_count: list[int],
) -> None:
    if depth > max_depth or node_count[0] >= max_nodes:
        return
    try:
        name = (ctrl.Name or "").strip()
        ctype = ctrl.ControlTypeName
        cls = (ctrl.ClassName or "")[:60]
        aid = ""
        try:
            aid = (ctrl.AutomationId or "")[:40]
        except Exception:
            pass
    except Exception:
        return
    node_count[0] += 1
    interesting_types = (
        "ButtonControl",
        "EditControl",
        "ComboBoxControl",
        "ListControl",
        "ListItemControl",
        "TreeItemControl",
        "CheckBoxControl",
        "TabItemControl",
        "MenuItemControl",
        "HyperlinkControl",
        "DataItemControl",
        "TextControl",
        "ToolBarControl",
    )
    if name and (ctype in interesting_types or depth <= 1):
        out.append(
            ControlHit(
                depth=depth,
                control_type=ctype,
                name=name[:120],
                class_name=cls,
                automation_id=aid,
            )
        )
    try:
        children = ctrl.GetChildren()
    except Exception:
        return
    for ch in children[:80]:
        _walk_named(ch, depth + 1, max_depth, out, max_nodes, node_count)


def classify_dialog(ctrl: auto.Control, max_depth: int = 6) -> DialogReport:
    name = (ctrl.Name or "").strip() or "(unnamed)"
    cls = ctrl.ClassName or ""
    hits: list[ControlHit] = []
    node_count = [0]
    _walk_named(ctrl, 0, max_depth, hits, max_nodes=400, node_count=node_count)

    chrome_names = {"Minimize", "Maximize", "Close", "System", "WindowTitle"}
    content = [
        h
        for h in hits
        if h.name not in chrome_names
        and h.control_type
        in (
            "ButtonControl",
            "EditControl",
            "ComboBoxControl",
            "ListControl",
            "ListItemControl",
            "TreeItemControl",
            "CheckBoxControl",
            "TabItemControl",
            "DataItemControl",
        )
    ]
    if content:
        verdict = "useful"
    elif node_count[0] <= 5:
        verdict = "empty"
    else:
        verdict = "chrome_only"

    if not content and any("MStnChild" in (h.class_name or "") for h in hits):
        verdict = "chrome_only"

    return DialogReport(
        name=name[:160],
        class_name=cls[:80],
        node_count=node_count[0],
        named_controls=content[:40],
        verdict=verdict,
    )


def find_window_by_name(name: str) -> Optional[auto.Control]:
    """Find a window by exact or prefix Name (desktop or ORD-owned)."""
    for top in auto.GetRootControl().GetChildren():
        try:
            n = top.Name or ""
        except Exception:
            continue
        if n == name or n.startswith(name):
            return top
    main = find_ord_main()
    if main is None:
        return None
    try:
        for ch in main.GetChildren():
            n = ch.Name or ""
            if n == name or n.startswith(name):
                return ch
    except Exception:
        pass
    return None


def ribbon_tab_names(main: Optional[auto.Control] = None) -> list[str]:
    main = main or find_ord_main()
    if main is None:
        return []
    tabs: list[str] = []
    hits: list[ControlHit] = []
    _walk_named(main, 0, 4, hits, max_nodes=300, node_count=[0])
    for h in hits:
        if h.control_type == "TabItemControl" and h.name:
            tabs.append(h.name)
    return tabs


def probe_surface(max_dialogs: int = 30) -> dict[str, Any]:
    """Build a JSON-serializable surface map of the live ORD UI."""
    main = find_ord_main()
    report: dict[str, Any] = {
        "main_window": None,
        "dialogs": [],
        "ribbon_tabs": [],
        "summary": {},
    }
    if main is None:
        report["summary"] = {"error": "ORD main window not found (is OpenRoadsDesigner running?)"}
        return report

    report["main_window"] = {
        "name": (main.Name or "")[:160],
        "class": main.ClassName or "",
        "hwnd": find_ord_hwnd(),
    }
    report["ribbon_tabs"] = ribbon_tab_names(main)

    seen: set[str] = set()
    useful = chrome = empty = 0
    for w in iter_owned_windows():
        try:
            key = f"{w.Name}|{w.ClassName}"
        except Exception:
            continue
        if key in seen:
            continue
        seen.add(key)
        dlg = classify_dialog(w)
        report["dialogs"].append(
            {
                "name": dlg.name,
                "class": dlg.class_name,
                "node_count": dlg.node_count,
                "verdict": dlg.verdict,
                "named_controls": [asdict(c) for c in dlg.named_controls],
            }
        )
        if dlg.verdict == "useful":
            useful += 1
        elif dlg.verdict == "chrome_only":
            chrome += 1
        else:
            empty += 1
        if len(report["dialogs"]) >= max_dialogs:
            break

    report["summary"] = {
        "dialog_count": len(report["dialogs"]),
        "useful": useful,
        "chrome_only": chrome,
        "empty": empty,
        "ribbon_tab_count": len(report["ribbon_tabs"]),
        "catalog": str(CATALOG_PATH),
    }
    return report


def write_probe(path: Optional[Path] = None) -> Path:
    path = path or DEFAULT_REPORT
    path.parent.mkdir(parents=True, exist_ok=True)
    data = probe_surface()
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    md = path.with_suffix(".md")
    lines = [
        "# ORD UI Automation surface",
        "",
        f"- Main: `{data.get('main_window')}`",
        f"- Ribbon tabs: {', '.join(data.get('ribbon_tabs') or [])}",
        f"- Summary: {data.get('summary')}",
        "",
        "## Dialogs",
        "",
    ]
    for d in data.get("dialogs") or []:
        lines.append(
            f"- **{d['name']}** — `{d['verdict']}` "
            f"(nodes={d['node_count']}, named={len(d.get('named_controls') or [])})"
        )
    md.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path


# --- Key-in -------------------------------------------------------------------


def find_keyin_hwnd() -> Optional[int]:
    found: list[int] = []

    def enum_tl(h, _):
        try:
            if (win32gui.GetWindowText(h) or "") == "Key-in":
                found.append(h)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(enum_tl, None)
    return found[0] if found else None


def open_keyin(settle_s: float = 1.0) -> dict[str, Any]:
    """Open the Key-in Browser. Does not type CAD commands into it.

    Uses Ctrl+D ribbon search → type Key-in → Enter (same as dump_keyin_browser.py).
    F9 is often remapped by AECOM Barmenu — do not rely on it.

    CadInputQueue is intentionally NOT used here so the QA pretool hook can run
    this module without a drafting task. Agents with an open QA task may still
    send ``DIALOG KEYIN OPEN`` via the bridge / ms_connect separately.
    """
    hwnd = find_keyin_hwnd()
    if hwnd and win32gui.IsWindowVisible(hwnd):
        return {"ok": True, "hwnd": int(hwnd), "method": "already_open"}

    steps: list[dict[str, Any]] = []
    act = activate_ord()
    steps.append({"activate": act})
    time.sleep(0.3)
    auto.SendKeys("{ESC}")
    time.sleep(0.15)
    auto.SendKeys("^d")
    time.sleep(0.9)
    auto.SendKeys("^a")
    time.sleep(0.05)
    auto.SendKeys("Key-in")
    time.sleep(1.0)
    auto.SendKeys("{ENTER}")
    time.sleep(settle_s)
    hwnd = find_keyin_hwnd()
    if not hwnd:
        auto.SendKeys("{DOWN}{ENTER}")
        time.sleep(1.0)
        hwnd = find_keyin_hwnd()
    if not hwnd:
        return {"ok": False, "error": "Key-in failed to open", "steps": steps}
    return {"ok": True, "hwnd": int(hwnd), "method": "Ctrl+D Key-in", "steps": steps}


# --- VBA OK dismiss -----------------------------------------------------------


def list_vba_dialogs() -> list[tuple[int, str]]:
    out: list[tuple[int, str]] = []

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd) or ""
        if any(t in title for t in VBA_ERROR_TITLES) or title == "Microsoft Visual Basic for Applications":
            out.append((hwnd, title))
        return True

    win32gui.EnumWindows(cb, None)
    return out


def click_button(dialog_hwnd: int, caption: str = "OK") -> bool:
    """BM_CLICK a classic Win32 Button child (works for VBA error OK)."""
    found: list[int] = []

    def cb(hwnd, _):
        try:
            cls = win32gui.GetClassName(hwnd)
            txt = (win32gui.GetWindowText(hwnd) or "").strip()
        except Exception:
            return True
        if cls == "Button" and txt == caption:
            found.append(hwnd)
        return True

    win32gui.EnumChildWindows(dialog_hwnd, cb, None)
    if not found:
        return False
    win32gui.SendMessage(found[0], win32con.BM_CLICK, 0, 0)
    return True


def dismiss_vba_ok(caption: str = "OK") -> dict[str, Any]:
    """Dismiss visible VBA error dialogs via BM_CLICK OK."""
    dialogs = list_vba_dialogs()
    # Prefer exact bare error title
    targets = [h for h, t in dialogs if t == "Microsoft Visual Basic for Applications"]
    if not targets:
        targets = [h for h, _ in dialogs]
    clicked: list[int] = []
    for hwnd in targets:
        if click_button(hwnd, caption=caption):
            clicked.append(int(hwnd))
            time.sleep(0.2)
    return {
        "ok": bool(clicked) or not dialogs,
        "found": [{"hwnd": h, "title": t} for h, t in dialogs],
        "clicked": clicked,
        "nothing_to_dismiss": not dialogs,
    }


def find_vba_ide_hwnd() -> Optional[int]:
    """VBA IDE main window (title usually contains 'Visual Basic' and a .mvba path)."""
    best: Optional[int] = None
    for hwnd, title in list_vba_dialogs():
        if title == "Microsoft Visual Basic for Applications":
            continue  # error/message box, not the IDE
        if "Visual Basic" in title:
            if best is None or ".mvba" in title.lower():
                best = hwnd
            if "[break]" in title.lower() or "[running]" in title.lower():
                return hwnd
    return best


def vba_ide_state() -> dict[str, Any]:
    """Parse idle / break / running from the VBA IDE title bar."""
    hwnd = find_vba_ide_hwnd()
    title = win32gui.GetWindowText(hwnd) if hwnd else ""
    low = title.lower()
    state = "missing"
    if hwnd:
        if "[break]" in low:
            state = "break"
        elif "[running]" in low:
            state = "running"
        else:
            state = "idle"
    return {"hwnd": hwnd, "title": title[:200], "state": state}


def _activate_hwnd(hwnd: int, settle_s: float = 0.35) -> bool:
    user32 = windll.user32
    try:
        user32.AllowSetForegroundWindow(-1)
    except Exception:
        pass
    fg = win32gui.GetForegroundWindow()
    tid_fg = win32process.GetWindowThreadProcessId(fg)[0] if fg else 0
    tid = win32process.GetWindowThreadProcessId(hwnd)[0]
    attached = False
    ok = False
    try:
        if tid_fg and tid_fg != tid:
            attached = bool(user32.AttachThreadInput(tid_fg, tid, True))
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        ok = bool(win32gui.SetForegroundWindow(hwnd))
        try:
            auto.ControlFromHandle(hwnd).SetActive()
        except Exception:
            pass
    finally:
        if attached:
            user32.AttachThreadInput(tid_fg, tid, False)
    time.sleep(settle_s)
    return ok or int(win32gui.GetForegroundWindow() or 0) == int(hwnd)


def _send_keys_unicode(keys: str) -> None:
    """Send a short key sequence to the foreground window (e.g. '%r' Alt+R, 'r')."""
    # Prefer SendKeys via win32com if available; fall back to keybd_event.
    try:
        import win32com.client  # type: ignore

        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys(keys)
        return
    except Exception:
        pass
    import win32api

    # Minimal: keys like %r = Alt+R then r for Reset
    i = 0
    while i < len(keys):
        ch = keys[i]
        if ch == "%" and i + 1 < len(keys):
            vk = ord(keys[i + 1].upper())
            win32api.keybd_event(0x12, 0, 0, 0)  # VK_MENU
            win32api.keybd_event(vk, 0, 0, 0)
            win32api.keybd_event(vk, 0, 2, 0)
            win32api.keybd_event(0x12, 0, 2, 0)
            i += 2
            time.sleep(0.2)
            continue
        vk = ord(ch.upper())
        win32api.keybd_event(vk, 0, 0, 0)
        win32api.keybd_event(vk, 0, 2, 0)
        i += 1
        time.sleep(0.15)


def _vbe_commandbar_reset() -> dict[str, Any]:
    """Run→Reset via VBIDE CommandBars. Prefer project named Test (WZTC bridge)."""
    out: dict[str, Any] = {"ok": False}
    try:
        import win32com.client  # type: ignore

        app = win32com.client.GetActiveObject("MicroStationDGN.Application")
        vbe = app.VBE
        preferred = None
        for i in range(1, int(vbe.VBProjects.Count) + 1):
            p = vbe.VBProjects.Item(i)
            name = str(getattr(p, "Name", "") or "")
            if name == "Test" or name.upper().startswith("WZTC"):
                preferred = p
                break
        if preferred is not None:
            try:
                vbe.ActiveVBProject = preferred
                out["project"] = str(preferred.Name)
            except Exception as ex:
                out["project_err"] = f"{type(ex).__name__}: {ex}"
        bar = vbe.CommandBars("Menu Bar")
        for c in bar.Controls:
            if "Run" not in str(getattr(c, "Caption", "")):
                continue
            for sc in c.Controls:
                cap = str(getattr(sc, "Caption", "")).replace("&", "")
                if cap.lower().startswith("reset"):
                    sc.Execute()
                    out["ok"] = True
                    out["via"] = "CommandBars"
                    return out
        out["error"] = "Reset control not found"
    except Exception as ex:
        out["error"] = f"{type(ex).__name__}: {ex}"
    return out


def reset_vba(settle_s: float = 0.8) -> dict[str, Any]:
    """Dismiss VBA OK dialogs and always Run→Reset (agent-owned; never ask engineer).

    Clears [break]/[running] and sticky status-bar macro errors (e.g. 0x800adf09).
    Prefer COM CommandBars Reset on Test.mvba; fall back to Alt+R, R on the IDE.
    """
    dismissed = dismiss_vba_ok()
    time.sleep(0.25)
    before = vba_ide_state()
    steps: list[str] = ["dismiss"]
    com = _vbe_commandbar_reset()
    steps.append({"commandbar_reset": com})
    if before["hwnd"] and (before["state"] in ("break", "running") or not com.get("ok")):
        hwnd = int(before["hwnd"])
        if _activate_hwnd(hwnd):
            fg = win32gui.GetForegroundWindow()
            fg_title = win32gui.GetWindowText(fg) if fg else ""
            if "Visual Basic" in fg_title:
                steps.append("activate")
                _send_keys_unicode("%r")  # Run menu
                time.sleep(0.25)
                _send_keys_unicode("r")  # Reset
                steps.append("run_reset_keys")
                time.sleep(settle_s)
            else:
                steps.append({"skip_keys": f"foreground not VBA: {fg_title[:80]!r}"})
        else:
            steps.append({"skip_keys": "could not foreground VBA IDE"})
    elif com.get("ok"):
        time.sleep(settle_s)
    after = vba_ide_state()
    ok = after["state"] not in ("break", "running")
    if after["state"] == "missing" and before["state"] == "missing":
        ok = True
    return {
        "ok": ok,
        "before": before,
        "after": after,
        "steps": steps,
        "dismissed": dismissed,
        "commandbar": com,
    }


def ensure_vba_idle(max_attempts: int = 2) -> dict[str, Any]:
    """Ensure VBA IDE is idle (not [break]/[running]) for hot-reload / bridge ops."""
    attempts: list[dict[str, Any]] = []
    for _ in range(max(1, max_attempts)):
        st = vba_ide_state()
        if st["state"] == "idle" or st["state"] == "missing":
            # still clear any bare error boxes
            d = dismiss_vba_ok()
            return {"ok": True, "state": st, "attempts": attempts, "dismissed": d}
        attempts.append(reset_vba())
        time.sleep(0.3)
    final = vba_ide_state()
    return {
        "ok": final["state"] == "idle",
        "state": final,
        "attempts": attempts,
    }


def status() -> dict[str, Any]:
    hwnd = find_ord_hwnd()
    main = find_ord_main()
    return {
        "ord_hwnd": hwnd,
        "ord_title": (win32gui.GetWindowText(hwnd)[:160] if hwnd else None),
        "uia_main": bool(main),
        "keyin_hwnd": find_keyin_hwnd(),
        "vba_dialogs": [{"hwnd": h, "title": t} for h, t in list_vba_dialogs()],
        "owned_window_names": [
            (getattr(w, "Name", None) or "")[:80] for w in (iter_owned_windows() or [])
        ][:20],
    }


# --- CLI ----------------------------------------------------------------------


def main(argv: Optional[list[str]] = None) -> int:
    p = argparse.ArgumentParser(description="ORD UI Automation helpers")
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("probe", help="Write surface map to outputs/uia/")
    sub.add_parser("status", help="Print ORD / Key-in / VBA dialog status")
    sub.add_parser("focus", help="Bring ORD to foreground")
    sub.add_parser("open-keyin", help="Open Key-in Browser (no CAD command typed)")
    sub.add_parser("dismiss-vba", help="BM_CLICK OK on VBA error dialogs")
    sub.add_parser("reset-vba", help="Dismiss OK + Run→Reset if [break]/[running]")
    sub.add_parser("ensure-vba-idle", help="Reset until VBA IDE is idle (for hot-reload)")

    args = p.parse_args(argv)

    if args.cmd == "probe":
        out = write_probe()
        print(f"wrote {out}")
        print(f"wrote {out.with_suffix('.md')}")
        data = json.loads(out.read_text(encoding="utf-8"))
        print("summary:", data.get("summary"))
        return 0
    if args.cmd == "status":
        print(json.dumps(status(), indent=2))
        return 0
    if args.cmd == "focus":
        print(json.dumps(activate_ord(), indent=2))
        return 0
    if args.cmd == "open-keyin":
        print(json.dumps(open_keyin(), indent=2))
        return 0
    if args.cmd == "dismiss-vba":
        print(json.dumps(dismiss_vba_ok(), indent=2))
        return 0
    if args.cmd == "reset-vba":
        print(json.dumps(reset_vba(), indent=2))
        return 0
    if args.cmd == "ensure-vba-idle":
        print(json.dumps(ensure_vba_idle(), indent=2))
        return 0
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
