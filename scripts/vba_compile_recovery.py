"""Dismiss VBA error dialog, Reset, Compile Test project."""
from __future__ import annotations

import sys
import time
from pathlib import Path

import pythoncom
import win32con
import win32gui

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "mcp-server"))
import ms_connect  # noqa: E402


def _enum_top():
    out = []

    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                out.append((hwnd, title))
        return True

    win32gui.EnumWindows(cb, None)
    return out


def _click_ok(dialog_hwnd: int) -> bool:
    found = []

    def cb(hwnd, _):
        cls = win32gui.GetClassName(hwnd)
        txt = win32gui.GetWindowText(hwnd)
        if cls == "Button" and txt.strip() == "OK":
            found.append(hwnd)
        return True

    win32gui.EnumChildWindows(dialog_hwnd, cb, None)
    if not found:
        return False
    win32gui.SendMessage(found[0], win32con.BM_CLICK, 0, 0)
    return True


def _find_control(bar, control_id: int):
    for i in range(1, bar.Controls.Count + 1):
        c = bar.Controls.Item(i)
        try:
            if int(c.Id) == control_id:
                return c
        except Exception:
            pass
        try:
            if c.Type == 10:  # popup
                sub = _find_control(c.CommandBar, control_id)
                if sub is not None:
                    return sub
        except Exception:
            pass
    return None


def main() -> None:
    tops = _enum_top()
    print("VBA windows:")
    for hwnd, title in tops:
        if "Visual Basic" in title or "VBA" in title:
            print(f"  {hwnd}: {title!r}")

    dialogs = [h for h, t in tops if t == "Microsoft Visual Basic for Applications"]
    if dialogs:
        print("dismissing error dialog", dialogs[0])
        _click_ok(dialogs[0])
        time.sleep(0.5)
    else:
        print("no bare VBA error dialog")

    pythoncom.CoInitialize()
    app = ms_connect.get_microstation_app("Test")
    vbe = app.VBE

    try:
        cp = vbe.ActiveCodePane
        if cp is not None:
            a, b, c, d = cp.GetSelection()
            print("selection", a, b, c, d)
            line = cp.CodeModule.Lines(a, 1)
            print("line:", line)
            print("module:", cp.CodeModule.Parent.Name)
    except Exception as e:
        print("ActiveCodePane:", e)

    # Reset Id=228
    try:
        std = vbe.CommandBars("Standard")
        reset = _find_control(std, 228)
        if reset is not None:
            print("Reset:", reset.Caption)
            reset.Execute()
            time.sleep(0.5)
        else:
            print("Reset control not found")
    except Exception as e:
        print("Reset failed:", e)

    # Compile project Id=578
    try:
        menu = vbe.CommandBars("Menu Bar")
        compile_ctl = _find_control(menu, 578)
        if compile_ctl is not None:
            print("Compile:", compile_ctl.Caption)
            compile_ctl.Execute()
            time.sleep(1.0)
        else:
            print("Compile control not found; listing Debug menu...")
            for i in range(1, menu.Controls.Count + 1):
                c = menu.Controls.Item(i)
                if "Debug" in str(getattr(c, "Caption", "")):
                    print(" ", c.Caption, c.Id)
                    try:
                        for j in range(1, c.CommandBar.Controls.Count + 1):
                            s = c.CommandBar.Controls.Item(j)
                            print("   ", s.Caption, s.Id)
                    except Exception as e:
                        print("   suberr", e)
    except Exception as e:
        print("Compile failed:", e)

    # Re-check dialog after compile
    time.sleep(0.5)
    tops = _enum_top()
    dialogs = [h for h, t in tops if t == "Microsoft Visual Basic for Applications"]
    if dialogs:
        print("compile still errors — reading selection")
        try:
            cp = vbe.ActiveCodePane
            a, b, c, d = cp.GetSelection()
            print("ERR LINE", a, cp.CodeModule.Lines(a, 3))
            print("module", cp.CodeModule.Parent.Name)
        except Exception as e:
            print("read err", e)
        _click_ok(dialogs[0])
    else:
        print("no error dialog after compile — likely clean")

    # Ping
    from pathlib import Path
    ping = Path(r"c:\repos\microstation-vba-project\Bridge\_ping.txt")
    if ping.exists():
        ping.unlink()
    app.CadInputQueue.SendKeyin("VBA RUN [Test]WZTCBridge.CursorBridgePing")
    time.sleep(1.5)
    print("ping", ping.exists(), ping.read_text() if ping.exists() else "")
    pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
