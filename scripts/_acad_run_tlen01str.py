"""Load TLEN01STR.lsp and run it so the AutoCAD alert appears on screen."""
from __future__ import annotations

import time
import pythoncom
from win32com.client import GetObject

try:
    import win32gui
    import win32con
except ImportError:
    win32gui = None

LSP = r"c:/repos/microstation-vba-project/scripts/TLEN01STR.lsp"


def bring_acad_forward(app) -> None:
    if win32gui is None:
        return
    try:
        hwnd = int(app.HWND)
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception as ex:
        print("foreground_err", ex)


def main() -> None:
    pythoncom.CoInitialize()
    app = GetObject(None, "AutoCAD.Application")
    doc = app.ActiveDocument
    print("file", doc.FullName)
    app.Visible = True
    bring_acad_forward(app)
    # Single SendCommand: load then invoke (alert shows in AutoCAD)
    doc.SendCommand(f'(progn (load "{LSP}") (C:TLEN01STR))\n')
    print("Sent TLEN01STR — check AutoCAD for the alert.")


if __name__ == "__main__":
    main()
