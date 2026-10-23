"""Compute Loc01 Straight LF (TLEN logic) and show alert in AutoCAD."""
from __future__ import annotations

import math
import pythoncom
from win32com.client import GetObject
from win32com.client.dynamic import Dispatch

LAYER = "Straight Steel Faced Concrete Curb (18'' Deep)"
X0, X1 = 1028440.0, 1028620.0
Y0, Y1 = 187390.0, 187530.0


def bring_forward(app) -> None:
    try:
        import win32gui
        import win32con

        hwnd = int(app.HWND)
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass


def line_len(e) -> float:
    """Late-bind Length / StartPoint-EndPoint (early-bound IAcadEntity hides Length)."""
    try:
        d = Dispatch(e)
        try:
            return float(d.Length)
        except Exception:
            pass
        try:
            s, t = d.StartPoint, d.EndPoint
            return math.hypot(float(t[0]) - float(s[0]), float(t[1]) - float(s[1]))
        except Exception:
            return 0.0
    except Exception:
        return 0.0


def main() -> None:
    pythoncom.CoInitialize()
    app = Dispatch(GetObject(None, "AutoCAD.Application"))
    doc = app.ActiveDocument
    print("file", doc.FullName)
    app.Visible = True
    bring_forward(app)

    tl = 0.0
    n = 0
    ms = doc.ModelSpace
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.Layer != LAYER:
                continue
            if e.ObjectName not in ("AcDbLine", "AcDbArc"):
                # Skip leftover FL on straight layer; TLEN for LINE/ARC only here
                if e.ObjectName != "AcDbLine":
                    continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2.0
            cy = (float(mn[1]) + float(mx[1])) / 2.0
            if not (X0 < cx < X1 and Y0 < cy < Y1):
                continue
            L = line_len(e)
            if L <= 0.0:
                continue
            tl += L
            n += 1
        except Exception:
            continue

    msg = f"Loc01 Straight — Total length of selected objects is {tl:.4f}  (count {n})"
    print(msg)
    # Show on AutoCAD screen (same as TLEN.LSP alert)
    safe = msg.replace("\\", "\\\\").replace('"', "'")
    doc.SendCommand(f'(alert "{safe}")\n')
    try:
        doc.Utility.Prompt("\n" + msg + "\n")
    except Exception:
        pass


if __name__ == "__main__":
    main()
