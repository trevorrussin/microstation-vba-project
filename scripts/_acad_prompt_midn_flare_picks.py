"""Load MidN flare ask LISP into AutoCAD, wait for picks, place solid.

User sees AutoCAD (alert) dialogs + yellow MTEXT + getpoint crosshairs.
"""
from __future__ import annotations

import math
import os
import time
import pythoncom
from win32com.client import Dispatch, VARIANT

LAYER = "PROP-D-ROAD-RAMPS"
MIDN = (1028501.0, 187485.0)
OUT = r"c:\repos\microstation-vba-project\scripts\_midn_flare_picks.txt"
LSP = r"c:/repos/microstation-vba-project/scripts/_midn_flare_ask.lsp"


def vpt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def retry_com(fn, tries=12, delay=0.5):
    last = None
    for _ in range(tries):
        try:
            return fn()
        except Exception as ex:
            last = ex
            time.sleep(delay)
    raise last


def delete_midn(doc):
    ms = doc.ModelSpace
    handles = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbSolid" or e.Layer != LAYER:
                continue
            mn, mx = e.GetBoundingBox()
            cx = (float(mn[0]) + float(mx[0])) / 2
            cy = (float(mn[1]) + float(mx[1])) / 2
            if cx < 1028488:
                continue
            if math.hypot(cx - MIDN[0], cy - MIDN[1]) < 30:
                handles.append(e.Handle)
        except Exception:
            continue
    for h in handles:
        try:
            doc.HandleToObject(h).Delete()
        except Exception:
            pass
    return len(handles)


def wait_picks(path, timeout=300.0):
    t0 = time.time()
    while time.time() - t0 < timeout:
        if os.path.exists(path) and os.path.getsize(path) > 20:
            time.sleep(0.35)
            with open(path, encoding="utf-8") as f:
                txt = f.read().strip()
            if "tip," in txt and "tan," in txt:
                return txt
        time.sleep(0.4)
    raise TimeoutError("no picks — did you cancel the AutoCAD prompts?")


def main():
    print("Connecting to AutoCAD…")
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    acad.Visible = True
    try:
        import win32gui
        import win32con

        win32gui.ShowWindow(int(acad.HWND), win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(int(acad.HWND))
    except Exception:
        pass

    doc = acad.ActiveDocument
    ms = doc.ModelSpace
    print("Active:", doc.Name)

    if os.path.exists(OUT):
        try:
            os.remove(OUT)
        except Exception:
            pass

    print("Cleaning old MidN solids…")
    n = retry_com(lambda: delete_midn(doc))
    print("deleted", n)

    # Load + run LISP (alert + getpoint happen inside AutoCAD)
    print(">>> Switch to AutoCAD — click OK on the alert, then pick TIP / LEFT / TAN")
    cmd = f'(load "{LSP}") '
    retry_com(lambda: doc.SendCommand(cmd) or True)

    txt = wait_picks(OUT)
    print(txt)
    pts = {}
    for line in txt.splitlines():
        a = line.split(",")
        if len(a) >= 3 and a[0] in ("tip", "left", "tan"):
            pts[a[0]] = (float(a[1]), float(a[2]))

    tip, left, tan = pts["tip"], pts["left"], pts["tan"]
    sol = retry_com(lambda: ms.AddSolid(vpt(*tip), vpt(*left), vpt(*tan), vpt(*tan)))
    sol.Layer = LAYER
    sol.Color = 2
    try:
        doc.Regen(1)
    except Exception:
        pass

    done = (
        f"Placed yellow flare\\n"
        f"tip ({tip[0]:.2f},{tip[1]:.2f})\\n"
        f"left ({left[0]:.2f},{left[1]:.2f})\\n"
        f"tan ({tan[0]:.2f},{tan[1]:.2f})\\n"
        f"handle {sol.Handle}"
    )
    try:
        doc.SendCommand(f'(alert "{done}") ')
    except Exception:
        pass
    print("HANDLE", sol.Handle)
    print("DONE")


if __name__ == "__main__":
    main()
