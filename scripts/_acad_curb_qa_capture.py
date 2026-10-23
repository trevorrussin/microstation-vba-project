"""Visual QA capture for curb classify (AutoCAD, active or named open DWG).

Turns OFF survey-xref layers that clutter the view (keeps survey curb lines),
zooms to a window around --center, captures the AutoCAD window to PNG.
Layer On/Off only — no geometry change. Engineer approved turning off
survey layers that get in the way (2026-10-01).

Usage:
  python scripts/_acad_curb_qa_capture.py --name Location_20_SGD --center 1029162 192952 --half 45 --out outputs/qa/loc20.png
"""
from __future__ import annotations

import argparse
import os
import time

import pythoncom
import win32gui
from PIL import ImageGrab
from pywintypes import com_error
from win32com.client import GetObject, VARIANT

BUSY = -2147418111
KEEP_SURVEY = ("CURB-LINE",)
# Host-resident survey/utility layers that clutter curb QA (layer On/Off only)
HOST_OFF_PREFIX = ("S-", "U-", "ST-", "D-TOPO", "E-", "Z-", "C-TINN", "C-TOPO-EG", "D-DIM")
HOST_KEEP = ("S-CURB-LINE",)


def R(fn, tries=100):
    for _ in range(tries):
        try:
            return fn()
        except com_error as e:
            if e.args and e.args[0] == BUSY:
                time.sleep(0.4)
                continue
            raise
    return fn()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--name", required=True, help="substring of open DWG name")
    ap.add_argument("--center", nargs=2, type=float, required=True)
    ap.add_argument("--half", type=float, default=45.0)
    ap.add_argument("--out", required=True)
    ap.add_argument("--keep-survey", action="store_true", help="do not turn layers off")
    a = ap.parse_args()

    pythoncom.CoInitialize()
    acad = GetObject(None, "AutoCAD.Application")
    doc = None
    for i in range(R(lambda: acad.Documents.Count)):
        d = R(lambda i=i: acad.Documents.Item(i))
        if a.name.lower() in d.Name.lower():
            doc = d
    if doc is None:
        print("doc not open:", a.name)
        return 2
    R(lambda: doc.Activate())
    off = 0
    if not a.keep_survey:
        surveys = []
        for j in range(R(lambda: doc.Blocks.Count)):
            b = R(lambda j=j: doc.Blocks.Item(j))
            if R(lambda: b.IsXRef):
                nm = R(lambda: b.Name)
                if "T4318" in nm.upper() or nm.upper().startswith("SURVEY"):
                    surveys.append(nm.upper() + "|")
        for k in range(R(lambda: doc.Layers.Count)):
            L = R(lambda k=k: doc.Layers.Item(k))
            n = R(lambda: L.Name).upper()
            survey_hit = any(n.startswith(s) for s in surveys) and not any(x in n for x in KEEP_SURVEY)
            host_hit = "|" not in n and n.startswith(HOST_OFF_PREFIX) and n not in HOST_KEEP
            if survey_hit or host_hit:
                if R(lambda: L.LayerOn):
                    R(lambda: setattr(L, "LayerOn", False))
                    off += 1
    cx, cy = a.center
    h = a.half
    p1 = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (cx - h, cy - h, 0.0))
    p2 = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (cx + h, cy + h, 0.0))
    R(lambda: acad.ZoomWindow(p1, p2))
    R(lambda: doc.Regen(1))
    time.sleep(2.0)
    hwnd = R(lambda: acad.HWND)
    win32gui.ShowWindow(hwnd, 3)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    time.sleep(1.5)
    rect = win32gui.GetWindowRect(hwnd)
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    ImageGrab.grab(bbox=rect, all_screens=True).save(a.out)
    print(f"layers_off={off} saved {a.out} rect={rect}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
