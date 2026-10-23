"""Test vlax-curve sampling of feature line to temp file."""
import os
import time
import pythoncom
from win32com.client import Dispatch

OUT = "C:/Users/RussinT/AppData/Local/Temp/fl_sample.txt"


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    if os.path.exists(OUT):
        os.remove(OUT)
    lisp = (
        f'(progn (vl-load-com) (setq f (open "{OUT}" "w")) '
        f'(setq e (handent "13D822")) '
        f'(if e (progn (write-line "OK" f) '
        f'(write-line (vl-princ-to-string (vlax-curve-getEndParam e)) f) '
        f'(write-line (rtos (vlax-curve-getDistAtParam e (vlax-curve-getEndParam e)) 2 4) f) '
        f'(setq p (vlax-curve-getPointAtDist e 0.0)) '
        f'(write-line (vl-princ-to-string p) f)) '
        f'(write-line "NOENT" f)) (close f) (princ)) '
    )
    doc.SendCommand(lisp)
    time.sleep(2.0)
    print("exists", os.path.exists(OUT))
    if os.path.exists(OUT):
        print(open(OUT, encoding="utf-8", errors="replace").read())


if __name__ == "__main__":
    main()
