"""Extract feature line vertices via AutoLISP entget / dump to file."""
from __future__ import annotations

import os
import time
import pythoncom
from win32com.client import Dispatch

OUT = r"c:\repos\microstation-vba-project\scripts\_fl_pts.txt"


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    if os.path.exists(OUT):
        os.remove(OUT)

    # Dump DXF group codes for feature line via lisp
    lisp = f"""
(progn
  (setq f (open "{OUT.replace(chr(92), chr(92)+chr(92))}" "w"))
  (foreach h (list "13D822" "13D824")
    (setq e (handent h))
    (write-line (strcat "HANDLE " h) f)
    (if e
      (progn
        (setq ed (entget e))
        (foreach p ed
          (write-line (vl-princ-to-string p) f)
        )
        (write-line "----" f)
      )
      (write-line "MISSING" f)
    )
  )
  (close f)
  (princ)
)
"""
    # simplify path escaping for lisp
    out_lisp = OUT.replace("\\", "/")
    lisp = f"""
(progn
  (setq f (open "{out_lisp}" "w"))
  (foreach h (list "13D822" "13D824")
    (setq e (handent h))
    (write-line (strcat "HANDLE " h) f)
    (if e
      (progn
        (setq ed (entget e '("*")))
        (foreach p ed
          (write-line (vl-princ-to-string p) f)
        )
        (write-line "====" f)
      )
      (write-line "MISSING" f)
    )
  )
  (close f)
  (princ "DONE")
)
"""
    doc.SendCommand(lisp.replace("\n", " ") + " ")
    time.sleep(2.0)
    if os.path.exists(OUT):
        print(open(OUT, encoding="utf-8", errors="replace").read()[:5000])
    else:
        print("no output file")


if __name__ == "__main__":
    main()
