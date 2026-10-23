"""Try EXPLODE feature line via AutoCAD SendCommand."""
import time
import pythoncom
from win32com.client import Dispatch


def main():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    before = doc.ModelSpace.Count
    # Use AutoLISP via SendCommand carefully
    lisp = '(progn (setq e (handent "13D822")) (if e (command "_.EXPLODE" e) (princ "NOENT")) (princ)) '
    doc.SendCommand(lisp)
    time.sleep(1.5)
    after = doc.ModelSpace.Count
    print(f"count {before} -> {after}")
    try:
        e = doc.HandleToObject("13D822")
        print("13D822 still", e.ObjectName, e.Layer)
    except Exception as ex:
        print("13D822 gone", ex)


if __name__ == "__main__":
    main()
