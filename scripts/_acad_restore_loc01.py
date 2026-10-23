"""Restore loc01 sources after partial v3 run."""
import time
import pythoncom
from win32com.client import Dispatch

LAYERS = {
    "Straight Steel Faced Concrete Curb (18'' Deep)",
    "Depressed Steel Faced Concrete Curb (18'' Deep)",
    "Corner Steel Faced Concrete Curb (18'' Deep)",
}

pythoncom.CoInitialize()
acad = Dispatch("AutoCAD.Application")
doc = acad.ActiveDocument
try:
    doc.SendCommand(chr(27) + chr(27))
except Exception:
    pass
time.sleep(0.3)

for i in range(25):
    doc.SendCommand("_U ")
    time.sleep(0.35)
    ok = False
    try:
        doc.HandleToObject("13D986")
        ok = True
    except Exception:
        pass
    print(f"undo {i+1}: 986={ok} count={doc.ModelSpace.Count}")
    if ok:
        break

ms = doc.ModelSpace
to_del = []
for i in range(ms.Count):
    try:
        e = ms.Item(i)
        if e.Layer in LAYERS:
            to_del.append(e.Handle)
    except Exception:
        pass
n = 0
for h in to_del:
    try:
        doc.HandleToObject(h).Delete()
        n += 1
    except Exception:
        pass
print("cleared steel", n)

for h in ["13D986", "13D822", "13D824", "13D25C", "13D25E", "13D7DC", "13D7E8", "13CEC7", "13CED2", "13D80A"]:
    try:
        e = doc.HandleToObject(h)
        mn, mx = e.GetBoundingBox()
        print(
            h,
            e.ObjectName,
            f"c=({(float(mn[0])+float(mx[0]))/2:.1f},{(float(mn[1])+float(mx[1]))/2:.1f})",
            f"wh=({float(mx[0])-float(mn[0]):.1f},{float(mx[1])-float(mn[1]):.1f})",
        )
    except Exception:
        print(h, "MISSING")
