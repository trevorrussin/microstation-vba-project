import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
mr = app.ActiveModelReference
want = {8737, 8764}
found = {}


def walk(el, depth=0):
    try:
        if el.IsTextNodeElement and el.ID.Low in want:
            tn = el.AsTextNodeElement
            rng = el.Range
            o = tn.Origin
            found[el.ID.Low] = {
                "gg": el.GraphicGroup,
                "ox": float(o.X),
                "oy": float(o.Y),
                "lx": float(rng.Low.X),
                "hx": float(rng.High.X),
                "ly": float(rng.Low.Y),
                "hy": float(rng.High.Y),
                "lines": [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)],
                "depth": depth,
                "type": int(el.Type),
            }
        if int(el.Type) in (2, 7):
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk(ee.Current, depth + 1)
    except Exception:
        pass


o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
while ee.MoveNext():
    walk(ee.Current)

for tid in sorted(found):
    f = found[tid]
    print(f"id={tid} gg={f['gg']} depth={f['depth']}")
    print(f"  origin=({f['ox']:.2f},{f['oy']:.2f})")
    print(f"  range X=[{f['lx']:.1f},{f['hx']:.1f}] Y=[{f['ly']:.1f},{f['hy']:.1f}]")
    for ln in f["lines"]:
        print(" ", repr(ln))
    print()

if 8737 in found and 8764 in found:
    a, b = found[8737], found[8764]
    print(f"Y separation: {abs(a['oy']-b['oy']):.2f} master units")
    print(f"X separation: {abs(a['ox']-b['ox']):.2f} master units")
