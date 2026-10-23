import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
mr = app.ActiveModelReference
print("file", app.ActiveDesignFile.FullName)

o = Dispatch("MicroStationDGN.ElementScanCriteria")
o.ExcludeNonGraphical()
ee = mr.Scan(o)
hits = []
while ee.MoveNext():
    el = ee.Current
    try:
        if not el.IsTextNodeElement:
            continue
        tn = el.AsTextNodeElement
        lines = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
        t = "\n".join(lines)
        u = t.upper()
        rid = el.ID.Low
        if rid in (6280, 6117) or "APS W/POST" in u or (
            "CONTROL CABINET" in u and "APS PUSH" in u
        ):
            rng = el.Range
            hits.append(
                (
                    rid,
                    float(rng.Low.Y),
                    float(rng.High.Y),
                    float(rng.Low.X),
                    lines,
                    el,
                )
            )
    except Exception:
        continue

for rid, ly, hy, lx, lines, el in sorted(hits, key=lambda x: (-x[2], x[0])):
    print(f"id={rid} y=[{ly:.1f},{hy:.1f}] x0={lx:.1f}")
    for ln in lines:
        print(" ", repr(ln))
    print()

# Desired end state per engineer:
# 6280 should read CONTROL CABINET text
# 6117 should read APS W/POST text
# If already that way, report; else swap again.

by_id = {h[0]: h for h in hits}
a = by_id.get(6280)
b = by_id.get(6117)
if not a or not b:
    # find by content
    print("ID lookup failed; using content match")
    sys.exit(1)

lines_a = a[4]
lines_b = b[4]
text_a = "\n".join(lines_a).upper()
text_b = "\n".join(lines_b).upper()
print("6280 has CONTROL?", "CONTROL CABINET" in text_a)
print("6117 has APS W/POST?", "APS W/POST" in text_b)

# Engineer wants: 6280 = CONTROL CABINET, 6117 = APS W/POST
need_6280_cabinet = "CONTROL CABINET" in text_a
need_6117_post = "APS W/POST" in text_b
if need_6280_cabinet and need_6117_post:
    print("ALREADY_CORRECT")
else:
    print("NEED_SWAP_AGAIN — applying now")
    ela, elb = a[5], b[5]
    ta, tb = ela.AsTextNodeElement, elb.AsTextNodeElement

    def set_lines(tn, lines):
        n = tn.TextLinesCount
        for i, line in enumerate(lines, start=1):
            if i <= n:
                tn.SetTextLine(i, line)
            else:
                tn.AddTextLine(line)
        for i in range(len(lines) + 1, n + 1):
            tn.SetTextLine(i, " ")
        tn.Rewrite()

    set_lines(ta, lines_b)
    set_lines(tb, lines_a)
    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.4)
    print("AFTER 6280", repr("\n".join(ta.TextLine(i) for i in range(1, ta.TextLinesCount + 1))))
    print("AFTER 6117", repr("\n".join(tb.TextLine(i) for i in range(1, tb.TextLinesCount + 1))))
