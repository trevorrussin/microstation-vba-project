"""Edit live conduit callout #11 (id=7912) to match markup CABLE(11) text.

Authority: r1 FreeText [64] body, with CABLE (6)->(11) per engineer;
item mid-digits use 09.18 corrections (7310714 / 7310514 / 75070011).
Text-only: SetTextLine / AddTextLine; SAVE DESIGN.
"""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

TARGET_ID = 7912  # live #11 top->bottom
AFTER = [
    'TRAFFIC SIGNAL CONDUIT (1)',
    '(RIGID PLASTIC, CLASS 2, 2") (ITEM NO. 680.520606)',
    "TRAFFIC SIGNAL CABLE (11)",
    "(7 CONDUCTORS, 14 AWG) (ITEM NO. 680.7310714) (2)",
    "(5 CONDUCTORS, 14 AWG) (ITEM NO. 680.7310514) (1)",
    "(SINGLE CONDUCTOR CABLE, NUMBER 10 GAGE) (ITEM NO. 670.7006) (2)",
    "(GROUND WIRE 10 AWG) (ITEM NO. 670.75070011) (1)",
]


def walk(el, out):
    try:
        if el.IsTextNodeElement and el.ID.Low == TARGET_ID:
            out.append(el)
        if int(el.Type) in (2, 7):
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk(ee.Current, out)
    except Exception:
        pass


def set_lines(tn_el, lines):
    tn = tn_el.AsTextNodeElement
    n = int(tn.TextLinesCount)
    for i, line in enumerate(lines, start=1):
        if i <= n:
            tn.SetTextLine(i, line)
        else:
            tn.AddTextLine(line)
    for i in range(len(lines) + 1, n + 1):
        tn.SetTextLine(i, " ")
    tn_el.Rewrite()


def main():
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = app.ActiveModelReference
    print("DGN", app.ActiveDesignFile.FullName)

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    found = []
    while ee.MoveNext():
        walk(ee.Current, found)
    if not found:
        raise SystemExit(f"id={TARGET_ID} not found")
    el = found[0]
    tn = el.AsTextNodeElement
    before = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
    print("BEFORE:")
    for ln in before:
        print(f"  {ln!r}")
    print("AFTER:")
    for ln in AFTER:
        print(f"  {ln!r}")

    set_lines(el, AFTER)
    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.4)

    # re-read
    o2 = Dispatch("MicroStationDGN.ElementScanCriteria")
    o2.ExcludeNonGraphical()
    ee2 = mr.Scan(o2)
    found2 = []
    while ee2.MoveNext():
        walk(ee2.Current, found2)
    tn2 = found2[0].AsTextNodeElement
    now = [tn2.TextLine(i) for i in range(1, tn2.TextLinesCount + 1)]
    print("VERIFY:")
    for ln in now:
        if ln.strip():
            print(f"  {ln!r}")
    ok = [ln.rstrip() for ln in now if ln.strip()] == AFTER
    print(f"match={ok} SAVED")


if __name__ == "__main__":
    main()
