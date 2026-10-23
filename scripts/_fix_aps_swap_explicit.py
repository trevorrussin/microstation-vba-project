"""Fix: 6280 = CONTROL CABINET, 6117 = APS W/POST (explicit, not swap-from-self)."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

CABINET = [
    "APS PUSH BUTTON ON ",
    "EXIST. CONTROL CABINET",
    "(ITEM NO. 680.8151)",
]
POST = [
    "APS W/POST",
    "(ITEM NOS. 680.5001,",
    " 680.8152 (1))",
]


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


def main():
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = app.ActiveModelReference
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    by_id = {}
    while ee.MoveNext():
        el = ee.Current
        try:
            if el.IsTextNodeElement and el.ID.Low in (6280, 6117):
                by_id[el.ID.Low] = el
        except Exception:
            continue
    if 6280 not in by_id or 6117 not in by_id:
        raise SystemExit(f"missing ids: {sorted(by_id)}")

    e6280, e6117 = by_id[6280], by_id[6117]
    set_lines(e6280.AsTextNodeElement, CABINET)
    set_lines(e6117.AsTextNodeElement, POST)
    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.5)

    # verify from fresh scan
    o2 = Dispatch("MicroStationDGN.ElementScanCriteria")
    o2.ExcludeNonGraphical()
    ee2 = mr.Scan(o2)
    while ee2.MoveNext():
        el = ee2.Current
        try:
            if not el.IsTextNodeElement or el.ID.Low not in (6280, 6117):
                continue
            tn = el.AsTextNodeElement
            lines = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
            print(f"id={el.ID.Low}")
            for ln in lines:
                print(" ", repr(ln))
        except Exception:
            continue
    print("SAVED")


if __name__ == "__main__":
    main()
