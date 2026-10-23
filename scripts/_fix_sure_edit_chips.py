"""Apply sure red/green small EDIT chip fixes on non-conduit callouts."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

# Sure finals from markup green chips + prior finish script intent
UPDATES = {
    # VEHICULAR: green [0] adds 680.8223 (1) — currently missing
    8698: [
        "VEHICULAR SIGNAL HEADS AND APS PUSH BUTTON",
        "(ITEM NOS. 680.818122011 (1), 680.81822011 (1), 680.8223 (1), 680.8151 (1))",
    ],
    # STREETLIGHT: green [44] adds 680.83080011 (1) — currently missing
    7781: [
        "STREETLIGHT W/PEDESTRIAN TRAFFIC SIGNALS AND APS PUSH BUTTON",
        "(ITEM NOS. 680.8141 (2), 680.815002 (2), 680.8151 (1), 680.83080011 (1))",
    ],
}


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


def find_tn(mr, tid):
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)

    def walk(el):
        try:
            if el.IsTextNodeElement and el.ID.Low == tid:
                return el
            if int(el.Type) in (2, 7):
                sub = el.AsComplexElement.GetSubElements()
                while sub.MoveNext():
                    hit = walk(sub.Current)
                    if hit is not None:
                        return hit
        except Exception:
            return None
        return None

    while ee.MoveNext():
        hit = walk(ee.Current)
        if hit is not None:
            return hit
    return None


def main():
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = app.ActiveModelReference
    print("DGN", app.ActiveDesignFile.FullName)

    for tid, lines in UPDATES.items():
        el = find_tn(mr, tid)
        if el is None:
            print(f"MISSING {tid}")
            continue
        before = [
            el.AsTextNodeElement.TextLine(i)
            for i in range(1, el.AsTextNodeElement.TextLinesCount + 1)
        ]
        set_lines(el, lines)
        after = [
            el.AsTextNodeElement.TextLine(i)
            for i in range(1, el.AsTextNodeElement.TextLinesCount + 1)
        ]
        print(f"\nUPDATED {tid}")
        print(" before:", before)
        print(" after: ", after)

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.4)
    print("\nSAVED")


if __name__ == "__main__":
    main()
