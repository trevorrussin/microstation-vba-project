"""Apply sure 09.18 item-number fixes on all conduit/cable callouts.

Markup (09.18) strikes wrong mid-digits and greens replacements; estimate
table confirms:
  5-cond  -> 680.730514
  7-cond  -> 680.730714
  13-cond -> 680.731314
  ground  -> 670.75070011  (was 670.70011)
"""
from __future__ import annotations

import re
import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

REPLACERS = [
    # order matters — more specific first
    (
        re.compile(r"\(5 CONDUCTORS, 14 AWG\) \(ITEM NO\. 680\.\d+\)"),
        "(5 CONDUCTORS, 14 AWG) (ITEM NO. 680.730514)",
    ),
    (
        re.compile(r"\(7 CONDUCTORS, 14 AWG\) \(ITEM NO\. 680\.\d+\)"),
        "(7 CONDUCTORS, 14 AWG) (ITEM NO. 680.730714)",
    ),
    (
        re.compile(r"\(13 CONDUCTORS, 14 AWG\) \(ITEM NO\. 680\.\d+\)"),
        "(13 CONDUCTORS, 14 AWG) (ITEM NO. 680.731314)",
    ),
    (
        re.compile(r"\(GROUND WIRE 10 AWG\) \(ITEM NO\. 670\.\d+\)"),
        "(GROUND WIRE 10 AWG) (ITEM NO. 670.75070011)",
    ),
]


def fix_line(line: str) -> str:
    out = line
    for rx, repl in REPLACERS:
        # keep trailing qty / SEE NOTE after item
        m = rx.search(out)
        if not m:
            continue
        # If pattern is only the item part, replace match; preserve suffix after match
        out = out[: m.start()] + repl + out[m.end() :]
    return out


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


def walk(el, out):
    try:
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            lines = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
            t = "\n".join(lines).upper()
            if "TRAFFIC SIGNAL CONDUIT" in t or (
                "TRAFFIC SIGNAL CABLE" in t and "680." in t
            ):
                out.append(el)
        if int(el.Type) in (2, 7):
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk(ee.Current, out)
    except Exception:
        pass


def main():
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = app.ActiveModelReference
    print("DGN", app.ActiveDesignFile.FullName)
    print("PDF authority: Bally's_Hutch_09.18.26_Traffic_BTS.pdf + estimate table")

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    nodes = []
    while ee.MoveNext():
        walk(ee.Current, nodes)
    # dedupe
    by_id = {el.ID.Low: el for el in nodes}
    nodes = list(by_id.values())
    print(f"callouts={len(nodes)}")

    total_line_edits = 0
    callouts_touched = 0
    for el in nodes:
        tn = el.AsTextNodeElement
        before = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
        after = [fix_line(ln) for ln in before]
        if before == after:
            continue
        callouts_touched += 1
        diffs = [(a, b) for a, b in zip(before, after) if a != b]
        total_line_edits += len(diffs)
        print(f"\nid={el.ID.Low} edits={len(diffs)}")
        for a, b in diffs:
            print(f"  - {a!r}")
            print(f"  + {b!r}")
        set_lines(el, after)

    if callouts_touched:
        app.CadInputQueue.SendKeyin("SAVE DESIGN")
        time.sleep(0.4)
    print(
        f"\nDONE callouts_touched={callouts_touched} "
        f"line_edits={total_line_edits}"
    )
    print(
        "NOTE: qty chips ((2)/(3)/(11)/(16) etc.) from 09.18 not applied yet — "
        "confirm those next."
    )


if __name__ == "__main__":
    main()
