"""Swap two BTS callout texts (keep positions / leaders)."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom  # noqa: E402
import ms_connect  # noqa: E402
from win32com.client import Dispatch  # noqa: E402

A_KEY = "APS W/POST"
B_KEY = "CONTROL CABINET"


def _node_text(tn) -> str:
    return "\n".join(tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1))


def _set_lines(tn, lines: list[str]) -> None:
    # SAFE: SetTextLine only; add lines if needed; never DeleteAllTextLines
    n = tn.TextLinesCount
    for i, line in enumerate(lines, start=1):
        if i <= n:
            tn.SetTextLine(i, line)
        else:
            tn.AddTextLine(line)
    # If old had more lines, blank extras (avoid DeleteAllTextLines)
    for i in range(len(lines) + 1, n + 1):
        tn.SetTextLine(i, " ")
    tn.Rewrite()


def main() -> None:
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = app.ActiveModelReference
    print("file", app.ActiveDesignFile.FullName)
    print("model", mr.Name)

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)

    a = b = None
    while ee.MoveNext():
        el = ee.Current
        try:
            if not el.IsTextNodeElement:
                continue
        except Exception:
            continue
        tn = el.AsTextNodeElement
        t = _node_text(tn).upper().replace("  ", " ")
        if A_KEY in t and "680.8152" in t:
            a = el
        elif B_KEY in t and "680.8151" in t and "APS PUSH" in t:
            b = el

    if a is None or b is None:
        raise SystemExit(f"missing callouts a={a is not None} b={b is not None}")

    ta = a.AsTextNodeElement
    tb = b.AsTextNodeElement
    lines_a = [ta.TextLine(i) for i in range(1, ta.TextLinesCount + 1)]
    lines_b = [tb.TextLine(i) for i in range(1, tb.TextLinesCount + 1)]
    print("BEFORE A id=", a.ID.Low)
    for ln in lines_a:
        print(" ", repr(ln))
    print("BEFORE B id=", b.ID.Low)
    for ln in lines_b:
        print(" ", repr(ln))

    # Swap wording only — keep origins / leaders
    _set_lines(ta, lines_b)
    _set_lines(tb, lines_a)

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.4)

    # Re-read
    ta2 = a.AsTextNodeElement
    tb2 = b.AsTextNodeElement
    print("AFTER A", a.ID.Low, repr(_node_text(ta2)))
    print("AFTER B", b.ID.Low, repr(_node_text(tb2)))

    # Overlap gate on these two ranges vs nearby callout text
    ra, rb = a.Range, b.Range

    def box(r):
        return (
            float(r.Low.X),
            float(r.Low.Y),
            float(r.High.X),
            float(r.High.Y),
        )

    ba, bb = box(ra), box(rb)
    pad = 0.25

    def overlap(p, q):
        return not (
            p[2] + pad < q[0]
            or q[2] + pad < p[0]
            or p[3] + pad < q[1]
            or q[3] + pad < p[1]
        )

    clashes = []
    if overlap(ba, bb):
        clashes.append(("A↔B", ba, bb))

    o2 = Dispatch("MicroStationDGN.ElementScanCriteria")
    o2.ExcludeNonGraphical()
    ee2 = mr.Scan(o2)
    others = []
    while ee2.MoveNext():
        el = ee2.Current
        try:
            if not el.IsTextNodeElement:
                continue
            if el.ID.Low in (a.ID.Low, b.ID.Low):
                continue
            rng = el.Range
            lx, hx = float(rng.Low.X), float(rng.High.X)
            ly, hy = float(rng.Low.Y), float(rng.High.Y)
            # neighborhood of both callouts
            if hx < 1029300 or lx > 1029500:
                continue
            if hy < 239880 or ly > 239970:
                continue
            others.append((el.ID.Low, (lx, ly, hx, hy), _node_text(el.AsTextNodeElement)[:60]))
        except Exception:
            continue

    for label, box1 in (("A", ba), ("B", bb)):
        for oid, ob, ot in others:
            if overlap(box1, ob):
                clashes.append((f"{label}↔{oid}", ot, box1, ob))

    print(f"overlap_clashes={len(clashes)}")
    for c in clashes:
        print(" CLASH", c)
    print("DONE")


if __name__ == "__main__":
    main()
