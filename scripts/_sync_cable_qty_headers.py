"""Final sure qty-chip pass for 09.18: sync CABLE headers to line qty sums;
apply chip 98 13-cond (5) on nearest callout 8653.
"""
from __future__ import annotations

import re
import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch


def walk(el, out):
    try:
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            lines = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
            t = "\n".join(lines).upper()
            if "TRAFFIC SIGNAL CONDUIT" in t or (
                "TRAFFIC SIGNAL CABLE" in t and "680." in t
            ):
                out.append({"id": el.ID.Low, "el": el, "lines": lines})
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


def qty_of(line: str) -> int:
    m = re.search(r"\((\d+)\)(\s*\(SEE NOTE \d+\))?$", line)
    return int(m.group(1)) if m else 0


def cable_sum(lines) -> int:
    total = 0
    for ln in lines:
        u = ln.upper()
        if u.startswith("TRAFFIC SIGNAL CABLE") or u.startswith(
            "TRAFFIC SIGNAL CONDUIT"
        ):
            continue
        if any(
            k in u
            for k in (
                "CONDUCTORS",
                "GROUND WIRE",
                "SINGLE CONDUCTOR",
            )
        ):
            total += qty_of(ln)
    return total


def set_qty(line: str, q: int) -> str:
    m = re.search(r"\((\d+)\)(\s*\(SEE NOTE \d+\))?$", line)
    if not m:
        return line
    return line[: m.start()] + f"({q})" + (m.group(2) or "")


def main():
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = app.ActiveModelReference
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    live = []
    while ee.MoveNext():
        walk(ee.Current, live)
    by_id = {x["id"]: x for x in live}

    # chip 98: 13-cond qty (5) on 8653 (nearest)
    lines = list(by_id[8653]["lines"])
    for i, ln in enumerate(lines):
        if "13 CONDUCTORS" in ln.upper():
            lines[i] = set_qty(ln, 5)
    by_id[8653]["lines"] = lines  # pending

    edits = 0
    touched = 0
    for tid, L in by_id.items():
        before = [
            L["el"].AsTextNodeElement.TextLine(i)
            for i in range(1, L["el"].AsTextNodeElement.TextLinesCount + 1)
        ]
        after = list(L["lines"]) if tid != 8653 else lines
        if tid == 8653:
            after = lines
        else:
            after = list(before)

        # sync CABLE (N) to sum of detail qtys
        s = cable_sum(after)
        for i, ln in enumerate(after):
            if ln.upper().startswith("TRAFFIC SIGNAL CABLE"):
                new_h = f"TRAFFIC SIGNAL CABLE ({s})"
                if after[i] != new_h:
                    after[i] = new_h
                break

        if after == before:
            continue
        touched += 1
        print(f"\nid={tid}")
        for a, b in zip(before, after):
            if a != b:
                edits += 1
                print(f"  - {a!r}")
                print(f"  + {b!r}")
        set_lines(L["el"], after)

    if touched:
        app.CadInputQueue.SendKeyin("SAVE DESIGN")
        time.sleep(0.4)
    print(f"\nSAVED callouts={touched} line_edits={edits}")

    # verify sums
    print("\n=== VERIFY CABLE headers ===")
    o2 = Dispatch("MicroStationDGN.ElementScanCriteria")
    o2.ExcludeNonGraphical()
    ee2 = mr.Scan(o2)
    live2 = []
    while ee2.MoveNext():
        walk(ee2.Current, live2)
    for L in sorted({x["id"]: x for x in live2}.values(), key=lambda x: x["id"]):
        s = cable_sum(L["lines"])
        hdr = next(
            (ln for ln in L["lines"] if ln.upper().startswith("TRAFFIC SIGNAL CABLE")),
            "",
        )
        ok = hdr.endswith(f"({s})")
        print(f"{'OK' if ok else 'BAD'} id={L['id']} {hdr} sum={s}")


if __name__ == "__main__":
    main()
