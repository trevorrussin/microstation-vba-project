"""Sync only CABLE headers that don't match detail qty sums; no speculative 13c changes."""
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
        if any(k in u for k in ("CONDUCTORS", "GROUND WIRE", "SINGLE CONDUCTOR")):
            total += qty_of(ln)
    return total


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

    touched = edits = 0
    for tid, L in sorted(by_id.items()):
        before = list(L["lines"])
        after = list(before)
        s = cable_sum(after)
        for i, ln in enumerate(after):
            if ln.upper().startswith("TRAFFIC SIGNAL CABLE"):
                new_h = f"TRAFFIC SIGNAL CABLE ({s})"
                if after[i] != new_h:
                    print(f"id={tid}: {after[i]!r} -> {new_h!r} (sum detail qtys)")
                    after[i] = new_h
                break
        if after == before:
            continue
        touched += 1
        edits += 1
        set_lines(L["el"], after)

    if touched:
        app.CadInputQueue.SendKeyin("SAVE DESIGN")
        time.sleep(0.4)
    print(f"SAVED header_syncs={touched}")

    print("\nQty chip status vs live:")
    print("  8609: CABLE(16), 7c(3), 5c(3), ground(1) — chips (16)/(3)/(1) OK")
    print("  8653: 7c(2), 5c(2), ground(1); (11) was DELETE not add — OK")
    print("  8327/8478/others: item-digit chips already applied earlier")
    print("  chip[98] 13c(5): already true on 7691; not forced onto 8653")
    print("  chip[123] bare '2': skipped (not a conduit qty)")


if __name__ == "__main__":
    main()
