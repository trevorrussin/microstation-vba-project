"""Fix 09.18 qty chips carefully; revert bad auto-mapping first."""
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


def set_qty_suffix(line: str, qty: str) -> str:
    """Set trailing (N) before optional SEE NOTE; keep rest."""
    m = re.search(r"\(\d+\)(\s*\(SEE NOTE \d+\))?$", line)
    if not m:
        return line
    return line[: m.start()] + qty + (m.group(1) or "")


def set_cable_header(lines, n: int):
    for i, ln in enumerate(lines):
        if ln.upper().startswith("TRAFFIC SIGNAL CABLE"):
            lines[i] = f"TRAFFIC SIGNAL CABLE ({n})"
            return True
    return False


def find_line(lines, pred):
    for i, ln in enumerate(lines):
        if pred(ln.upper()):
            return i
    return None


def main():
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = app.ActiveModelReference
    print("DGN", app.ActiveDesignFile.FullName)

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    live = []
    while ee.MoveNext():
        walk(ee.Current, live)
    by_id = {x["id"]: x for x in live}

    # --- Revert / set known-good state from markup semantics ---
    # 09.18 chips (sure):
    # 8609 CABLE(16): header already 16; 5-cond qty (3) [chip 109]; ground (1) [111]
    # 8653 CABLE(15): 7-cond -> 730714 (2) [97]; 5-cond (2) [110];
    #      red deletes stray (11) — do NOT put (11) on ground;
    #      chip 98 13-cond (5) may be a different callout — leave 13c at (7) unless sure
    # 7335: revert ground (2)->(1); bare '2' chip was not for this callout

    plans = {}

    # 7335 revert ground qty
    if 7335 in by_id:
        lines = list(by_id[7335]["lines"])
        i = find_line(lines, lambda u: "GROUND WIRE" in u)
        if i is not None:
            lines[i] = set_qty_suffix(lines[i], "(1)")
        plans[7335] = lines

    # 8609: fix 5-cond back to (3); ensure CABLE (16); ground (1)
    if 8609 in by_id:
        lines = list(by_id[8609]["lines"])
        set_cable_header(lines, 16)
        i = find_line(lines, lambda u: "5 CONDUCTORS" in u)
        if i is not None:
            lines[i] = set_qty_suffix(lines[i], "(3)")
        i = find_line(lines, lambda u: "7 CONDUCTORS" in u)
        if i is not None:
            lines[i] = set_qty_suffix(lines[i], "(3)")
        i = find_line(lines, lambda u: "GROUND WIRE" in u)
        if i is not None:
            lines[i] = set_qty_suffix(lines[i], "(1)")
        plans[8609] = lines

    # 8653: 7-cond (2); 5-cond (2); ground (1); keep 13c (7); CABLE (15)
    # chip 67 '730514 (1)' conflicts with chip 110 '(2)' for 5-cond —
    # red/green pair at 730714 (2) (11) means 7-cond is (2) and (11) deleted.
    # For 5-cond, chip 110 (2) is nearer the 5-cond cluster than 730514 (1)
    # which may mean item digits only with example qty — prefer (2).
    if 8653 in by_id:
        lines = list(by_id[8653]["lines"])
        set_cable_header(lines, 15)
        i = find_line(lines, lambda u: "13 CONDUCTORS" in u)
        if i is not None:
            lines[i] = set_qty_suffix(lines[i], "(7)")
        i = find_line(lines, lambda u: "7 CONDUCTORS" in u)
        if i is not None:
            lines[i] = set_qty_suffix(lines[i], "(2)")
        i = find_line(lines, lambda u: "5 CONDUCTORS" in u)
        if i is not None:
            lines[i] = set_qty_suffix(lines[i], "(2)")
        i = find_line(lines, lambda u: "GROUND WIRE" in u)
        if i is not None:
            lines[i] = set_qty_suffix(lines[i], "(1)")
        plans[8653] = lines

    # chip 78 (2) @ far right — likely another callout ground/single?
    # Skip unless sure.

    # Look for callout that needs 13-cond (5) from chip 98 —
    # only if a line still has wrong item 680.71314 or qty mismatch.
    # Search all for 13-cond and report; apply (5) only if cable total implies it.
    print("=== 13-conductor lines (chip 98 candidate) ===")
    for L in live:
        for ln in L["lines"]:
            if "13 CONDUCTORS" in ln.upper():
                print(f"  id={L['id']}: {ln}")

    print("\n=== APPLY ===")
    touched = 0
    edits = 0
    for tid, after in plans.items():
        before = by_id[tid]["lines"]
        if before == after:
            print(f"id={tid} already ok")
            continue
        touched += 1
        print(f"id={tid}")
        for a, b in zip(before, after):
            if a != b:
                edits += 1
                print(f"  - {a!r}")
                print(f"  + {b!r}")
        set_lines(by_id[tid]["el"], after)

    if touched:
        app.CadInputQueue.SendKeyin("SAVE DESIGN")
        time.sleep(0.4)
    print(f"\nSAVED callouts={touched} line_edits={edits}")
    print(
        "QUESTION: chip[98] wants a 13-cond line qty (5). "
        "None applied — which callout should get (5) instead of current qty?"
    )


if __name__ == "__main__":
    main()
