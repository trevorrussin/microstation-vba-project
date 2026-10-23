"""Find table bottom and left callout cluster bbox below it."""
from __future__ import annotations

import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch


def text_of(el):
    try:
        if el.IsTextElement:
            return el.AsTextElement.Text
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            return "\n".join(tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1))
    except Exception:
        pass
    return None


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    mr = app.ActiveModelReference

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)

    table_bits = []
    callout_tns = []
    while ee.MoveNext():
        el = ee.Current
        try:
            lvl = el.Level.Name if el.Level else ""
        except Exception:
            lvl = ""
        try:
            r = el.Range
            lx, ly = float(r.Low.X), float(r.Low.Y)
            hx, hy = float(r.High.X), float(r.High.Y)
        except Exception:
            continue

        # table-ish on left
        if "Table" in lvl or "Details" in lvl or "O_Details" in lvl:
            if hx < 1029450 and hy > 239700:
                table_bits.append(
                    (el.ID.Low, int(el.Type), lvl, lx, ly, hx, hy, hy - ly, hx - lx)
                )

        # TextNodes that look like left callouts (below table zone)
        try:
            if el.IsTextNodeElement and lx < 1029440:
                t = text_of(el)
                if t and len(t) > 15:
                    callout_tns.append(
                        {
                            "id": el.ID.Low,
                            "t": t.replace("\n", " | ")[:60],
                            "lx": lx,
                            "ly": ly,
                            "hx": hx,
                            "hy": hy,
                            "cy": (ly + hy) / 2,
                            "lvl": lvl,
                        }
                    )
        except Exception:
            pass

    print("=== table-ish elements (left, high Y) ===")
    for row in sorted(table_bits, key=lambda x: -x[6])[:30]:
        print(
            f"  id={row[0]} type={row[1]} lvl={row[2][:40]} "
            f"[{row[3]:.0f},{row[4]:.0f}]-[{row[5]:.0f},{row[6]:.0f}] "
            f"h={row[7]:.0f} w={row[8]:.0f}"
        )

    # Left callout column: X mostly < 1029435, cy below table
    # From prior scan, stack is ~239625-239774, X ~1029316-1029432
    left_stack = [
        c
        for c in callout_tns
        if c["hx"] < 1029445 and c["cy"] < 239790 and c["cy"] > 239600
    ]
    print(f"\n=== left stack TextNodes ({len(left_stack)}) ===")
    if not left_stack:
        return
    for c in sorted(left_stack, key=lambda x: -x["cy"]):
        print(
            f"  id={c['id']} cy={c['cy']:.1f} "
            f"[{c['lx']:.1f},{c['ly']:.1f}]-[{c['hx']:.1f},{c['hy']:.1f}] "
            f"{c['t']!r}"
        )

    lx = min(c["lx"] for c in left_stack)
    ly = min(c["ly"] for c in left_stack)
    hx = max(c["hx"] for c in left_stack)
    hy = max(c["hy"] for c in left_stack)
    print(f"\nUNION bbox textnodes: [{lx:.2f},{ly:.2f}]-[{hx:.2f},{hy:.2f}]")
    print(f"  w={hx-lx:.1f} h={hy-ly:.1f}")

    # Also include associated TE overlays in same column (duplicates of TN lines)
    # and stubs/dims? User said "callouts" - wrap the text blocks.
    # Check GraphicGroups / dims whose text end is in this box.


if __name__ == "__main__":
    main()
