"""Restore 6 NOTES on tsp_02 / XS Stacked Views (bottom-left).

Does NOT open/switch files. Aborts if wrong DGN is active.
Authority: BTS markup captures + r1 FreeText [78] for notes 5-6.
"""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

# Post-markup NOTES (1-4 from notes_markup.png; 5-6 from r1 FreeText [78])
NOTES_LINES = [
    "NOTES",
    "1. REFER TO DWG. TSN-01 FOR TRAFFIC SIGNAL NOTES AND LEGEND AND DETAILS.",
    "2. ENSURE THAT EACH CONDUIT DOES NOT HAVE MORE THAN 9 LIVE CABLES.",
    "3. FOR MAST ARM FOOTING, (3'-0\" X 3'-0\" WITH 5'-0\" DEPTH) REFER TO NYCDOT TRAFFIC SIGNAL STANDARD DWG. F-002. FOR S-1 POLE FOOTING, REFER TO NYCDOT TRAFFIC SIGNAL STANDARD 01.",
    "4. FURNISH AND INSTALL THAT MEETS ALL REQUIREMENTS OF NYCDOT TRAFFIC SIDEWALK BOX DETAIL DRAWING NUMBERS MISC-008_1 AND MISC-008_2.",
    "5. PULL IN EXISTING CABLES TO NEW CONDUIT.",
    "6. PULL IN NEW CABLES TO EXISTING CONDUIT.",
]


def walk(el, out):
    try:
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            lines = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
            t = "\n".join(lines).upper()
            if "NOTES" in t or "REFER TO DWG" in t or "PULL IN" in t or "ENSURE THAT EACH" in t:
                rng = el.Range
                out.append(
                    {
                        "el": el,
                        "id": el.ID.Low,
                        "lines": lines,
                        "cx": (float(rng.Low.X) + float(rng.High.X)) / 2,
                        "cy": (float(rng.Low.Y) + float(rng.High.Y)) / 2,
                        "x0": float(rng.Low.X),
                        "y0": float(rng.Low.Y),
                        "y1": float(rng.High.Y),
                    }
                )
        elif el.IsTextElement:
            te = el.AsTextElement
            t = te.Text.upper()
            if any(
                k in t
                for k in (
                    "NOTES",
                    "REFER TO DWG",
                    "PULL IN",
                    "ENSURE THAT EACH",
                    "SIDEWALK BOX",
                    "MAST ARM FOOTING",
                )
            ):
                rng = el.Range
                out.append(
                    {
                        "el": el,
                        "id": el.ID.Low,
                        "lines": [te.Text],
                        "cx": (float(rng.Low.X) + float(rng.High.X)) / 2,
                        "cy": (float(rng.Low.Y) + float(rng.High.Y)) / 2,
                        "x0": float(rng.Low.X),
                        "y0": float(rng.Low.Y),
                        "y1": float(rng.High.Y),
                        "is_te": True,
                    }
                )
        if int(el.Type) in (2, 7):
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk(ee.Current, out)
    except Exception:
        pass


def main():
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    try:
        path = app.ActiveDesignFile.FullName
    except Exception:
        raise SystemExit("No design file active — open tsp_02 yourself, then re-run.")
    print("DGN", path)
    if "tsp_02" not in path.lower():
        raise SystemExit(
            f"Wrong file active ({path}). Open tsp_02 yourself — will not switch."
        )

    mr = None
    for m in app.ActiveDesignFile.Models:
        if m.Name == "XS Stacked Views":
            mr = m
            break
    if mr is None:
        mr = app.ActiveModelReference
        print("WARN using active model", mr.Name)

    # Prefer activated model if already XS Stacked Views
    if app.ActiveModelReference.Name == "XS Stacked Views":
        mr = app.ActiveModelReference

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    found = []
    while ee.MoveNext():
        walk(ee.Current, found)

    print(f"note-ish elements: {len(found)}")
    for f in sorted(found, key=lambda x: -x["cy"]):
        head = " | ".join(ln[:50] for ln in f["lines"][:2])
        print(f"  id={f['id']} cy={f['cy']:.1f} cx={f['cx']:.0f} {head!r}")

    # Find NOTES title or first note as template (left-justified)
    title = next(
        (f for f in found if any(ln.strip().upper() == "NOTES" for ln in f["lines"])),
        None,
    )
    if title is None:
        # leftmost note-ish TextNode
        tns = [f for f in found if not f.get("is_te")]
        title = min(tns, key=lambda x: x["x0"]) if tns else None
    if title is None:
        raise SystemExit(
            "No NOTES template found on sheet. Need an existing left-justified "
            "text element to clone — cannot invent style."
        )

    print("TEMPLATE id=", title["id"], "lines=", title["lines"][:3])

    # If a multi-line NOTES TextNode exists, rewrite it; else report need for place
    el = title["el"]
    if el.IsTextNodeElement:
        tn = el.AsTextNodeElement
        # If this node is ONLY the title, we may need sibling lines —
        # prefer rewriting a fat node that already held note body.
        body_nodes = [
            f
            for f in found
            if not f.get("is_te")
            and f["id"] != title["id"]
            and abs(f["cx"] - title["cx"]) < 80
        ]
        # Pick the largest TextNode near title as the notes body host
        host = title
        if body_nodes:
            host = max(
                [title] + body_nodes,
                key=lambda f: len([ln for ln in f["lines"] if ln.strip()]),
            )
        print("HOST id=", host["id"], "nlines=", len(host["lines"]))
        tn = host["el"].AsTextNodeElement
        n = int(tn.TextLinesCount)
        for i, line in enumerate(NOTES_LINES, start=1):
            if i <= n:
                tn.SetTextLine(i, line)
            else:
                tn.AddTextLine(line)
        for i in range(len(NOTES_LINES) + 1, n + 1):
            tn.SetTextLine(i, " ")
        host["el"].Rewrite()
        app.CadInputQueue.SendKeyin("SAVE DESIGN")
        time.sleep(0.4)
        print("RESTORED into TextNode", host["id"])
        for ln in NOTES_LINES:
            print(" ", ln)
        return

    raise SystemExit("NOTES title is TextElement — need manual place path; aborting.")


if __name__ == "__main__":
    main()
