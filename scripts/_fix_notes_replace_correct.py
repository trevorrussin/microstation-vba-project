"""Fix mistaken table overwrite; put 6 NOTES on correct left corner mast-arm text."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

NOTES_LINES = [
    "NOTES",
    "1. REFER TO DWG. TSN-01 FOR TRAFFIC SIGNAL NOTES AND LEGEND AND DETAILS.",
    "2. ENSURE THAT EACH CONDUIT DOES NOT HAVE MORE THAN 9 LIVE CABLES.",
    "3. FOR MAST ARM FOOTING, (3'-0\" X 3'-0\" WITH 5'-0\" DEPTH) REFER TO NYCDOT TRAFFIC SIGNAL STANDARD DWG. F-002. FOR S-1 POLE FOOTING, REFER TO NYCDOT TRAFFIC SIGNAL STANDARD 01.",
    "4. FURNISH AND INSTALL THAT MEETS ALL REQUIREMENTS OF NYCDOT TRAFFIC SIDEWALK BOX DETAIL DRAWING NUMBERS MISC-008_1 AND MISC-008_2.",
    "5. PULL IN EXISTING CABLES TO NEW CONDUIT.",
    "6. PULL IN NEW CABLES TO EXISTING CONDUIT.",
]


def walk(el, out, depth=0):
    try:
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            lines = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
            t = "\n".join(lines)
            rng = el.Range
            out.append(
                {
                    "el": el,
                    "id": el.ID.Low,
                    "kind": "TN",
                    "lines": lines,
                    "t": t,
                    "tu": t.upper(),
                    "cx": (float(rng.Low.X) + float(rng.High.X)) / 2,
                    "cy": (float(rng.Low.Y) + float(rng.High.Y)) / 2,
                    "x0": float(rng.Low.X),
                }
            )
        elif el.IsTextElement:
            te = el.AsTextElement
            t = te.Text
            rng = el.Range
            out.append(
                {
                    "el": el,
                    "id": el.ID.Low,
                    "kind": "TE",
                    "lines": [t],
                    "t": t,
                    "tu": t.upper(),
                    "cx": (float(rng.Low.X) + float(rng.High.X)) / 2,
                    "cy": (float(rng.Low.Y) + float(rng.High.Y)) / 2,
                    "x0": float(rng.Low.X),
                }
            )
        if int(el.Type) in (2, 7) and depth < 5:
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk(ee.Current, out, depth + 1)
    except Exception:
        pass


def scan(mr):
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    out = []
    while ee.MoveNext():
        walk(ee.Current, out)
    return out


def set_node_lines(el, lines):
    tn = el.AsTextNodeElement
    n = int(tn.TextLinesCount)
    for i, line in enumerate(lines, start=1):
        if i <= n:
            tn.SetTextLine(i, line)
        else:
            tn.AddTextLine(line)
    for i in range(len(lines) + 1, n + 1):
        tn.SetTextLine(i, " ")
    el.Rewrite()


def activate_xs(app):
    for m in app.ActiveDesignFile.Models:
        if m.Name == "XS Stacked Views":
            try:
                m.Activate()
            except Exception:
                app.CadInputQueue.SendKeyin('MODEL ACTIVE "XS Stacked Views"')
            time.sleep(0.5)
            break
    print("model", app.ActiveModelReference.Name)


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    print("file", app.ActiveDesignFile.FullName)
    activate_xs(app)
    mr = app.ActiveModelReference
    texts = scan(mr)

    # 1) Restore table desc line if still corrupted
    for x in texts:
        if x["id"] == 67947 or (
            abs(x["cx"] - 1030273) < 5
            and "NOTES" in x["tu"]
            and "REFER TO DWG" in x["tu"]
        ):
            print("fixing table cell id", x["id"], "was", x["t"][:60])
            if x["kind"] == "TE":
                x["el"].AsTextElement.Text = "20 FT MAST ARM (NYC)"
                x["el"].Rewrite()
            else:
                set_node_lines(x["el"], ["20 FT MAST ARM (NYC)"])

    texts = scan(mr)

    # 2) List ALL mast-arm related outside table X band
    print("--- candidates outside table (x0 < 1030150) ---")
    outside = []
    for x in texts:
        if x["x0"] >= 1030150:
            continue
        tu = x["tu"]
        if any(
            k in tu
            for k in (
                "MAST ARM",
                "20 FT",
                "25'",
                "NOTES",
                "REFER TO",
                "PULL IN",
                "ENSURE",
            )
        ):
            outside.append(x)
            print(
                f"  id={x['id']} {x['kind']} cy={x['cy']:.1f} cx={x['cx']:.0f} "
                f"{x['t'][:100]!r}"
            )

    # Also show anything with MAST anywhere for diagnosis
    print("--- all MAST ARM texts ---")
    for x in texts:
        if "MAST ARM" in x["tu"]:
            print(
                f"  id={x['id']} {x['kind']} cy={x['cy']:.1f} cx={x['cx']:.0f} "
                f"{x['t'][:100]!r}"
            )

    # Prefer left-of-table 20 FT MAST (not table DESC_X)
    targets = [
        x
        for x in texts
        if "MAST ARM" in x["tu"]
        and "20" in x["tu"]
        and x["x0"] < 1030150
        and "NOTES\n" not in x["t"]
    ]
    # If notes corner used a longer callout containing 20 FT SINGLE MAST
    if not targets:
        targets = [
            x
            for x in texts
            if "20 FT" in x["tu"]
            and "MAST" in x["tu"]
            and abs(x["cx"] - 1030273) > 20  # not desc column
        ]

    if not targets:
        app.CadInputQueue.SendKeyin("SAVE DESIGN")
        raise SystemExit(
            "Table restore done if needed; no left-corner 20ft mast arm found to replace"
        )

    targets.sort(key=lambda z: (z["cx"], z["cy"]))
    target = targets[0]
    print("REPLACE id", target["id"], target["t"][:80])

    if target["kind"] == "TN":
        set_node_lines(target["el"], NOTES_LINES)
    else:
        # TextElement: place as one string — engineer asked for notes text here.
        # Better: if nearby TextNode template exists left-justified, use clone.
        # For now SetText with newlines (ORD often keeps as one TE).
        target["el"].AsTextElement.Text = "\n".join(NOTES_LINES)
        target["el"].Rewrite()

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.4)
    print("SAVED")


if __name__ == "__main__":
    main()
