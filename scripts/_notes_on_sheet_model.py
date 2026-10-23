"""Find 20ft mast arm on nyu_sheet_model (sheet) and replace with 6 NOTES."""
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
                    "y0": float(rng.Low.Y),
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
                    "y0": float(rng.Low.Y),
                }
            )
        if int(el.Type) in (2, 7) and depth < 6:
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


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    print("file", app.ActiveDesignFile.FullName)

    # Activate sheet model
    for m in app.ActiveDesignFile.Models:
        if m.Name == "nyu_sheet_model":
            try:
                m.Activate()
            except Exception:
                app.CadInputQueue.SendKeyin('MODEL ACTIVE "nyu_sheet_model"')
            time.sleep(0.5)
            break
    mr = app.ActiveModelReference
    print("model", mr.Name)

    texts = scan(mr)
    print("texts", len(texts))

    print("--- MAST / 20 FT / NOTES ---")
    hits = []
    for x in texts:
        tu = x["tu"]
        if any(
            k in tu
            for k in (
                "MAST ARM",
                "20 FT",
                "NOTES",
                "REFER TO DWG",
                "PULL IN",
                "ENSURE THAT",
                "25'",
                "SINGLE MAST",
            )
        ):
            hits.append(x)
            print(
                f"  id={x['id']} {x['kind']} cy={x['cy']:.2f} cx={x['cx']:.2f} "
                f"{x['t'][:120]!r}"
            )

    # Target: 20 FT mast arm (not a long estimate string)
    targets = [
        x
        for x in hits
        if "20" in x["tu"]
        and "MAST" in x["tu"]
        and "NOTES" not in x["tu"]
        and "FURNISH AND INSTALL 20 FT SINGLE" not in x["tu"]  # table row
    ]
    if not targets:
        targets = [x for x in hits if "20 FT MAST ARM" in x["tu"]]

    if not targets:
        raise SystemExit("no sheet-model 20ft mast arm target")

    # Prefer shortest / notes-corner style
    targets.sort(key=lambda z: (len(z["t"]), z["cx"], z["cy"]))
    target = targets[0]
    print("TARGET", target["id"], target["t"][:100])

    if target["kind"] == "TN":
        set_node_lines(target["el"], NOTES_LINES)
    else:
        # CreateTextElement won't wrap as notes block well.
        # Replace TE text with full notes as multiline if supported;
        # also try converting by using a TextNode clone if any NOTES TN nearby.
        tns = [h for h in hits if h["kind"] == "TN"]
        tmpl = tns[0]["el"] if tns else None
        if tmpl is not None:
            # Move template clone to target origin, set lines, delete old TE
            rot = tmpl.AsTextNodeElement.Rotation
            origin = app.Point3dFromXYZ(target["x0"], target["y0"], 0)
            # Simpler: just SetText on TE — user asked replace text
            target["el"].AsTextElement.Text = "\n".join(NOTES_LINES)
            target["el"].Rewrite()
        else:
            target["el"].AsTextElement.Text = "\n".join(NOTES_LINES)
            target["el"].Rewrite()

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.4)
    print("SAVED into", target["id"])


if __name__ == "__main__":
    main()
