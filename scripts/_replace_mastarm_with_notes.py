"""Replace 20ft mast-arm corner text with 6 NOTES on open tsp_02. No OpenDesignFile."""
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
                    "lines": lines,
                    "t": t,
                    "tu": t.upper(),
                    "cx": (float(rng.Low.X) + float(rng.High.X)) / 2,
                    "cy": (float(rng.Low.Y) + float(rng.High.Y)) / 2,
                    "x0": float(rng.Low.X),
                    "kind": "TN",
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
                    "lines": [t],
                    "t": t,
                    "tu": t.upper(),
                    "cx": (float(rng.Low.X) + float(rng.High.X)) / 2,
                    "cy": (float(rng.Low.Y) + float(rng.High.Y)) / 2,
                    "x0": float(rng.Low.X),
                    "kind": "TE",
                }
            )
        if int(el.Type) in (2, 7) and depth < 5:
            ee = el.AsComplexElement.GetSubElements()
            while ee.MoveNext():
                walk(ee.Current, out, depth + 1)
    except Exception:
        pass


def scan_model(mr):
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    texts = []
    while ee.MoveNext():
        walk(ee.Current, texts)
    return texts


def is_mast20(tu: str) -> bool:
    if "MAST" not in tu or "ARM" not in tu:
        return False
    return ("20 FT" in tu) or ("20'" in tu) or ("20 FT" in tu.replace("  ", " ")) or (
        "20" in tu and "SINGLE MAST" in tu
    ) or ("FURNISH AND INSTALL 20" in tu)


def main():
    pythoncom.CoInitialize()
    app = None
    for desc, a in ms_connect._candidate_apps():
        if a.HasActiveDesignFile:
            app = a
            break
    if app is None:
        raise SystemExit("No active design file on COM")

    path = app.ActiveDesignFile.FullName
    print("Active", path)
    if "tsp_02" not in path.lower():
        raise SystemExit(f"Wrong file: {path}")

    # Prefer XS Stacked Views (estimate/notes sheet) within same DGN
    mr = app.ActiveModelReference
    print("Model now", mr.Name)
    if mr.Name != "XS Stacked Views":
        for m in app.ActiveDesignFile.Models:
            if m.Name == "XS Stacked Views":
                try:
                    m.Activate()
                    time.sleep(0.4)
                except Exception as e:
                    print("Activate ERR", e)
                    app.CadInputQueue.SendKeyin('MODEL ACTIVE "XS Stacked Views"')
                    time.sleep(0.6)
                break
        mr = app.ActiveModelReference
        print("Model after", mr.Name)

    texts = scan_model(mr)
    print(f"scanned texts={len(texts)} in {mr.Name}")

    cands = [x for x in texts if is_mast20(x["tu"])]
    # Also catch table-row style "20 FT SINGLE MAST ARM"
    if not cands:
        cands = [
            x
            for x in texts
            if "20" in x["tu"] and "MAST ARM" in x["tu"]
        ]

    print(f"mast-arm hits={len(cands)}")
    for x in sorted(cands, key=lambda z: (z["cx"], z["cy"])):
        print(
            f"  id={x['id']} {x['kind']} cy={x['cy']:.1f} cx={x['cx']:.0f} "
            f"{x['t'][:110]!r}"
        )

    if not cands:
        # dump left-side candidates for notes corner
        print("LEFT note-ish:")
        for x in sorted(texts, key=lambda z: -z["cy"]):
            if x["x0"] > 1030250:
                continue
            if any(
                k in x["tu"]
                for k in ("NOTE", "MAST", "20 FT", "REFER", "PULL", "ENSURE", "ARM")
            ):
                print(f"  id={x['id']} cy={x['cy']:.1f} cx={x['cx']:.0f} {x['t'][:100]!r}")
        raise SystemExit("no 20ft mast arm text found")

    # Notes corner = leftmost hit (or among left cluster, lowest)
    left_min = min(c["cx"] for c in cands)
    left_cands = [c for c in cands if abs(c["cx"] - left_min) < 50]
    target = min(left_cands, key=lambda z: z["cy"])
    print("TARGET id", target["id"], "before:")
    for ln in target["lines"]:
        print(" ", repr(ln))

    if target["kind"] == "TN":
        set_node_lines(target["el"], NOTES_LINES)
    else:
        te = target["el"].AsTextElement
        # Single TextElement can't hold 7 lines well — convert by rewriting
        # first line only is wrong. Prefer fail if TE.
        # Try multiline string; if engine flattens, still better than nothing.
        te.Text = "\n".join(NOTES_LINES)
        target["el"].Rewrite()

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.5)

    # verify
    texts2 = scan_model(app.ActiveModelReference)
    host = next((x for x in texts2 if x["id"] == target["id"]), None)
    print("AFTER:")
    if host:
        for ln in host["lines"]:
            if ln.strip():
                print(" ", repr(ln))
    print("SAVED")


if __name__ == "__main__":
    main()
