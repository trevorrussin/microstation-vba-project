"""Delete NOTES TEs and recreate left-justified stack on nyu_sheet_model."""
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

LEFT_JUST = 2
X0 = 0.05
Y_TOP = 0.278
LINE_PITCH = 0.0142


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    if "tsp_02" not in app.ActiveDesignFile.FullName.lower():
        raise SystemExit("wrong file")

    for m in app.ActiveDesignFile.Models:
        if m.Name == "nyu_sheet_model":
            try:
                m.Activate()
            except Exception:
                app.CadInputQueue.SendKeyin('MODEL ACTIVE "nyu_sheet_model"')
            time.sleep(0.3)
            break
    mr = app.ActiveModelReference
    print("model", mr.Name)

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    to_remove = []
    tmpl = None
    rot = None
    while ee.MoveNext():
        el = ee.Current
        try:
            if not el.IsTextElement:
                continue
            te = el.AsTextElement
            t = te.Text
            if el.ID.Low in range(69446, 69453) or t.startswith("NOTES") or t.startswith(
                "1. REFER TO DWG"
            ) or t.startswith("2. ENSURE") or t.startswith("3. FOR MAST") or t.startswith(
                "4. FURNISH"
            ) or t.startswith("5. PULL IN") or t.startswith("6. PULL IN"):
                to_remove.append(el)
            if tmpl is None and len(t.strip()) <= 6:
                tmpl = el
                rot = te.Rotation
        except Exception:
            continue

    if tmpl is None:
        raise SystemExit("no template")
    if rot is None:
        rot = tmpl.AsTextElement.Rotation

    print(f"removing {len(to_remove)} old note TEs")
    for el in to_remove:
        try:
            mr.RemoveElement(el)
        except Exception as e:
            print("  remove fail", el.ID.Low, e)

    placed = []
    for i, line in enumerate(NOTES_LINES):
        y = Y_TOP - i * LINE_PITCH
        origin = app.Point3dFromXYZ(X0, y, 0)
        created = app.CreateTextElement1(tmpl, line, origin, rot)
        if isinstance(created, tuple):
            created = created[0]
        ts = created.AsTextElement.TextStyle
        ts.Justification = LEFT_JUST
        try:
            created.SetTextStyle(ts)
        except Exception:
            created.AsTextElement.TextStyle = ts
        mr.AddElement(created)
        created.Rewrite()
        r = created.Range
        placed.append(
            (
                created.ID.Low,
                line[:48],
                float(r.Low.X),
                float(r.High.X),
                float(r.Low.Y),
            )
        )

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.3)
    print(f"placed={len(placed)} SAVED")
    for row in placed:
        print(f"  id={row[0]} x=[{row[2]:.3f},{row[3]:.3f}] y={row[4]:.3f} {row[1]!r}")


if __name__ == "__main__":
    main()
