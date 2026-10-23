"""Rebuild sheet NOTES as stacked TextElements (left-justified) on nyu_sheet_model.

Replaces the single TE id=69445 that held literal newlines in one line box.
"""
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

# Paper-space line pitch (sheet model units). Text height ~0.01.
LINE_PITCH = 0.014


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    print("file", app.ActiveDesignFile.FullName)
    if "tsp_02" not in app.ActiveDesignFile.FullName.lower():
        raise SystemExit("wrong file")

    for m in app.ActiveDesignFile.Models:
        if m.Name == "nyu_sheet_model":
            try:
                m.Activate()
            except Exception:
                app.CadInputQueue.SendKeyin('MODEL ACTIVE "nyu_sheet_model"')
            time.sleep(0.4)
            break
    mr = app.ActiveModelReference
    print("model", mr.Name)

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    host = None
    tmpl = None
    while ee.MoveNext():
        el = ee.Current
        try:
            if not el.IsTextElement:
                continue
            te = el.AsTextElement
            t = te.Text
            if el.ID.Low == 69445 or t.startswith("NOTES"):
                host = el
            if tmpl is None and len(t.strip()) <= 4:
                tmpl = el  # small scale label as font template fallback
            if tmpl is None:
                tmpl = el
        except Exception:
            continue

    if host is None:
        raise SystemExit("NOTES host TE not found")
    if tmpl is None:
        tmpl = host

    hr = host.Range
    # Left origin: original Low.X was negative (-0.795) after wide string —
    # use a sensible notes column X near sheet left. Prefer original mast-arm
    # corner: scale labels sit around x=0.43; notes typically left of plan.
    # Use host Low.X if reasonable, else 0.05.
    x0 = float(hr.Low.X)
    y_top = float(hr.High.Y)
    if x0 < -0.2:
        # Wide corrupted range — put notes at typical left margin near scale bar
        x0 = 0.05
    print(f"place origin x0={x0:.3f} y_top={y_top:.3f} host={host.ID.Low}")

    rot = host.AsTextElement.Rotation
    # Capture height from template
    tr = tmpl.Range
    th = float(tr.High.Y) - float(tr.Low.Y)
    pitch = max(LINE_PITCH, th * 1.35)
    print(f"tmpl_h={th:.4f} pitch={pitch:.4f}")

    # Remove old host
    mr.RemoveElement(host)

    placed = []
    for i, line in enumerate(NOTES_LINES):
        y = y_top - i * pitch
        origin = app.Point3dFromXYZ(x0, y, 0)
        created = app.CreateTextElement1(tmpl, line, origin, rot)
        if isinstance(created, tuple):
            created = created[0]
        # Left justify if property exists
        try:
            te = created.AsTextElement
            if hasattr(te, "Justification"):
                te.Justification = 0  # Left
            created.Rewrite()
        except Exception:
            pass
        mr.AddElement(created)
        created.Rewrite()
        placed.append((created.ID.Low, line[:50], y))

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.4)

    print(f"placed={len(placed)} SAVED")
    for eid, ln, y in placed:
        print(f"  id={eid} y={y:.3f} {ln!r}")


if __name__ == "__main__":
    main()
