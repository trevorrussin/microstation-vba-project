"""Left-justify NOTES TEs 69446-69452 on nyu_sheet_model; save."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

# Bentley: 0-5 = left* variants (origin at left edge); 7 = center
LEFT_JUST = 2

IDS = {69446, 69447, 69448, 69449, 69450, 69451, 69452}
X0 = 0.05  # left margin origin


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

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    fixed = []
    while ee.MoveNext():
        el = ee.Current
        try:
            if el.ID.Low not in IDS or not el.IsTextElement:
                continue
            te = el.AsTextElement
            # Move origin to left margin (Y unchanged)
            origin = te.Origin
            new_origin = app.Point3dFromXYZ(X0, float(origin.Y), float(origin.Z))
            # CreateTextElement doesn't expose Move easily — use MoveElement or rewrite origin
            try:
                te.Origin = new_origin
            except Exception:
                # fallback: Move by delta
                dx = X0 - float(origin.X)
                app.CadInputQueue.SendKeyin("NULL")  # no-op placeholder
                el.Move(app.Point3dFromXYZ(dx, 0, 0))

            ts = te.TextStyle
            ts.Justification = LEFT_JUST
            try:
                el.SetTextStyle(ts)
            except Exception:
                te.TextStyle = ts
            el.Rewrite()
            r = el.Range
            fixed.append(
                (
                    el.ID.Low,
                    te.Text[:40],
                    float(r.Low.X),
                    float(r.High.X),
                    float(r.Low.Y),
                )
            )
        except Exception as e:
            print("fail", getattr(el, "ID", None), e)

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.3)
    print(f"fixed={len(fixed)} SAVED left_just={LEFT_JUST} x0={X0}")
    for row in sorted(fixed, key=lambda x: -x[4]):
        print(f"  id={row[0]} x=[{row[2]:.3f},{row[3]:.3f}] y={row[4]:.3f} {row[1]!r}")


if __name__ == "__main__":
    main()
