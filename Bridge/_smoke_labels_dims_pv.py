"""Rebuild order table, adopt south align, smoke labels/dims/PV."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import bridge_client as bc
import view_capture
import wztc_ops as ops

ops.set_bridge(bc.bridge)

OUTWARD = -1.0
ALIGN_ID = 55431  # south ~3000ft E-W
SHEET_ELEMS = "MergingTaper|DownstreamTaper|ProtectiveVehicle|ArrowPanel|ChannelizingDevices"


def main() -> int:
    print("build order table", flush=True)
    ot = ops.build_wztc_order_table(
        speed=45,
        road_type="Non-Freeway",
        lane_width=12,
        shoulder_width="< 5 ft",
        sheet_num="619-311",
        sign_rows=[
            {"align_idx": 1, "sign_num": "W20-01RA", "side": "One Side"},
            {"align_idx": 1, "sign_num": "W20-05RA", "side": "One Side"},
            {"align_idx": 1, "sign_num": "W04-02R", "side": "One Side"},
            {"align_idx": 2, "sign_num": "G20-02", "side": "One Side"},
        ],
    )
    print("order", ot.get("status"), "rows", len(ot.get("rows") or []), flush=True)

    print("adopt align", flush=True)
    ad = bc.bridge.call("ADOPT_ALIGNMENT_ELEMENT", alignIdx=1, elementId=str(ALIGN_ID))
    print(ad, flush=True)
    if ad.get("status") != "OK":
        return 1

    print("labels", flush=True)
    r1 = bc.bridge.call("PLACE_ORDER_TABLE_LABELS", alignIdx=1, outwardSign=OUTWARD, textExtraAlong=20)
    print(r1.get("status"), r1.get("note"), "ids", r1.get("createdElementIds"), flush=True)
    print("resultFile rows", len(r1.get("rows") or []), flush=True)

    print("dims", flush=True)
    r2 = bc.bridge.call("PLACE_ORDER_TABLE_DIMENSIONS", alignIdx=1, outwardSign=OUTWARD, offsetDist=25)
    print(r2.get("status"), "ids", r2.get("createdElementIds"), "nrows", len(r2.get("rows") or []), flush=True)

    print("symbols", flush=True)
    r3 = bc.bridge.call(
        "PLACE_SHEET_SYMBOL_CELLS",
        alignIdx=1,
        sheetElements=SHEET_ELEMS,
        outwardSign=OUTWARD,
    )
    print(r3.get("status"), r3.get("rows") or r3, flush=True)

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MicroStationDGN.Application")
    try:
        v = app.ActiveDesignFile.Views(1)
        v.Origin = app.Point3dFromXYZ(1019800, 216880, 0)
        v.Extents = app.Point3dFromXYZ(3100, 280, 1)
        v.Redraw()
    except Exception as ex:
        print("view", ex, flush=True)
    out = Path(r"c:\repos\microstation-vba-project\Bridge\captures\review_labels_dims_pv.png")
    print("capture", view_capture.capture_microstation(out), flush=True)

    ok = r1.get("status") == "OK" and r2.get("status") == "OK" and r3.get("status") == "OK"
    return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
