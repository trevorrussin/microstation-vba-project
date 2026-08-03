"""Rebuild state, place workspace + channelizing + sheet symbols (PV/AP)."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import bridge_client as bc
import view_capture
import wztc_ops as ops

ops.set_bridge(bc.bridge)

OUTWARD = -1.0
ALIGN_ID = 55431
SHEET_ELEMS = (
    "MergingTaper|ShoulderTaper|DownstreamTaper|"
    "ProtectiveVehicle|ArrowPanel|ChannelizingDevices"
)


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
    if ot.get("status") != "OK":
        print(ot, flush=True)
        return 1

    print("adopt align", flush=True)
    ad = bc.bridge.call("ADOPT_ALIGNMENT_ELEMENT", alignIdx=1, elementId=str(ALIGN_ID))
    print(ad, flush=True)
    if ad.get("status") != "OK":
        return 1

    print("workspace", flush=True)
    r_ws = bc.bridge.call(
        "PLACE_ORDER_TABLE_WORKSPACE", alignIdx=1, outwardSign=OUTWARD, laneWidth=12
    )
    print(r_ws, flush=True)

    print("channelizing", flush=True)
    r_ch = bc.bridge.call(
        "PLACE_ORDER_TABLE_CHANNELIZING", alignIdx=1, outwardSign=OUTWARD, laneWidth=12
    )
    print(r_ch, flush=True)

    print("symbols PV/AP", flush=True)
    r_sym = bc.bridge.call(
        "PLACE_SHEET_SYMBOL_CELLS",
        alignIdx=1,
        sheetElements=SHEET_ELEMS,
        outwardSign=OUTWARD,
    )
    print(r_sym.get("status"), r_sym.get("rows") or r_sym, flush=True)

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

    out = Path(r"c:\repos\microstation-vba-project\Bridge\captures\review_ws_chan_ap.png")
    print("capture", view_capture.capture_microstation(out), flush=True)

    # close-up: path start → VS / PV / WS
    try:
        v.Origin = app.Point3dFromXYZ(1019700, 216980, 0)
        v.Extents = app.Point3dFromXYZ(350, 120, 1)
        v.Redraw()
        out2 = Path(r"c:\repos\microstation-vba-project\Bridge\captures\review_ws_pv_closeup.png")
        print("capture", view_capture.capture_microstation(out2), flush=True)
    except Exception as ex:
        print("closeup1", ex, flush=True)

    # close-up: shoulder taper / AP
    try:
        v.Origin = app.Point3dFromXYZ(1020800, 216980, 0)
        v.Extents = app.Point3dFromXYZ(350, 120, 1)
        v.Redraw()
        out3 = Path(r"c:\repos\microstation-vba-project\Bridge\captures\review_ap_shoulder_closeup.png")
        print("capture", view_capture.capture_microstation(out3), flush=True)
    except Exception as ex:
        print("closeup2", ex, flush=True)

    ok = (
        r_ws.get("status") == "OK"
        and r_ch.get("status") == "OK"
        and r_sym.get("status") == "OK"
    )
    return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
