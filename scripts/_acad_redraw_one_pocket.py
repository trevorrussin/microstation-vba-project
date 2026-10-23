"""Redraw one flare pocket from engineer TIP/LEFT/TAN picks (no guessing).

Usage:
  python scripts/_acad_redraw_one_pocket.py MidN-north
"""
from __future__ import annotations

import math
import sys
import pythoncom
from win32com.client import Dispatch, VARIANT

sys.path.insert(0, r"c:\repos\microstation-vba-project\scripts")
from acad_ask import ask_message, ask_points

LAYER = "PROP-D-ROAD-RAMPS"
NW_CENTER = (1028470.0, 187478.0)

POCKETS = {
    "MidN-north": {
        "label": "MidN NORTH (35.01 / over green arrow)",
        "zoom": (1028498.0, 187488.0, 12.0),
        "center": (1028497.0, 187488.0),
        "radius": 8.0,
        "hint": (
            "35.01 north pocket only (do not click the SE end).\n"
            "P1/P2/P3 = three corners of the yellow triangle.\n"
            "Right side should follow the tan curb past the green arrow."
        ),
    },
    "MidN-south": {
        "label": "MidN SOUTH",
        "zoom": (1028506.0, 187478.0, 12.0),
        "center": (1028507.0, 187479.0),
        "radius": 7.0,
        "hint": "Three corners of the south MidN yellow triangle (P1/P2/P3).",
    },
    "E-north": {
        "label": "East NORTH end",
        "zoom": (1028585.0, 187451.0, 14.0),
        "center": (1028585.0, 187451.0),
        "radius": 7.0,
        "hint": "Three corners of the east-north yellow triangle (P1/P2/P3).",
    },
    "E-south": {
        "label": "East SOUTH end",
        "zoom": (1028592.0, 187442.0, 14.0),
        "center": (1028592.0, 187442.0),
        "radius": 7.0,
        "hint": "Three corners of the east-south yellow triangle (P1/P2/P3).",
    },
    "MidS-north": {
        "label": "MidS NORTH end",
        "zoom": (1028525.0, 187435.0, 14.0),
        "center": (1028526.0, 187435.0),
        "radius": 12.0,
        "hint": (
            "TIP and LEFT on real corners (topo / curb).\n"
            "TAN may be empty space if there is no third vertex — "
            "click where the pocket should close so it is a real triangle."
        ),
    },
    "MidS-south": {
        "label": "MidS SOUTH end",
        "zoom": (1028524.0, 187420.0, 14.0),
        "center": (1028524.0, 187420.0),
        "radius": 8.0,
        "hint": "Three corners of the MidS south yellow triangle (P1/P2/P3).",
    },
}


def vpt(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z)))


def delete_near(doc, cx, cy, radius):
    ms = doc.ModelSpace
    handles = []
    for i in range(ms.Count):
        try:
            e = ms.Item(i)
            if e.ObjectName != "AcDbSolid" or e.Layer != LAYER:
                continue
            mn, mx = e.GetBoundingBox()
            x = (float(mn[0]) + float(mx[0])) / 2
            y = (float(mn[1]) + float(mx[1])) / 2
            if math.hypot(x - NW_CENTER[0], y - NW_CENTER[1]) < 28:
                continue
            if math.hypot(x - cx, y - cy) <= radius:
                handles.append(e.Handle)
        except Exception:
            continue
    for h in handles:
        try:
            doc.HandleToObject(h).Delete()
        except Exception:
            pass
    return len(handles)


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in POCKETS:
        print("Usage: _acad_redraw_one_pocket.py <id>")
        print("ids:", ", ".join(POCKETS))
        sys.exit(2)
    pid = sys.argv[1]
    pocket = POCKETS[pid]

    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument

    notes = ask_message(
        title=pocket["label"],
        message=pocket["hint"] + "\n\nClick OK (or type continue) to pick 3 points in AutoCAD.\n"
        "Type stop/cancel in the box to abort.",
    )
    if notes.lower().strip() in ("stop", "cancel", "abort", "stop / cancel"):
        print("ABORTED_BY_USER", notes, flush=True)
        return
    n = delete_near(doc, pocket["center"][0], pocket["center"][1], pocket["radius"])
    print("deleted", n, flush=True)

    # Reduce OSNAP forcing duplicate corners
    try:
        doc.SendCommand("(setvar \"OSMODE\" 0) ")
    except Exception:
        pass

    p1, p2, p3 = ask_points(
        title=pocket["label"],
        message=pocket["hint"]
        + "\n\nPick any 3 corners of the yellow triangle (P1, P2, P3).",
        prompts=[
            "1/3 Point 1",
            "2/3 Point 2",
            "3/3 Point 3 (must differ from 1 and 2)",
        ],
        zoom=pocket["zoom"],
        banner=True,
        rubber_band_from_first=True,
    )
    # Reject near-degenerate
    def area(a, b, c):
        return abs(
            (a[0] * (b[1] - c[1]) + b[0] * (c[1] - a[1]) + c[0] * (a[1] - b[1])) / 2
        )

    if area(p1, p2, p3) < 0.05:
        print("DEGENERATE", p1, p2, p3, flush=True)
        ask_message(
            title="Degenerate — not placed",
            message="Those 3 points have ~zero area (two were the same).\n"
            "Nothing new was kept from this attempt beyond deletes.",
        )
        sys.exit(3)

    sol = doc.ModelSpace.AddSolid(vpt(*p1), vpt(*p2), vpt(*p3), vpt(*p3))
    sol.Layer = LAYER
    sol.Color = 2
    try:
        doc.Regen(1)
    except Exception:
        pass
    print("POCKET", pid, flush=True)
    print("HANDLE", sol.Handle, flush=True)
    print("P1", p1, flush=True)
    print("P2", p2, flush=True)
    print("P3", p3, flush=True)
    ask_message(
        title="Placed " + pocket["label"],
        message=f"handle {sol.Handle}\nP1 {p1}\nP2 {p2}\nP3 {p3}",
    )


if __name__ == "__main__":
    main()
