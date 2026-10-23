"""Rebuild last 3 estimate rows as SAMPLE 2-line wraps; extend col frames.

Does not touch rows above 680.8223. Authority: SAMPLE wraps in
_fix_bts_table_format.py (823 / 830 / 837 all two lines).
"""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom  # noqa: E402
import ms_connect  # noqa: E402
from win32com.client import Dispatch  # noqa: E402
from view_capture import capture_window  # noqa: E402

ITEM_X = 1030191.0
DESC_X = 1030272.88
UNIT_X = 1030347.37
QTY_X = 1030367.63
LINE_PITCH = 5.80
H2 = 12.80
# Bottom of 680.8223 band — last-3 stack starts here
TOP = 239568.534

ROWS = [
    (
        "680.82310011",
        ["FURNISH AND INSTALL TRAFFIC SIGNAL POLE", "TYPE M-2A (NYC)"],
        "EA",
        "2",
    ),
    (
        "680.83080011",
        [
            "SHAFT EXTENSION AND SINGLE STREET LIGHT LUMINAIRE",
            "SUPPORT ARM FOR M-2A TYPE TRAFFIC SIGNAL POLE (NYC)",
        ],
        "EA",
        "1",
    ),
    (
        "680.83720011",
        [
            "FURNISH AND INSTALL 5 FT EXTENSION ARM FOR",
            "20 FT MAST ARM (NYC)",
        ],
        "EA",
        "2",
    ),
]


def _make_band(app, mr, tmpl, x0, x1, y0, y1):
    pts = (
        app.Point3dFromXYZ(x0, y0, 0),
        app.Point3dFromXYZ(x1, y0, 0),
        app.Point3dFromXYZ(x1, y1, 0),
        app.Point3dFromXYZ(x0, y1, 0),
    )
    created = app.CreateShapeElement1(tmpl, pts)
    if isinstance(created, tuple):
        created = created[0]
    mr.AddElement(created)
    created.Rewrite()
    return created


def _set_node_lines(tn, lines):
    tn.DeleteAllTextLines()
    for line in lines:
        tn.AddTextLine(line if line.strip() else " ")
    tn.Rewrite()


def main() -> None:
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = None
    for m in app.ActiveDesignFile.Models:
        if m.Name == "XS Stacked Views":
            mr = m
            break
    if mr is None:
        raise SystemExit("XS Stacked Views not found")

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)

    remove = []
    tall_cols = []
    tmpl_item = tmpl_unit = tmpl_qty = tmpl_desc = band_tmpl = None
    band_x0 = band_x1 = None

    while ee.MoveNext():
        el = ee.Current
        try:
            rng = el.Range
            lvl = el.Level.Name
        except Exception:
            continue
        cy = (float(rng.High.Y) + float(rng.Low.Y)) / 2
        cx = (float(rng.High.X) + float(rng.Low.X)) / 2
        if float(rng.Low.X) < 1030150 or float(rng.High.X) > 1030400:
            continue

        if "Table Border" in lvl and int(el.Type) == 6:
            h = float(rng.High.Y) - float(rng.Low.Y)
            w = float(rng.High.X) - float(rng.Low.X)
            y0 = float(rng.Low.Y)
            y1 = float(rng.High.Y)
            if w > 180 and 5 < h < 14 and y1 <= TOP + 0.05 and y0 < TOP:
                remove.append(el)
                if band_tmpl is None:
                    band_tmpl = el
                    band_x0, band_x1 = float(rng.Low.X), float(rng.High.X)
            elif h > 20 and y1 > 239700:
                tall_cols.append(el)
            continue

        if lvl != "TCSH_P_T":
            continue
        # Body text belonging to last-3 bands (below TOP)
        if cy >= TOP - 0.05:
            continue
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            t = " ".join(
                tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)
            ).strip()
            if t in (
                "ITEM",
                "DESCRIPTION",
                "UNIT",
                "QUANTITY",
                "ESTIMATE OF QUANTITIES",
            ):
                continue
            if abs(cx - ITEM_X) < 8 and tmpl_item is None:
                tmpl_item = el
            if abs(cx - UNIT_X) < 8 and tmpl_unit is None:
                tmpl_unit = el
            if abs(cx - QTY_X) < 8 and tmpl_qty is None:
                tmpl_qty = el
            remove.append(el)
        elif el.IsTextElement and abs(cx - DESC_X) < 60:
            if tmpl_desc is None:
                tmpl_desc = el
            remove.append(el)

    # Also grab templates from rows above if missing
    if not all([tmpl_item, tmpl_unit, tmpl_qty, tmpl_desc, band_tmpl]):
        o2 = Dispatch("MicroStationDGN.ElementScanCriteria")
        o2.ExcludeNonGraphical()
        ee2 = mr.Scan(o2)
        while ee2.MoveNext():
            el = ee2.Current
            try:
                rng = el.Range
                lvl = el.Level.Name
            except Exception:
                continue
            cy = (float(rng.High.Y) + float(rng.Low.Y)) / 2
            cx = (float(rng.High.X) + float(rng.Low.X)) / 2
            if float(rng.Low.X) < 1030150 or float(rng.High.X) > 1030400:
                continue
            if lvl == "O_Details_Table Border_P" and int(el.Type) == 6:
                h = float(rng.High.Y) - float(rng.Low.Y)
                w = float(rng.High.X) - float(rng.Low.X)
                if band_tmpl is None and w > 180 and 5 < h < 14 and cy < 239707:
                    band_tmpl = el
                    band_x0 = float(rng.Low.X)
                    band_x1 = float(rng.High.X)
            if lvl != "TCSH_P_T" or cy < TOP or cy > 239707:
                continue
            if el.IsTextNodeElement:
                if abs(cx - ITEM_X) < 8 and tmpl_item is None:
                    tmpl_item = el
                if abs(cx - UNIT_X) < 8 and tmpl_unit is None:
                    tmpl_unit = el
                if abs(cx - QTY_X) < 8 and tmpl_qty is None:
                    tmpl_qty = el
            elif el.IsTextElement and abs(cx - DESC_X) < 60 and tmpl_desc is None:
                tmpl_desc = el

    if not all([tmpl_item, tmpl_unit, tmpl_qty, tmpl_desc, band_tmpl]):
        raise SystemExit(
            f"missing tmpl item={tmpl_item} unit={tmpl_unit} "
            f"qty={tmpl_qty} desc={tmpl_desc} band={band_tmpl}"
        )

    print(f"remove={len(remove)} tall_cols={len(tall_cols)}")
    for el in remove:
        try:
            mr.RemoveElement(el)
        except Exception as ex:
            print("rm", ex)

    bottom = TOP - len(ROWS) * H2
    row_ys = [TOP - H2 / 2 - i * H2 for i in range(len(ROWS))]
    print(f"bottom={bottom:.3f} row_ys={[round(y, 3) for y in row_ys]}")

    for cy in row_ys:
        _make_band(app, mr, band_tmpl, band_x0, band_x1, cy - H2 / 2, cy + H2 / 2)

    # Extend / replace tall columns to new bottom
    for el in tall_cols:
        r = el.Range
        x0, x1 = float(r.Low.X), float(r.High.X)
        y1 = float(r.High.Y)
        _make_band(app, mr, el, x0, x1, bottom, y1)
        mr.RemoveElement(el)
        print(f"col -> y=[{bottom:.3f},{y1:.3f}] x=[{x0:.1f},{x1:.1f}]")

    rot = tmpl_item.AsTextNodeElement.Rotation

    def place_node(tmpl, x, y, text):
        c = tmpl.Clone()
        r = c.Range
        cx = (float(r.High.X) + float(r.Low.X)) / 2
        cy = (float(r.High.Y) + float(r.Low.Y)) / 2
        c.Move(app.Point3dFromXYZ(x - cx, y - cy, 0))
        mr.AddElement(c)
        _set_node_lines(c.AsTextNodeElement, [text])

    for (item, lines, unit, qty), cy in zip(ROWS, row_ys):
        place_node(tmpl_item, ITEM_X, cy, item)
        place_node(tmpl_unit, UNIT_X, cy, unit)
        place_node(tmpl_qty, QTY_X, cy, qty)
        ys = [cy + LINE_PITCH / 2, cy - LINE_PITCH / 2]
        for text, y in zip(lines, ys):
            origin = app.Point3dFromXYZ(DESC_X, y, 0)
            te = app.CreateTextElement1(tmpl_desc, text, origin, rot)[0]
            mr.AddElement(te)
            te.Rewrite()
        print(f"{item} cy={cy:.3f} lines=2")

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.4)

    v = app.ActiveDesignFile.Views(1)
    v.Center = app.Point3dFromXYZ(1030275, 239550, 0)
    v.Extents = app.Point3dFromXYZ(220, 80, 0)
    app.CadInputQueue.SendKeyin("UPDATE VIEW EXTENDED 1")
    time.sleep(2.2)
    p = capture_window(
        "MicroStation",
        out_path=r"c:\repos\microstation-vba-project\scripts\_bts_fmt\table_last3_wrapped.png",
    )
    print("capture", p)


if __name__ == "__main__":
    main()
