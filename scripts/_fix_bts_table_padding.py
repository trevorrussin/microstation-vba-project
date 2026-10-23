"""Grow estimate-table row padding: contiguous bands, text clear of blue lines.

Safe geometry only: CreateShapeElement1 / Clone+Move. Never ModifyVertex.
"""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom  # noqa: E402
import ms_connect  # noqa: E402
from win32com.client import Dispatch  # noqa: E402

ITEM_X = 1030191.0
DESC_X = 1030272.88
UNIT_X = 1030347.37
QTY_X = 1030367.63
# TextElement Range height is ~5.0 here — pitch must exceed that or
# the two line boxes overlap (live 2026-09-17: pitch 3.4 → box_clear=-1.6).
LINE_PITCH = 5.80
CELL_PAD = 0.75
# Just tall enough for pad + text (1-line / 2-line)
H1 = 7.40
# 2 lines: top-of-upper to bottom-of-lower ≈ LINE_PITCH + text_h(~5) + 2*pad
H2 = 12.80
HEADER_BOTTOM = 239707.134

# Prefer one line when length matches existing 1-line rows in this column.
# Keep 2 lines only for the longest SAMPLE-style wraps.
DESCS = {
    "680.5001": ["POLE EXCAVATION AND CONCRETE FOUNDATION"],
    "680.510501": ["PULLBOX, RECTANGULAR, 26 X 18 INCH, REINFORCED CONCRETE"],
    "680.51100610": ["CLEAN EXISTING PULLBOX"],
    "680.520606": ['TRAFFIC SIGNAL CONDUIT, RIGID PLASTIC, CLASS 2, 2"'],
    "680.730514": ["SIGNAL CABLE 5 CONDUCTORS, 14 AWG"],
    "680.730714": ["SIGNAL CABLE 7 CONDUCTORS, 14 AWG"],
    "680.731314": ["SIGNAL CABLE 13 CONDUCTORS, 14 AWG"],
    "680.807200011": [
        "FURNISH AND INSTALL TRAFFIC SIGNAL POLE TYPE S-1A (NYC)"
    ],
    "680.8141": ["PEDESTRIAN SIGNAL BRACKET MOUNT ASSEMBLY"],
    "680.815002": [
        "PEDESTRIAN SIGNAL MODULE 16-INCH X 18-INCH,",
        "BI-MODAL HAND/MAN SYMBOLS WITH COUNTDOWN TIMER, LED",
    ],
    "680.8151": ["ACCESSIBLE PEDESTRIAN SIGNAL (APS) WITHOUT POST"],
    "680.8152": ["ACCESSIBLE PEDESTRIAN SIGNAL (APS) W POST"],
    "680.81812011": ["FURNISH 3 SECTION TRAFFIC SIGNAL LENS, 12 INCH (NYC)"],
    "680.81822011": [
        "INSTALL TRAFFIC SIGNAL LENS, 8 IN AND 12 INCH (NYC)"
    ],
    "680.820500511": ["FURNISH AND INSTALL 20 FT SINGLE MAST ARM ASSEMBLY"],
    "680.8223": ["BREAKAWAY TRANSFORMER BASE"],
    "680.82310011": [
        "FURNISH AND INSTALL TRAFFIC SIGNAL POLE",
        "TYPE M-2A (NYC)",
    ],
    "680.83080011": [
        "SHAFT EXTENSION AND SINGLE STREET LIGHT LUMINAIRE",
        "SUPPORT ARM FOR M-2A TYPE TRAFFIC SIGNAL POLE (NYC)",
    ],
    "680.83720011": [
        "FURNISH AND INSTALL 5 FT EXTENSION ARM FOR",
        "20 FT MAST ARM (NYC)",
    ],
}

# item -> (unit, qty) from current live table intent
META = {
    "680.5001": ("CY", "12.1"),
    "680.510501": ("EA", "4"),
    "680.51100610": ("EA", "1"),
    "680.520606": ("LF", "1400"),
    "680.730514": ("LF", " "),
    "680.730714": ("LF", " "),
    "680.731314": ("LF", "1400"),
    "680.807200011": ("EA", "1"),
    "680.8141": ("EA", "4"),
    "680.815002": ("EA", "4"),
    "680.8151": ("EA", "5"),
    "680.8152": ("EA", "3"),
    "680.81812011": ("EA", "12"),
    "680.81822011": ("EA", "12"),
    "680.820500511": ("EA", "3"),
    "680.8223": ("EA", "3"),
    "680.82310011": ("EA", "2"),
    "680.83080011": ("EA", "1"),
    "680.83720011": ("EA", "2"),
}

ORDER = list(DESCS.keys())


def _set_node_lines(tn, lines):
    tn.DeleteAllTextLines()
    for line in lines:
        tn.AddTextLine(line if line.strip() else " ")
    tn.Rewrite()


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

    body_text = []
    old_bands = []
    tall_cols = []
    band_tmpl = None
    tmpl_item = tmpl_unit = tmpl_qty = tmpl_desc = None

    while ee.MoveNext():
        el = ee.Current
        try:
            rng = el.Range
            lvl = el.Level.Name
        except Exception:
            continue
        cy = (float(rng.High.Y) + float(rng.Low.Y)) / 2
        cx = (float(rng.High.X) + float(rng.Low.X)) / 2
        # Include a little below the prior bottom so old 2-line orphans get wiped.
        if cy < 239545 or cy > 239725:
            continue
        if float(rng.Low.X) < 1030150 or float(rng.High.X) > 1030400:
            continue

        if lvl == "O_Details_Table Border_P":
            h = float(rng.High.Y) - float(rng.Low.Y)
            w = float(rng.High.X) - float(rng.Low.X)
            if int(el.Type) == 6 and w > 150 and 5 < h < 14 and cy < HEADER_BOTTOM:
                old_bands.append(el)
                if band_tmpl is None:
                    band_tmpl = el
            elif int(el.Type) == 3 and 5 < h < 14 and cy < HEADER_BOTTOM:
                old_bands.append(el)  # short verts with bands
            elif int(el.Type) == 6 and h > 20:
                tall_cols.append(el)
            continue

        if lvl != "TCSH_P_T":
            continue
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            j = " ".join(
                tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)
            ).strip()
            if j in (
                "ITEM",
                "DESCRIPTION",
                "UNIT",
                "QUANTITY",
                "ESTIMATE OF QUANTITIES",
            ):
                continue
            if cy > HEADER_BOTTOM:
                continue
            if abs(cx - ITEM_X) < 5 and tmpl_item is None:
                tmpl_item = el
            if abs(cx - UNIT_X) < 5 and tmpl_unit is None:
                tmpl_unit = el
            if abs(cx - QTY_X) < 5 and tmpl_qty is None:
                tmpl_qty = el
            body_text.append(el)
        elif el.IsTextElement and abs(cx - DESC_X) < 50 and cy < HEADER_BOTTOM:
            if tmpl_desc is None and cy > 239560:
                tmpl_desc = el
            body_text.append(el)

    if not all([tmpl_item, tmpl_unit, tmpl_qty, tmpl_desc, band_tmpl]):
        raise SystemExit("missing templates")

    # Band X span from template
    br = band_tmpl.Range
    x0, x1 = float(br.Low.X), float(br.High.X)

    def park(el, x, y):
        c = el.Clone()
        r = c.Range
        cx = (float(r.High.X) + float(r.Low.X)) / 2
        cy = (float(r.High.Y) + float(r.Low.Y)) / 2
        c.Move(app.Point3dFromXYZ(x - cx, y - cy, 0))
        mr.AddElement(c)
        c.Rewrite()
        return c

    py = 239740.0
    p_item = park(tmpl_item, ITEM_X, py)
    p_unit = park(tmpl_unit, UNIT_X, py)
    p_qty = park(tmpl_qty, QTY_X, py)
    p_desc = park(tmpl_desc, DESC_X, py)
    p_band = park(band_tmpl, (x0 + x1) / 2, py)
    rot = p_item.AsTextNodeElement.Rotation

    print(f"remove bands={len(old_bands)} texts={len(body_text)}")
    for el in old_bands + body_text:
        try:
            mr.RemoveElement(el)
        except Exception as ex:
            print("rm", ex)

    # Contiguous row centers top -> bottom
    heights = [H2 if len(DESCS[it]) > 1 else H1 for it in ORDER]
    row_ys = []
    y_hi = HEADER_BOTTOM
    for h in heights:
        cy = y_hi - h / 2
        row_ys.append(cy)
        y_hi = y_hi - h
    bottom = y_hi
    print(f"bottom={bottom:.3f} rows={len(ORDER)}")

    for cy, h in zip(row_ys, heights):
        _make_band(app, mr, p_band, x0, x1, cy - h / 2, cy + h / 2)

    # Extend tall columns to new bottom via replace (axis-aligned)
    for el in tall_cols:
        r = el.Range
        tx0, tx1 = float(r.Low.X), float(r.High.X)
        ty1 = float(r.High.Y)
        _make_band(app, mr, el, tx0, tx1, bottom, ty1)
        try:
            mr.RemoveElement(el)
        except Exception:
            pass
        print(f"col extended bottom->{bottom:.3f}")

    def place_node(tmpl, x, y, text):
        c = tmpl.Clone()
        r = c.Range
        cx = (float(r.High.X) + float(r.Low.X)) / 2
        cy = (float(r.High.Y) + float(r.Low.Y)) / 2
        c.Move(app.Point3dFromXYZ(x - cx, y - cy, 0))
        mr.AddElement(c)
        _set_node_lines(c.AsTextNodeElement, [text])

    placed = 0
    for item, cy, h in zip(ORDER, row_ys, heights):
        unit, qty = META[item]
        lines = DESCS[item]
        place_node(p_item, ITEM_X, cy, item)
        place_node(p_unit, UNIT_X, cy, unit)
        place_node(p_qty, QTY_X, cy, qty)
        n = len(lines)
        if n == 1:
            ys = [cy]
        else:
            ys = [cy + LINE_PITCH / 2, cy - LINE_PITCH / 2]
        for text, y in zip(lines, ys):
            origin = app.Point3dFromXYZ(DESC_X, y, 0)
            te = app.CreateTextElement1(p_desc, text, origin, rot)[0]
            mr.AddElement(te)
            te.Rewrite()
            placed += 1
        print(f"{item} cy={cy:.3f} h={h:.2f} lines={n}")

    for el in (p_item, p_unit, p_qty, p_desc, p_band):
        try:
            mr.RemoveElement(el)
        except Exception:
            pass

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.4)

    # Peek check
    o2 = Dispatch("MicroStationDGN.ElementScanCriteria")
    o2.ExcludeNonGraphical()
    ee2 = mr.Scan(o2)
    bands = []
    descs = []
    while ee2.MoveNext():
        el = ee2.Current
        try:
            rng = el.Range
            lvl = el.Level.Name
        except Exception:
            continue
        cy = (float(rng.High.Y) + float(rng.Low.Y)) / 2
        cx = (float(rng.High.X) + float(rng.Low.X)) / 2
        if cy < bottom - 1 or cy > HEADER_BOTTOM + 1:
            continue
        if float(rng.Low.X) < 1030150 or float(rng.High.X) > 1030400:
            continue
        if lvl == "O_Details_Table Border_P" and int(el.Type) == 6:
            w = float(rng.High.X) - float(rng.Low.X)
            h = float(rng.High.Y) - float(rng.Low.Y)
            if w > 150 and 5 < h < 14:
                bands.append((float(rng.Low.Y), float(rng.High.Y), cy, h))
        if lvl == "TCSH_P_T" and el.IsTextElement and abs(cx - DESC_X) < 50:
            descs.append(
                (
                    el.AsTextElement.Text.strip()[:40],
                    float(rng.Low.Y),
                    float(rng.High.Y),
                    cy,
                )
            )
    peeks = 0
    for txt, ly, hy, cy in descs:
        band = min(bands, key=lambda b: abs(b[2] - cy))
        ot = hy - (band[1] - CELL_PAD)
        ob = (band[0] + CELL_PAD) - ly
        if ot > 0.05 or ob > 0.05:
            peeks += 1
            print(f"PEEK {txt!r} ot={ot:.2f} ob={ob:.2f} h={band[3]:.2f}")
    print(f"placed={placed} peek_vs_pad={peeks} H1={H1} H2={H2} SAVED")


if __name__ == "__main__":
    main()
