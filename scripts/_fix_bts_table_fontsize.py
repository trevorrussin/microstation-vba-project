"""Rebuild estimate table body text at header font size (fix microscopic text)."""
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
LINE_PITCH = 2.70
HEADER_BOTTOM = 239707.134

DESCS = {
    "680.5001": ["POLE EXCAVATION AND CONCRETE FOUNDATION"],
    "680.510501": ["PULLBOX, RECTANGULAR, 26 X 18 INCH, REINFORCED CONCRETE"],
    "680.51100610": ["CLEAN EXISTING PULLBOX"],
    "680.520606": [
        "TRAFFIC SIGNAL CONDUIT, RIGID PLASTIC,",
        'CLASS 2, 2"',
    ],
    "680.730514": ["SIGNAL CABLE 5 CONDUCTORS, 14 AWG"],
    "680.730714": ["SIGNAL CABLE 7 CONDUCTORS, 14 AWG"],
    "680.731314": ["SIGNAL CABLE 13 CONDUCTORS, 14 AWG"],
    "680.807200011": [
        "FURNISH AND INSTALL TRAFFIC SIGNAL POLE",
        "TYPE S-1A (NYC)",
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
        "INSTALL TRAFFIC SIGNAL LENS,",
        "8 IN AND 12 INCH (NYC)",
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


def main() -> None:
    pythoncom.CoInitialize()
    _, app = next(ms_connect._candidate_apps())
    mr = None
    for m in app.ActiveDesignFile.Models:
        if m.Name == "XS Stacked Views":
            mr = m
            break
    if mr is None:
        raise SystemExit("no XS Stacked Views")

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)

    bands = []  # (cy, h)
    body_text = []
    hdr = {}  # name -> element (TextNode)

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
        if lvl == "O_Details_Table Border_P" and int(el.Type) == 6:
            w = float(rng.High.X) - float(rng.Low.X)
            h = float(rng.High.Y) - float(rng.Low.Y)
            if w > 150 and 5 < h < 14 and cy < HEADER_BOTTOM:
                bands.append((cy, h))
            continue
        if lvl != "TCSH_P_T":
            continue
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            j = tn.TextLine(1).strip()
            if j in ("ITEM", "DESCRIPTION", "UNIT", "QUANTITY", "ESTIMATE OF QUANTITIES"):
                hdr[j] = el
                continue
            if cy < HEADER_BOTTOM and cy > 239540:
                body_text.append(el)
        elif el.IsTextElement and abs(cx - DESC_X) < 50 and cy < HEADER_BOTTOM:
            body_text.append(el)

    if "ITEM" not in hdr or "DESCRIPTION" not in hdr:
        raise SystemExit(f"header templates missing: {list(hdr)}")

    bands.sort(key=lambda b: -b[0])
    if len(bands) != len(ORDER):
        print(f"WARN bands={len(bands)} rows={len(ORDER)}")

    # Use header TextNodes as style templates (rangeH~5). Clone them for body.
    # For DESCRIPTION column use CreateTextElement1 from DESCRIPTION header's
    # text component if possible.
    desc_hdr = hdr["DESCRIPTION"]
    item_hdr = hdr["ITEM"]
    unit_hdr = hdr["UNIT"]
    qty_hdr = hdr["QUANTITY"]

    # Extract TextElement template from DESCRIPTION header complex
    te_tmpl = None
    try:
        sub = desc_hdr.AsComplexElement.GetSubElements()
        while sub.MoveNext():
            c = sub.Current
            if c.IsTextElement:
                te_tmpl = c
                break
    except Exception:
        pass
    if te_tmpl is None:
        # fall back: create from header via CreateTextElement1(header_node, ...)
        te_tmpl = desc_hdr

    r = te_tmpl.Range
    print(
        f"desc tmpl rangeH={float(r.High.Y)-float(r.Low.Y):.3f} "
        f"rangeW={float(r.High.X)-float(r.Low.X):.3f}"
    )

    # Park header clones for item/unit/qty (TextNodes)
    def park_node(el, x, y):
        c = el.Clone()
        rr = c.Range
        cx = (float(rr.High.X) + float(rr.Low.X)) / 2
        cy = (float(rr.High.Y) + float(rr.Low.Y)) / 2
        c.Move(app.Point3dFromXYZ(x - cx, y - cy, 0))
        mr.AddElement(c)
        c.Rewrite()
        return c

    py = 239740.0
    p_item = park_node(item_hdr, ITEM_X, py)
    p_unit = park_node(unit_hdr, UNIT_X, py)
    p_qty = park_node(qty_hdr, QTY_X, py)
    # park a desc TextElement
    if te_tmpl.IsTextElement:
        p_desc = te_tmpl.Clone()
    else:
        # create one line from header
        origin = app.Point3dFromXYZ(DESC_X, py, 0)
        rot = desc_hdr.AsTextNodeElement.Rotation
        p_desc = app.CreateTextElement1(te_tmpl, "TEMPLATE", origin, rot)
        if isinstance(p_desc, tuple):
            p_desc = p_desc[0]
    rr = p_desc.Range
    cx = (float(rr.High.X) + float(rr.Low.X)) / 2
    cy = (float(rr.High.Y) + float(rr.Low.Y)) / 2
    p_desc.Move(app.Point3dFromXYZ(DESC_X - cx, py - cy, 0))
    mr.AddElement(p_desc)
    p_desc.Rewrite()
    rot = p_item.AsTextNodeElement.Rotation

    print(f"removing {len(body_text)} micro body texts")
    for el in body_text:
        try:
            mr.RemoveElement(el)
        except Exception:
            pass

    def place_node(tmpl, x, y, text):
        c = tmpl.Clone()
        rr = c.Range
        cx = (float(rr.High.X) + float(rr.Low.X)) / 2
        cy = (float(rr.High.Y) + float(rr.Low.Y)) / 2
        c.Move(app.Point3dFromXYZ(x - cx, y - cy, 0))
        mr.AddElement(c)
        _set_node_lines(c.AsTextNodeElement, [text])
        return c

    placed = 0
    for i, item in enumerate(ORDER):
        if i >= len(bands):
            break
        cy, h = bands[i]
        unit, qty = META[item]
        lines = DESCS[item]
        place_node(p_item, ITEM_X, cy, item)
        place_node(p_unit, UNIT_X, cy, unit)
        place_node(p_qty, QTY_X, cy, qty)
        if len(lines) == 1:
            ys = [cy]
        else:
            ys = [cy + LINE_PITCH / 2, cy - LINE_PITCH / 2]
        for text, y in zip(lines, ys):
            origin = app.Point3dFromXYZ(DESC_X, y, 0)
            te = app.CreateTextElement1(p_desc, text, origin, rot)
            if isinstance(te, tuple):
                te = te[0]
            mr.AddElement(te)
            te.Rewrite()
            rh = float(te.Range.High.Y) - float(te.Range.Low.Y)
            rw = float(te.Range.High.X) - float(te.Range.Low.X)
            if placed < 3:
                print(f"  sample desc rangeH={rh:.3f} rangeW={rw:.2f} {text[:40]}")
            placed += 1
        print(f"row {item} cy={cy:.3f}")

    for el in (p_item, p_unit, p_qty, p_desc):
        try:
            mr.RemoveElement(el)
        except Exception:
            pass

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.4)

    # verify sizes
    o2 = Dispatch("MicroStationDGN.ElementScanCriteria")
    o2.ExcludeNonGraphical()
    ee2 = mr.Scan(o2)
    tiny = 0
    ok = 0
    while ee2.MoveNext():
        el = ee2.Current
        try:
            if el.Level.Name != "TCSH_P_T":
                continue
            rng = el.Range
            cy = (float(rng.High.Y) + float(rng.Low.Y)) / 2
            if cy > HEADER_BOTTOM or cy < 239540:
                continue
            rh = float(rng.High.Y) - float(rng.Low.Y)
            if el.IsTextElement or el.IsTextNodeElement:
                if rh < 1.0:
                    tiny += 1
                else:
                    ok += 1
        except Exception:
            pass
    print(f"SAVED placed_desc={placed} ok_size={ok} tiny={tiny}")


if __name__ == "__main__":
    main()
