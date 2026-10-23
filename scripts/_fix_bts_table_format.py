"""Replace ESTIMATE OF QUANTITIES description TextNodes with TextElements.

Existing TextNodes auto-wrap at ~22 chars (too narrow). SAMPLE LIST OF PAY
ITEMS keeps short/medium descriptions on ONE centered line and only wraps at
natural phrase breaks. Markup SHOW AS 1 LINE requires BREAKAWAY on one line.
"""
from __future__ import annotations

import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom  # noqa: E402
import ms_connect  # noqa: E402
from win32com.client import Dispatch  # noqa: E402

# SAMPLE-style wraps (page 4 LIST OF PAY ITEMS). One string = one visual line.
ROWS = [
    ("680.5001", ["POLE EXCAVATION AND CONCRETE FOUNDATION"], "CY", "12.1"),
    (
        "680.510501",
        ["PULLBOX, RECTANGULAR, 26 X 18 INCH, REINFORCED CONCRETE"],
        "EA",
        "4",
    ),
    ("680.51100610", ["CLEAN EXISTING PULLBOX"], "EA", "1"),
    (
        "680.520606",
        ["TRAFFIC SIGNAL CONDUIT, RIGID PLASTIC,", 'CLASS 2, 2"'],
        "LF",
        "1400",
    ),
    ("680.730514", ["SIGNAL CABLE 5 CONDUCTORS, 14 AWG"], "LF", ""),
    ("680.730714", ["SIGNAL CABLE 7 CONDUCTORS, 14 AWG"], "LF", ""),
    ("680.731314", ["SIGNAL CABLE 13 CONDUCTORS, 14 AWG"], "LF", "1400"),
    (
        "680.807200011",
        ["FURNISH AND INSTALL TRAFFIC SIGNAL POLE", "TYPE S-1A (NYC)"],
        "EA",
        "1",
    ),
    ("680.8141", ["PEDESTRIAN SIGNAL BRACKET MOUNT ASSEMBLY"], "EA", "4"),
    (
        "680.815002",
        [
            "PEDESTRIAN SIGNAL MODULE 16-INCH X 18-INCH,",
            "BI-MODAL HAND/MAN SYMBOLS WITH COUNTDOWN TIMER, LED",
        ],
        "EA",
        "4",
    ),
    (
        "680.8151",
        ["ACCESSIBLE PEDESTRIAN SIGNAL (APS) WITHOUT POST"],
        "EA",
        "5",
    ),
    (
        "680.8152",
        ["ACCESSIBLE PEDESTRIAN SIGNAL (APS) W POST"],
        "EA",
        "3",
    ),
    (
        "680.81812011",
        ["FURNISH 3 SECTION TRAFFIC SIGNAL LENS, 12 INCH (NYC)"],
        "EA",
        "12",
    ),
    (
        "680.81822011",
        ["INSTALL TRAFFIC SIGNAL LENS,", "8 IN AND 12 INCH (NYC)"],
        "EA",
        "12",
    ),
    (
        "680.820500511",
        ["FURNISH AND INSTALL 20 FT SINGLE MAST ARM ASSEMBLY"],
        "EA",
        "3",
    ),
    ("680.8223", ["BREAKAWAY TRANSFORMER BASE"], "EA", "3"),
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
        ["FURNISH AND INSTALL 5 FT EXTENSION ARM FOR", "20 FT MAST ARM (NYC)"],
        "EA",
        "2",
    ),
]

ITEM_X = 1030191.0
DESC_X = 1030272.88
UNIT_X = 1030347.37
QTY_X = 1030367.63
LINE_PITCH = 2.70  # smidge above prior ~2.5; keep cell pad vs blue borders


def _set_node_lines(tn, lines: list[str]) -> None:
    tn.DeleteAllTextLines()
    for line in lines:
        tn.AddTextLine(line)
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

    # Template text element from CLEAN EXISTING component
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    tmpl_node = None
    body = []  # (y, x, el, kind, joined)
    while ee.MoveNext():
        el = ee.Current
        try:
            if el.Level.Name != "TCSH_P_T":
                continue
            if not (el.IsTextNodeElement or el.IsTextElement):
                continue
        except Exception:
            continue
        rng = el.Range
        y = (float(rng.High.Y) + float(rng.Low.Y)) / 2
        x = (float(rng.High.X) + float(rng.Low.X)) / 2
        # leftover tests above title
        if y > 239720:
            mr.RemoveElement(el)
            print("removed stray", y)
            continue
        if y > 239708 or y < 239598:
            continue
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            lines = [tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1)]
            joined = " ".join(lines).strip()
            if joined in (
                "ITEM",
                "DESCRIPTION",
                "UNIT",
                "QUANTITY",
                "ESTIMATE OF QUANTITIES",
            ):
                continue
            if tmpl_node is None and joined == "CLEAN EXISTING PULLBOX":
                tmpl_node = el
            kind = "desc"
            if abs(x - ITEM_X) < 5:
                kind = "item"
            elif abs(x - UNIT_X) < 5:
                kind = "unit"
            elif abs(x - QTY_X) < 5:
                kind = "qty"
            elif abs(x - DESC_X) < 20:
                kind = "desc"
            body.append((y, x, el, kind, joined))
        elif el.IsTextElement:
            # prior partial TextElement placements in body
            if abs(x - DESC_X) < 20:
                body.append((y, x, el, "desc_te", el.AsTextElement.Text.strip()))

    if tmpl_node is None:
        raise SystemExit("CLEAN EXISTING template not found")

    ce = tmpl_node.AsComplexElement
    sub = ce.GetSubElements()
    sub.MoveNext()
    te_tmpl = sub.Current
    rot = tmpl_node.AsTextNodeElement.Rotation

    # Group by row Y
    body.sort(key=lambda t: (-t[0], t[1]))
    row_ys: list[float] = []
    for y, *_ in body:
        if not row_ys or abs(row_ys[-1] - y) > 0.3:
            # desc_te may be offset from row center — only seed from nodes
            pass
    # Seed row centers from item column only (stable)
    for y, x, el, kind, joined in body:
        if kind != "item":
            continue
        if not row_ys or abs(row_ys[-1] - y) > 0.3:
            row_ys.append(y)
    print(f"rows={len(row_ys)} expected={len(ROWS)}")
    if len(row_ys) != len(ROWS):
        raise SystemExit(f"row count mismatch {len(row_ys)} vs {len(ROWS)}")

    placed = 0
    for i, y_row in enumerate(row_ys):
        item, desc_lines, unit, qty = ROWS[i]
        row_els = [b for b in body if abs(b[0] - y_row) < 1.8]
        # tighter for item/unit/qty; desc TextNodes centered on row
        item_el = next((b for b in row_els if b[3] == "item" and abs(b[0] - y_row) < 0.3), None)
        unit_el = next((b for b in row_els if b[3] == "unit" and abs(b[0] - y_row) < 0.3), None)
        qty_el = next((b for b in row_els if b[3] == "qty" and abs(b[0] - y_row) < 0.3), None)
        desc_nodes = [
            b
            for b in body
            if b[3] == "desc" and abs(b[0] - y_row) < 0.3
        ]
        # any TextElements near this row in desc column (including multi-line offsets)
        desc_tes = [
            b
            for b in body
            if b[3] == "desc_te" and abs(b[0] - y_row) < (LINE_PITCH * 2.5)
        ]

        if not item_el or not unit_el or not qty_el:
            print(f"SKIP row {i} missing cols")
            continue

        # Update item / unit / qty TextNodes (short — no wrap issue)
        for el_info, lines in (
            (item_el, [item]),
            (unit_el, [unit]),
            (qty_el, [qty]),
        ):
            el = el_info[2]
            tn = el.AsTextNodeElement
            cur = " ".join(tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1))
            if cur != lines[0]:
                _set_node_lines(tn, lines)
                print(f"row{i} {el_info[3]}: {cur} -> {lines[0]}")

        # Remove old description TextNodes and TextElements
        for b in desc_nodes + desc_tes:
            try:
                mr.RemoveElement(b[2])
            except Exception as e:
                print("remove fail", e, b[4][:40])

        # Place SAMPLE-style TextElement lines centered on DESC_X
        n = len(desc_lines)
        if n == 1:
            ys = [y_row]
        else:
            # center stack on row Y
            total = (n - 1) * LINE_PITCH
            top = y_row + total / 2
            ys = [top - k * LINE_PITCH for k in range(n)]

        for text, y in zip(desc_lines, ys):
            origin = app.Point3dFromXYZ(DESC_X, y, 0)
            te = app.CreateTextElement1(te_tmpl, text, origin, rot)[0]
            mr.AddElement(te)
            te.Rewrite()
            placed += 1
        print(f"row{i} {item}: {n}L {desc_lines}")

    print(f"done placed_desc_lines={placed}")


if __name__ == "__main__":
    main()
