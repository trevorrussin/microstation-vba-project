"""Replace PB# circles with fitted rectangles — preserve text via Clone.

Does not Drop cells (Drop orphans children). For each TCSHL cell:
  1) Clone the PB# TextElement
  2) Place an axis-aligned rectangle around text Range + PAD
  3) Remove the cell (circle + original text)
  4) Add the cloned text back
"""
from __future__ import annotations

import re
import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

PAD = 0.30


def find_pb_cells(mr):
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    out = []
    while ee.MoveNext():
        el = ee.Current
        try:
            if not el.IsCellElement:
                continue
            if el.AsCellElement.Name.upper() != "TCSHL":
                continue
            ellipse = None
            text_el = None
            label = None
            sub = el.AsComplexElement.GetSubElements()
            while sub.MoveNext():
                s = sub.Current
                try:
                    if s.IsEllipseElement:
                        ellipse = s
                    elif s.IsTextElement:
                        t = s.AsTextElement.Text.strip()
                        if re.match(r"^PB\d+$", t, re.I):
                            text_el = s
                            label = t
                except Exception:
                    continue
            if ellipse is not None and text_el is not None:
                out.append(
                    {
                        "cell": el,
                        "ellipse": ellipse,
                        "text": text_el,
                        "label": label,
                        "cell_id": el.ID.Low,
                    }
                )
        except Exception:
            continue
    return out


def make_rect(app, mr, tmpl, lx, ly, hx, hy):
    pts = (
        app.Point3dFromXYZ(lx, ly, 0),
        app.Point3dFromXYZ(hx, ly, 0),
        app.Point3dFromXYZ(hx, hy, 0),
        app.Point3dFromXYZ(lx, hy, 0),
    )
    created = app.CreateShapeElement1(tmpl, pts)
    if isinstance(created, tuple):
        created = created[0]
    mr.AddElement(created)
    created.Rewrite()
    return created


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    path = app.ActiveDesignFile.FullName
    print("file", path)
    if "tsp" not in path.lower():
        raise SystemExit(f"Unexpected active file: {path}")

    mr = app.ActiveModelReference
    print("model", mr.Name)

    cells = find_pb_cells(mr)
    print(f"PB cells with circle: {len(cells)}")
    if not cells:
        raise SystemExit("No PB# + circle cells found")

    results = []
    for item in sorted(cells, key=lambda x: x["label"]):
        cell = item["cell"]
        ellipse = item["ellipse"]
        text_el = item["text"]
        label = item["label"]

        tr = text_el.Range
        lx = float(tr.Low.X) - PAD
        ly = float(tr.Low.Y) - PAD
        hx = float(tr.High.X) + PAD
        hy = float(tr.High.Y) + PAD
        tw = float(tr.High.X) - float(tr.Low.X)
        th = float(tr.High.Y) - float(tr.Low.Y)
        ew = float(ellipse.Range.High.X) - float(ellipse.Range.Low.X)

        cloned = text_el.Clone()
        if isinstance(cloned, tuple):
            cloned = cloned[0]

        rect = make_rect(app, mr, ellipse, lx, ly, hx, hy)
        mr.RemoveElement(cell)
        mr.AddElement(cloned)
        cloned.Rewrite()

        results.append(
            {
                "label": label,
                "txt_id": cloned.ID.Low,
                "rect_id": rect.ID.Low,
                "rect_wh": (
                    float(rect.Range.High.X) - float(rect.Range.Low.X),
                    float(rect.Range.High.Y) - float(rect.Range.Low.Y),
                ),
                "text_wh": (tw, th),
                "old_circle": ew,
            }
        )
        r = results[-1]
        print(
            f"  {label}: circle {ew:.2f}sq -> rect "
            f"{r['rect_wh'][0]:.2f}x{r['rect_wh'][1]:.2f} "
            f"(text {tw:.2f}x{th:.2f}) text_id={r['txt_id']} rect_id={r['rect_id']}"
        )

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.4)

    # verify
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    pb_texts = []
    while ee.MoveNext():
        el = ee.Current
        try:
            if el.IsTextElement:
                t = el.AsTextElement.Text.strip()
                if re.match(r"^PB\d+$", t, re.I):
                    pb_texts.append((el.ID.Low, t))
        except Exception:
            pass
    print(f"DONE replaced={len(results)} SAVED surviving_PB_texts={len(pb_texts)}")
    for eid, t in sorted(pb_texts, key=lambda x: x[1]):
        print(f"  {t} id={eid}")


if __name__ == "__main__":
    main()
