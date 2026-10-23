"""Deep scan for PB texts / rects / leftover circles after replace."""
from __future__ import annotations

import re
import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch


def walk(el, texts, ellipses, shapes, depth=0, parent=None):
    try:
        if el.IsTextElement:
            t = el.AsTextElement.Text.strip()
            if re.search(r"PB", t, re.I):
                r = el.Range
                texts.append(
                    (
                        el.ID.Low,
                        t,
                        depth,
                        parent,
                        float(r.Low.X),
                        float(r.Low.Y),
                        float(r.High.X),
                        float(r.High.Y),
                    )
                )
        if el.IsEllipseElement:
            r = el.Range
            w = float(r.High.X) - float(r.Low.X)
            if w < 15:
                ellipses.append((el.ID.Low, w, float(r.Low.X), float(r.Low.Y), depth, parent))
        if el.IsShapeElement:
            r = el.Range
            w = float(r.High.X) - float(r.Low.X)
            h = float(r.High.Y) - float(r.Low.Y)
            if max(w, h) < 20:
                shapes.append(
                    (el.ID.Low, w, h, float(r.Low.X), float(r.Low.Y), float(r.High.X), float(r.High.Y), depth, parent)
                )
        t = int(el.Type)
        if t in (2, 6, 4, 7) or depth < 5:
            try:
                sub = el.AsComplexElement.GetSubElements()
                while sub.MoveNext():
                    walk(sub.Current, texts, ellipses, shapes, depth + 1, el.ID.Low)
            except Exception:
                pass
    except Exception:
        pass


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    mr = app.ActiveModelReference
    print("file", app.ActiveDesignFile.FullName)
    print("model", mr.Name)

    # direct ID lookup for known text/rect ids
    for eid in [6791, 6316, 6319, 6322, 6325, 6328, 6331, 6334, 9604, 9605, 9606, 9607]:
        try:
            el = mr.GetElementByID64(eid)
            kind = "?"
            txt = ""
            try:
                if el.IsTextElement:
                    kind = "text"
                    txt = el.AsTextElement.Text
                elif el.IsEllipseElement:
                    kind = "ellipse"
                elif el.IsShapeElement:
                    kind = "shape"
                elif el.IsCellElement:
                    kind = "cell " + el.AsCellElement.Name
            except Exception as e:
                kind = f"err {e}"
            print(f"  lookup {eid}: {kind} {txt!r}")
        except Exception as e:
            print(f"  lookup {eid}: FAIL {e}")

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    texts, ellipses, shapes = [], [], []
    n = 0
    while ee.MoveNext():
        n += 1
        walk(ee.Current, texts, ellipses, shapes)
    print(f"roots={n} PB-ish texts={len(texts)} small ell={len(ellipses)} small shapes={len(shapes)}")
    for row in texts:
        print("  TEXT", row)
    # shapes near old PB coords
    print("recent rects 9604-9611:")
    for s in shapes:
        if 9604 <= s[0] <= 9611:
            print(" ", s)


if __name__ == "__main__":
    main()
