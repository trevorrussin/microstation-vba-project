"""Place a rectangle wrapping left-below-table callouts on fea_tsp."""
from __future__ import annotations

import re
import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

PAD = 2.0  # master units around callout text union


def text_of(el):
    try:
        if el.IsTextElement:
            return el.AsTextElement.Text
        if el.IsTextNodeElement:
            tn = el.AsTextNodeElement
            return "\n".join(tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1))
    except Exception:
        pass
    return None


def walk(el, out, depth=0):
    try:
        t = text_of(el)
        if t and t.strip():
            r = el.Range
            out.append(
                {
                    "id": el.ID.Low,
                    "text": t,
                    "lx": float(r.Low.X),
                    "ly": float(r.Low.Y),
                    "hx": float(r.High.X),
                    "hy": float(r.High.Y),
                    "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                }
            )
        if int(el.Type) in (2, 6, 4, 7) and depth < 5:
            try:
                sub = el.AsComplexElement.GetSubElements()
                while sub.MoveNext():
                    walk(sub.Current, out, depth + 1)
            except Exception:
                pass
    except Exception:
        pass


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    path = app.ActiveDesignFile.FullName
    print("file", path)
    if "fea_tsp" not in path.lower():
        raise SystemExit(f"Unexpected file: {path}")

    mr = app.ActiveModelReference
    print("model", mr.Name)

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    texts = []
    while ee.MoveNext():
        walk(ee.Current, texts)

    stack = [
        t
        for t in texts
        if t["hx"] <= 1029440
        and t["lx"] >= 1029280
        and 239615 <= t["cy"] <= 239780
        and len(t["text"].strip()) > 3
        and not re.fullmatch(r"\d+", t["text"].strip())
    ]
    if not stack:
        raise SystemExit("No left-below-table callouts found")

    lx = min(t["lx"] for t in stack) - PAD
    ly = min(t["ly"] for t in stack) - PAD
    hx = max(t["hx"] for t in stack) + PAD
    hy = max(t["hy"] for t in stack) + PAD
    print(
        f"wrap n={len(stack)} pad={PAD} "
        f"[{lx:.2f},{ly:.2f}]-[{hx:.2f},{hy:.2f}] "
        f"w={hx-lx:.1f} h={hy-ly:.1f}"
    )

    # Template: prefer an existing small shape on TCPP_P_T / O_Details; else any shape
    ee = mr.Scan(o)
    tmpl = None
    fallback = None
    while ee.MoveNext():
        el = ee.Current
        try:
            if not el.IsShapeElement:
                continue
            fallback = el
            lvl = el.Level.Name if el.Level else ""
            if "TCPP" in lvl or "Details" in lvl or "TCSH" in lvl:
                tmpl = el
                break
        except Exception:
            continue
    if tmpl is None:
        tmpl = fallback
    if tmpl is None:
        raise SystemExit("No shape template for CreateShapeElement1")

    pts = (
        app.Point3dFromXYZ(lx, ly, 0),
        app.Point3dFromXYZ(hx, ly, 0),
        app.Point3dFromXYZ(hx, hy, 0),
        app.Point3dFromXYZ(lx, hy, 0),
    )
    created = app.CreateShapeElement1(tmpl, pts)
    if isinstance(created, tuple):
        created = created[0]

    # Match callout annotation look: TCPP_P_T, visible weight
    try:
        lvl = app.ActiveDesignFile.Levels("TCPP_P_T")
        created.Level = lvl
    except Exception as e:
        print("level set fail", e)
    try:
        created.Color = 0  # bylevel / black-white
        created.LineWeight = 2
    except Exception as e:
        print("sym fail", e)

    mr.AddElement(created)
    created.Rewrite()

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.3)

    r = created.Range
    print(
        f"PLACED shape id={created.ID.Low} "
        f"[{float(r.Low.X):.2f},{float(r.Low.Y):.2f}]-"
        f"[{float(r.High.X):.2f},{float(r.High.Y):.2f}] "
        f"lvl={created.Level.Name if created.Level else '?'} "
        f"col={created.Color} wt={created.LineWeight} SAVED"
    )


if __name__ == "__main__":
    main()
