"""Verify wrap 9693 contains all left-below-table callouts; grow if needed."""
from __future__ import annotations

import re
import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

PAD = 2.5
WRAP_ID = 9693


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
                    "text": t.replace("\n", " | ")[:50],
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
    mr = app.ActiveModelReference

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

    wrap = mr.GetElementByID64(WRAP_ID)
    wr = wrap.Range
    wlx, wly = float(wr.Low.X), float(wr.Low.Y)
    whx, why = float(wr.High.X), float(wr.High.Y)
    print(f"wrap [{wlx:.2f},{wly:.2f}]-[{whx:.2f},{why:.2f}]")

    outside = []
    for t in stack:
        if t["lx"] < wlx or t["ly"] < wly or t["hx"] > whx or t["hy"] > why:
            outside.append(t)
            print(
                f"  OUT id={t['id']} [{t['lx']:.1f},{t['ly']:.1f}]-"
                f"[{t['hx']:.1f},{t['hy']:.1f}] {t['text']!r}"
            )

    print(f"stack={len(stack)} outside={len(outside)}")

    # Always rebuild to PAD=2.5 for a bit more air; use color 3 (red) so it's obvious
    lx = min(t["lx"] for t in stack) - PAD
    ly = min(t["ly"] for t in stack) - PAD
    hx = max(t["hx"] for t in stack) + PAD
    hy = max(t["hy"] for t in stack) + PAD

    # Need replace: remove old, place new (never ModifyVertex)
    tmpl = wrap
    mr.RemoveElement(wrap)

    pts = (
        app.Point3dFromXYZ(lx, ly, 0),
        app.Point3dFromXYZ(hx, ly, 0),
        app.Point3dFromXYZ(hx, hy, 0),
        app.Point3dFromXYZ(lx, hy, 0),
    )
    created = app.CreateShapeElement1(tmpl, pts)
    if isinstance(created, tuple):
        created = created[0]
    try:
        created.Level = app.ActiveDesignFile.Levels("TCPP_P_T")
    except Exception:
        pass
    created.Color = 3  # red — matches callout leader color family
    created.LineWeight = 3
    mr.AddElement(created)
    created.Rewrite()
    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.3)

    r = created.Range
    print(
        f"REPLACED id={created.ID.Low} "
        f"[{float(r.Low.X):.2f},{float(r.Low.Y):.2f}]-"
        f"[{float(r.High.X):.2f},{float(r.High.Y):.2f}] "
        f"col={created.Color} wt={created.LineWeight} SAVED"
    )


if __name__ == "__main__":
    main()
