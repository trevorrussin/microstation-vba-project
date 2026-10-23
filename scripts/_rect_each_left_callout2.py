"""Rebuild per-callout rects using TextNode Range only (no TE merge)."""
from __future__ import annotations

import re
import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

PAD = 0.50


def text_of_tn(el):
    tn = el.AsTextNodeElement
    return "\n".join(tn.TextLine(i) for i in range(1, tn.TextLinesCount + 1))


def walk_tns(el, tns, depth=0):
    try:
        if el.IsTextNodeElement:
            t = text_of_tn(el)
            if t and t.strip():
                r = el.Range
                tns.append(
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
                    walk_tns(sub.Current, tns, depth + 1)
            except Exception:
                pass
    except Exception:
        pass


def main():
    pythoncom.CoInitialize()
    app = next(a for _, a in ms_connect._candidate_apps() if a.HasActiveDesignFile)
    if "fea_tsp" not in app.ActiveDesignFile.FullName.lower():
        raise SystemExit("wrong file")
    mr = app.ActiveModelReference

    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()

    # Remove previous per-callout rects (9694-9704) and any leftover group wrap
    removed = []
    ee = mr.Scan(o)
    to_del = []
    while ee.MoveNext():
        el = ee.Current
        try:
            if not el.IsShapeElement:
                continue
            eid = el.ID.Low
            r = el.Range
            w = float(r.High.X) - float(r.Low.X)
            h = float(r.High.Y) - float(r.Low.Y)
            lx = float(r.Low.X)
            # our batch ids or large group wrap in left column
            if 9694 <= eid <= 9704:
                to_del.append(eid)
            elif (
                w > 80
                and h > 80
                and lx < 1029350
                and float(r.High.X) > 1029420
                and 239600 < float(r.Low.Y) < 239650
            ):
                to_del.append(eid)
        except Exception:
            continue
    for eid in to_del:
        try:
            mr.RemoveElement(mr.GetElementByID64(eid))
            removed.append(eid)
        except Exception as e:
            print("del fail", eid, e)
    print(f"removed {removed}")

    ee = mr.Scan(o)
    tns = []
    while ee.MoveNext():
        walk_tns(ee.Current, tns)

    callouts = [
        t
        for t in tns
        if t["hx"] <= 1029440
        and t["lx"] >= 1029280
        and 239615 <= t["cy"] <= 239780
        and len(t["text"].strip()) > 3
        and not re.fullmatch(r"\d+", t["text"].strip())
    ]
    callouts.sort(key=lambda x: -x["cy"])
    print(f"callouts={len(callouts)}")

    # template
    ee = mr.Scan(o)
    tmpl = None
    while ee.MoveNext():
        el = ee.Current
        try:
            if el.IsShapeElement:
                tmpl = el
                break
        except Exception:
            continue
    if tmpl is None:
        for sid in (9626, 9624, 9622, 9612):
            try:
                tmpl = mr.GetElementByID64(sid)
                if tmpl.IsShapeElement:
                    break
            except Exception:
                tmpl = None
    if tmpl is None:
        raise SystemExit("no template")

    placed = []
    for c in callouts:
        lx, ly = c["lx"] - PAD, c["ly"] - PAD
        hx, hy = c["hx"] + PAD, c["hy"] + PAD
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
        created.Color = 3
        created.LineWeight = 2
        mr.AddElement(created)
        created.Rewrite()
        head = c["text"].replace("\n", " | ")[:48]
        placed.append((c["id"], created.ID.Low, hx - lx, hy - ly, head))
        print(
            f"  TN {c['id']} rect={created.ID.Low} "
            f"{hx-lx:.1f}x{hy-ly:.1f} {head!r}"
        )

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.3)
    print(f"DONE placed={len(placed)} SAVED")


if __name__ == "__main__":
    main()
