"""Per-callout rectangles for left-below-table callouts; remove group wrap."""
from __future__ import annotations

import re
import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

PAD = 0.60  # tight fit around each callout text block


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


def walk_collect(el, tns, tes, shapes, depth=0):
    try:
        if el.IsTextNodeElement:
            t = text_of(el)
            if t and t.strip():
                r = el.Range
                tns.append(
                    {
                        "el": el,
                        "id": el.ID.Low,
                        "text": t,
                        "lx": float(r.Low.X),
                        "ly": float(r.Low.Y),
                        "hx": float(r.High.X),
                        "hy": float(r.High.Y),
                        "cx": (float(r.Low.X) + float(r.High.X)) / 2,
                        "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                        "lvl": el.Level.Name if el.Level else "?",
                    }
                )
        elif el.IsTextElement:
            t = text_of(el)
            if t and t.strip():
                r = el.Range
                tes.append(
                    {
                        "id": el.ID.Low,
                        "text": t,
                        "lx": float(r.Low.X),
                        "ly": float(r.Low.Y),
                        "hx": float(r.High.X),
                        "hy": float(r.High.Y),
                        "cx": (float(r.Low.X) + float(r.High.X)) / 2,
                        "cy": (float(r.Low.Y) + float(r.High.Y)) / 2,
                    }
                )
        elif el.IsShapeElement:
            r = el.Range
            shapes.append(
                {
                    "el": el,
                    "id": el.ID.Low,
                    "lx": float(r.Low.X),
                    "ly": float(r.Low.Y),
                    "hx": float(r.High.X),
                    "hy": float(r.High.Y),
                    "w": float(r.High.X) - float(r.Low.X),
                    "h": float(r.High.Y) - float(r.Low.Y),
                    "lvl": el.Level.Name if el.Level else "?",
                    "col": el.Color,
                    "wt": el.LineWeight,
                }
            )
        if int(el.Type) in (2, 6, 4, 7) and depth < 5:
            try:
                sub = el.AsComplexElement.GetSubElements()
                while sub.MoveNext():
                    walk_collect(sub.Current, tns, tes, shapes, depth + 1)
            except Exception:
                pass
    except Exception:
        pass


def in_left_stack(item):
    return (
        item["hx"] <= 1029440
        and item["lx"] >= 1029280
        and 239615 <= item["cy"] <= 239780
        and len(item["text"].strip()) > 3
        and not re.fullmatch(r"\d+", item["text"].strip())
    )


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
    try:
        created.Level = app.ActiveDesignFile.Levels("TCPP_P_T")
    except Exception:
        pass
    created.Color = 3
    created.LineWeight = 2
    mr.AddElement(created)
    created.Rewrite()
    return created


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
    tns, tes, shapes = [], [], []
    while ee.MoveNext():
        walk_collect(ee.Current, tns, tes, shapes)

    # Remove prior group wrap(s): huge shapes covering the whole stack
    removed = []
    for s in shapes:
        if s["w"] > 80 and s["h"] > 80 and s["lx"] < 1029350 and s["hx"] > 1029420:
            if 239600 < s["ly"] < 239650 and s["hy"] > 239750:
                try:
                    mr.RemoveElement(mr.GetElementByID64(s["id"]))
                    removed.append(s["id"])
                except Exception as e:
                    print("remove fail", s["id"], e)
    print(f"removed group wraps: {removed}")

    # One callout = one TextNode in the left stack (body of each callout)
    callouts = [t for t in tns if in_left_stack(t)]
    callouts.sort(key=lambda x: -x["cy"])
    print(f"callout TextNodes: {len(callouts)}")
    if not callouts:
        raise SystemExit("No left callouts found")

    # Template shape (any remaining small shape, or recreate after remove)
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
        # last resort: the PB rects we placed earlier
        for sid in range(9626, 9603, -1):
            try:
                tmpl = mr.GetElementByID64(sid)
                if tmpl and tmpl.IsShapeElement:
                    break
            except Exception:
                tmpl = None
    if tmpl is None:
        raise SystemExit("No shape template")

    placed = []
    for c in callouts:
        # Prefer union of this TN + overlapping TE lines that belong to same block
        # (TE overlays often duplicate TN lines and may extend the visual box)
        lx, ly, hx, hy = c["lx"], c["ly"], c["hx"], c["hy"]
        for te in tes:
            if not in_left_stack(te):
                continue
            # TE is part of this callout if vertically overlapping / near and X-aligned
            if abs(te["cx"] - c["cx"]) > 40:
                continue
            if te["ly"] > hy + 1 or te["hy"] < ly - 1:
                continue
            # only absorb TEs whose text appears in the TN
            te_head = te["text"].strip()[:20].upper()
            if te_head and te_head not in c["text"].upper() and len(te["text"]) > 8:
                # still allow if tightly inside TN Y span
                if not (ly - 0.5 <= te["cy"] <= hy + 0.5):
                    continue
            lx = min(lx, te["lx"])
            ly = min(ly, te["ly"])
            hx = max(hx, te["hx"])
            hy = max(hy, te["hy"])

        rl = lx - PAD
        rb = ly - PAD
        rr = hx + PAD
        rt = hy + PAD
        rect = make_rect(app, mr, tmpl, rl, rb, rr, rt)
        head = c["text"].replace("\n", " | ")[:50]
        placed.append((c["id"], rect.ID.Low, head, rr - rl, rt - rb))
        print(
            f"  TN {c['id']} -> rect {rect.ID.Low} "
            f"{rr-rl:.1f}x{rt-rb:.1f} {head!r}"
        )

    app.CadInputQueue.SendKeyin("SAVE DESIGN")
    time.sleep(0.3)
    print(f"DONE placed={len(placed)} removed_group={len(removed)} SAVED")


if __name__ == "__main__":
    main()
