"""Wrap P1-P8 and PB1-PB8 tags in TCSHL-style circles (clone number bubbles)."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

import pythoncom
import ms_connect
from win32com.client import Dispatch

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
mr = app.ActiveModelReference
q = app.CadInputQueue

# Existing plain tags from earlier markup work
TAGS = {
    "P1": 5771,
    "P2": 5772,
    "P3": 5773,
    "P4": 5774,
    "P5": 5775,
    "P6": 5776,
    "P7": 5777,
    "P8": 5778,
    "PB1": 5779,
    "PB2": 5780,
    "PB3": 5781,
    "PB4": 5782,
    "PB5": 5783,
    "PB6": 5784,
    "PB7": 5785,
    "PB8": 5786,
}

TEMPLATE_TCSHL = 3696  # number bubble "1"


def rng_center(e):
    r = e.Range
    return (
        0.5 * (float(r.Low.X) + float(r.High.X)),
        0.5 * (float(r.Low.Y) + float(r.High.Y)),
    )


def set_cell_text(cell_el, text: str, height: float | None = None):
    sub = cell_el.AsComplexElement.GetSubElements()
    while sub.MoveNext():
        s = sub.Current
        if not s.IsTextElement:
            continue
        te = s.AsTextElement
        te.Text = text
        if height is not None:
            try:
                te.TextStyle.Height = height
                te.TextStyle.Width = height
            except Exception:
                pass
        s.Rewrite()
        return True
    return False


tmpl = mr.GetElementByID64(TEMPLATE_TCSHL)
tmpl_cell = tmpl.AsCellElement
tmpl_origin = (float(tmpl_cell.Origin.X), float(tmpl_cell.Origin.Y))
tmpl_scale = tmpl_cell.Scale
print("template", TEMPLATE_TCSHL, "origin", tmpl_origin, "scale", tmpl_scale.X, tmpl_scale.Y)

created = []
for label, tid in TAGS.items():
    tag = mr.GetElementByID64(tid)
    cx, cy = rng_center(tag)
    # Clone template cell
    clone = tmpl.Clone()
    # Move from template origin to tag center
    dx = cx - tmpl_origin[0]
    dy = cy - tmpl_origin[1]
    clone.Move(app.Point3dFromXYZ(dx, dy, 0))
    # PB labels are 3 chars — slightly larger bubble / slightly smaller text
    if label.startswith("PB"):
        try:
            # ScaleUniform about origin after move: bump ~20%
            clone.ScaleUniform(app.Point3dFromXYZ(cx, cy, 0), 1.2)
        except Exception as e:
            print("scale", label, e)
    mr.AddElement(clone)
    clone.Rewrite()
    nid = int(clone.ID64)
    # Edit text inside (re-get after add)
    cell = mr.GetElementByID64(nid)
    h = 1.7 if label.startswith("PB") else 2.0
    ok = set_cell_text(cell, label, height=h)
    # Ensure level matches TCSHL
    try:
        cell.Level = app.ActiveDesignFile.Levels("TCSH_P_T")
        cell.Rewrite()
    except Exception as e:
        print("level", label, e)
    print("placed", label, "cell", nid, "at", round(cx, 2), round(cy, 2), "text_ok", ok)
    created.append((label, nid, tid, cx, cy))

# Delete old plain TextElements
for label, nid, tid, cx, cy in created:
    try:
        mr.RemoveElement(mr.GetElementByID64(tid))
        print("deleted old text", tid, label)
    except Exception as e:
        print("delete fail", tid, e)

# Overlap gate: Range boxes among new cells + remaining callout textnodes
def ranges():
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    out = []
    while ee.MoveNext():
        e = ee.Current
        try:
            eid = int(e.ID64)
            r = e.Range
            box = (float(r.Low.X), float(r.Low.Y), float(r.High.X), float(r.High.Y))
        except Exception:
            continue
        keep = False
        label = None
        if e.Type == 2:
            try:
                if e.AsCellElement.Name == "TCSHL":
                    # only our new ones or all TCSHL
                    keep = True
                    sub = e.AsComplexElement.GetSubElements()
                    while sub.MoveNext():
                        s = sub.Current
                        if s.IsTextElement:
                            label = s.AsTextElement.Text
                            break
            except Exception:
                pass
        if e.IsTextNodeElement:
            try:
                tn = e.AsTextNodeElement
                t = tn.TextLine(1)
                if any(k in t.upper() for k in ("APS", "STREETLIGHT", "VEHICULAR", "PUSH BUTTON")):
                    keep = True
                    label = t[:30]
            except Exception:
                pass
        if keep:
            out.append((eid, label, box))
    return out


def overlaps(a, b):
    return not (a[2] <= b[0] or b[2] <= a[0] or a[3] <= b[1] or b[3] <= a[1])


items = ranges()
clashes = []
for i, a in enumerate(items):
    for b in items[i + 1 :]:
        # only flag if at least one is a P/PB label cell
        labs = {(a[1] or ""), (b[1] or "")}
        if not any(x.startswith("P") for x in labs):
            continue
        if overlaps(a[2], b[2]):
            clashes.append((a[0], a[1], b[0], b[1]))

print("overlap clashes involving P/PB cells:", len(clashes))
for c in clashes[:20]:
    print(" ", c)

q.SendKeyin("SAVE DESIGN")
time.sleep(0.3)
print("created", len(created))
print("done")
