"""Fix the live agent sign label: single inch marks, 48x48, color 0."""
from __future__ import annotations

import pythoncom
import win32com.client
from win32com.client import Dispatch

pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("MicroStationDGN.Application")
mr = app.ActiveModelReference

oScan = Dispatch("MicroStationDGN.ElementScanCriteria")
oScan.ExcludeNonGraphical()
ee = mr.Scan(oScan)
fixed = 0
while ee.MoveNext():
    el = ee.Current
    try:
        is_node = el.IsTextNodeElement
        is_text = el.IsTextElement
    except Exception:
        continue
    rng = el.Range
    cx = (float(rng.High.X) + float(rng.Low.X)) / 2
    cy = (float(rng.High.Y) + float(rng.Low.Y)) / 2
    if abs(cx - 1022196) > 50 or abs(cy - 217347) > 30:
        continue
    if is_node:
        tn = el.AsTextNodeElement
        # Replace lines if API allows
        try:
            n = tn.TextLinesCount
            print("node lines", n)
            for i in range(1, n + 1):
                print("  before", i, repr(tn.TextLine(i)))
        except Exception as e:
            print("list err", e)
        # Try delete + re-place is cleaner for text nodes
        try:
            el.Color = 0
            el.Rewrite()
            print("set node color 0")
        except Exception as e:
            print("color err", e)
        fixed += 1
    elif is_text:
        te = el.AsTextElement
        print("text before", repr(te.Text), "color", el.Color)
        t = te.Text
        if '""' in t or "36" in t:
            te.Text = t.replace('""', '"').replace("36", "48")
            el.Color = 0
            el.Rewrite()
            print("text after", repr(te.Text), "color", el.Color)
            fixed += 1

print("fixed", fixed)

# Also set face/post/line color 0 near agent assembly
oScan = Dispatch("MicroStationDGN.ElementScanCriteria")
oScan.ExcludeNonGraphical()
ee = mr.Scan(oScan)
sf = None
for i in range(1, app.ActiveDesignFile.Levels.Count + 1):
    lv = app.ActiveDesignFile.Levels(i)
    if lv.Name == "SF_P":
        sf = lv
        break
while ee.MoveNext():
    el = ee.Current
    hit = False
    if el.IsCellElement:
        c = el.AsCellElement
        if c.Name == "W20-01RA" and abs(float(c.Origin.X) - 1022130.87) < 1:
            hit = True
        if c.Name == "TWZSGN_P" and abs(float(c.Origin.X) - 1022110.87) < 1 and abs(float(c.Origin.Y) - 217353.75) < 1:
            hit = True
    elif el.IsLineElement:
        try:
            s = el.AsLineElement.StartPoint
            e = el.AsLineElement.EndPoint
            if abs(float(s.X) - 1022110.87) < 0.1 and abs(float(s.Y) - 217353.75) < 0.1 and abs(float(e.X) - 1022130.87) < 0.1:
                hit = True
        except Exception:
            pass
    if hit:
        if sf is not None:
            el.Level = sf
        el.Color = 0
        el.LineWeight = 3
        el.Rewrite()
        print("symbology fixed on element")
