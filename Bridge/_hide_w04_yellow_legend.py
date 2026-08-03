"""Hide W04 SF_P legend duplicates that obscure black SFB_P merge symbol."""
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

while ee.MoveNext():
    el = ee.Current
    if not el.IsCellElement:
        continue
    c = el.AsCellElement
    if c.Name != "W04-02R":
        continue
    if abs(float(c.Origin.Y) - 216917.4) > 2:
        continue
    print("cell", int(el.ID.Low))
    # enumerate via cell's own enumerator
    c.ResetElementEnumeration()
    i = 0
    while c.MoveToNextElement():
        try:
            # Current component
            # Use GetSubElements instead for identity
            pass
        except Exception:
            pass
        i += 1

    ce = c.GetSubElements()
    i = 0
    hidden = 0
    while ce.MoveNext():
        sub = ce.Current
        try:
            lvl = sub.Level.Name
            col = int(sub.Color)
        except Exception:
            i += 1
            continue
        # Yellow SF_P components after the diamond fill/border obscure black legend
        if lvl == "SF_P" and col == 4 and i >= 9:
            try:
                sub.IsHidden = True
                sub.Rewrite()
                hidden += 1
                print(f"  hid sub[{i}]")
            except Exception as ex:
                try:
                    sub.DisplayPriority = -1000
                    sub.Rewrite()
                    print(f"  priority sub[{i}]", ex)
                except Exception as ex2:
                    print(f"  fail sub[{i}]", ex, ex2)
        elif lvl == "SFB_P":
            try:
                sub.DisplayPriority = 1000
                sub.Rewrite()
                print(f"  raise SFB_P sub[{i}]")
            except Exception as ex:
                print(f"  SFB_P fail {i}", ex)
        i += 1
    try:
        el.Rewrite()
    except Exception as ex:
        print("cell rewrite", ex)
    print("hidden", hidden)
