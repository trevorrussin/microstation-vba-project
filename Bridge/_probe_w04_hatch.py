"""Probe W04 fill/priority; fix hatch to unfilled+visible stripes; report order."""
from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client
from win32com.client import Dispatch

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import bridge_client as bc
import view_capture

pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("MicroStationDGN.Application")
mr = app.ActiveModelReference

# Probe W04 for FillMode / Priority
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
    print("W04 cell", int(el.ID.Low))
    ce = c.GetSubElements()
    i = 0
    while ce.MoveNext() and i < 15:
        sub = ce.Current
        props = {}
        for name in ("FillMode", "FillColor", "Priority", "Filled", "Area"):
            try:
                props[name] = getattr(sub, name)
            except Exception:
                pass
        try:
            lvl = sub.Level.Name
            col = int(sub.Color)
        except Exception:
            lvl, col = "?", -1
        print(f"  sub[{i}] type={sub.Type} color={col} lvl={lvl} {props}")
        i += 1

# Fix workspace 56626: try remove fill via change symbology / re-hatch with bridge
# First try CHANGE via keyin or ExecHatch with fill none by recreating workspace

print("\nRecreate workspace unfilled with denser hatch...")
# delete old shape
try:
    print(bc.bridge.call("DELETE_ELEMENT", elementId="56626", ownElementOnly="N", reason="replace solid-looking hatch"))
except Exception as ex:
    print("del", ex)

# Place larger/taller workspace so 10ft hatch shows, OR we'll fix VBA to use NotFilled + spacing 2
# For now place via bridge then fix in VBA hotreload

# Probe msdFillMode constants via pythoncom if available
try:
    from win32com.client import constants
    print("constants sample", [x for x in dir(constants) if "Fill" in x or "fill" in x][:30])
except Exception as ex:
    print("no constants", ex)
