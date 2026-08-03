"""Try fixing W04 by raising SFB_P priority; recreate workspace unfilled."""
from __future__ import annotations

import sys
from pathlib import Path

import pythoncom
import win32com.client
from win32com.client import constants, gencache

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import bridge_client as bc
import view_capture

pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("MicroStationDGN.Application")

# Discover fill mode enum values from type lib if possible
try:
    # Common Bentley: msdFillModeNone=0, Filled=1, Outlined=2
    for name in ("msdFillModeNone", "msdFillModeNotFilled", "msdFillModeFilled", "msdFillModeOutlined"):
        try:
            print(name, getattr(win32com.client.constants, name))
        except Exception:
            pass
except Exception as ex:
    print(ex)

# Fix W04: set high Priority on SFB_P subs, low on SF_P legend duplicates
mr = app.ActiveModelReference
from win32com.client import Dispatch
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
    print("fixing W04", int(el.ID.Low))
    # Drop to components? Try Priority on cell first
    try:
        el.Priority = 0
    except Exception as ex:
        print(" cell priority", ex)
    ce = c.GetSubElements()
    i = 0
    while ce.MoveNext():
        sub = ce.Current
        try:
            lvl = sub.Level.Name
        except Exception:
            lvl = "?"
        try:
            if lvl == "SFB_P":
                sub.Priority = 500
                sub.Rewrite()
                print(f"  sub[{i}] SFB_P priority->500")
            elif lvl == "SF_P" and i >= 9:
                # yellow legend duplicates — try hide via display priority low + transparency
                try:
                    sub.Priority = -500
                    sub.Rewrite()
                    print(f"  sub[{i}] SF_P legend priority->-500")
                except Exception as ex:
                    print(f"  sub[{i}] fail", ex)
        except Exception as ex:
            print(f"  sub[{i}] err", ex)
        i += 1
    try:
        el.Rewrite()
    except Exception:
        pass

# Recreate workspace via updated VBA after we patch it — first patch then call
print("done probe")
