"""Attach to ORD instance that has tsp_02 open; find 20ft mast arm in notes corner."""
from __future__ import annotations

import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import win32com.client
import ms_connect

pythoncom.CoInitialize()

LOCAL = r"c:\pwworking\usny\d0172869\60785527_cph_tsp_02.dgn"

print("=== GetObject(local path) ===")
try:
    obj = win32com.client.GetObject(LOCAL)
    print("ok", type(obj))
    for attr in ("FullName", "Name", "PathName"):
        try:
            print(attr, getattr(obj, attr))
        except Exception as e:
            print(attr, "ERR", e)
    try:
        app = obj.Application
        print("App.Caption", app.Caption)
        print("HasActive", app.HasActiveDesignFile)
        if app.HasActiveDesignFile:
            print("Active", app.ActiveDesignFile.FullName)
    except Exception as e:
        print("App ERR", e)
except Exception as e:
    print("FAIL", type(e).__name__, e)

print("\n=== ROT apps again ===")
for i, (desc, app) in enumerate(ms_connect._candidate_apps()):
    print(i, "Caption", repr(app.Caption), "HasActive", app.HasActiveDesignFile)
    # Caption sometimes includes filename when maximized differently
    try:
        # Force focus?
        app.Visible = True
    except Exception:
        pass

# Try matching caption containing tsp
print("\n=== retry ActiveDesignFile after Visible ===")
import time

time.sleep(0.5)
for i, (desc, app) in enumerate(ms_connect._candidate_apps()):
    print(i, "HasActive", app.HasActiveDesignFile)
    try:
        print("  ", app.ActiveDesignFile.FullName)
    except Exception as e:
        print("  ERR", e)
