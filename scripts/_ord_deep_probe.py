"""Deeper ORD COM probe + try keyins to activate model."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect

pythoncom.CoInitialize()

for i, (desc, app) in enumerate(ms_connect._candidate_apps()):
    print("=" * 40, i)
    # list interesting members containing File/Design/Model
    names = [n for n in dir(app) if any(k in n for k in ("File", "Design", "Model", "Document", "Project"))]
    print("members sample", names[:40])
    for n in (
        "ActiveDesignFile",
        "HasActiveDesignFile",
        "OpenDesignFile",
        "OpenDesignFileForProgram",
        "CreateDesignFile",
        "ActiveModelReference",
        "DefaultModelReference",
        "ActiveWorkspace",
        "VBE",
        "CadInputQueue",
    ):
        try:
            v = getattr(app, n)
            print(n, "=", type(v).__name__, getattr(v, "Name", getattr(v, "Count", "")) )
        except Exception as e:
            print(n, "ERR", str(e)[:120])

    # Try keyins that might attach active file
    for cmd in (
        'MODEL ACTIVE "XS Stacked Views"',
        "DIALOG MESSAGECENTER",
    ):
        try:
            app.CadInputQueue.SendKeyin(cmd)
            print("sent", cmd)
            time.sleep(0.3)
        except Exception as e:
            print("keyin fail", cmd, e)
    print("HasActive after keyin", app.HasActiveDesignFile)
    try:
        print("file", app.ActiveDesignFile.FullName)
    except Exception as e:
        print("still no file", e)
