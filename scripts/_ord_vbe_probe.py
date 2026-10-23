"""Write a tiny VBA dump via keyin if possible; else probe both ORD apps harder."""
from __future__ import annotations

import sys
import time
from pathlib import Path

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect

OUT = Path(r"c:\repos\microstation-vba-project\scripts\_ord_active_dump.txt")
pythoncom.CoInitialize()

# Try: use CadInputQueue Python eval? No.
# Try AttachToActiveEntity style by reading Application from each and
# calling OpenDesignFileForProgram on the LOCAL path WITHOUT closing —
# FORBIDDEN by never-switch rule if it opens a second copy.
# Instead: ask VBA in-process via keyin VBA RUN if Test project exists.

for i, (desc, app) in enumerate(ms_connect._candidate_apps()):
    print("==== app", i, "Caption", app.Caption)
    print("HasActive", app.HasActiveDesignFile)
    # VBE / VBProject
    try:
        vbe = app.VBE
        print("VBE", vbe)
        print("VBProjects", vbe.VBProjects.Count)
        for j in range(1, vbe.VBProjects.Count + 1):
            p = vbe.VBProjects.Item(j)
            print("  project", j, p.Name, getattr(p, "FileName", ""))
    except Exception as e:
        print("VBE ERR", e)

    # Try keyin that writes session info — MicroStation has:
    #  expand echo $(FILE_NAME)
    #  expand echo $(MODEL_NAME)
    try:
        # Some builds support:
        app.CadInputQueue.SendKeyin('EXPAND ECHO $(DGNFILE)')
        print("sent EXPAND ECHO DGNFILE")
    except Exception as e:
        print("expand", e)

print("done")
