"""Force MicroStation Move tool Make Copy OFF and verify GG lock."""
from __future__ import annotations

import sys
import time

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")

import pythoncom
import ms_connect

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
q = app.CadInputQueue
as_ = app.ActiveSettings

print("file", app.ActiveDesignFile.FullName)
print("model", app.ActiveModelReference.Name)

# Ensure GG lock on for move-as-unit
as_.GraphicGroupLockEnabled = True
print("GG lock", as_.GraphicGroupLockEnabled)

# Dump ActiveSettings copy/move/lock-ish props
for n in sorted(dir(as_)):
    if any(k in n.lower() for k in ("opy", "move", "lock", "fence")):
        try:
            print(f"AS.{n} =", getattr(as_, n))
        except Exception as e:
            print(f"AS.{n} ERR {e}")

# Activate Move, then try every plausible TOOLSETTINGS name
q.SendKeyin("NULL")
time.sleep(0.2)
q.SendKeyin("MOVE ICON")
time.sleep(0.3)

names = [
    "MAKECOPY",
    "MAKECOPYTOGGLE",
    "COPY",
    "COPYTOGGLE",
    "COPIES",
    "USECOPY",
    "COPYMODE",
    "COPYELEMENT",
    "MAKECOPIES",
    "CREATECOPY",
    "KEEPORIGINAL",
    "KEEPORIGINALS",
    "DUPLICATE",
    "MOVECOPY",
    "MOVECOPYTOGGLE",
    "COPYON",
    "ONCOPY",
]
for name in names:
    for val in ("0", "FALSE", "OFF"):
        cmd = f"SET ITEM TOOLSETTINGS {name}={val}"
        try:
            q.SendKeyin(cmd)
            print("sent", cmd)
        except Exception as e:
            print("fail", cmd, e)

# Also chained form from Bentley docs
for cmd in [
    "MOVE ICON;SET ITEM TOOLSETTINGS MAKECOPY=0",
    "MOVE ELEMENT;SET ITEM TOOLSETTINGS MAKECOPY=0",
    "MOVE;SET ITEM TOOLSETTINGS MAKECOPY=0",
]:
    try:
        q.SendKeyin(cmd)
        print("chained", cmd)
    except Exception as e:
        print("chained fail", cmd, e)

q.SendReset()
print("done")
