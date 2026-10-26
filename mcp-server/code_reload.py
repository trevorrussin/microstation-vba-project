"""Reload edited mcp-server modules inside the RUNNING MCP server, so a code change does not need a restart.

Exposed as the `reload_code` MCP tool (server.py). Needs a restart only for:
  - server.py itself (new tools, changed tool signatures or docstrings), this file
  - a module not in MODULES
Not needed for: bridge_client.py (the live Bridge instances are kept and given the new class; thread-local COM handles, the request-id counter and
BridgeError keep their identity), wztc_ops.py, workflow_contract.py, task_guard.py, retrieval.py, op_recorder.py, qa_tools.py, plan_workflow.py,
the contract YAML (read on every call), the C# add-in (impl DLL swaps itself), VBA (scripts/hot_reload.py).

Safe because every other module calls these as `import x` / `x.func()` (lazy imports), and importlib.reload updates the module object
in place. Every file is compiled BEFORE anything is reloaded: a syntax error reloads nothing. Live state is carried over.
"""
from __future__ import annotations

import importlib
import os
import sys
import time

MODULES = ["bridge_client", "workflow_contract", "task_guard", "retrieval", "op_recorder", "qa_tools", "plan_workflow", "cif_catalog", "wztc_ops"]   # dependency order
BRIDGE_KEEP = ["_tls", "_req_counter", "BridgeError", "bridge", "chat_bridge"]  # state a reload of bridge_client must not reset
KEEP = {"bridge_client": BRIDGE_KEEP, "wztc_ops": ["_PLAN_SESSION", "_LAST_PLACED_ROAD", "_LEVEL_ALIASES_CACHE", "_LEVEL_CATEGORIES_CACHE"]}     # live session state


def reload_code(bridge, extra: tuple = ()) -> dict:
    names = [m for m in list(MODULES) + list(extra) if m in sys.modules]
    errors = {}
    for n in names:
        f = getattr(sys.modules[n], "__file__", None)
        if not f:
            continue
        try:
            with open(f, encoding="utf-8") as fh:
                compile(fh.read(), f, "exec")
        except SyntaxError as ex:
            errors[n] = f"{ex.msg} (line {ex.lineno})"
    if errors:
        return {"status": "ERROR", "reloaded": [], "syntax_errors": errors, "note": "nothing was reloaded; fix the file and call reload_code again"}
    done, failed = [], {}
    for n in names:
        mod = sys.modules[n]
        kept = {k: getattr(mod, k) for k in KEEP.get(n, []) if hasattr(mod, k)}
        try:
            importlib.reload(mod)
        except Exception as ex:                       # module state may be half-updated: say so
            failed[n] = f"{type(ex).__name__}: {ex}"
            continue
        for k, v in kept.items():
            if v is not None and k == "_PLAN_SESSION":
                try:
                    v.__class__ = mod.PlanSession       # keep the live object, give it the new methods
                except Exception:
                    pass
            setattr(mod, k, v)
        if n == "bridge_client":
            for k in ("bridge", "chat_bridge"):         # the live instances (server.py and wztc_ops hold them) get the new methods
                inst = kept.get(k)
                if inst is not None:
                    try:
                        inst.__class__ = mod.Bridge
                    except Exception as ex:
                        failed[n] = f"could not re-class {k}: {ex}"
        done.append(n)
    wo = sys.modules.get("wztc_ops")
    if wo is not None and bridge is not None:
        wo.set_bridge(bridge)
    out = {"status": "OK" if not failed else "PARTIAL", "reloaded": done, "at": time.strftime("%H:%M:%S")}
    if failed:
        out["failed"] = failed
        out["note"] = "a module failed to reload and may be half-updated; restart the MCP server if tools misbehave"
    return out
