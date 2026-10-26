"""Preflight a workflow contract against the engineer's inputs file.

    python scripts/workflow_preflight.py profile-design [inputs.yaml] [--offline]

Prints READY / WAITING_FOR_INPUT / BLOCKED / UNVERIFIED with what is missing. Exit 0 = READY, 2 = WAITING_FOR_INPUT, 1 = BLOCKED/UNVERIFIED.
Runs nothing in MicroStation. Online (default) it also reads the alignment's stationing from OpenRoads through cif_read.
"""
from __future__ import annotations

import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "mcp-server"))
os.environ.setdefault("QA_ALLOW_NO_TASK", "1")   # reads only

import workflow_contract as wc  # noqa: E402


def main(argv: list[str]) -> int:
    args = [a for a in argv if not a.startswith("--")]
    offline = "--offline" in argv
    wf = args[0] if args else "profile-design"
    path = args[1] if len(args) > 1 else str(wc.INPUTS / f"{wf}.yaml")
    if not os.path.exists(path):
        print(json.dumps({"status": wc.WAITING, "workflow": wf, "missing_inputs": [f"{path} (copy {wf}.template.yaml and fill it in)"]}, indent=2))
        return 2
    reader = None
    if not offline:
        try:
            import wztc_ops
            from bridge_client import bridge
            wztc_ops.set_bridge(bridge)
            reader = lambda cmd, a: wztc_ops.cif_read(cmd, **a).get("rows", [])  # noqa: E731
        except Exception as ex:
            print(f"(online checks unavailable: {ex})", file=sys.stderr)
    out = wc.preflight_and_arm(wf, wc.load_inputs(path), reader)
    print(json.dumps(out, indent=2))
    return {wc.READY: 0, wc.WAITING: 2}.get(out["status"], 1)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
