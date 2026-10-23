"""CALC-01: engineering calculations are reproducible and traceable.

Task card `calcs`: [{id, fn: "hydraulics.manning_full_pipe_q", inputs: {...}, result: 12.3,
                     units: "cfs", criterion: "hdm/chapt_08.pdf p94"}]
The gate re-runs `engcalc.<fn>(**inputs)` and compares to `result` (0.5% tolerance).
Missing inputs, units, or criterion -> FAIL. The LLM never is the calculator of record.
"""
from __future__ import annotations

import importlib
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, ROOT)


def verify(task, ctx):
    calcs = task.get("calcs") or []
    if not calcs:
        return {"status": "FAIL", "observed": "no calcs recorded", "expected": ">=1 calc row",
                "evidence": [], "fix": "Record each calc: qa_gate calc --task ID --fn hydraulics.manning_full_pipe_q --inputs '{...}' --result X --units cfs --criterion 'doc pN'"}
    bad, ok = [], []
    for c in calcs:
        cid = c.get("id", "?")
        if not c.get("units") or not c.get("criterion"):
            bad.append(f"{cid}: missing units or criterion")
            continue
        mod, _, fn = c.get("fn", "").rpartition(".")
        try:
            f = getattr(importlib.import_module(f"engcalc.{mod}"), fn)
            val = f(**(c.get("inputs") or {}))
            got = float(val["value"] if isinstance(val, dict) else val)
            want = float(c.get("result"))
            if abs(got - want) <= max(abs(want) * 0.005, 1e-9):
                ok.append(f"{cid}: {got:.4g} {c['units']} reproduced")
            else:
                bad.append(f"{cid}: recorded {want:.4g} but {c['fn']} gives {got:.4g}")
            ctx.setdefault("results", {})[cid] = got
        except Exception as ex:
            bad.append(f"{cid}: cannot reproduce ({ex})")
    if bad:
        return {"status": "FAIL", "observed": f"{len(bad)} calc problems", "expected": "all reproducible",
                "evidence": bad + ok, "fix": "Correct inputs/results or add the missing units/criterion."}
    return {"status": "PASS", "observed": f"{len(ok)} calcs reproduced", "expected": "", "evidence": ok, "fix": ""}
