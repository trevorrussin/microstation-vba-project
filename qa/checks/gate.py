"""Gate-runner-owned checks. STATE-01 is evaluated by scripts/qa_gate.py (_state_row) after every other check,
because it compares the drawing at the end of the run with the drawing the checks read; this stub keeps the
criteria registry resolvable."""
from __future__ import annotations


def state(task, ctx):
    return {"status": "NA", "observed": "evaluated by the gate runner after all other checks", "expected": "", "evidence": [], "fix": ""}
