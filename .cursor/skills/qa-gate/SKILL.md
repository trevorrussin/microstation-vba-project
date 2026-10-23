---
name: qa-gate
description: Run the QA gate for any MicroStation drafting, AutoCAD/CAD classification, or engineering calculation task - standards lookup first, measurable expectations, deterministic CAD-standards checks, geometry render + independent visual review, fix-and-rerun loop until PASS. Use at the start of every drafting/engineering task and before claiming done.
---

# QA gate workflow

Authority: `CLAUDE.md` → "QA gate + standards"; plan `dev-notes/qaqc-standards-plan.md`.

## 1. Start (before touching the drawing)
```
python scripts/qa_gate.py start --type cad.drafting --desc "<what the engineer asked>" --extent lowX,lowY,highX,highY
```
Read what it prints: criteria, known failure modes, likely governing standards.

## 2. Standards
```
python mcp-server/standards_index.py search "<requirement question>" --agency NYSDOT
python mcp-server/standards_index.py section hdm/chapt_08.pdf 8.7.5.1
python mcp-server/standards_index.py toc hdm/chapt_16.pdf
```
Read the whole section, follow `refs`, then cite:
`python scripts/qa_gate.py cite --doc hdm/chapt_08.pdf --page 95 --quote "<exact words>"`

## 3. Expectations (measurable)
`python scripts/qa_gate.py expect --id E1 --desc "drums on merging taper" --where level=TWZCD --count ">=12"`

## 4. Do the work, then run → fix → re-run
```
python scripts/qa_gate.py run
```
- Open every image it lists (Read the PNG). Compare with the expectations.
- Record: `qa_gate review --ok --checked "<what you verified>"` or `--issue "<problem>"`.
- Vision reviewer issues you believe are false: `qa_gate review --dispute N "<evidence>"` (shown to engineer).
- Fix every XX / ?? row, re-run. Do not stop on FAIL/UNCERTAIN (the Stop hook will send you back).

## 5. Finish
- PASS → `qa_gate close`, report the record path + key evidence lines.
- Cannot pass from evidence → `qa_gate escalate --reason "<exact blocker>"`, tell the engineer.

## Calculations
Use `engcalc/` functions; record each with `qa_gate calc --id Q1 --fn hydraulics.manning_full_pipe_q --inputs '{...}' --result X --units cfs --criterion "<doc pN>"`. The gate re-runs them.

## After an engineer correction
Add a row to `qa/failure_modes.yaml` and a known-bad fixture in `tests/qa/`; `python -m pytest tests/qa -q` must pass.
