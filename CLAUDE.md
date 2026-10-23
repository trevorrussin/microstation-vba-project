# NYSDOT WZTC Designer — Instructions for Claude Code

This file is read automatically at the start of every Claude Code session.
It contains project rules, patterns, and constraints that apply to all work in this repo.

---

## Project Summary

MicroStation 2023 CONNECT VBA tool for NYSDOT Workzone Traffic Control plan design.
Language: VBA. Files: `.bas` (modules), `.frm` (user forms), `.cls` (class modules).
There is no standalone compiler — code is run inside the MicroStation VBA IDE (Alt+F11).

**8-step sequential workflow:**
WZTCDesigner → DrawWorkSpace → AlignDraw → PlacePerp → frmSignSubColors → PlaceSign → PlaceElements → PlaceCells

State persists across form unloads via 40+ public variables in `SharedState.bas`.
All forms are modeless (`vbModeless`).

---

## Rules — Always Follow

### Before Writing Any CadInputQueue Code
Read the relevant Legacy file first. These contain the exact working patterns for
MicroStation VBA command sequences that are known to work:
- `Legacy Files/LegacyPrototype.bas` — sign face / PLACE CELL ICON / ATTACH LIBRARY pattern
- `Legacy Files/LegacyElements.bas` — HATCH ICON / PLACE SHAPE CONSTRAINED pattern

Do not guess CadInputQueue sequences. Match the Legacy pattern.

### MicroStation VBA API Constraints
- `ElementScanCriteria` valid methods: `ExcludeNonGraphical`, `ExcludeAllTypes`, `IncludeType`
  — `IncludeOnlyType` does NOT exist and will cause a "variable not defined" compile error
- `SendCommand` is for tool activation: `PLACE LINE CONSTRAINED`, `HATCH ICON`, `PLACE CELL ICON`
- `SendKeyin` is for settings: `ACTIVE LEVEL`, `ACTIVE COLOR`, `ACTIVE WEIGHT`, `ATTACH LIBRARY`, `TEXTEDITOR PLAYCOMMAND`
- `ae.CenterPoint` works in MicroStation 2023 CONNECT; `TryComputeCenter()` in PerpPlacement.bas is the fallback
- Do NOT add `Attribute VB_Name` to any module — not needed in MicroStation VBA

### Element Level / Color / Weight — Do Not Mix Up
| Element type | Level | Color | Weight |
|---|---|---|---|
| Alignment lines, perp tick lines | Default | 0 (white) | 0 |
| Dimension elements | Default | 2 (yellow) | 0 |
| WZTC drawing elements (Work Space, Channelizing Devices, Barriers, etc.) | Per DrawElements.bas GetElementLevel() | 6 | 2 |

When fixing a bug or adding a feature, only change the element properties of the element being worked on.
Do NOT change the level/color/weight of other element types.

### Form Behavior
- All forms are modeless — the GetInput loop works while the form is visible
- Only hide the form for cell placement (PlaceSelectedCell) and callout placement (TEXTEDITOR PLACENOTE)
- Standard drawing commands (PLACE LINE CONSTRAINED, HATCH ICON) do NOT require hiding the form

### Never switch MicroStation DGN files
- Do **not** open, close, or switch DGN files/models (`OpenDesignFile`, `RD=`, etc.).
- If the wrong file is active, stop and ask the engineer to switch.
- Full rule: `.cursor/rules/never-switch-dgn.mdc`.

### Code Style
- Make minimal changes — do not refactor code surrounding a bug fix
- Do not add docstrings, comments, or type annotations to code that was not changed
- Do not add error handling for scenarios that cannot happen in the workflow
- Status messages go in `lblStatus.Caption` on the active form

### Communication (STE-inspired)
Applies to Claude Code replies in this repo (Cursor also loads `.cursor/rules/ste-writing.mdc`).
This is separate from `mcp-server/prompts.py` (in-MicroStation WZTC chat agent only).
- Short imperative sentences. Prefer: place, delete, refuse, ask, verify, report, fix.
- Avoid: try to, somehow, if needed, appropriately, should work, might want to.
- Lead with the answer; then evidence; then next action.
- Before claiming done/fixed: cite what you verified (test, path, log, tool result).
- On failure: name the next step. Do not invent designer inputs or sheet features.

### Implement-all (no short-win batching)
When the engineer says **implement all**, **fix all**, **do everything**,
**wire it all**, or equivalent (including “the should-fixes” / a list of gaps
you just proposed):

1. Treat the **full list** as one job. Finish every item before stopping.
2. Do **not** stop at a partial win or “phase 1 / rest later” unless they
   explicitly defer an item. Prefer one larger complete improvement over
   several thin batches.
3. **Wire end-to-end** — chat_driver tools, prompts, MCP server, hot-reload,
   driver restart, tests as required. A helper nobody calls is not done.
4. **Verify** before claiming done; cite evidence.
5. If blocked, name the block, then finish every unblocked item in the same
   effort — do not abandon the rest of the list.

Mirror of `.cursor/rules/implement-all.mdc` (Cursor alwaysApply). Same bar
in both tools.

### Engineering loop (plan → verify → learn)
For non-trivial work: understand → investigate → plan (with verification) →
implement → **run verify** → on FAIL/UNCERTAIN enter
`.cursor/rules/engineering-failure-protocol.mdc` (**BLOCKED** — diagnose,
do not advance) → on PASS continue → update durable rules after engineer
corrections. Execution loop: `.cursor/rules/agent-engineering-loop.mdc`.
Fresh-context review: `.cursor/skills/fresh-context-review/SKILL.md`.
**Evidence proportionality** (same protocol file): sample similar items
first (representative/random + edges — not only first-N); do not exhaustively
open every instance unless further checks could change the conclusion, stakes
require it, cheap automation covers all, or the engineer asks for every.
State sample-based scope when not exhaustive.

### QA gate + standards (every drafting / CAD / engineering task)
Gate of record: `python scripts/qa_gate.py` (plan: `dev-notes/qaqc-standards-plan.md`).
1. **Start** before drafting: `qa_gate start --type cad.drafting|cad.classify.curb|engineering.calc|generic --desc "..." [--extent ...] [--files ...]`.
   It prints the criteria, the known failure modes (`qa/failure_modes.yaml`) and the standards likely governing the task.
2. **Standards backcheck before placing** (not after the engineer catches it):
   `python mcp-server/standards_index.py search "..." [--agency NYSDOT|NYC-DDC|NYC-DEP]`
   (or MCP `search_standards` / `get_standard_section` / `read_standard_pages` / `follow_standard_ref`).
   Library: `Standards/` (registry `Standards/registry.json`, re-fetch `scripts/standards_fetch.py`).
   Search for **every class of element you are about to draw** (levels, weights, text style,
   text clearance vs other text **and** linework, dims, cells). Read the whole clause — do not
   invent a narrower reading. Example: HDM Ch 20 p9(h) text-vs-text **and** p9(i) place text
   off linear features (leader if needed). Never quote a criterion from memory; `qa_gate cite` it.
3. **Expectations**: `qa_gate expect ...` — measurable (counts by level/type/cell, values with limits) from the sheet/standard/request.
4. **Run → fix → re-run** until PASS: `qa_gate run` (exit 0/1/2). Look at every image it lists; record `qa_gate review --ok --checked "..."` or `--issue`.
   The Stop hook (`scripts/qa_stop_hook.py`) blocks ending a turn while an open task is not PASS.
5. **Close** only on PASS (`qa_gate close`). If evidence cannot resolve it: `qa_gate escalate --reason ...` and tell the engineer exactly what is blocked.
6. Every engineer correction → new row in `qa/failure_modes.yaml` + a known-bad fixture in `tests/qa/` (`python -m pytest tests/qa -q` must stay green).
7. **Completeness** (2026-10-23): `start` snapshots the model and lists candidate standards — `qa_gate dispose --cite ... --applies yes|no --reason ...` each one. Editing/moving/deleting EXISTING elements needs `start --may-modify` (else CAD-CHG-01 FAILs). `qa/applicability.yaml` maps element kinds → requirements; read the COV-01 list and tell the engineer which items are engineer-review/unverified — a PASS means the implemented checks passed, not full NYSDOT/DDC compliance.
8. **Operation contracts** (2026-10-23): bridge edit ops are recorded with before/after geometry and re-measured (OP-01): offset distance/side, move delta, rotation, extend, break, text edits, placements. Put the engineer's actual numbers and directions in `--desc` (a mismatch with the ops is UNCERTAIN). Edits made by direct COM scripts have no recorded op — declare `qa_gate expect --rel offset|length|angle --a ID --b ID --value N`. The gate prints "CAD standards checks" and "Requested edit verified" separately; report both.
Calculations go through `engcalc/` (the gate re-runs them); the LLM is never the calculator of record.

### Loc01 AutoCAD curb layers (HWP24QC1)
Authority: NW visual QA + `.cursor/skills/loc01-curb-classify/SKILL.md`
(also `.claude/skills/loc01-curb-classify/`).

**Agent control (all tasks, not only curb):** see
`.cursor/rules/engineering-failure-protocol.mdc` — first line: never confuse
batch progress with engineering progress. FAIL / UNCERTAIN → BLOCKED →
local|systemic|plan investigation → prove the fix before PASS. Domain geometry
rules below answer *how to classify* when the plan is working.

| Layer | Rule |
|-------|------|
| Depressed | INNER flare-to-flare per ped ramp; split verts at flare edges |
| Corner | Outside depressed; short-chord larger curvature only |
| Straight | Outside depressed; colinear including long segs (≥~2.5 ft) |

Never touch OUTER FLs `13D80A/13D25C/13D7DC/13CEC7/13D2EA/13D7BC`. Prefer hide
sources over Delete. No flare drawing unless asked. Dense 0.3 ft FL sampling
for Corner/Straight is **forbidden**. Done gate:
`python scripts/_acad_verify_loc01_curb.py` (must exit 0 before claiming done).

**Loc 05–98 batch (same rules):** ped-ramp flares may be on `PROP-D-ROAD-RAMPS`
**or** `PRO-S-CURB-LINE` / `S-CURB-DROP` inserts with ≥2 solids (Loc10). Every
sheet: if ramp/flare blocks exist and Depressed LF ≈ 0 → **VERIFY FAIL**, do
**not** save, do **not** leave Straight in drop zones.

**Engineering harness (Screenshots sandbox):** prefer
`python scripts/engineering_harness.py status|verify|restore-steel|classify`
(`--sample` / `--sheet` / `--all`). Manifest authority:
`outputs/engineering-harness-manifest.json` (`PASSED|BLOCKED|ESCALATED`).
Per-sheet gate: `scripts/_acad_verify_location_sheet.py`. No `pwworking`/`d017*`
xrefs. Batch helper still: `scripts/_acad_batch_classify_locations.py` (exits on
first fail).

---

## Tutorial playbooks (engineers' working procedures)

`Standards/Tutorials/` (26 recorded training videos, 4 decks, the 584-page ORD manual, the workflow sheet) is ingested into `Standards/derived/tutorials/` (speech + on-screen OCR with
timestamps; half the videos are silent screen recordings, read via OCR + keyframes). `Data/tutorial-playbooks/*.md` distils them into 14 playbooks: each step is `who: agent` (has a test) or
`who: engineer` (GUI dialogs the agent must not drive, with the reason), cited to file + timestamp/slide/page.

- Before drafting or editing: `list_tutorial_playbooks` -> `get_tutorial_playbook(topic)`; `check_playbook(topic)` preflights this workspace (levels, cells, styles, models, coordinate system); `search_tutorials` for anything a playbook lacks.
- Use THIS workspace's names (ny_barscale.cel, ny_Prop Title), not the recording's (nyu_*, CP_Normal).
- Prove a playbook still works: `python scripts/playbook_runner.py --all --live --report outputs/playbooks/report.md` (scratch elements, read-back verified, removed afterwards; do not switch models while it runs).
- Rebuild after adding material: `python scripts/tutorials_ingest.py` (resumable; ONE worker only - two GPU workers thrash 12 GB), `python scripts/tutorial_digest.py`, `python mcp-server/tutorial_kb.py build`. `pytest tests/qa` checks every citation resolves and every tutorial file is covered.
- Measurable rules are in the QA gate: CAD-SCALE-01 (scale bar = annotation scale), CAD-MODEL-01 (annotation in the Drawing model), OP-01 sign cross-check (ORD offsets are negative LEFT; copy_parallel is + LEFT), `qa/applicability.yaml` requirements carry `playbook: topic#step`.

---

## File Sync Protocol

**After Claude edits an existing `.bas`/`.frm`/`.cls` module on disk (standard path):**
Run `python mcp-server/hot_reload.py <path(s)>` to push the change into the open MicroStation VBA
project in place — do not ask the user to manually delete+re-import for routine code updates.
Never hot-reload while VBA is `[running]`/`[break]` (a bridge op, keyin batch, or GetInput loop
still executing) — this can wedge the project or crash MicroStation. If a call may still be in
flight, stop and ask the user to Reset (or Ctrl+Break) in the VBA IDE first.

Manual delete+Import File is still required, not hot-reload, when:
- the module/form does not exist in the IDE yet (brand-new file)
- adding or changing **UserForm Designer controls** (layout is not part of the CodeModule)
- hot-reload fails after a clean Reset (fall back to manual import for the affected files)

`mcp-server/*.py` changes (chat_driver.py, wztc_ops.py, etc.) are not picked up by hot-reload.
Restart it yourself — run `python mcp-server/restart_chat_driver.py` (it checks the process is
idle before stopping it, then relaunches in a new visible console window). Do not just tell the
user to restart it manually; this is the same "agent restarts, doesn't ask" rule as VBA hot-reload
above.

**After user edits in the MicroStation VBA IDE:**
The disk file is out of date. User must:
1. Right-click module → Export File
2. Tell Claude what changed before asking for further edits

If the user describes a manual change they made in the IDE, treat it as authoritative — do not
revert or overwrite it. Ask for clarification if the change is unclear.

---

## Sheet spec sync (after every named-sheet fix)

When you fix or learn anything about a **619-xxx** sheet build (compiler, VBA,
engineer QA), update **in the same effort**:

1. `Data/sheet-specs/<sheet>.json` — notes the compiler/agent must obey
2. `Data/sheet-specs/<sheet>.build.md` — playbook tip + Do-not row
   (`sheet.buildGuide`; loaded by `get_sheet_requirements` /
   `get_sheet_build_guide` for the in-MicroStation agent)
3. `dev-notes/agent-log.md` — cross-tool bridge

Do **not** leave tips only in chat or only in the agent-log — the next WZTC
agent turn will miss them. Mirror of `.cursor/rules/sheet-spec-sync.mdc`.

---

## Cross-Tool Work Log (Cursor + Claude Code)

This project is sometimes worked on from Cursor as well as Claude Code. Neither tool can see
the other's session history or persistent memory, so `dev-notes/agent-log.md` is a manual
bridge: Cursor is instructed (via `.cursor/rules/agent-log.mdc`) to append an entry there after
finishing non-trivial work; Claude Code should do the same.

**Check `dev-notes/agent-log.md` whenever the user mentions work happened in Cursor** (or another
tool) since your last session, or at the start of a session if it's been a while. Fold anything
load-bearing into your own persistent memory the same way you would if the user had described it
directly in chat — don't just leave it sitting in the log unprocessed.

---

## Debugging Tips for This Project

When the user reports a bug, ask for:
1. The exact error message (if any) from the VBA IDE
2. The output from the Immediate window (Debug.Print lines)
3. Which step in the 8-step workflow triggered the issue
4. Which file/sub the issue is in, if known

Do not guess the root cause without this information for CadInputQueue timing issues —
these are notoriously hard to diagnose from symptoms alone.

---

## Future Development (Planned in ROADMAP.md)
- v1.2.0: Cost estimate tool
- v1.3.0: Detour plan designer
- v1.4.0: Workzone summary tables
- v1.5.0: Workzone cross sections

For any new workflow feature (new form, new module), use plan mode before writing code.