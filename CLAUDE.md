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

### Evidence-grounded claims (all tasks)
Never claim more than the evidence supports. Never say "I think" about file, app, drawing or code state:
check, then state `I know. I checked: <what>, result: <result>`, or say the check was not run.
1. **Evidence, not assumption.** Claims about files, code, app state, tool results, and engineering outcomes need a real observation or valid existing evidence. Never imply a check ran when it did not.
2. **Reuse fresh evidence.** Do not repeat a check while the earlier result still applies. Any state-changing call (model switch, edit, hot-reload, view change, save) makes earlier observations of that state stale. Revalidate then, or when the stakes need more.
3. **Label the claim.** `qa_gate claim --level observed|inferred|supported|hypothesis|verified|unknown|contradicted`. Verified = passed its specified check, not universally true. `close` blocks on hypothesis, unknown, contradicted; inferred and supported are reported as NOT verified.
4. **Proportional to risk.** General knowledge: no tool call. Current app state: fresh observation. File contents: read or search. Compiles / tests pass: the real build or test result. CAD edit succeeded: independent post-op read-back. Engineering correctness: deterministic checks plus standards. QA gate steps always apply.
5. **Define success first** for state changes (expected outcome, invariants), then verify the resulting state.
6. **Report precisely:** what was checked, how, what it establishes. Lead with the answer.
7. **Stop when evidence is insufficient:** report UNCERTAIN or BLOCKED, never success.
8. **No verification loops:** do not re-inspect unchanged evidence or search unrelated resources without a reason.
Enforced in code: claim ledger, receipts, drawing fingerprint, stop hook, and `scripts/qa_claims_hook.py` (Claude Code `.claude/settings.json` Stop; Cursor `.cursor/hooks.json` stop with `--cursor`): once per turn it blocks "I checked / verified" with no tool call in the turn, "I think / probably" about state, and (Claude Code only) "N passed" or "<check id> PASS" that no tool result in the turn shows. Not enforced: whether the evidence is the right evidence; chat claims that name no test count or check id are not in the ledger.

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
9. **Release integrity** (2026-10-23, branch `claude/qa-gate-integrity`): a PASS is bound to the exact drawing. `run` fingerprints the model (every element's checked fields + a vertex hash per linear element, read ONE op per request: multi-op batches misattribute vertex rows) and STATE-01 fails if it changes mid-run; `close` recomputes it and is BLOCKED if the drawing (or active model) changed since the PASS. Edit after a PASS = re-run. The task stores the model it started in; a different active model makes CAD-CHG-01 UNCERTAIN.
10. **Independence**: the agent's own `review --ok` never passes VIS-01. VIS-01 passes on an independent vision review (it now sees existing linework, the deterministic results and cell content levels) or on the engineer's `qa_gate signoff` (interactive terminal only, HMAC key in `~/.qa_signoff_key`, bound to the drawing state; never run it or read the key). A vision issue is disputed only with `review --dispute N "why" --evidence CAD-LVL-01,...` naming checks that PASSED this run. Issues that contradict a deterministic PASS without element ids are auto-discarded.
11. **Coverage blocks**: COV-01 is blocking. `mandatory: true` requirements in `qa/applicability.yaml` must have a PASSing check (or an engineer sign-off if manual: R-VCLR, R-WZ-GEOM); a candidate marked `applies yes` that no registry row verifies becomes `GAP-n` (needs sign-off or a registry entry); a candidate dismissed as "no" that the registry ties to the changed elements is a contradiction (STD-DISP-01 UNCERTAIN). Changed elements with no recorded op block (OP-01) unless a `--rel` expectation covers them or the engineer signs off `OP-01`. `latest.json` carries `provenance` (gate commit, criteria/registry hashes, standards document sha256, model, fingerprint, sign-off).
12. **Tool boundary** (round 2): the bridge refuses drawing edits (`mcp-server/task_guard.py`, hooked in `Bridge.call_batch`) unless a `cad.drafting` task is open: start one first. Reads are always allowed. Expired (>12 h idle) and blocked (hook gave up) tasks do not count: `qa_gate reopen --task ID`. Scratch/maintenance scripts that are not drawing work set `QA_ALLOW_NO_TASK=1` (the playbook runner does). The MCP server must be restarted to load this.
13. **Evidence integrity**: COM-01 re-reads every changed element (+ a seeded sample) through MicroStation's COM model and requires range + level to match the bridge. The bridge client no longer trusts a shared result slot in a batch (VBA reuses a small slot pool and reads results after the whole batch ran): read ops are re-run alone, edit ops raise. `latest.json` has `stage_seconds`; a run still going after 240 s dumps all stacks to `<task>/stall-trace.txt`; the vision call has a 180 s timeout and a timeout is UNCERTAIN, never a clean review.
14. **Intent and geometry validators**: `start` prints the intent the gate derived from `--desc` (quantities, directions, ORD offset sign). A move/offset/rotate/extend/copy whose request names no number is UNCERTAIN (INTENT; engineer sign-off). `qa_gate expect --rel` kinds: `offset`, `length`, `angle`, `offset_curve` (constant offset from a CURVED reference, side-aware), `distance` (min gap box-to-polyline: `--params '{"min":0.5}'`), `side` (`{"target":"above|below|left|right"}`), `vclear` (`{"x":[stations],"ve":5,"min":14.5}`; the minimum comes from the engineer/HDM, never assumed). Requirements flagged `coverage: partial` in `qa/applicability.yaml` report "partial", never "verified". A sign-off keeps the original status in `signed_off_from`.
15. **Measured, not claimed**: `python scripts/qa_corpus_report.py` runs the regression corpus (`tests/qa/corpus_cases.py`, 53 cases) and prints false-pass / false-fail rates; `pytest` requires both 0. Add a case for every engineer correction. Fingerprint scope (active model only, vertex hashes for the linear elements that return vertices, cap 1500, references not covered) is stated in `latest.json` provenance.
16. **Research before work, enforced** (2026-10-23). Playbooks are not the only source: standards (`search_standards`, `get_standard_section`, `read_standard_pages`), the written tutorial sources the playbooks cite (NYSDOT profiles PDF, decks, ORD manual via `read_tutorial`), and the QA failure modes all establish the problem, the steps and the rules. Reading is recorded as a receipt (`outputs/qa/receipts.jsonl`; only non-empty reads count; search alone is not a read). `qa_gate start` lists a task's candidates (standards, playbooks, cited written sources). The bridge refuses drawing edits (`PlaybookNotRead`) until every candidate is dispositioned and each `yes` was read: `qa_gate dispose --cite ... --applies yes|no --reason ...` (`yes` is refused without the read; `no` needs only the reason). Chat answers: the `UserPromptSubmit` hook (`scripts/retrieval_prompt_hook.py`) names the matched playbooks, and the Stop hook holds the turn until they are read and a standards/tutorial search ran (2 holds, then it states the answer is not source-backed). Before presenting or running a plan call `check_plan_against_playbook(topic, plan)` and resolve REVIEW rows. A receipt proves the read, not that you followed it; recordings are searched on demand, and a read of any page of a large PDF counts. Code: `mcp-server/retrieval.py`; tests: `tests/qa/test_retrieval_gate.py`. MCP restart needed.
17. **Process discipline in code** (2026-10-24, after the "turn off the circles" run). PROC-01 reads the task's event log (task before edit, read before edit, at most 2 failed edits in a row, replan needs new reading, `blocked` needs an escalation reason). The bridge logs every edit attempt (`outputs/qa/attempts.jsonl`); after 2 failures in a row `task_guard` refuses edits (`ReplanRequired`) until `qa_gate replan --task ID --note "..."` (needs a read since the failure) or `qa_gate escalate`. `scripts/qa_pretool_hook.py` (PreToolUse, Bash|PowerShell) applies the same rules to Python/COM scripts that attach to MicroStation and call a mutating method; read-only scripts pass, `QA_ALLOW_NO_TASK=1` for scratch. Statements get evidence labels: `qa_gate claim --level observed|inferred|hypothesis|verified|unknown --text ... --evidence ...` (`verified` must cite a check that PASSED; `close` refuses open hypotheses and lists inferred claims as NOT verified). Display, level and view changes are drawing state: unwanted or missing linework starts at playbook references-levels-models-measuring S3 (Level Display, `who: engineer`), not at a COM script. The retrieval prompt hook ignores pasted material and discussion questions, needs 2 keyword hits, demands at most 2 playbooks, and never repeats a demand the Stop hook gave up on. MCP restart needed for `task_guard`/`bridge_client` changes (`reload_code` covers task_guard).
18. **Diagnostic ladders, experience memory, A/B experiments, trajectory** (2026-10-24, after the circles run; no MCP restart needed). `Data/diagnostic-ladders/*.yaml` order the checks for a known symptom (first: unwanted or missing graphics, rungs R1-R5, cheap and read-only first, GUI rungs `who: engineer`). `qa_gate start` attaches a ladder when the description matches and prints it; record each result with `qa_gate rung --id R# --status negative|positive|handed_off|na --evidence "..."`. The bridge guard refuses edits while a cheaper rung has no result (a `positive` rung ends the diagnosis) and PROC-01 rule P6 flags a scripted edit made before them. The UserPromptSubmit hook injects the ladder for a matching prompt. `Data/experience/experience.jsonl` (tracked) records what past tasks tried: status resolved|unresolved|handed_off, findings labelled observed|reported|inferred, no hypotheses stored as findings. `qa_gate escalate` and `close` write a row from the task, `qa_gate start` prints the closest rows, `qa_gate experience add|search` works by hand. An unresolved or handed-off row is not a solution. `qa/experiment.py run_ab` is the reversible A/B (snapshot, capture, change ONE thing, capture, always restore, require the state equal the start): verdicts SUPPORTED, REFUTED, SIDE_EFFECT, INCONCLUSIVE, RESTORE_FAILED. The bridge has no per-view Level Display read/write op yet, so view toggles stay engineer steps until one exists. `tests/qa/test_metamorphic.py` checks relations without hand-written answers (rotate/mirror round trips, translation and reversal invariance, the validator's verdict under translation, move-then-inverse). `python scripts/qa_trajectory_report.py` reports verified-success, escalation, read-before-edit, failed and redundant edit rates and process violations from the guards' own logs; read it with the corpus report. Add a ladder file for each new symptom class an engineer correction reveals.
19. **Modes and the capability catalog** (2026-10-24; `find_capability` MCP tool needs `/mcp` reconnect, everything else loads per call). Every task has a mode. `production` (default) changes the engineer's drawing with established capabilities only: while one is open, `scripts/qa_mode_hook.py` (PreToolUse on Edit|Write|MultiEdit|NotebookEdit|Bash|PowerShell, registered in the git-ignored `.claude/settings.json`) refuses edits to the tooling (`mcp-server/`, `Modules/`, `ord-addin/`, `qa/`, `tests/`, the key-in registry, workflow contracts, ladders, `.claude/`) and loader commands (hot_reload, ord-addin build/register). `development` builds or investigates a capability: `qa_gate start --mode development --scratch <file.dgn|none> --mode-reason "..."`. A named scratch must be the ACTIVE file at start and at close; `--scratch none` is code-only and `task_guard` refuses every drawing edit. Never start development on the engineer's real file; if the wrong file is open, ask the engineer to switch (never-switch-DGN rule). With no task open nothing is restricted. The hook is a static look at the tool call; it stops the accidental case, not a determined one. `mcp-server/capabilities.py` (`find_capability`, `python scripts/capability.py "..."`, and printed by `qa_gate start`) searches workflows, playbooks, bridge tools, key-ins and ladders and returns status, method, `production_ok`, verification and next call, or UNSUPPORTED plus the mode's instruction (production: stop and tell the engineer; development: investigate on a scratch file, register as a candidate with a test, promotion needs the engineer's approval). Methods come from the deterministic table `Data/capabilities/policy.yaml`, with per-capability overrides; no model picks the method and no success-rate data exists yet. Matching is lexical: a `LOW CONFIDENCE` match covered under 75% of the request words, so read it before relying on it. Add an override only when a capability's best method differs from its kind default. Do not add a YAML flow-list item containing `: ` unquoted (it becomes a mapping).
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

**MCP server (Claude Code tools) code: call the `reload_code` MCP tool instead of asking for a restart.** It reloads bridge_client (live Bridge instances are kept and re-classed), wztc_ops, workflow_contract,
task_guard, retrieval, op_recorder, qa_tools and plan_workflow inside the running server (syntax-checks first, keeps the plan session). Restart
only for server.py / code_reload.py changes (new tools, changed signatures): `mcp-server/proxy.py` (the process Claude Code owns, see `.mcp.json`)
restarts the server child on its own when those files change and sends `tools/list_changed`; touch `mcp-server/.proxy-restart` to force it.
Verified 2026-10-25: after a `server.py` edit the live `place_text_label` schema showed the new arguments within a minute. Check the live tool schema
(ToolSearch `select:<tool>`) before claiming a tool change is live. Only a change to `proxy.py` itself needs the engineer to reconnect `wztc-designer` (`/mcp`);
the agent cannot reconnect (`restart_chat_driver.py` is a different process). C# add-in ops (`ord-addin\build.ps1`) and
`Data/workflow-contracts/*.yaml` need neither. (`reload_code` itself needs one restart after it is first added.)

`mcp-server/*.py` changes to chat_driver.py are not picked up by hot-reload.
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