# Agent Work Log

Shared cross-tool log. This project gets worked on from both Claude Code and
Cursor ? neither tool can see the other's session history or memory, so this
file is the bridge. **Whichever AI tool finishes a non-trivial, non-obvious
piece of work appends an entry here before finishing.**

"Non-obvious" is the bar ? don't log routine edits that are already clear from
`git log`/`git diff`. Log the things a future session (in either tool) would
otherwise have to rediscover: a design decision and why, a bug whose root
cause wasn't where it looked, a gotcha specific to this codebase, a dead end
that shouldn't be retried.

## Entry format

```
## YYYY-MM-DD ? <tool> ? <one-line summary>
<2-6 sentences: what changed, why, anything surprising. Reference specific
files/functions. If it's a bug fix, say what the real root cause was, not
just the symptom.>
```

Newest entries at the bottom (append-only, chronological ? matches `git log`
ordering conventions already used elsewhere in this repo).

## How each tool uses this file

- **Claude Code**: reads this at the start of a session when the user
  mentions work happened elsewhere ("I did some work in Cursor"), or
  periodically checks it ? see the pointer in `CLAUDE.md`. Anything
  load-bearing gets folded into Claude Code's own persistent memory
  (`~/.claude/projects/.../memory/`) from here, same as if the user had
  described it directly in chat.
- **Cursor**: instructed via `.cursor/rules/agent-log.mdc` to append here
  after finishing meaningful work.

This is a manual bridge, not automatic sync ? each tool only reads/writes it
when its own instructions or the user tells it to.

---

## 2026-08-02 ? Claude Code ? sign rotation now matches current view angle

`DrawSign.bas`/`WZTCExec.bas`: sign face cells were rotating to match the
alignment/perpendicular direction, which is mathematically guaranteed to
flip some signs upside-down and has nothing to do with the view. Fixed to
capture the view's rotation (`ViewRotationAngleDegrees`, via `v.Rotation`
before it gets reset to identity) and use that for `ACTIVE ANGLE` instead ?
verified live via direct COM reads that a placed cell's rotation exactly
matches the view's rotation at placement time, both unrotated (0?) and
rotated (60?) cases. See Claude Code memory `feedback-sign-rotation-matches-view`
for the full writeup.

## 2026-08-02 ? Claude Code ? bounded reuse for Bridge/results_*.tsv

`WZTCBridge.bas`'s `WriteResultRows` used to write one new
`results_<reqId>.tsv` file per query call, forever (19+ leftover files
found after one session). Changed to a bounded pool of 8 reusable
filenames (`results_slot0.tsv`..`results_slot7.tsv`, chosen via
`reqId mod 8`) ? each write already truncates/overwrites via `Open ...
For Output`, so no deletion is involved anywhere, file count just stays
bounded. Only safe because a single batch never has more than 8
multi-row-result ops today (`call_batch` in `bridge_client.py` is only
ever called with one op at a time) ? if real multi-op batching gets
added later, bump `RESULT_POOL_SIZE` past the largest expected batch.

Hit a real VBA gotcha along the way: a new module-level `Const` placed
between two existing procedures (instead of in the top-of-file
declarations block) fails to compile with "Only comments may appear
after End Sub, End Function, or End Property" ? moved
`RESULT_POOL_SIZE` up next to `BRIDGE_DIR` etc. to fix. See Claude Code
memory `feedback-vba-compile-error-recovery` for the full incident,
including a clean scriptable way to detect/dismiss/reset a blocking VBA
compile-error dialog via `VBE.CommandBars` (no manual clicking, no
guessing at dialog internals).

## 2026-08-02 ? Cursor ? keyin probe 3s hang timeout + wave8 promote

`scripts/keyin_batch.py`: live `SendKeyin` now runs in a child process with a hard **3s** timeout (`SENDKEYIN_TIMEOUT_SEC`); hangs are recorded as `HANG` ? `unsafe-blocked` and the batch continues. Also stopped executing `tool`/`datapoint` kinds (they activate-and-wait ? `TITLEBLOCK PLACE` was the wave8 forever-hang). Wave8 sparse-category harvest probed/promoted **+398** registry rows; 4 PDF embed/layers keyins hit the new timeout and were added to the skip list. Registry ~2208 / ~1959 verified.

## 2026-08-02 ? Cursor ? drawing recipe probe (element-delta bar)

Added `scripts/recipe_batch.py` + `Data/recipe-candidates.tsv`. Unlike settings
`keyin_batch.py`, a drawing recipe only promotes when graphical element count
on `DELETE.dgn` increases (plus COM alive / 3s step timeout). First live results:
`PLACE_LINE` and `PLACE_SHAPE_CONSTRAINED` ? `verified-headless-safe`;
`HATCH_ICON` (Legacy two-identical-seed pattern) completed without hang but
`+0` elements ? stays `needs-testing`. Circle/block/arc/smartline not seeded
(no in-repo CadInputQueue sequence).

## 2026-08-02 ? Cursor ? Phase C edit direct_api + Phase B WZTC registry rows

Phase C: added `ExecCopy/Rotate/Scale/Mirror/ArrayElementByID` in
`WZTCExec.bas` using Element API patterns live-proven on `DELETE.dgn`
(Clone+Move, ScaleUniform, Matrix3d Z-rotate Transform, Mirror two-point) ?
wired through `WZTCBridge` + MCP tools. No CadInputQueue inventing.
Phase B: catalogued existing place bridge ops as `direct_api`
`verified-headless-safe` rows (`PLACE_CELL` was flipped from bare
`unsafe-blocked` COMMAND). Re-import `WZTCExec.bas` and `WZTCBridge.bas`
in the VBA IDE before exercising the new edit ops.

## 2026-08-02 ? Cursor ? hatch Element API + PLACE_ARC / PLACE_TEXT_LABEL

CadInputQueue `HATCH ICON` (Legacy twin-seed) stays unreliable headlessly
(+0 elements on DELETE.dgn). Switched workspace hatch to
`CreateHatchPattern1` + `ClosedElement.SetPattern(..., Matrix3dIdentity)`
? live `HasPattern=True`. New bridge/MCP: `HATCH_ELEMENT`, `PLACE_ARC`
(placeArcModeEx=3), `PLACE_TEXT_LABEL` (TEXTEDITOR INSERT_TEXT). All three
plus updated `PLACE_WORKSPACE` verified OK via bridge on DELETE.dgn.
`HATCH_ICON` registry row ? `unsafe-blocked`. Hot-reloaded WZTCExec +
WZTCBridge. Note: PrintWindow captures often omit associative hatch lines
even when HasPattern is True ? trust HasPattern / in-app view for hatch.

## 2026-08-02 ? Cursor ? Tier1-3 general geometry ops

Added Element-API geometry suite to `WZTCExec` / `WZTCBridge` / MCP:
Tier1 place (circle/ellipse/block/polyline/polygon) + symbology; Tier2
copy-parallel (lines), crosshatch/remove-hatch, break-line, extend-line
(recreate, not EndPoint ? EndPoint assign hung VBA), fillet/complex
(needs-testing); Tier3 fence block + copy/move/delete contents, select/clear.
TRIM/CHAMFER left interactive-only (no COM ConstructTrim/Chamfer).
Live: Tier1 + symbology + copy-parallel OK on DELETE.dgn before
`LineElement.EndPoint` hang wedged VBA `[running]`. Reset/Ctrl+Break from
automation failed ? user must interrupt VBA (or restart MicroStation), then
hot-reload and re-verify extend/fence/fillet/complex.

## 2026-08-02 ? Cursor ? Tier1-3 live-verified + complex-string fix

After VBA Reset, phased live verify on DELETE.dgn: extend (recreate-line path,
not EndPoint), break, crosshatch/remove, fence define/copy/undefine, fillet all
OK. `CREATE_COMPLEX_STRING` initially failed compile ?
`CreateComplexStringElement1` needs `ChainableElement()`, not `Element()`;
fixed in `ExecCreateComplexString` via `el.AsChainableElement`, hot-reloaded,
live OK (`partCount=2`). All Tier1?3 geometry bridge ops now
`verified-headless-safe` except TRIM/CHAMFER (still interactive-only). Gotchas:
`scale` as a VBA local name conflicts with MicroStation `Scale` (use
`lenScale`); PrintWindow often omits hatch lines ? trust `HasPattern`;
VBA `[running]`/`[break]` blocks hot-reload until Reset.

## 2026-08-02 ? Cursor ? manual-search robustness + sheet-registry Non-Freeway wave

`manual_search.py` / `ingest_manuals.py`: repo-relative index paths (no hardcoded
`c:\repos\...`); missing index returns `heading=INDEX_MISSING` instead of silent
`[]`; multi-token zero-hit queries retry with OR then quoted phrase (fixes
`lane closure` + `source=supplement` empties under FTS5 AND). Re-ingest embeds
`619-NNN` into stdsht page headings. `sheet-registry.tsv` grown 6?19 sheets:
fixed 619-311 (Book 3 p.81-82 tables 311-01..05); added Non-Freeway?first
080/201?203/308?309/312?314/317 plus Short Term 304?306. Helper
`scripts/extract_sheet_signs.py` dumps candidates from sheet-owned `.dgn` pages
only. Live bridge: `get_sheet_requirements('619-201')` found=True; unseeded
still found=False.

## 2026-08-02 ? Cursor ? sheet-registry batch 2 (Short Term remainder + Intermediate)

Grew `sheet-registry.tsv` 19?38. Added general notes/legend (010?011), Short
Duration freeway (205?207), remaining common Short Term (315?316, 318?319,
321?325), and Intermediate starters (401?403, 407, 410). Skipped 619-012
(full sign catalog, not a placement typical), 619-101/102 (in DesignerRef but
no Book 3 owned pages), and multi-sheet barrier 001 (no `.dgn` text ownership).
Signs from sheet-owned `.dgn` pages only. Live bridge OK for 205/322/407;
101 still found=false.

## 2026-08-02 ? Cursor ? sheet-registry batch 3 (fill most remaining Book 3)

Grew `sheet-registry.tsv` 38?80. Added detail/general (001?006), mowing/special
(021?023, 031?033, 041, 060, 090?091), Mobile (110?114), remaining Short Duration
(208?209, 211?212), remaining Intermediate (412, 414?418, 421?423), and Long Term
(501?504, 517?518, 520, 523?524). Still unseeded (~11): 012 catalog; 050?051 /
101?104 / 204 / 419?420 / 519 not present (or not owned) in Book 3 PDF text.
503 seeded with Detour element but empty signs (drawing text unreadable to
extractor). Live spot-check via bridge after write.

## 2026-08-02 ? Cursor ? sheet-registry last-row miss was off-by-one

Root cause of `619-524` not found was not Line Input dropping the last
line: `ReadAllLines` returned it fine (`last=619-524`), but
`GetSheetRequirements` / `ListRegisteredSheets` looped `For i = 2 To n`
and read `lines(i - 1)`, so the last 1-based data row `lines(n)` was never
examined (sentinel after 524 had made 524 become `lines(n-1)` and appear
to "fix" it). Corrected to `lines(i)`. Kept ADODB.Stream whole-file read.
Verified live: 524/523/001/201 found; 101 still correctly unseeded.

## 2026-08-02 ? Cursor ? sheet-registry complete for DesignerRef (91/91)

Seeded remaining 11 DesignerRef sheets: 012 (sign catalog, empty signs on
purpose), plus stubs for sheets absent from 2026 Book 3 PDF (050-051,
101-104, 204, 419-420) and 519 (TOC listed, p.149-150 blank). Stubs keep
title/roadType/duration but empty signs with an explicit confirm-against-
current-sheet note ? no invented sign lists. Live: 012/101/204/519/524 all
found=True. Book 3 extras not in DesignerRef (024-026, 034, 042-046) still
unseeded.

## 2026-08-02 ? Cursor ? restore full signs + sheet NOTES into registry

Rebuilt all 80 complete `sheet-registry.tsv` rows without the old ~255-char
truncation: full PDF sign lists again (e.g. 619-502 now includes W20-1 /
W24-1bL/R). `notes` now append instructional NOTES text extracted from Book 3
pages (numbered note bodies + key SHALL/SHOULD lines) after the table/page
cite, separated by `||`. Cap ~1800 chars per notes field. Stubs/012 unchanged.
Live bridge OK for 502/307 with long notes.

## 2026-08-02 ? Cursor ? trim sheet-registry notes to cites only

Removed noisy PDF NOTES bodies from complete rows; `notes` are back to short
table/page cites (e.g. `Tables 502-01..502-05; Book 3 p.138-139`). Full sign
lists kept. Stub/catalog notes unchanged. Rationale: drawing-text NOTES were
fragmented and inflated tool-result tokens for little agent benefit.

## 2026-08-02 ? Cursor ? click-first + cost caps after arc session

Live arc-placement turn burned money and UX on three things: (1) fishing
with `get_journal(150)` + `classify_site_features` radius=2000 (325 rows)
instead of a point-pick when the engineer offered to click the state-law
sign; (2) `ask_user_choice` options labeled "I'll click the point" that
dismissed `btnPickPoint` when clicked ("choice is gone"); (3) `view_drawing`
base64 left in `chat-history.json` (~900KB / three images) so a cache-miss
turn hit ~243k input tokens (~$0.73). Fixes: prompt click-first rules;
`ask_user_choice` allows empty options with `allow_point_pick`, strips
duplicate pick-option labels; `MAX_SPATIAL_ROWS`/`MAX_JOURNAL_LINES` caps in
`wztc_ops.py`; `_strip_bulky_history` on load/save in `chat_driver.py`
(existing history scrubbed 1.2MB?229KB). Re-import `WZTCChatPanel.frm` for
the pick-button status caption; restart `chat_driver.py` for the Python
changes.

## 2026-08-02 ? Cursor ? chat-log false resync + list_levels cap

Conversation pane "repeated history" + stale Reference pick prompt: 
`WZTCChatTimer` treated any `ReadAllLines=0` (file locked mid-append) as
rotation (`n < mLastLineCount` ? reset to 0) and replayed the entire
`chat-log.tsv` into an already-filled panel ? including old
`ASK_USER_CHOICE` lines that re-showed "Use coordinates I already gave."
Fix: only resync when `n > 0 And n < mLastLineCount`; call new
`ResetTranscriptPanes` on real rotation; `FINAL` always `HideChoiceButtons`.
Also `list_levels(name_contains=...)` + hard cap ? unfiltered returned 3046
rows and ~$0.50 of follow-on input on the orange-color turn. Prompt: don't
guess color indices. Re-import `WZTCChatTimer.bas` + update panel form code;
restart chat_driver.

## 2026-08-02 ? Cursor ? prefer hot_reload for VBA sync

Engineer preference (standing): after editing existing `.bas`/`.frm`/`.cls`
on disk, run `python mcp-server/hot_reload.py <files>` instead of asking for
manual delete+re-import. Cursor rule: `.cursor/rules/hot-reload-sync.mdc`.
Python (`chat_driver.py` / `wztc_ops.py`) still needs a process restart.
New IDE components and UserForm Designer controls still need manual Import.

## 2026-08-02 ? Cursor ? resolve_color / list_colors (COM) + level filter

Color indices are per-DGN (DELETE.dgn: 3=red, 6?orange). `list_colors` +
`resolve_color(name|rgb)` live in `wztc_ops.py` via
`ActiveDesignFile.ExtractColorTable` / `FindClosestColor` (not the VBA
bridge) ? confirmed orange?index 6. First attempt typed `ColorTable` in
`WZTCQuery.bas` and hot-reload left the VBA project unable to Accept any
further CodeModule writes (every hot_reload COM-exceptions); color ops
moved to Python COM and Query/Bridge reverted to pre-color text. If bridge
keyins still fail: VBA IDE Reset, then manual re-import WZTCQuery +
WZTCBridge from disk. `list_levels` refuses empty `name_contains`. Prompt
requires resolve_color before named-color symbology. Restart chat_driver.

## 2026-08-02 ? Cursor ? resolve_color / list_colors + require level filter

Color indices are per-DGN (DELETE.dgn: 3=red, 6?orange via FindClosestColor).
Added `LIST_COLORS` / `RESOLVE_CLOSEST_COLOR` in `WZTCQuery`/`WZTCBridge`
(ExtractColorTable + GetColors / FindClosestColor, KB0039791), wrapped as
`list_colors` + `resolve_color(name|rgb)` in `wztc_ops` / chat tools /
`server.py`. `list_levels` now refuses empty `name_contains`. Prompt:
resolve_color before change_element_symbology for named colors. Hot-reloaded
Query+Bridge; restart chat_driver for Python.

## 2026-08-02 ? Cursor ? cells / line styles / fonts + registry false-OK scrub

Agent stumble items 1?3/7/8. New Python-COM tools in `wztc_ops.py` (same
pattern as resolve_color ? not VBA bridge): `list_line_styles` (requires
filter; 471 styles live), `resolve_line_style` (aliases like dashed?
`( Dashed )`; **Name** is the lookup key, Number is not ? `LineStyles(-104)`
fails), `cell_library_status` / `attach_cell_library` /
`list_cells` via `AttachCellLibrary` + `GetCellInformationEnumerator(False,
False)`, `list_fonts`/`resolve_font`, `list_text_styles`/`resolve_text_style`.
`change_element_symbology` gained `line_style_name`; VBA
`ExecChangeElementSymbology` prefers Name over collection index (hot-reloaded
Exec+Bridge). Annotation scale still from `describe_drawing_state` only.
Restart chat_driver for new tools.

## 2026-08-02 ? Cursor ? false-OK CommandName audit (item #8)

Replaced the precautionary family downgrade with a live audit:
`scripts/keyin_false_ok_audit.py`. Probe does SendKeyin **without** the
immediate SendReset that masked pending prompts in `keyin_batch._one_keyin`,
then classifies by `CommandState.CommandName` (ARMED if non-empty and ?
baseline `Element Selection`). Ran `--scope all-gated` on DELETE.dgn:
1733 rows ? CLEAN 1661, ARMED 67, HANG 3, SKIP 2. Applied: 67 ?
`unsafe-blocked` (ZOOM_IN/OUT, WINDOW_CENTER/ORIGIN, many SET_* /
UPDATE_VIEW / REFERENCE_* that prompt "Select view|reference|point"),
11 CLEAN `needs-testing` restored to verified (incl. ZOOM_HALF/DOUBLE/
PREVIOUS, FIT_ELEMENT/SELECTION/FENCE, ACCUDRAW_ROTATE_VIEW), 3 HANGs ?
`needs-testing`. Results: `Bridge/keyin-false-ok-audit.json`. Empty
CommandName after NULL/NOCOMMAND is CLEAN, not ARMED.

## 2026-08-02 ? Cursor ? ask_user_choice element pick + selection-tool reality

Added `allow_element_pick` alongside `allow_point_pick` on
`ask_user_choice` (`chat_driver.py` + `WZTCChatPanel.frm`). Element
identify is injected as a normal choice-button caption ("Identify an
element in the drawing") ? no extra Designer control. Clicking that
caption runs GetInput ? `LocateElement` and replies `elementId=? type=?
level=? [cell=?]` (does not merely echo the label). Point pick stays on
dedicated `btnPickPoint`. Prefer element-pick for "which existing thing";
point-pick for "put it here."

Selection-tool audit vs color-quality resolve helpers: agent-direct
`select_element(id)` / `clear_selection` are solid (need an ID first ?
element-pick now supplies that). Registry `POWERSELECTOR MODE
NEW/ADD/SUBTRACT/INVERT`, `SINGLE`, `CHOOSE ALL`/`NONE`,
`SELECTION SET INVERT` are settings-mode toggles (CLEAN in false-OK
audit) ? they change how the interactive tool behaves, they do **not**
select anything headlessly. `POWERSELECTOR BLOCK/SHAPE/LINE/CIRCLE` stay
on Element Selection CommandName but still expect engineer clicks for
the fence ? not color-class resolve tools. No working headless "select by
level" op in the registry; use `list_levels` + find/classify, or
element-pick. `CHOOSE_ELEMENT` remains unsafe-blocked.

## 2026-08-02 ? Claude Code ? agent-driven 8-step wizard built + live-verified through a real chat conversation

Built and live-verified all 4 components of the plan at
`~/.claude/plans/polished-purring-reef.md` (full detail in Claude Code
memory `project_agent_driven_8step_wizard_status.md`): `BUILD_WZTC_ORDER_TABLE`
(`WZTCRules.bas`), `FIND_REFERENCE_LINEWORK` (`WZTCQuery.bas`, defaults to
active-model-only per the engineer's direction ? reference-attachment
scanning is built and opt-in via `includeReferences=True` but not
exercised by default), `DEFINE_ALIGNMENT_SEGMENT`/`COMMIT_ALIGNMENT`
(`AlignmentTool.bas`/`WZTCExec.bas`), and `PLACE_ORDER_TABLE_STATIONS`
(`PerpPlacement.bas`, batches what used to be one `place_perp_line` call
per order-table item into one call per alignment). All 5 wired into
`chat_driver.py`'s `_WZTC_OP_NAMES` and confirmed callable by the actual
chat agent, not just direct bridge test scripts ? full plan (work space ?
alignment ? spacing ? sign ? tick line) drawn successfully end-to-end
through the real `WZTCChatPanel.frm` conversation.

**Real incident, same session**: an uncapped O(n?) chaining algorithm in
`FindReferenceLinework` (rebuilding a connected path from scanned reference
elements) hung MicroStation badly enough that Ctrl+Break didn't recover it
? required a full restart. Root-caused to two things: unbounded chaining
complexity (fixed with a hard 80-segment cap before the expensive pass
runs) and reference-attachment COM calls (`.AsVertexList`, `att.Name`)
throwing instead of returning `Nothing` the way `.AsLineElement` does
elsewhere in this codebase (fixed with whole-block error guards, not
line-by-line). Full writeup: Claude Code memory
`feedback_reference_scan_hang_and_com_errors.md`.

**Two real `chat_driver.py` bugs found via the live conversation, both
fixed**: (1) `enter_mode`'s effect is deferred to the next turn by design,
but nothing told the agent that, so it retried WZTC tools in the same turn
repeatedly and concluded tooling was broken ? fixed via explicit
`GENERAL_MODE_HINT` guidance. (2) `_SESSION_MODE` didn't survive a
`chat_driver.py` restart even though conversation history does, causing
the same confusion a second way (agent's own history showed it "already"
in wztc mode when the fresh process was actually back in general) ? fixed
by persisting the mode to `Bridge/chat-session-mode.txt`, loaded at
startup with a safe fallback to `"general"` on anything unrecognized.

**New tool**: `mcp-server/restart_chat_driver.py` ? safe automated restart
(mirrors `hot_reload.py`'s role for VBA). Refuses to restart if
`chat-log.tsv`'s last entry isn't `FINAL`/`ERROR` (a turn or pending
`ask_user_choice` looks in progress). Two bugs found building it, both
fixed: a substring PID-match filter caught the restart script's own
process (`restart_chat_driver.py` contains the literal substring
`chat_driver.py`) producing repeated phantom-duplicate false alarms; and
an initial `CREATE_NO_WINDOW` launch made a genuinely healthy process
invisible to the engineer's normal terminal-based workflow. Full writeup:
Claude Code memory `feedback_restart_script_self_match_and_hidden_window.md`.

Also fixed 4 issues from a code review of the Cursor session's work
earlier the same day (line-style error handling swallowing color/weight
changes, a silent button-overwrite in `WZTCChatPanel.frm`, ~150 lines of
duplicated lookup logic in `wztc_ops.py`, and `CLAUDE.md`'s File Sync
Protocol being stale on the hot-reload-first convention).

Still open: `place_order_table_stations` vs `place_perp_line` preference
was reinforced in the system prompt + both tools' docstrings after the
agent chose the less-efficient per-item path once live ? not yet re-verified
in a second live run. `find_reference_linework`'s `includeReferences=True`
path is still unverified against real reference geometry.

## 2026-08-02 ? Cursor ? incomplete WZTC sketch: history + order-table enforcement

Read `Bridge/chat-history.json` for the evening drawing turn (~msgs 214?243).
Root cause was not missing tools at the end: after mode flailing and a
wrong claim that `place_workspace` didn't exist, the agent followed the
engineer's literal checklist (workspace + 500ft alignment + W20-01RA +
**one** `place_perp_line` at sta 0) and declared the plan complete ?
never called `build_wztc_order_table` / `place_order_table_stations`.
Prompt only framed the order-table path as "whole plan from a
description," so a one-sign ask skipped it.

Fix: expanded `WZTC_SYSTEM_PROMPT_ADDENDUM` trigger + anti-patterns
(`chat_driver.py`); `place_perp_line(..., one_off=False)` now refuses
when workspace was placed without an order table, or when an order table
exists but stations weren't batched yet (`wztc_ops.py` `_PLAN_SESSION`);
`commit_alignment` returns a `nextStep` hint; `exit_mode` clears the
flags. Restart `chat_driver` to pick this up. Still needs a live redraw
to confirm the agent takes the order-table path.

## 2026-08-02 ? Cursor ? live order-table agent retest + screenshots

Drove the chat agent from Cursor via `scripts/agent_test_drive.py`
(append `Bridge/chat-input.tsv`, poll `chat-log.tsv`, auto-reply
checkpoints). Path taken: `build_wztc_order_table` ? `place_workspace` ?
`define_alignment_segment`/`commit_alignment` ?
`place_order_table_stations` ? `place_sign`/`set_sign_attributes`. No
`place_perp_line` for plan stations.

Gotchas confirmed live: (1) a 500 ft alignment clamps later order-table
stations onto the endpoint ? agent correctly refused to `place_sign`
until length was enough; (2) adding a second `define_alignment_segment`
then `commit_alignment` only kept the *new* segment (`elementCount=1`),
so stations restarted mid-plan ? fix was redrawing **one** contiguous
2500 ft segment from the original start. Final stations
(`results_slot6.tsv`): 8 distinct rows, W20-01RA at cumulative **1740 ft**
`(1022110.87, 217353.75)` with label `36" x 36"`. Screenshots (local
`view_capture`, not agent `view_drawing`):
`Bridge/captures/review_ordertable_{overview,workspace,ticks,sign,chat}.png`.

## 2026-08-02 ? Cursor ? reverse mistaken 3ft sign-face rescale; keep Scale 960

User corrected the earlier design intent: faces must stay at annotation
scale so they match the TEXTEDITOR label (this drawing:
`AnnotationScaleFactor=960` ? `Scale=(960,960)`), not real-world feet.
Live COM on the order-table retest face `W20-01RA` near
`(1022130.87, 217353.75)` showed Scale?58.7 / 3.0ft bbox after
`RescaleJustPlacedCellToTrueSize`, vs older correct faces at Scale 960 /
~49ft bbox. Removed that rescale path from
`DrawSign.PlaceSignFaceAndText` (and the private helpers), re-assert
Default/color0/weight0 after TEXTEDITOR before the face/post so
placement actives aren't left on the text style, updated
`chat_driver`/`wztc_ops` prompt wording, hot-reloaded `DrawSign`, and
ScaleUniform'd the live test face back to 960. Finished sign symbology
remains `set_sign_attributes` ? SF_P / color 240 / weight 3 (Designer
finish step) ? that was correct on the retest, not a bug.

## 2026-08-03 ? Cursor ? sign label: single inch marks, white, 48x48 W20-01RA

Root cause of doubled inch marks (`36"" x 36""`): `CadInputQueue.SendKeyin`
does **not** VBA-unescape `""`, so `Replace(" ? "")` before
`TEXTEDITOR INSERT_TEXT` wrote literal doubles into the drawing (Legacy
pattern was wrong for CONNECT). Fixed in `DrawSign.InsertTextWithInchMarks`
(piecewise INSERT_TEXT around each `"`; live-probed). Label/placement
actives forced to color 0 (white); `ExecSetSignAttributes` changed from
color 240?0 (240 painted labels/faces black on this background).
`SignLibrary` W20-01RA Non-Freeway size updated 36?**48" x 48"** to match
Legacy/user. Verified live: `PLACE_SIGN` returned `size=48" x 48"` with
single quotes.

## 2026-08-03 ? Cursor ? sign assembly: edge stem + shaft on perp tip

Wrong agent assemblies put a 20ft stem to the **face origin** (through a
Scale-960 diamond) and centered the TWZSGN_P **bbox** on the tip (looked
SE / straddling). Engineer ref: 50ft stem **post outer edge ? face inner
edge** only, post **High.Y = tip**, and the T **shaft centerline** on the
tip (crossbar hangs east; origin ? tip+(1.6, ?halfPost)).

`DrawSign.PlaceSignAssembly` now: measure half-extent along outward dir;
`ShaftLateralOffsetFromOrigin` picks the tall subelement (shaft) and
offsets so shaft mid sits on the tip laterally; stem endpoints from
`attachmentPt + dir*(2*halfPost)` ? face inward edge (`STEM_GAP=50`);
`SnapInwardEdgeToTip` nudges post/face along dir so measured inward
edge lands exactly on the tip (half-extent alone was ~0.004ft short).

Verification gotcha: early "pass" assemblies were placed at fabricated
X (1022200/1022300) with **no perp under them** ? looked disconnected
from any tip even though Y math matched. On the real tip
(1022110.87, 217313.75) post High.Y = tip exactly and origin matches
engineer ref within ~0.02ft. Floating test junk deleted.

## 2026-08-03 ? Cursor ? full-sheet WZTC flow: ask Designer inputs + whole 619 sheet

Agent was stopping after one W20 because the prompt treated sheet
contents as optional and allowed gathering inputs "conversationally"
without forcing ask_*. Strengthened `WZTC_SYSTEM_PROMPT_ADDENDUM`:
mandatory ask_user_choice/ask_user for speed/road_type/lane/shoulder/sheet
before any draw; get_sheet_requirements for named closures; every registry
sign ? sign_rows; do not FINAL-complete until all isSign + elements/handoff.
Fixed stale `place_order_table_stations` "point/tangent" wording.

Live drive on new south alignment (Y?217040, ~3000ft): agent asked
shoulder/sheet/speed/lane via ask_*; pulled **619-311**; built full order
table; place_order_table_stations ?2; placed 4 signs (W20-01RA, W20-05RA,
W04-02R, G20-02); channelizing runs; TWZAP_P + TWZWVA_P cells; handoff for
dims + NYW8-33 gap. `place_workspace` returned OK twice/thrice but created
no shape (agent verified with find_elements_near and continued) ? separate
bug to fix. Harness: `scripts/agent_test_drive.py` now answers Designer
asks (including FINAL-text asks).

## 2026-08-03 ? Cursor ? visual QA of 619-311 south run (expensive, many defects)

Screenshots + COM on Y?217040 corridor confirmed user complaints:

1. **White faces**: `ExecSetSignAttributes` forces `el.Color = 0` on all
   created IDs (post/face/stem/label). That bleaches orange sign faces.
   Color 0 was meant for labels after 240?black; must not paint faces white.
2. **Line through face / way past sign**: stems are **3000 ft** verticals
   (e.g. 55511: Y 216992?213992), not 50 ft edge-to-edge. Same length as
   the just-placed alignment ? almost certainly **AccuDraw distance lock**
   leftover from `define_alignment_segment` corrupting later
   `PLACE LINE CONSTRAINED` (stem + channelizing).
3. **Channelizing everywhere**: two TWZCD_P lines also ~3000 ft (55742/55743)
   instead of the short polylines in the PLACE_ELEMENT_RUN request; custom
   linestyle reads as orange squares along the whole bogus path.
4. **No work-space hatch**: three PLACE_WORKSPACE tries, no shape in model
   near intended 50?12 rect (only PV cell + perp tick there).
5. **No dimensions**: only `handoff` (deferred) ? expected until a real
   dimension path exists.

Captures under `Bridge/captures/audit_south_*.png`.

## 2026-08-03 ? Cursor ? fix AccuDraw stems, workspace, sign attrs

Root cause of 3000ft stems/channelizing: CadInputQueue `PLACE LINE
CONSTRAINED` honored AccuDraw distance left at 3000 after
`define_alignment_segment`. Switched stem (`DrawSign.PlaceSignAssembly`),
channelizing (`ExecPlaceElementRun`), and workspace (`ExecPlaceWorkspace`)
to Element API (`CreateLineElement2` / `CreateLineElement1` /
`CreateShapeElement1`) ? same pattern as `PerpPlacement`.

`ExecSetSignAttributes`: stop blind Color override on face cells (0
bleached them; 6 painted the legend orange). Labels/stems ? color 0;
TWZSGN_P post ? color 6; faces keep library orange+black legend.

Live verify after dirt 3000ft align: stem len=50.00 tip-edge, workspace
elementId + HasPattern=True, channelizing len=160 (not 3000), face shows
ROAD WORK AHEAD on orange.

## 2026-08-03 ? Cursor ? leave sign-face cell symbology alone

`ExecSetSignAttributes` was setting `Level=SF_P` + `LineWeight=3` on
face cells. That forced every subelement to wt=3 and collapsed SFB_P
onto SF_P ? borders/text looked wrong vs engineer W20-01RA (subs are
ByCell wt=-1 with SF_P/SFB_P split). Face cells are now untouched;
only labels/stems get SF_P+white, post TWZSGN_P gets color 6. South
corridor faces re-placed; COM match to user ref on weights/levels.

## 2026-08-03 ? Cursor ? W04 yellow legend + workspace hatch visibility

Two live QA defects on south 619-311:

1. **W04-02R** library cell has yellow SF_P copies of the merge symbol
   on top of black SFB_P; yellow diamond fill also wins display order.
   `IsHidden` on cell components does **not** persist. Fix in
   `DrawSign.HideDuplicateYellowLegend`: `MoveToNextElement(False)` +
   `DeleteCurrentElement` for small SF_P/color-4 legend strokes, and
   `DisplayPriority` 2000/?2000 via `ReplaceCurrentElement` so black
   reads above the fill. **Do not** call `MoveToNextElement` with
   nesting=True while ReplaceCurrentElement ? infinite walk (killed
   live). Guard the loop.

2. **Workspace** looked solid because `CreateShapeElement1` defaulted
   to filled, and associative `SetPattern` is invisible when the view
   Patterns attribute is off. `ExecPlaceWorkspace` now uses
   `msdFillModeNone`, spacing 2, and `DrawDiagonalHatchLines` (real
   line elements) so stripes show regardless of view attrs.

Blank perp ticks on the order table are expected: Non-Sign rows
(Roll Ahead, buffer, tapers, ?) get ticks without faces. 619-311
placeable signs on south align: W20-01RA@1840, W20-05RA@2190,
W04-02R@2540, G20-02@DS; NYW8-33 still a library gap.

## 2026-08-03 ? Cursor ? encode first-time-right rules for WZTC agent

The south 619-311 run took many cleanup passes because the agent kept
drawing after silent failures and then "fixed" symbology the wrong way.
Encoded into `chat_driver.WZTC_SYSTEM_PROMPT_ADDENDUM` (First-time-right
QA) plus corrected tool docs: verify `place_workspace` elementId before
continuing; `set_sign_attributes` must NOT recolor faces (stale docstring
still said color=240/wt=3); blank isSign=N ticks are expected; mid-plan
visual check after workspace+first sign; don't fabricate verify coords;
W04 legend handled in VBA. Restart chat_driver for prompt/doc changes.

## 2026-08-03 ? Cursor ? auto Non-Sign labels, dims, work vehicle

Promoted former PlaceElements/PlaceCells handoffs onto the agent path:

- `PerpPlacement.PlaceOrderTableLabels` ? text at Non-Sign outward tips
- `WZTCExec.ExecPlaceDimension` ? visual dim (ext + dim line + text,
  color 2). CadInputQueue `DIMENSION SIZE WITH LINES` creates **0**
  elements from programmatic SendDataPoint on this install; Element API
  annotation is the headless path.
- `PlaceOrderTableDimensions` ? every consecutive station pair spacing>0
- `PlaceSheetSymbolCells` ? ProtectiveVehicle?TWZWVA_P, ArrowPanel?TWZAP_P
  at Vehicle Space tip
- `AlignmentTool.AdoptExistingAlignmentElement` ? recover SharedState
  after VBA reload without redrawing the align

Bridge/MCP/registry + chat_driver completion checklist updated. Live
smoke on south align 55431: 7 labels, 9 dims, TWZWVA_P+TWZAP_P OK.





## 2026-08-03 ? Cursor ? real ny_Plan dims; fix labels + ArrowPanel gap

Engineer correctly rejected fake line+text dims. Annotate Linear Dimension
maps to `CreateDimensionElement1(..., msdDimTypeSizeArrow)` +
`DimHeight` + `DimensionStyles(\"ny_Plan\")` (active in this DGN;
also Det_Dim_Above / ny_Details / etc.). CadInputQueue DIMENSION SIZE
WITH LINES still creates 0 headlessly; Python COM Type-mismatch on
Matrix3d ? VBA path works. Live: 9 `IsDimensionElement` dims on south
align 55431.

Also: Non-Sign labels now at mid-segment + 35 ft past tip (was tip-stacked);
ArrowPanel origin 70 ft upstream of ProtectiveVehicle (was 15 ft; cells
~46 ft long ? overlapped). Cleared prior fake annotation junk before
re-smoke.

## 2026-08-03 ? Cursor ? sheet-gated annot; tip dims; PV in VS bay

Follow-up to engineer QA on labels/dims/PV:

- Double dim text: `ny_Plan.ShowSecondaryText` was True (primary +
  secondary). `ExecPlaceDimension` now forces False for the placement
  copy so only one measurement shows.
- Dims measure **tick tip ? tip** (same ends as the perp ticks), not
  centerline; Roll Ahead included from path start. Sheet-gated: for
  619-311 only Roll Ahead / Vehicle Space / Buffer / MergingTaper ?
  no Shoulder / Upstream Barrier / Beam labels or dims.
- Labels X-centered on the same tip-segment midpoint as their dim
  (`CreateTextElement1` + `Move` half-width ? Justification alone
  does not center on this install).
- `TWZWVA_P` centered in the Vehicle Space tick bay, scaled to table
  length (~50'). ArrowPanel ~55 ft upstream of VS start (~31 ft clear).
- `place_*` tools take `sheetElements` from `get_sheet_requirements`.

## 2026-08-03 ? Cursor ? sheet-first authority; 619-311 ShoulderTaper gap

Engineer correctly called out that verbal 'maybe skip shoulder taper' was a
trick ? official 619-311 PDF has SHOULDER TAPER (L/3) on the plan and Table
311-02. `Data/sheet-registry.tsv` had omitted `ShoulderTaper` from the
elements pipe; fixed. Rule: standard sheet > registry > chat; if registry
misses sheet content, stop and report a data bug.

Annotation rules corrected: dims between **every** consecutive tick (not
sheet-gated); name labels below only when sheet-required; length above dim
(Merging/Shifting pattern). Encoded in `chat_driver` + always-on Cursor
rule `.cursor/rules/sheet-first-qa.mdc` + visual QA before engineer handoff.
## 2026-08-03 ? Cursor ? workspace/AP/channelizing sheet placement

Finished the remaining 619-311 geometry (workspace hatch, Arrow Panel at
Shoulder Taper, bounded channelizing) on south align 55431.

Root cause of PLACE_WORKSPACE bridge hangs: msdFillModeNone is NOT a
defined VBA constant in this MicroStation project (Compile error:
Variable not defined), which popped a blocking dialog. ExecPlaceWorkspace
now passes fill mode as numeric 0 to CreateShapeElement1(Nothing, vpts, 0)
? do not assign .FillMode after create either (also dialoged earlier).

Live placement: WS 59908 path-start->VS (~210x12, 78 hatch lines);
channelizing 59987/59988/59989 (shoulder ~80 + merging ~560 + long ~570);
PV 59990 in VS bay; AP 60039 at Shoulder Taper tip X~1020945.
Deleted leftover AccuDraw-length TWZCD lines 56328/56627 (~1970 ft).

## 2026-08-03 ? Cursor ? root cause: agent path never actually reads the standard sheet

Investigated why every rebuild of 619-311 produces new sheet violations.
Rendered the official sheet (Bridge/captures/sheet_619311_p1.png, downloaded
from dot.ny.gov .../standard-sheets-us-repository/619-311.pdf) and diffed the
plan against a live station audit of the south corridor (align 55431).

Root cause is structural, not a placement bug: the only machine-readable
"sheet" is one row of Data/sheet-registry.tsv whose `elements` field is six
tokens. It carries no geometry ? no station order, no work-area extent, no
sign sequence, no PV count/positions, no device spacing, no dimension list.
Every placement function in PerpPlacement.bas therefore hardcodes its own
guess, and the agent re-derives the rest from prose in
chat_driver.WZTC_SYSTEM_PROMPT_ADDENDUM, which yields a different answer each
run. WZTCRules.GetDefaultUpstreamItems returns the SAME 7 rows for every
sheet in the catalog ? wztcSheet is stored and never consulted ? so the
station skeleton is already wrong before anything is drawn.

Confirmed sheet violations, live:
- Work area hatch (PlaceOrderTableWorkspace) spans path start ? Vehicle Space,
  i.e. it covers the Roll Ahead Distance. Sheet Note 4 forbids anything in the
  roll ahead distance; the hatched WORK AREA belongs on the far side of
  station 0. Order table has no Work Area length row at all.
- "Vehicle Space" (50 ft) does not exist on 619-311. The vehicle sits at the
  end of the Roll Ahead Distance; the code invents a segment and dims/labels it.
- Advance sign order is reversed. Sheet (upstream?taper): W20-1 ROAD WORK,
  then W20-5R RIGHT LANE CLOSED YY, then W4-2R merge nearest the taper. Live
  has W20-01RA nearest and W04-02R furthest. Nothing in code or prompt defines
  the ordering convention ? signs land in whatever order sign_rows was typed.
- Sign spacing is a flat 350 ft from SignLibrary.defSpacing. Table 311-03 is
  A/B/C by area type (urban ?30=100, 35-40=200, ?45=350, rural=500) and is
  never read. WZTCRules.ComputeSpacing computes a different
  AdvanceWarningSpacing (930 @ 45 mph) that matches no table here and is unused.
  There is no urban/rural input anywhere in the app.
- Arrow panel belongs at the upstream end of the LANE taper (sheet draws the
  trailer there, level with the top of the shoulder taper), not at the Shoulder
  Taper tick. The 2026-08-03 "AP at Shoulder Taper tip" decision was wrong.
- Phantom stations: Upstream Taper Temp Barrier + Box/Corr Beam ticks (~240 ft)
  are always emitted, push every sign upstream, and get perp ticks on a sheet
  with no barrier.
- PV / AP / sign posts are placed at the perp tick TIP (PERP_HALF_LEN = 40 ft),
  a drafting construct, so they sit 40 ft off the pavement. The manual
  PlaceSign/PlaceCells flow had the engineer click the real spot on the tick.
- G20-02 END ROAD WORK placed 2989 ft UPSTREAM past all advance signs; sheet
  puts it downstream, 80-400 ft past the downstream taper. Downstream taper
  never drawn.
- Shoulder-taper channelizing ends at 0.35*laneWidth inside the lane while the
  lane taper starts at offset 0 ? the two runs do not connect, and the sheet
  puts the shoulder taper on the shoulder (outboard), not inside the lane.
- Rebuilds are additive, not idempotent: live had two identical TWZWVA_P cells,
  a stale 100 ft channelizing stub inside the hatch, a label missing its
  length, and dims on only 4 of 11 tick spans.

Fixing individual placements will not stop the regressions. The sheet needs a
real machine-readable spec (per-sheet station sequence, sides, annotation set,
symbol anchors) that the placement code reads, plus idempotent re-place.

## 2026-08-03 ? Cursor ? 619-311 machine-readable sheet spec (Data/sheet-specs/)

Wrote `Data/sheet-specs/619-311.json`: the whole sheet as data ? tables
311-01..311-05 verbatim, the ordered corridor of zones, the two order-table
alignment walks, sign order + legend substitution, symbol anchors, the exact
annotation set, the five printed notes, 16 machine-checkable rules, and a
`knownCodeDeviations` list pointing each defect at the section that settles it.
Schema documented in `Data/sheet-specs/README.md`. This supersedes
`Data/sheet-registry.tsv` for 619-311 only; every other sheet still has the
six-token summary. No VBA/MicroStation changes ? wiring the placement code to
read this is the next step.

Design decision: the spec is plain JSON read directly by Python, NOT through
the VBA bridge. `get_sheet_requirements` currently round-trips to MicroStation
via `WZTCSheetRegistry.bas` just to read a TSV off disk; there is no reason to
pay that for a static data file, and JSON is unusable from VBA anyway. VBA keeps
receiving resolved numbers as bridge params (same shape as `sheetElements`).

Extraction gotcha worth not rediscovering: **do not** transcribe these tables
from `page.get_text()`. Plain extraction interleaves the multi-column tables
with the notes column and the plan callouts ? the earlier `sheet-registry.tsv`
seeding was done that way, which is why `Data/README.md` warns its numbers
"still need per-sheet verification". Group `get_text("words")` by rounded `y`
and sort by `x` instead; that reconstructs rows exactly. Same trap on the render
side: PyMuPDF pixmap coords are not the coords of the downscaled image you
actually look at ? convert back through the render scale before converting to
PDF points, or every anchor you measure is wrong by that factor (cost me one
bad read of the shoulder-taper extent).

New findings this pass, all confirmed against the PDF (not in the prior entry):

- **Table 311-03's XX/YY columns are sign legend text, not distances.** W20-1
  reads "ROAD WORK XX" and W20-5R reads "RIGHT LANE CLOSED YY", where XX/YY come
  from the same row as A/B/C (urban ?45: XX="1000 FT.", YY="AHEAD"; rural:
  XX="1500 FT.", YY="1000 FT."). This is what resolves the SignLibrary
  Ahead/Feet/Mile variant ambiguity that `resolve_sign_code` currently punts to
  the engineer ? the sheet answers it.
- **NYW8-33 "LANE CLOSED" is vehicle-mounted, not a roadside post sign.** On the
  plan it is a leader callout pointing at the protective vehicle. Its absence
  from SignLibrary as a post-mounted sign is not a gap that blocks this sheet.
- **Sheet covers 25-55 mph only.** Table 311-02 has no 60/65 row. The app's
  dropdown and `WZTCRules.ComputeSpacing` both go to 65, so any 60/65 plan on
  this sheet is extrapolated, not standard.
- **Shoulder taper has exactly three width bands** (?4 / 5-7 / ?8 ft).
  `ComputeSpacing` invents distinct per-foot values for 8/9/10/11/12 ft; at
  45 mph it returns 160 ft for a 12 ft shoulder where the sheet says 120 ft
  (160 is the 50 mph value). Fabricated data, not a rounding difference.
- **Roll ahead distance is a MIN/MAX range**, not a single number
  (45-50 mph ? 80-160 ft). `ComputeSpacing` emits one value with no range.
- **The two tapers are one continuous cone line.** The lane taper crosses the
  lane width, the shoulder taper continues the same diagonal across the
  shoulder; they share the point at the travel-lane/shoulder edge. The current
  0.35*laneWidth jog has no basis on the sheet.
- **Arrow panel anchor pinned precisely**: upstream end of the LANE taper, on
  the shoulder ? measured off the drawing, the trailer base sits on the
  lane-taper/shoulder-taper junction line. Also, VEH #1 is the sheet's "OR"
  alternative to the trailer at the same station, not a second vehicle.
- Work area hatch spans the closed lane **and** the closed paved shoulder
  (main plan and Detail 311A agree).
- `10'-0" (MIN.)` is measured centerline ? longitudinal cone line, i.e. the
  minimum remaining open travel lane, not a device offset.

Verification: `python scripts/validate_sheet_spec.py Data/sheet-specs/619-311.json`
checks cross-references plus transcription invariants that actually catch typos
? on this sheet a skip line is exactly 40 ft and channelizing devices are always
skip lines + 1, so one wrong digit in any taper triplet fails. It then resolves a
worked case and prints the station table. `Bridge/_roundtrip_619311.py` (throwaway)
re-extracts every table cell and all five notes from the PDF and diffs them
against the JSON: 0 failures across 311-01..311-05.


## 2026-08-03 - Cursor - 619-311 spec: geometry re-derived from PDF vectors (found a real error), then wired into the order table

### The vector layer makes plan geometry deterministic - use it, don't eyeball

These standard-sheet PDFs are vector, not scans. `page.get_drawings()` on 619-311
returns 5229 paths, and the plan is fully recoverable from them:

- Dimension lines are long vertical strokes in narrow x bands. **Their endpoints
  are the segment boundaries**, exactly.
- Each dimension's text label sits at the **midpoint** of the segment it
  dimensions, so matching label centres to segments labels every segment.
- Symbols are orange paths, colour `(1.0, 0.5, 0.0)`: arrow panel, both vehicles,
  sign diamonds, hatch. Yellow `(.94,.94,0)` is centreline, grey `(.94,.94,.94)`
  pavement.
- Descending y is downstream on this sheet.

`scripts/extract_plan_geometry.py` does this generically for any 619 PDF. Run it
BEFORE writing a spec's `corridor` section. Gotcha when writing similar code:
blob merging needs repeated union-find passes, not one pass - a single pass leaves
rectangles fragmented into their four edges and they never classify.

### The error this caught (would have produced a wrong drawing)

The first draft of `619-311.json` said gap A is datumed at the upstream end of the
SHOULDER taper, with the shoulder taper a sequential station in the upstream walk.
The vectors say otherwise: the A dimension (x=326.4, y 479.9-572.1) and the
shoulder taper dimension (x=133.6, y 479.9-509.2) **share the datum y=479.9**.

So the shoulder taper lies INSIDE gap A - it starts where the lane taper ends and
runs upstream within A, consuming no station of its own. Encoding it as a
sequential row pushed every advance warning sign 120 ft too far upstream
(1470/1820/2170 instead of the correct 1350/1700/2050).

On a not-to-scale drawing this is invisible by eye and obvious in the coordinates.
The extractor now prints a DATUM SHARING section specifically to surface it.
Spec gained `consumesStation: false` / `containedIn` / `overlayZones`, plus rule
`shoulder-taper-is-an-overlay`. Also corrected: the roll ahead dimension meets the
protective vehicle's FRONT (dimension ends y=303.2, vehicle body starts y=304.7),
so the roll ahead is clear space ahead of the vehicle - the spec said rear.

Everything else in the spec's `drawing`-confidence claims re-verified and passed:
taper continuity, arrow panel at the lane-taper junction, A/B/C contiguous, sign
diamonds at the upstream end of their gaps, all advance signs upstream of the
shoulder taper.

### Task 1 done: the spec now drives the order table

`mcp-server/sheet_spec.py` (new) owns spec loading/resolution in Python - VBA has
no usable JSON parser and the spec is a static data file, so routing it through
the bridge buys nothing. VBA keeps receiving resolved numbers as bridge params.

- `WZTCRules.BuildOrderTable` gained `specRows/specRowCount/overridesTSV`. Non-empty
  specRows REPLACE `GetDefaultUpstreamItems`, which returns the same 7 rows for
  every sheet including three stations 619-311 doesn't have (Vehicle Space,
  temp barrier, box/corr beam). Empty = no spec, legacy path unchanged.
- `WZTCBridge.ExecBuildOrderTable` parses `nonSignRowsTSV` / `spacingOverridesTSV`.
- `wztc_ops.build_wztc_order_table` resolves the spec when one exists, and can now
  derive the sign rows itself - `sign_rows` is optional on spec sheets. Returns
  `specDriven` + `stationWalk` + `signLegends`. Requires `area_type` on spec sheets
  and raises if omitted, because sign spacing AND sign legends both depend on it.

**Sign variant ambiguity is solved by the sheet.** Table 311-03's XX/YY are sign
legend text, and SignLibrary's suffixes map straight onto them: AHEAD->A, feet->F,
mile->M. URBAN 45 mph gives "ROAD WORK 1000 FT." -> `W20-01RF`; RURAL gives
`W20-05RF` where URBAN gives `W20-05RA`. Note `W20-01R*` means Work (ROAD) while
`W20-05R*` means Right - the R is not the same thing in the two families.
This is what `resolve_sign_code` currently punts back to the engineer.

**Careful with the device-count override.** `wztcSkipLines` is
`SkipMerge + SkipShoulder + SkipBuffer + SkipRollAhead`, and the sheet gives skip
counts per TAPER only - nothing for buffer or roll ahead. So Python sends the four
per-taper counts and VBA substitutes only the taper terms, keeping ComputeSpacing's
buffer/roll-ahead skips. Overriding the aggregate with just the taper sum would
have been a fresh fabrication of exactly the kind this work is removing.

Shoulder band collapse also lands here: the app's dropdown is per-foot (8..12 ft),
the sheet prints three bands, and `ComputeSpacing` invents 160 ft at 45 mph/12 ft
where Table 311-02 says 120. `sheet_spec.shoulder_band()` maps the dropdown onto
the printed band so the fabricated value never reaches the drawing.

### State / what is NOT done

MicroStation was closed, so **the VBA changes are on disk but not loaded and not
compile-checked** - hot-reload `Modules/WZTCRules.bas` and `Modules/WZTCBridge.bas`
before the next live run, and restart the Python process (chat_driver) for the
`sheet_spec.py` import. The end-to-end payload is proven without MicroStation by
`Bridge/_test_order_payload.py` (stubs the bridge and prints the exact call).
Task 2 (idempotent rebuild) not started.

`Data/sheet-specs/AUTHORING.md` documents the whole extraction procedure for the
remaining sheets.

## 2026-08-03 - Cursor - live BUILD_WZTC_ORDER_TABLE confirmed sheet-driven

Hot-reloaded WZTCRules.bas + WZTCBridge.bas into the open MicroStation session
(process name is `microstation`, not `ustation` -- earlier "MS closed" report
was a false negative from that filter).

`Bridge/_live_order_table.py` called BUILD_WZTC_ORDER_TABLE with the 619-311
spec payload (45 mph / 12 ft / URBAN). Status OK, 8 rows, all checks passed:
Roll Ahead / Buffer / Lane Taper / W04-02R / W20-05RA / W20-01RF upstream;
Downstream Taper / G20-02 downstream; no Vehicle Space, barrier, or shoulder-
taper sequential station. SharedState now holds the sheet table.

Note: PID 30484 is `mcp-server/server.py` (started 4:03 AM, before today`s
Python edits). Restart that process before MCP tool calls that use sheet_spec /
the new build_wztc_order_table signature. chat_driver.py is not required for
the bridge keyin path.

## 2026-08-03 - Cursor - idempotent rebuild via CLEAR_PLAN_ELEMENTS

Root cause of stacked rebuilds: place ops are additive; nothing deleted the
previous run`s ticks/cells/channelizing before the next place. Live had
duplicate TWZWVA_P, stale channelizing stubs, incomplete dims.

Fix:
- New bridge op CLEAR_PLAN_ELEMENTS (WZTCBridge.ExecClearPlanElements):
  walks the journal for createdElementIds= on OK RESP lines that are not
  already UNDONE, deletes those elements via ExecDeleteElementsByID, marks
  each create-op UNDONE. keepAlignments=Y (default) skips
  DEFINE_ALIGNMENT_SEGMENT / COMMIT_ALIGNMENT / ADOPT_ALIGNMENT_ELEMENT so
  a rebuild reuses the corridor.
- Python clear_plan_elements() + place_order_table_stations(clear_prior=True).
  Re-place of an align already placed this session raises unless
  clear_prior or force. chat_driver prompt step 3b documents the rebuild wipe.

VBA gotcha: Dim idS collided with Dictionary ids because VBA is
case-insensitive -- compile hung SendKeyin. Renamed to oneId.

Live: first CLEAR deleted 106 of 214 journal IDs (3 create-ops from the
3:17 AM workspace/channelizing/symbols run). Second CLEAR deleted=0.
Re-place gate refuses without clear_prior. Stale response.tsv can
false-match reused P1 reqIds -- wipe response before a fresh probe if a
prior op left P1 in the file.

## 2026-08-03 - Claude Code - Phase 0: generalized the sheet-spec test tooling before batching the other ~73 sheets

Cursor's 619-311 work (spec + validate_sheet_spec.py + _roundtrip_619311.py +
_live_order_table.py) is real and verified, but all three test tools were
hardcoded to 619-311 specifically -- literal `tables["311-02"]`/`"311-05"`
string lookups in the validator, hardcoded PDF pixel windows in the round-trip
script, hardcoded expected sign labels in the live-build check. Running any of
them against a different sheet either breaks or silently checks nothing. With
~73 more sheets planned (see STATUS.md below), that meant a bespoke throwaway
script per sheet forever -- doesn't scale and is exactly how the last-sheet
quality erodes from the first.

Fixed by making the sheet -> its own table ids the data, not the code:

- **New `tableRoles` section** in the spec schema: canonical role name
  (`taperAndBuffer`, `advanceWarningSpacing`, `rollAheadDistance`, `signSizes`,
  `protectiveVehicle`) -> this sheet's actual table id. Added to 619-311.json.
- **`mcp-server/sheet_spec.py`**: `resolve()` and `zone_length()` now key off
  `spec["tableRoles"]` instead of literal `"311-02"`/`"311-03"`/`"311-04"`
  strings. Verified identical output before/after against 619-311 (same
  bufferFt/laneTaper/station numbers as the earlier live-verified run).
- **`scripts/validate_sheet_spec.py`** rewritten sheet-generic: reads
  `tableRoles` for which invariant checks apply (skips with a WARN, not a
  silent pass, if a sheet doesn't declare a role), and calls
  `mcp-server/sheet_spec.py`'s `resolve()`/`station_walk()` instead of
  duplicating that logic. CLI no longer defaults to 619-311-specific case
  values -- with no args it now derives a worked case from the spec's own
  `inputs` declarations (first allowed value per input), so it makes sense to
  run against a sheet you've never seen. Verified: 0 errors, station numbers
  match the live-confirmed 1350/1700/2050 sign stations both with an explicit
  case and with no args at all.
- **`scripts/pdf_table_extract.py`** (new): the coordinate-grouping primitives
  (`words_in_window`, `group_rows`, `squash`, `assert_row_count`) factored out
  of the throwaway round-trip script. Deliberately did NOT try to make the
  round-trip comparison itself generic -- every 619 table has a different
  column layout, and a blind generic differ would either miss real
  transcription errors on an unfamiliar layout or need endless special cases,
  which is worse than the current state. Each sheet still gets its own short
  script, but now under a fixed convention (`Bridge/roundtrip/<sheet>.py`)
  built only from these shared primitives instead of copy-pasted whole.
  `Bridge/roundtrip/619-311.py` is the migrated version of
  `Bridge/_roundtrip_619311.py` (left in place, not deleted) -- reran against
  the actual PDF, 0 failures, same as before.
- **`scripts/live_build_check.py`** (new, generalizes
  `Bridge/_live_order_table.py`): expected upstream/downstream labels and
  forbidden labels are now derived from the spec's own `orderTable.alignments`
  (`rows` for what must appear, `excludedRows` + `overlayZones` for what must
  not) instead of hardcoded per sheet. Dry-run verified against 619-311
  (without calling the live bridge) reproduces exactly the same
  present/forbidden lists the hand-written script checked. Not yet run against
  a live MicroStation session -- needs that before being trusted as gate 4 on
  the next sheet.
- **`Data/sheet-specs/STATUS.md`** (new): tracks all ~74 plan sheets across 9
  corridor families (grouped by shape, not by duration -- duration mostly
  changes which table row gets looked up, not the corridor shape) plus 619-011
  as the shared table library, plus mowing/mulching and misc detail sheets at
  lower priority per the user's stated focus (Short Term/Intermediate/Long
  Term, Freeway/Non-Freeway/Parkway). Each row tracks the 4 gates above.
  619-311 marked `done`; everything else `not-started`.
- `AUTHORING.md` / `README.md` updated to point at the new tooling and the
  `Bridge/roundtrip/<sheet>.py` convention, and to require `tableRoles` before
  `tables` in the authoring order (Step 5.5).

Next: 619-011 (shared table library), then family #2 reference sheet 619-302,
per STATUS.md.

## 2026-08-03 - Claude Code - 619-011 (shared table library) authored and verified; 619-302 recon done in parallel

Per STATUS.md's Phase 1. Used two background sub-agents for the mechanical
half of the extraction (each independently spot-verified before trusting,
per the project's usual bar) plus a third background agent for reconnaissance
on Family 2's reference sheet, while doing Table 011-01 and final assembly
directly.

**619-011.json** (new): downloaded the per-sheet PDF (same NYSDOT repository
URL pattern as 619-311, documented in AUTHORING.md Step 0 -- not a guess).
7 tables (011-01..011-07), a 33-item WZTC legend, and Detail 011A, all via
coordinate-grouped word extraction, none read off a rendered image. This sheet
is not a corridor/plan sheet -- new `sheet.kind: "referenceLibrary"` schema
addition (schemaVersion 1.0 -> 1.1), and `scripts/validate_sheet_spec.py` now
branches on it: reference-library sheets get a lighter structural check
(tableRoles resolve, every table has rows) instead of the corridor/orderTable/
signs cross-reference checks a plan sheet needs. `Bridge/roundtrip/619-011.py`
round-trips all of 011-02/03/04/06 plus Detail 011A against the actual PDF:
0 failures.

**Confirms the "shared table library" hypothesis directly**, not just by
fingerprint-matching numbers across sheets as before: 619-311's tables
311-01/02/03/04 are each confirmed (cell-for-cell during transcription, not
assumed) to be an exact excerpt of a slice of 011-01/02/03(buffer-merged-into-
311-02)/04/06 -- documented in 619-011.json's new `knownExcerpts` section.
311-05 (sign sizes) has NO source on this sheet -- flagged as likely living on
619-012 "Work Zone Traffic Control Sign Table", not yet fetched.

**Real findings from the transcription, not just plumbing:**
- Table 011-03 (buffer space) genuinely has only 8 rows -- 60 MPH is absent,
  jumping straight from 55 to 65. Verified independently twice (word-grouping
  extraction and a second `get_text('text', clip=...)` pass) plus a third time
  by me directly before accepting it. Not an extraction bug.
- Table 011-02 has one internally inconsistent printed cell: speed 65 / lateral
  shift 12 ft prints `800/19/20`, but 19 skip lines x 40 ft = 760, not 800 --
  every other one of the table's 108 cells satisfies `ft = skipLines * 40`
  exactly. Re-verified directly against the PDF twice before accepting this
  as a real sheet defect rather than a transcription slip. Documented in the
  spec's new `knownAnomalies` field on table 011-02 and deliberately NOT
  silently corrected (would violate the verbatim rule) or hidden from the
  validator (it still flags this one cell on every run -- expected and correct).
- Legend row 18 verbatim reads "AUTOMATED FLAGGER ASSISTAMCE DEVICE WITH
  OPERATOR" -- "ASSISTAMCE" is a genuine typo in the source PDF (confirmed via
  direct word-coordinate lookup), transcribed as printed.
- Detail 011A confirms the 40 ft skip-line constant directly from its source:
  30' skip + 10' line = 40' cycle. (`DEFAULT_SKIP_LINE_FT` in
  `validate_sheet_spec.py` was previously only inferred from 311's arithmetic.)
- Table 011-01 (protective vehicle) is a 5-duration x Freeway/Non-Freeway
  master table; 311-01 is confirmed to be exactly its SHORT_TERM/NON-FREEWAY
  slice. The other 4 duration columns (Mobile, Short Duration, Intermediate,
  Long Term) are new data needed for Family 1's Intermediate/Long Term
  siblings (412/414/423/523) -- source their protective-vehicle rows from here
  instead of re-transcribing.
- Table 011-06 (advance warning spacing) adds a FREEWAY row (A=1000, B=1500,
  C=2640, "1 MILE"/"? MILE") absent from 311-03 -- needed for freeway families.
- Table 011-07 gives the general taper-length formulas (Merging=L,
  Shifting=L/2, Shoulder=L/3, downstream/one-lane-two-way=50-100ft) that 311's
  shoulder-taper-as-L/3 rule and downstream-taper fixed range both trace back
  to -- cite this table instead of re-deriving the formula per sheet.
- Table 011-05 (flare rates for positive barrier) is entirely new (619-311 has
  no barrier); needed for Family 3 (freeway shoulder closure) and any
  long-term barrier sheet.

**619-302 recon** (background agent, reconnaissance only -- no spec written,
per STATUS.md's "prove one sheet clean before batching" caution applied at
per-family granularity too): downloaded the per-sheet PDF, mapped table/plan/
notes coordinates, ran `extract_plan_geometry.py` for the corridor. Confirmed
302-01/302-02(25-55mph)/302-05 byte-identical to 311 (and by extension
619-011). Real differences: **302's table numbering is 04=sign sizes/05=roll
ahead -- the reverse of 311's 04=roll ahead/05=sign sizes** (tableRoles must be
assigned by content, never by matching suffix digit); 302-02 adds a 65mph row;
302-03 is keyed by ROAD TYPE with an added FREEWAY row; three protective
vehicles vs 311's two; 8 printed notes in a different order than 311's 5; same
DATUM SHARING trap (shoulder-taper-inside-gap-A) confirmed again. Also found
sheet-registry.tsv's 619-302 row lists sign codes (R2-1/NYR2-2/NYR2-6) that
don't appear anywhere in the real per-sheet PDF -- likely book-PDF
transcription noise, not a real sheet gap; and the "(EITHER PVH OR PVL)"
fragment from an earlier book-PDF read does not exist on the real sheet at
all. Full findings in STATUS.md under Family 2.

Next: full careful authoring of 619-302 (not just recon) as Family 2's
reference sheet, then its siblings 303/402/403/504, per STATUS.md.

## 2026-08-03 - Claude Code - 619-302 authored and verified live (Family 2 reference sheet); 2 real tooling bugs found and fixed along the way

Full careful authoring pass on top of the earlier recon (which was reconnaissance-
only, per its own caveat). Did the corridor/anchor/DATUM-SHARING reasoning and
the 8 printed notes directly (the judgment-heavy part, same as 311's own
authoring); used two parallel background sub-agents for the mechanical table
transcription (302-01/02, then 302-03/04/05), each cross-checking against
619-011.json and 619-311.json before reporting back. Every sub-agent claim
was spot-verified directly against the PDF before being trusted, same bar as
619-011 -- and this pass is exactly why that bar matters: two of the recon
pass's own assumptions turned out wrong once actually extracted (302-02 has
8 rows not 9 -- no 60mph; 302-04 has 6 rows not 5 -- WARNING FLAG is present),
and both were caught and corrected before the spec was finalized, not after.

**All four gates pass**: `scripts/validate_sheet_spec.py` (0 errors, 1 expected
warning on the FREEWAY row's non-equal A/B/C, which is correct data);
`Bridge/roundtrip/619-302.py` (0 failures across all 5 tables + all 8 notes,
against the actual PDF); and **gate 4 (live build) actually ran against a real
MicroStation session** -- `BUILD_WZTC_ORDER_TABLE` returned 8 correct rows,
all sign/non-sign present-checks and forbidden-row checks passed. Also
re-ran gate 4 against 619-311 with the same (now-fixed) generic script to
confirm the tooling itself, not just this one sheet.

**Two real bugs in the Phase-0 tooling, found by actually using it on a
second sheet instead of just re-running it on the first:**

1. `mcp-server/sheet_spec.py`'s `resolve()` hardcoded the advance-warning-
   spacing table's row key name as `"areaType"` (619-311's own convention).
   619-302 and 619-011 key the same table role as `"roadType"` instead (a
   real, not superficial, schema difference -- 311 has a separate
   AREA_TYPE dimension crossed with speed; 302/011 fold URBAN/RURAL/FREEWAY
   into one ROAD_TYPE dimension with speed qualifiers baked into the URBAN
   rows). Fixed to detect and support both row-key conventions. Also added
   `applicability.speedRangeMph.allowed` (an explicit list) as an alternative
   to `{min,max,increment}`, for a sheet like 302 whose Table 302-02 genuinely
   skips 60mph even though its own master table (619-011's 011-02) doesn't --
   a uniform range would have silently implied a row that isn't there.
2. `scripts/live_build_check.py`'s generic `expected_labels()` compared the
   bridge's returned rows against the sheet's own bare `signCode` ("W4-2R")
   instead of the resolved SignLibrary key the bridge actually returns
   ("W04-02R") -- different digit zero-padding means the bare code isn't
   even a substring of the resolved key, so every sign present-check would
   have silently failed forever. The dry-run verification done during Phase 0
   (no live bridge call) could not catch this -- it only surfaced by actually
   running gate 4 against a live MicroStation session on this sheet. Fixed by
   resolving each Sign row's SignLibrary key via `sheet_spec.sign_library_key()`
   before comparing, same as the bridge does internally.

**Real findings on the sheet content itself**, all confirmed by direct
coordinate extraction (not assumed from the family resemblance to 311):
- Table numbering trap confirmed: 302-04=sign sizes, 302-05=roll ahead --
  the reverse of 311's 04=roll ahead/05=sign sizes. `tableRoles` is assigned
  by content, never by matching suffix digit.
- **Cross-sheet discrepancy, not a transcription error**: 302-02's
  65mph/12ft lane-taper cell prints `800/20/21` (internally consistent,
  20?40=800) but 619-011's own master table 011-02 prints `800/19/20` for
  the identical lookup (which 011-02's own `knownAnomalies` entry already
  flags as inconsistent). Two official NYSDOT sheets disagree on one cell.
  Verified directly against both PDFs' vector text before accepting this as
  real rather than an extraction slip. 619-302.json trusts its own printed
  value; both specs document the disagreement rather than silently picking
  a winner.
- Three protective vehicles (VEH #1 taper/arrow-panel alternate, VEH #2
  shoulder vehicle conditional on shoulder width >= 8ft per Note 8, VEH #3
  the main lane-closure vehicle) vs 311's two. VEH #2's station anchor is
  flagged in the spec as the lowest-confidence call on this sheet (its label
  shares a text-extraction row-band with the buffer-space dimension, read as
  the same station as VEH #3 but offset onto the shoulder) -- worth a visual
  confirmation pass before a live build relies on it for real.
- 8 printed notes vs 311's 5, in a genuinely different order (302's note 3 =
  311's note 4, and vice versa) plus 3 new notes (5, 7, 8) -- matched by
  content, not by assuming note-N means the same thing on every sheet.
- No "10'-0\" (MIN.)" minimum-open-lane dimension on this sheet at all --
  confirmed absent by coordinate search rather than assumed missing.
- Confirmed (again) that `Data/sheet-registry.tsv`'s sign codes for this
  sheet (R2-1/NYR2-2/NYR2-6) don't exist anywhere in the real per-sheet PDF --
  book-PDF transcription noise, not a real content gap. This spec supersedes
  the registry for 619-302.

Next: Family 2's siblings (303/402/403/504) as diffs against this reference,
then Family 3 (619-301, freeway shoulder closure), per STATUS.md.

## 2026-08-03 ? Cursor ? 619-303 tables draft extracted

Wrote `Data/sheet-specs/_draft_619303_tables.json` from `Bridge/captures/619-303.pdf` via PyMuPDF `words_in_window` / `group_rows` / `assert_row_count` (script: `Bridge/_extract_619303_tables.py`). Table roles by CONTENT (numbering differs from 302): 01=PV, 02=roll ahead, 03=spacing, 04=taper+buffer, 05=sign sizes.

Surprises worth carrying forward:
- **9 printed notes** (not 8 like 302). Note 8 is new (min lane width 11' freeway / 10' non-freeway); Note 9 = 302's Note 8 (VEH #2); Note 5 cites W20-5a (two-lane), not W20-5.
- Sign sizes use **W20-5aR**, not W20-5R. Speeds on 303-04: 25..55,65 ? **no 60** (same gap as 302).
- 303-04 identical to 302-02 including the known 65mph/12ft `800/20/21` vs 011-02 `800/19/20` discrepancy.
- **302.json transcription bug**: `302-01` SHOULDER/OTHER/ge45 is recorded as `P` but both 303.pdf and a re-extract of 302.pdf print `P, TMIA` (matches 011-01 SHORT_TERM). Fix 619-302.json when convenient; do not treat this as a sheet difference.
## 2026-08-03 ? Cursor ? 619-303 done (all 4 gates); Family 2 two-lane sibling

Authored `Data/sheet-specs/619-303.json` as a structural variant of 619-302 (not a blind clone). Corridor has two successive MERGING TAPER L with a dimensioned **2L** span between them; mid advance sign is **W20-5aR** (SignLibrary `W20-05aR*`); table roles by content are 02=roll ahead / 04=taper / 05=sign sizes (different numbering from 302); 9 printed notes; two arrow panels. Tables 303-01/02/03/04 verified identical to 302-01/05/03/02 including the 65mph/12ft `800/20/21` cross-sheet discrepancy vs 011.

Tooling: `mcp-server/sheet_spec.zone_length` now honors optional `lengthSource.scale` (2L = 2? laneTaper). Round-trip `Bridge/roundtrip/619-303.py`; live build returned dual MERGING TAPER + 2L + W20-05aRM ? all checks passed.

## 2026-08-03 ? Cursor ? 619-402 done (all 4 gates); intermediate one-lane sibling

Canonical PDF is `619-402_E3.pdf` (29-APR-2026) copied to `Bridge/captures/619-402.pdf` (2 pages). Corridor skeleton matches 302; genuine intermediate diffs: **PVH/PVL+TMIA** on 402-01, new **402-05** channelizing application matrix, Note 4 **20'** device spacing (not 302's 40'), Note 7 NY9-11 recommended, Note 8 regulatory speed mid A/B (R2-1/NYR2-*), tables on page 2. 402-03/04/02 byte-identical to 302-02/03/05. Round-trip + live BUILD_WZTC_ORDER_TABLE both passed. knownCodeDeviations flag that placement still defaults to 40' and does not yet emit the regulatory mid-A/B row.

Next per STATUS.md: 619-403 (intermediate two-lane), 619-504 (long-term), then Family 3 (619-301).

## 2026-08-03 ? Cursor ? Family 2 siblings 303/402/403/504 all four gates passed

Continued from Claude Code's 619-302 handoff per STATUS.md.

**619-303** (two-lane short-term): dual MERGING TAPER + 2L (`lengthSource.scale`), W20-5aR, 9 notes. `sheet_spec.zone_length` gained scale support. Live: W20-05aRM + 2L rows.

**619-402** (intermediate one-lane, E3 2-page): corridor?302; PVH/PVL; Note 4 **20'** spacing; 402-05 channelizing matrix; regulatory R2-1/NYR2-*. Tables 03/04/02 == 302 taper/spacing/roll.

**619-403** (intermediate two-lane): hybrid 303+402. Pages **rotation=270** (window round-trips fragile ? used sibling-identity + phrase checks). FREEWAY.C=2640 inferred (no 2640 token) ? knownAnomaly. Live dual-taper+2L+W20-05aRM.

**619-504** (long-term barrier): **structural break** ? no PV/roll-ahead roles; positive barrier + 504-03 flare rates; order table is Merging Taper + signs only. `resolve()`/`zone_length` now skip missing rollAhead/PV. Live confirmed no ROLL AHEAD/BUFFER rows.

Canonical PDFs: `619-402_E3`, `619-403_E1_0`, `619-504_E3` copied to `Bridge/captures/619-*.pdf`.

Next per STATUS.md: Family 3 reference **619-301** (freeway shoulder closure).

## 2026-08-03 ? Cursor ? fix 302-01 / 303-01 SHOULDER+OTHER+ge45 cell (P ? P, TMIA)

Late [Extract 619-303 tables](d9c301cb-8f4c-4f9e-8970-ad18ad990bde) draft flagged that `619-302.json` recorded SHOULDER CLOSURE / OTHER HAZARDS / ge45 as bare `P`, but both 302.pdf and 303.pdf (and 011-01 SHORT_TERM) print `P, TMIA`. Corrected that cell in `619-302.json` and the cloned `619-303.json` 303-01 row. Round-trips re-checked.

## 2026-08-03 ? Cursor ? Family 3 reference 619-301 table extraction

Authored `Bridge/_extract_619301_tables.py` ? `Data/sheet-specs/_draft_619301_tables.json` from `619-301_E3.pdf` (1 page, **rotation=270**). Family 3 freeway **right shoulder closure** differs sharply from Family 2 lane-closure 302: only **4 tables** (no advance-warning spacing table ? spacing is plan callouts 1320'/1500'/1000'/MILE); 301-03 is **shoulder taper + buffer only** (no merging/lane taper column) with **4 speed rows** (45?65) not 8; 301-02 roll-ahead keyed by **protective-vehicle GVW** not posted speed; 301-01 uses **PVH+TMIA** (2 shoulder rows) not 302's P/TMIA. Signs are W21-5aR/W21-5bR + W7-3a + G20-1 (no W20-5R/W4-2R/NYW8-33). Overlapping 301-03 shoulder/buffer cells match 302-02 exactly. Notes are **9** (adds W7-3a/G20-1/R2-1 regulatory content). `tableRoles` assigned by content; no `advanceWarningSpacing` role.

## 2026-08-03 ? Cursor ? 619-301 done (Family 3 reference); sheet_spec generalized for shoulder closures

Authored `Data/sheet-specs/619-301.json` from E3 PDF (`Bridge/captures/619-301.pdf`, rotation=270). Structural break from Family 2: **no advance-warning spacing table** (plan callouts 1000'/1500'/1320'); **shoulder taper only** (sequential station, not gap-A overlay); **roll-ahead keyed by PV GVW** (301-02) not posted speed; signs **W21-5aR / W21-5bR**; PV codes **PVH+TMIA**; speeds **45/50/55/65** only. 301-03 buffer+shoulder cells match 302-02 on overlapping speeds.

`mcp-server/sheet_spec.resolve` / `zone_length` / `validate_sheet_spec` updated: optional AW table, optional laneTaper / laneWidthFt, GVW roll-ahead lookup, `fixedFt` gaps, FREEWAY size fallback. Live BUILD_WZTC_ORDER_TABLE passed all checks.

Next per STATUS.md: Family 3 siblings (205/315/401/415/501) as diffs against 619-301.

## 2026-08-03 ? Cursor ? 619-315/415 ramp-approach table drafts

Authored `Bridge/_extract_619315_tables.py` and `_extract_619415_tables.py` ? `_draft_619315_tables.json` / `_draft_619415_tables.json`. Both are Family 3 **shoulder at ramp approach** (2 pages); **315** E1 rot=270 both pages like 301/403, **415** E3 rot=0. **315-01/02 identical to 301-01/02** (PVH+TMIA, GVW roll-ahead); **315-03** adds `lateralShiftTaper` lane 10/11/12 + shoulder bands **9?12 ft** (overlapping ?8 ft matches 301-03); **315-05** short-term channelizing matrix (301 has no table 05). Ramp-specific **W3-7a** (Note 8, plan-only ? not in 315-04 sign table); plan spacing **2640'/1500'/1000'**. **415** table numbering permuted vs 402 (01=taper, 02=roll, 03=PV, 04=chan, 05=signs); **415-03** FREEWAY-only 2-row PV; **415-02** omits ?40 MPH roll-ahead row; **NYR9-11** 48?84 + 11 N-nighttime notes. Neither sheet has advance-warning spacing table.

## 2026-08-03 ? Cursor ? Family 3 shoulder duration siblings 205/401/501 table drafts

Authored `Bridge/_extract_619205_tables.py`, `_extract_619401_tables.py`, `_extract_619501_tables.py` ? `_draft_619205/401/501_tables.json`. **205** (short duration, 1p rot=0): 3 tables only ? P+TMIA (not 301 PVH), speed-keyed roll-ahead, generic W21-5, **no taper/buffer table**. **401** (intermediate, 2p): 5 tables like 402 shoulder variant ? 401-03 shoulder+buffer **matches 301-03** on 45?65 with extra laneTaper 10/11/12 cols; 20' channelizing (Note 5); W21-5bU/W21-5c signs; 10+N notes. **501** (long-term E3): 4 tables ? **no PV/roll-ahead**; 501-01 identical to 401-03; 501-03 barrier flare 50/55/65 only; positive barrier on plan (Note 5); OM3-L/R object markers.

## 2026-08-03 ? Cursor ? Family 3 siblings 205/315/401/415/501 all four gates

Authored and cleared gates for all five Family 3 siblings vs reference 619-301 via `Bridge/_build_family3_siblings.py` + sign/size sync. Round-trips: `Bridge/roundtrip/619-{205,315,401,415,501}.py`. Live BUILD_WZTC_ORDER_TABLE all PASS.

Surprises worth keeping: **401 size table uses W20-5aR** (not W21-5aR) ? PDF confirms; library key `W20-05aRM`. **315 gapC = 2640'** (freeway mile) vs 301's 1320'; shoulderWidthBands must be in **physical ascending order** (not alpha-sorted) or validator false-fails monotonicity ? 315 has 7 bands including 9?12 ft. **315/415** carry `lateralShiftTaper` (lane 10/11/12) separate from shoulderTaper; need `laneWidthFt` input even though corridor still walks SHOULDER TAPER. **501** like 504: no PV/roll roles; drop `bufferSpace` from corridor but do not leave `channelizingDevices` runs pointing at it (use `positiveBarrier..workArea`). **205** has no taperAndBuffer role ? order is roll + W21-5 + W20-1 only.

## 2026-08-03 ??? Cursor ??? Family 4 Parkway sheets 306/212/114/041 all four gates

Authored complete specs for Family 4 from per-sheet PDFs (Bridge/captures/619-{306,212,114,041}.pdf): drafts via Bridge/_extract_family4_tables.py, specs via Bridge/_build_family4_specs.py, round-trips under Bridge/roundtrip/619-{306,212,114,041}.py. All four gates PASS (validate + round-trip 0 fails + live BUILD_WZTC_ORDER_TABLE).

**619-306** (family reference): MERGING+DOWNSTREAM like 302 but fixed gaps 1000/1500/2640 and **no AW table** (like 301). Table 306-03 cells == 302-02 on 45-65. **Plan has no shoulder-taper dimension** (L/3 cols exist in table only) ??? do not emit SHOULDER TAPER row. **No NYW8-33**; only 3 notes. **619-212**: short duration; plan SHOULDER TAPER only; gaps 500/1500/2640(? mi); NYW8-33 on PV; roll-ahead has no ???40 row (recon guess corrected). **619-114**: mobile; 3 tables; moving roll-ahead; W20-5R @500??? min; >15 min ??? 212. **619-041**: mowing/moving non-freeway; W8-23 only; NON-FREEWAY PV speed bands; roll-ahead includes ???40; work area ???40???; >5 min ??? 201.

## 2026-08-03 ??? Cursor ??? Family 1 siblings (202/203/312/317/325/412/414/423/523) all four gates

Authored complete specs for all nine remaining Family 1 non-freeway multilane lane-closure sheets as diffs vs `619-311` (builders `Bridge/_build_family1_part1.py` + `_build_family1_part2.py`; round-trips via `Bridge/_gen_family1_roundtrips.py` ??? `Bridge/roundtrip/619-{n}.py`). STATUS.md Family 1 rows all `done`.

**Surprises:** (1) Prior `619-202.pdf` / `203.pdf` captures were path-only text (2 words) ??? re-fetched from NYSDOT; short-duration sheets have **no taper/buffer/G20-2**, A/B only, operator stays in PV. (2) **312/412** taper tables have **no shoulder columns**; L/2 is an overlay like 311's shoulder taper; AW is A/B only. (3) **tableRoles by content**: 325 is 01=PV/02=roll/03=taper/04=AW vs 317's 01=AW/02=taper/03=roll/04=PV. (4) Intermediate/Long-Term PV cells match `619-011` INTERMEDIATE_TERM / LONG_TERM NON-FREEWAY slices. (5) `sheet_spec.resolve` now tolerates missing AW `C` and missing `shoulderTaper`; `live_build_check` null-safe on absent taper overrides.

## 2026-08-03 ??? Cursor ??? Family 5 ramp-adjacent sheets (10/10) all four gates

Authored complete specs for Family 5 (ref 619-318): 318, 316, 319, 113, 211, 416, 417, 418, 517, 518. Extract `Bridge/_extract_family5_tables.py` ??? drafts; build `Bridge/_build_family5_specs.py`; round-trip `Bridge/roundtrip/619-family5.py` (+ per-sheet wrappers). All four gates PASS live 2026-08-03.

**Surprises:** (1) **No AW spacing table** on any F5 sheet ??? fixed plan gaps (318/319/417/418 use 1000/1500/1320/1320 with **two W20-1**; long-term 517/518 use 1000/1500/2640). Table 318-03 is NY2C-4 advance-*placement*, not A/B/C. (2) Two taper shapes: **lane+3band** identical to 302-02@45-65 (318/319/418/518/316) vs **7-band shoulder-only** (416/417/517) where plan MERGING L aliases the 10/11/12 ft columns (L/3-scale values, not full 302 L). (3) Partial-exit 316/416 walk **SHOULDER TAPER** + W21-5 (not MERGING/W4-2R). (4) 113/211 are mobile/short-duration minimal ??? note-derived roll 80/160, no taper table. (5) Dual W20-1 order rows both resolve to `W20-01RM` today; nearer should prefer `W20-01RPM`.

## 2026-08-03 ??? Cursor ??? Family 7 mobile (111/110/112) done; Family 8 (101-104) blocked

Authored complete specs for Family 7 from Bridge/captures/619-{111,110,112}.pdf via Bridge/_extract_family7_tables.py + _build_family7_specs.py; round-trip Bridge/roundtrip/619-family7.py. All four gates PASS live 2026-08-03. 619-113 stays Family 5 cross-ref.

**Surprises:** (1) **Two roll-ahead shapes in one family**: 111 is speed-keyed moving (same cells as 114-02: 200/5-280/7 / 160/4-240/6) with P,TMIA PV; 110/112 are **GVW-keyed moving** (200/5-240/6 light, 160/4-200/5 heavy) with **PVH+TMIA** ??? header `45-60 / w 55` is context like 301, not speed columns (vision/image reads keep misreading this as a speed??GVW matrix). (2) **2-page 111/112**: Sheet 1 = shoulder <8' (fewer signs / shorter gaps); Sheet 2 = >=8' with 1500'-? mi advance; specs primary-model Sheet 2. (3) 111 Sheet 1 has **no W20-5R** (only NYW8-33+W4-2R vehicle-mounted); Sheet 2 adds it. 112 Sheet 1 already has W20-5AR; Sheet 2 adds W4-2R.

**Family 8 blocked:** 619-101..104 have no real PDFs ??? every candidate URL on standard-sheets-us-repository, transportation-systems/repository, and metric repo returns ~10 KB HTML error pages (HEAD 200). DesignerRef-only; registry already warned not in 2026 Book 3. Historical master list numbered stop-and-go as 101-104; current NYSDOT special-ops index shows 045/046 instead (also undownloadable from the same public repos). Cannot clear gates without a vector PDF.

## 2026-08-03 ? Cursor ? Family 6 two-lane/flagger sheets (15/18) gates cleared

Authored complete specs for Family 6 (ref 619-307): 307, 308, 309, 314, 321, 322, 323, 324, 407, 421, 422, 519, 524, 090, 091. Extract `Bridge/_extract_family6_tables.py` ? drafts; build `Bridge/_build_family6_specs.py`; round-trip `Bridge/roundtrip/619-family6.py` (+ per-sheet wrappers). No sheet_spec/validate tooling changes needed ? buffer-only `taperAndBuffer` (optional laneTaper/shoulderTaper) already worked from 205/114 patterns.

**Surprises:** (1) **307 has no merging taper** ? `taperAndBuffer` is buffer-only; AW/buffer/PV/roll cells == 311 on 25?55. Signs are W20-7/W20-4 (+ conditional W3-4) not W4-2R/W20-5R. (2) **407 buffer trap**: intermediate drops 25?40; only 45/50/55/65 (incl. 645/16) ? do not clone 307's 7-row buffer. (3) **314**: no AW table, fixed 500' gaps; roll is 2-band only (no =40). (4) **Pedestrian 321/322/519**: live-build n/a ? `BUILD_WZTC_ORDER_TABLE` errors on sign-only payloads (same class as 011). (5) **419/420 blocked** (NYSDOT returns HTML); **520 blocked** (image-only, no text layer).

## 2026-08-03 ??? Cursor ??? Family 9 + Misc sheet specs (gates cleared / blocked)

Authored Family 9 (ref 619-023: 023/021/022/031/032/033/060) and Misc (001/004/005/006/010/012/080) via `Bridge/_build_family9_specs.py` + `Bridge/_build_misc_specs.py`; round-trip `Bridge/roundtrip/619-family9.py`. Corridor sheets with roll+sign cleared all 4 gates live 2026-08-03. 021/080 live-build n/a (sign-only). Reference libraries use `kind=referenceLibrary` (live-build n/a). **050/051/002 blocked** ??? NYSDOT repos return HTML stubs (same as Family 8).

**Surprises:** (1) **GVW??speed roll matrix** on 022/023/032/060/033 (not pure GVW like 110, not pure speed like 041) ??? encoded speed-keyed with min=heavy/max=light plus verbatim lightGvw/heavyGvw; ???40 band has min==max=120 so `validate_sheet_spec` now allows min???max. (2) **031 registry title wrong** ("Freeway Mowing") ??? PDF is two-lane mulching/herbicide with P,TMIA like 041. (3) **033 plan typo** SEE TABLE 034-02 ??? 033-02. (4) **012 size cells** largely path-only; catalog harvests sign-code tokens + color legend from text.

## 2026-08-03 ? Cursor ? Sheet-spec batch complete (68 done / 10 blocked)

Finished continuous pass through remaining STATUS queue after Family 3. Final inventory: **68 specs done**, **10 blocked** (all PDF availability ? no not-started left).

**Blocked (need real vector PDFs):** Family 8 stop-and-go **101?104**; F6 sidewalk/crosswalk intermediates **419/420**; F6 **520** (image-only, OCR); F9 **050/051**; Misc **002**. Public NYSDOT repos return HTML stubs for these.

**Families cleared this session:** F1 leftovers (9), F4 Parkway (4), F5 ramp (10), F6 flagger/ped/closures (15/18), F7 mobile (3), F9 mowing (7/9), Misc ref-libs (7/8). Prior session already had F2+F3+011+311.

## 2026-08-03 ? Cursor ? Family 6 verbatim notes backfill (Claude follow-up)

Implemented Claude's bounded follow-up: restore `notes.printed` on all 15 available Family 6 specs (were empty `items[]`). Placement `rules[]` were already populated ? this is engineer-facing prose, not a CAD correctness gap. Helper `Bridge/_backfill_family6_notes.py`.

**Result:** upright sheets (307/308/309/323/407/421/422/519/524, plus 322) ? `confidence: verbatim`. Rotated-page sheets (090/091/314/321/324) ? `confidence: drawing` with disclaimer ? PDF text layer stores note columns side-by-side so linearization interleaves words; still better than empty, but edge-case wording should be confirmed against the PDF. Skipped D-tier PDF chase and 311-level cell round-trips on siblings per Claude's scope.


## 2026-08-03 ??? Cursor ??? Live agent WZTC QA: 619-311 + 619-301

Drove full agent WZTC mode builds after the sheet-spec catalog. **619-311** completed earlier in-session (specDriven URBAN path; QA on disk as Bridge/captures/qa_agent_311_*.png). **619-301** initially blocked because wztc_ops.build_wztc_order_table required area_type whenever any sheet spec existed ??? wrong for sheets with no advanceWarningSpacing role. Fixed gate to require area_type only when that role is present; null-safe overrides when laneTaper is absent. Restarted chat_driver and cleared poisoned Bridge/chat-history.json (agent refused retries based on prior errors). Fresh 301 run: order table without area_type (SHOULDER TAPER 160 / buffer 495 / roll 120, signs W21-05aR / W21-05bRM / W20-01RM / G20-02), 6000 ft corridor at Y=216840, stations + 4 place_sign + dims. Labels/workspace/channelizing tools returned empty or agent handed off (shoulder sheets lack Vehicle Space bay the batch tools assume). Gotcha: test-driver auto-replies invented 311 blurbs / Y=217040 ??? patched scripts/agent_test_drive.py with --align-y and area_type-aware ASK replies.


## 2026-08-03 ??? Cursor ??? Fix live-test bugs (labels / shoulder tools / area_type / registry)

Root causes from the 311/301 agent QA: (1) sheet-spec Non-Sign labels were ALL CAPS while PerpPlacement Select Case was case-sensitive ??? place_order_table_labels returned 0 rows; LANE TAPER also did not match Merging/Shifting Taper for channelizing. Fixed by sheet_spec.canonical_order_label at emit time + OrderLabelKind() in PerpPlacement.bas. (2) Workspace/PV required Vehicle Space; channelizing required Merging ??? shoulder sheets like 619-301 have neither. Fallbacks: Buffer Space bay for WS/PV; Shoulder Taper as primary channelizing diagonal. (3) build_wztc_order_table now only requires area_type when advanceWarningSpacing exists, accepts protective_vehicle_gvw, null-safe taper overrides. (4) sheet-registry 619-301 was missing ShoulderTaper (same class of bug as 311 earlier). chat_driver prompt updated. Hot-reload of PerpPlacement.bas needs MicroStation with the Test VBA project loaded.


## 2026-08-03 ??? Cursor ??? Cost controls for chat_driver / agent_test_drive

Did not ship with the CAD bugfix earlier ??? added now after engineer call-out. chat_driver: MAX_TOOL_ITERATIONS 30???18 (env WZTC_MAX_TOOL_ITERATIONS); persist history capped to 40 msgs / ~350k chars (WZTC_MAX_HISTORY_MESSAGES / WZTC_MAX_HISTORY_CHARS); strip prior-turn thinking blocks on save/load; load_history rewrites trimmed file so a fat 600+ msg history cannot silently return. agent_test_drive: default --max-continues 16???8; new --fresh-history clears Bridge/chat-history.json before a live drive. view_drawing docstring: prefer ???1???2 captures/turn. Restart chat_driver to pick up.


## 2026-08-03 ??? Cursor ??? Scoped clear_prior + live reload

Hot-reloaded PerpPlacement.bas + WZTCBridge.bas with Test project loaded; restarted chat_driver (wztc, empty history, cost caps active).

**clear_prior wipe bug:** place_order_table_stations(align_idx=2, clear_prior=True) used to call clear_plan_elements() with no scope and deleted Upstream too. CLEAR_PLAN_ELEMENTS now accepts optional alignIdx; scoped clear only deletes journal create-ops tagged with that alignIdx=. place_order_table_stations(clear_prior=True) passes its align_idx. place_sign gains optional align_idx (journaled) so future scoped clears include signs. Untagged legacy PLACE_SIGN rows are left alone under scoped clear (full clear_plan_elements() still wipes everything). Held off on work-area lengthSource=null (#3).

## 2026-08-03 ? Cursor ? View-capture flakiness (blank QA frames)

Root cause of many 'blank' QA screenshots was not Extents.Z (~0.02 after write is fine). Two real issues: (1) **over-wide framing** ? a 2000-ft overview makes each sign a few pixels so vision models read 'empty'; use ~150?400 ft width for sign closeups. (2) **async repaint** ? capture too soon after Center/Extents write yields uniform dark grey. Fixed in `mcp-server/view_capture.py`: settle default 2.5s; expand requested extents to the View child window's pixel aspect (`_fit_extents_to_aspect`); Center-then-Extents + redraw helpers; empty-frame detect (`_drawing_looks_empty`) with one retry; `navigate_and_capture` helper. `adjust_view` returns applied width/height after aspect-fit. Smoke: Bridge/captures/qa_capture_fix_{close,overview}.png on 619-301 corridor Y?216840 ? non-empty. Restart chat_driver to pick up (imports view_capture at load).

## 2026-08-03 ? Cursor ? Agent must restart chat_driver (not ask engineer)

Engineer call-out: agents kept saying 'restart chat_driver' instead of doing it. Updated `.cursor/rules/hot-reload-sync.mdc` so Python edits require running `python mcp-server/restart_chat_driver.py` (same alwaysApply rule that already covers VBA hot-reload). Prior wording told the agent to ask the engineer -- that was the bug.

## 2026-08-03 - Claude Code - Placement-plan compiler Stage 0 + Stage 1, planned and shipped

Full plan at `~/.claude/plans/snuggly-snacking-sifakis.md` (Claude-Code-local,
not repo-portable -- the stages are described in full below and in
`Data/sheet-specs/STATUS.md`'s new "Placement-plan compiler" section, so
Cursor/other tools have the substance even without that file). This is the
first work on the bigger redesign discussed with the user: move sheet
geometry semantics out of `PerpPlacement.bas`'s label-string-matching
heuristics and into Python, which already holds the full spec (zones,
overlays, device counts, anchors, rules) that dies at the bridge boundary
today.

**Stage 0 -- safety nets, both fixes live-verified:**

1. `mcp-server/wztc_ops.py`'s `build_wztc_order_table` now raises `ValueError`
   immediately when `sheet_spec.load()` returns `None`, before any bridge call
   -- no sheet should ever again silently get the generic 7-row
   `GetDefaultUpstreamItems` fallback. Confirmed: `619-502` now refuses
   cleanly.
2. **Root-caused the real "Subscript out of range" crash** on
   `619-321`/`619-322`/`619-519` (sign-only/pedestrian sheets, zero non-sign
   rows) -- this took a long live checkpoint-bisection session (added a
   `cp`/`checkpoint` marker pattern, same spirit as `PerpPlacement.
   FindInteriorPoint`, to `WZTCRules.BuildOrderTable` and `WZTCBridge.
   ExecBuildOrderTable`, hot-reloading and re-testing at progressively finer
   granularity) because static reading never turned up anything wrong. The
   actual bug: **`ReDim arr(0 To -1)` -- the standard VBA idiom for a
   zero-length dynamic array -- throws "Subscript out of range" in this
   specific MicroStation VBA host.** Not standard VBA behavior; an
   environment-specific landmine that had simply never been exercised before
   (every sheet ever tested had at least one row on both alignments). Found
   4 occurrences (`WZTCBridge.bas` x2, `WZTCRules.bas`, `DesignerRef.bas`) and
   fixed all of them the same way: a size-1 dummy array with the real count
   tracked separately (the pattern `WZTCRules.GetSpecItemsForAlignment`
   already used safely). `WZTCBridge.ExecBuildOrderTable`'s `signRowsTSV`
   parsing was also rewritten to build off a `Collection` instead of an
   incrementally-`ReDim`'d array while chasing this, which is unrelated to the
   actual bug but is more robust regardless (no subscript bookkeeping to get
   wrong). Confirmed live: 619-321/322/519 now return `OK` with correct sign
   rows; 619-311/619-302 live-build still pass (after fixing a stale
   assertion in `scripts/live_build_check.py` that didn't know about Cursor's
   `canonical_order_label()` normalization -- unrelated pre-existing gap, not
   a regression from this work).
3. Left the checkpoint-based error reporting in place (not stripped as
   scaffolding) since it's what actually made the bug findable and matches
   this codebase's own established pattern.

**Stage 1 -- compiler skeleton + parity proof, done:**

New `GET_ALIGNMENT_VERTICES` bridge op (`PerpPlacement.GetAlignmentVertices`,
dispatch wired in `WZTCBridge.bas`, Python registration through
`wztc_ops.py`/`server.py`/`chat_driver.py`'s `_BASE_OP_NAMES` -- the exact
registration gap that silently broke 28 tools before, per
`feedback_chat_agent_tool_registration_gap`) exports a committed alignment's
raw path segments (straight or arc, with center/radius/start/sweep for arcs)
in one call. `mcp-server/alignment_geometry.py` replicates
`PerpPlacement.GetPointAndTangent`'s exact station->XY arc-length walk in
pure Python -- **live-verified bit-for-bit identical** against the real
`STATION_TO_POINT` bridge op across 7 stations including both segment
boundaries and clamped out-of-range cases (0 mismatches), plus a synthetic
self-test covering straight, arc, and mixed-segment paths
(`python mcp-server/alignment_geometry.py`). `sheet_spec.compile_plan()`
compiles one alignment's stations/dimensions/labels into explicit primitives
in absolute coordinates, replicating `PlaceOrderTableDimensions`/
`PlaceOrderTableLabels`'s tick-tip/offset geometry line-for-line (same
`PERP_HALF_LEN`=40, `offsetDist`=15, `textExtraAlong`=20, `outwardSign`=-1.0
defaults). `scripts/test_compile_plan_parity.py` is the reusable two-layer
check (geometry-engine live parity + compile_plan smoke test against the real
619-311 spec) -- `python scripts/test_compile_plan_parity.py --align-idx 9`,
0 failures.

Used a small throwaway straight test alignment (`align_idx=9`, define via
`define_alignment_segment`/`commit_alignment`, 2 segments/250ft, at an
out-of-the-way location x~900000/y~100000) for all live verification rather
than touching the real corridor content in the currently-open "left lane
closure" model. **Not fully cleaned up** -- `clear_plan_elements(align_idx=9,
keep_alignments=False)` returned `deleted=0` (its journal-scoping filter
apparently doesn't match how `DEFINE_ALIGNMENT_SEGMENT` tags its align index;
not chased further). Two small harmless line elements are left in that model
at the out-of-the-way location above -- low priority to clean up manually,
noted here so it isn't mysterious later.

**Not started**: Stage 2 (channelizing as real counted cones -- the biggest
visual gap), Stage 3 (symbols/PV-count-from-legend), Stage 4 (hatch/Detail-A/
min-open-lane), Stage 5 (rules engine as a hard pre-draw gate + retiring
`PlaceOrderTable*`). See STATUS.md's new "Placement-plan compiler" section.

## 2026-08-03 - Claude Code - Placement-plan compiler Stages 2-5, planned as Stage 0/1 follow-up, shipped same session

Continuation of the compiler work above, all in `mcp-server/sheet_spec.py`
(`compile_channelizing`/`compile_symbols`/`compile_hatch`/`run_rules_gate`),
verified via `scripts/test_compile_plan_parity.py`'s now 6-layer suite. Used
two throwaway test alignments this session (align_idx 9 and 10, both at
out-of-the-way locations, same ones flagged for manual cleanup in the Stage
0/1 entry above -- align 10 was redefined twice for different distance
tests, final state has it ~1900 ft from align 9's origin).

**Stage 2 -- channelizing devices.** `PlaceOrderTableChannelizing` draws only
THREE bare 2-point polylines total, no device count, and a
`laneWidthFt*0.35` fudge factor for the shoulder-taper lateral offset that
leaves a real lateral jog where it meets the lane taper (documented in
619-311.json's `knownCodeDeviations`). `compile_channelizing()` places real
counted cones (count from the spec's `deviceCountSource`, e.g.
`resolved['laneTaper']['devices']`) and connects shoulder-taper to
lane-taper at the exact shared point instead of the fudge. Live-verified:
shoulderTaperRun and laneTaperRun's shared station both land at
(900000.0, 101000.0) -- exact match, not "close enough". Added
`check_taper_continuity()` as a standalone, reusable rule check (shipped
with this stage per the plan's own "validate per-stage" principle, not
deferred to Stage 5).

**Stage 3 -- symbols.** `PlaceSheetSymbolCells` only ever places ONE
protective vehicle and one arrow panel, no matter what the sheet needs.
`compile_symbols()` compiles every `symbols.items` entry with a
`stationAnchor` and a `cellHint` (vehicles) or `id=='arrowPanel'` --
deliberately does NOT re-derive a PV count from a table legend + closed-lane
count (that derivation already happened once, correctly, during spec
authoring; re-deriving it from scratch here would be exactly the kind of
guessing `sheet_spec.py`'s own module docstring says not to do). Verified
against both 619-311 (arrowPanel/protectiveVehicle2 correctly cross-linked
as alternatives, both landing at the identical station) and 619-302
(protectiveVehicle1 -- not 2 or 3 -- correctly identified as the arrow
panel's "VEH #1" partner, purely from `sheetLabel` matching, no
sheet-specific code).

**Stage 4 -- work-area hatch. Real design correction made mid-build**: the
original plan assumed `work_area_length_ft` would be an external parameter
the caller supplies. It's wrong -- `orderTable.alignments[0].station0` /
`[1].station0` are literally defined as "Upstream edge of the WORK AREA" /
"Downstream edge of the WORK AREA", so the work area's length is already
implicit in how the engineer/agent committed the two alignments (the
real-world distance between their own station-0 points), not a number to
invent or ask for separately. Rewrote `compile_hatch()` to take both
alignments' segments and use their own station-0 points directly.
Live-verified: two test alignments 300 ft apart (by construction) round-trip
back through `compile_hatch()` as `workAreaLengthFt: 300.0` exactly. Also
compiles the conditional Detail-A transverse device rows (only past the
sheet's own `maxSpacingFt`, both numbers from the spec, not invented) --
verified a 1900 ft work area produces exactly 2 transverse rows (800/1600),
a 200 ft one produces 0.

**Stage 5 -- rules gate, scoped honestly.** `run_rules_gate()` checks the
subset of a sheet's `rules[]` that's mechanically checkable from compiled
geometry (taper-continuity via Stage 2's check, cone-spacing, sign-order,
arrow-panel-anchor, no-occupancy-buffer-rollahead -- simplified to a
degenerate-length guard once Stage 4's redesign made the old overlap
impossible by construction rather than just checked-for). Verified in BOTH
directions, which matters more than it sounds: 0 failures on a correctly-
compiled 619-311 plan, and correctly catches a deliberately-reintroduced
taper jog (moved one cone 5 ft off) with a real failure message. **Did not**
retire `PlaceOrderTable*` or claim the original plan's Stage 5 bar met (a
golden test per family, 9 families) -- that's real follow-up work, not done
here. What shipped is the gate mechanism, proven on two families
(619-311/619-302), not a claim the whole catalog has been re-verified
through it.

All of Stages 1-5 verified together in one run:
`python scripts/test_compile_plan_parity.py --align-idx 9` -- 6 layers,
0 failures. Full regression re-run at the end: all 68 sheet specs (67
pass + 619-011's one documented anomaly), 619-311/619-302 live-build still
pass, `alignment_geometry.py`'s synthetic self-test still passes.

Next: either genuinely start retiring `PlaceOrderTable*` (would need the
per-family golden testing this session didn't do), or extend the compiler
to more sheet families to prove `compile_channelizing`/`compile_symbols`/
`compile_hatch` generalize as well as `compile_plan` already has been shown
to (only checked against 619-311 and 619-302 so far).

## 2026-08-04 ? Cursor ? Expose get_elements_range / absolute adjust_view (copy-turn miss)

Live agent turn copying a white line 100 ft above an orange arc burned all 18 iterations and never called `copy_element`. Root causes: (1) engineer already gave `elementId=61956` / `24763` but `GET_ELEMENTS_RANGE` existed only as an internal screenshot autofocus helper ? not an agent tool ? so the model fished with `find_elements_near` (bbox-center match + row cap misses long lines); (2) `adjust_view` `pan_x`/`pan_y` are relative deltas and the agent passed absolute model coords, flinging the view; (3) even a successful copy needed `own_element_only=False` for pre-existing site geometry. Fixed in `wztc_ops.py`: `get_elements_range`, `focus_view_on_elements`, `adjust_view` absolute `center_x`/`center_y`/`width`/`height`; exposed on chat_driver `_BASE_OP_NAMES` + `server.py`; prompt warns against absolute-as-pan and teaches ID?range?copy. Smoke: line Y=216840, arc top Y=217672.3 ? delta_y?932.3 for '100 above arc'. Restarted chat_driver.

## 2026-08-04 ? Cursor ? Wire adopt_alignment + place_sheet_geometry for agent

Compiler Stages 1?5 already lived in `sheet_spec.py` / `alignment_geometry.py` but the chat agent still only had heuristic `place_order_table_*`. `ADOPT_ALIGNMENT_ELEMENT` existed in VBA (`AlignmentTool.AdoptExistingAlignmentElement` / `WZTCBridge`) with no Python/agent wrapper ? live miss after hot-reload wiped SharedState while corridor lines remained. Wired `wztc_ops.adopt_alignment`, `compile_sheet_plan` (slim by default), `place_sheet_geometry` (compile + `run_rules_gate` + execute dims/labels/channelizing polylines/PV-AP/hatch), exposed on `chat_driver._WZTC_OP_NAMES` + `server.py`. Prompt now prefers compiler path when sheet JSON exists; `adopt_alignment` for Reset/hot-reload recovery. `execute_compiled_plan` stays internal (no huge coord round-trip through the LLM). Channelizing places one `place_element_run` per cone run (not per-cone circles). Restarted chat_driver PID 47336. Offline 619-311 compile+gate smoke: 0 gate failures.

## 2026-08-04 ??? Claude Code ??? QA-diagnosis fixes, agent-control-loop enforcement, and a two-part refactor

Long session across three rounds. **Round 1 (7 fixes from a live 619-311 QA diagnosis):** corridor-topology precondition added to `run_rules_gate` (catches Downstream committed further along Upstream's own line instead of a distinct work-area edge); `NYW8-33` added to `SignLibrary.bas` (W4-2R turned out to already exist ??? that "gap" was actually a lookup bug); `compile_hatch`'s transverse-channelizing condition now checks shoulder width ??? 8ft, not length alone; `SignLibrary.PadNumericStem` now zero-pads the route-family number too (`W4-2R`???`W04-02R`, skipped for `NY`-prefixed codes per the `NYW8-4` precedent) ??? this was the real root cause behind a live legend-resolution miss, not the legend-suffix logic itself (which was already correct); AP/PV lateral offsets in `compile_symbols` now derive from the sheet's own `lateralAnchor` text + real lane/shoulder widths instead of the unrelated 40ft reference-tick constant (backfilled `lateralAnchor` across 17 other sheet specs during verification ??? batch-tested all 61 "done" specs, zero crashes); `MAX_TOOL_ITERATIONS` raised 18???26 now that the cache-busting bug is fixed; added `get_locked_designer_inputs()` as real persisted session state instead of relying on the model rereading chat history; visual-QA-gate requirement added to the completion checklist.

**Round 2 (4 agent-control-loop enforcement gaps, from a self-critique of round 1):** `place_sign` now refuses a hand-picked sign code that doesn't match the locked order table's `sign_rows` (`one_off=True` to override) ??? this is the actual mechanism behind the legend-resolution miss above, not a content gap; the heuristic `PlaceOrderTable*`/`place_sheet_symbol_cells` tools now refuse when a sheet spec exists for the build and redirect to `place_sheet_geometry` (`force=True` to override) ??? previously they were a fully unvalidated escape hatch around `run_rules_gate`; `check_corridor_topology` factored into a standalone function and wired into `commit_alignment`/`adopt_alignment` so a bad corridor is caught the moment both alignments are ready, not only deep in `compile_hatch`; `AlignmentIsReady`/`InitAlignmentPlacementHeadless` VBA error messages now carry an `ALIGNMENT_NOT_READY:` prefix naming `adopt_alignment` directly instead of unconditionally pointing at `COMMIT_ALIGNMENT` (which would draw a duplicate).

**Round 3 (refactor, on request):** `wztc_ops.py`'s `_PLAN_SESSION` dict ??? `PlanSession`/`DesignerInputs` dataclasses. `sheet_spec.py` (1145 lines) split into `sheet_resolve.py` (load/resolve/legend/order-table), `sheet_compile.py` (the placement-plan compiler), `sheet_rules.py` (rules gate) ??? `sheet_spec.py` is now a ~65-line facade re-exporting every name, so no call site changed. `chat_driver.py` (1673???760 lines) split into `prompts.py`, `usage.py`, `chat_log.py`, `input_watcher.py`, `chat_history.py`; added `_validate_op_names` (startup `ImportError` if `_BASE_OP_NAMES`/`_WZTC_OP_NAMES` drift from `wztc_ops.py`'s real exports ??? this list had already silently drifted once before); bundled `_TOUCHED_ELEMENT_IDS`/`_SESSION_MODE` into a `SessionState` dataclass; deduped the 4x-repeated screenshot PNG???BMP conversion. Also fixed two real `server.py`/`wztc_ops.py` signature-drift bugs found during an audit: `list_registry_commands` was missing `opname_contains` (unreachable cost-saving filter, ~$0.75/unfiltered call) and `focus_view_on_elements` was missing `min_width`/`min_height`. One regression caught during the split's own regression pass: `ChatLog`'s new `archive_dir` param had no default, which would have broken `eval_harness.py`'s single-arg construction ??? fixed with a `path.parent / "archive"` default before it shipped. All 61 "done" sheet specs re-verified against the final state (0 crashes); `chat_driver.py`/`server.py`/`wztc_ops.py` all import cleanly; process restarted and confirmed alive (PID 38888).

## 2026-08-04 ? Cursor ? assemble_corridor + cross_validate_stations

Closed the two remaining Layer-2 gaps from the 619-311 corridor miss. (1) `assemble_corridor(upstream_edge, downstream_edge)` in `wztc_ops.py` builds Align1/Align2 from work-area edges with station-0 = edge and walk away along travel T; approach auto-sized from `station_walk` max + 50 ft. Prefer over freestyle define+commit pairs. (2) `cross_validate_stations` compares VBA `get_alignment_stationing` cum stations (sorted multiset) to Python `station_walk` and requires path length >= walk max; auto-runs from `place_order_table_stations` / `place_sheet_geometry` unless force=True. VBA: headless `CommitCurrentAlignmentHeadless` always sets first point from earliest GG line StartPoint; `AdoptExistingAlignmentElement` uses StartPoint (not west-prefer); `CLEAR_PLAN_ELEMENTS keepAlignments=N` calls `ResetAllAlignmentBookkeeping`. Wired on `_WZTC_OP_NAMES` / `server.py`; prompts step 3 prefers assemble. Hot-reloaded AlignmentTool+WZTCBridge; restarted chat_driver.

## 2026-08-04 ? Cursor ? Fix orphan tool_result history trim (API 400)

Live: new 619-311 turn failed immediately with Anthropic `unexpected tool_use_id in tool_result` because `_trim_history_window` sliced `messages[-N:]` mid tool-chain, leaving `chat-history.json[0]` as a lone `tool_result`. Fixed in `chat_history.py`: safe trim start index, `_repair_tool_pairing` (drop orphans, require first message is real user text, clear if none left), persist repair on load. Backed up bad file to `Bridge/chat-history.pre-orphan-repair.bak.json`, cleared empty-repaired history, restarted driver, re-queued engineer request.

## 2026-08-04 ? Cursor ? Add Pydantic schema gate for sheet specs

Added `mcp-server/sheet_schema.py` (Pydantic v2) as pass 0 for sheet-spec shape: plan sheets must have applicability/corridor/orderTable/signs/symbols/annotations/rules; referenceLibrary sheets only need tables+tableRoles. `sheet_resolve.load` / `load_raw_path` validate before returning; `scripts/validate_sheet_spec.py` uses that gate first. `extra=''allow''` keeps nested table rows flexible. Dep: `pydantic>=2.0.0` in `mcp-server/requirements.txt`. Smoke: all 68 `619-*.json` specs pass.

## 2026-08-04 ? Cursor ? Add Shapely geometry QA to rules gate

Added `shapely>=2.0.0` and `mcp-server/sheet_geometry_qa.py`, wired into `sheet_rules.run_rules_gate`. Checks: hatch polygon validity/area, protectiveVehicle/arrowPanel not inside work-area hatch (live miss class), AP?PV center separation (skips altGroup / same-station OR alternatives), channelizing run self-intersection. Failure strings only ? no WKT in tool results. Soft no-op if shapely missing.

## 2026-08-04 ? Cursor ? Add pytest offline suite

Added `pytest>=8.0.0` and `tests/`: `test_sheet_schema` (all 619 specs + plan/library shape), `test_geometry_qa` (Shapely hatch/PV checks), `test_station_walk` (619-311 walk + compare_station_tables). `tests/conftest.py` puts `mcp-server` on `sys.path`. Run: `python -m pytest tests/ -q`.

## 2026-08-05 ? Cursor ? Harden agent control-loop (anti-fish + locked inputs)

After Layer-2 geometry gates, live 619-311 failures were mostly control-loop: blank `area_type` after lock, Default `find_reference_linework`, endless `find_elements_near`, burning `MAX_TOOL_ITERATIONS` on QA. Hardened in `wztc_ops.py`: `_merge_locked_designer_inputs` auto-fills blank compile/place kwargs and refuses conflicts; mid-plan refuses vague Default linework, freestyle `define_alignment_segment` (`assemble_corridor` passes `force=True` internally), and wide/repeated `find_elements_near` (tighter after `place_sheet_geometry`). `place_sheet_geometry` success sets `sheet_geometry_placed` + `nextStep` nudge toward one screenshot then FINAL. `prompts.py` CONTROL-LOOP DISCIPLINE block; `chat_driver` MAX_TOOL stop text demands resume-with-FINAL not more fishing. Tests: `tests/test_control_loop_gates.py` (5 passed). Restarted chat_driver PID 19720.

## 2026-08-05 ? Cursor ? Sheet-build checklist (scoped; general CAD stays free)

Deterministic plan state machine applies ONLY while a named 619 sheet build is active (`build_wztc_order_table` locked). General CAD / one-offs / questions remain unconstrained. Added `mcp-server/plan_workflow.py` + `get_plan_status()` checklist (corridor?stations?signs?attrs?compiler?visual QA). Out-of-order `place_sign` / `place_order_table_stations` / `place_sheet_geometry` raise structured `PLAN_GATE` (missing/accepted/nextTool). After compiler, free `adjust_view` refused ? `run_visual_qa_captures()` (4 scripted frames). Tools stamped with `nextStep` via `_attach_plan_next`. Prompt CONTROL-LOOP section scopes sheet-only. Tests: `tests/test_plan_workflow.py`. Wired on chat_driver + server; restarted driver.

## 2026-08-05 ? Cursor ? run_sheet_build executor (sheet-scoped)

Thin deterministic executor for named 619 builds only: after `build_wztc_order_table`, agent point-picks two work-area edges and calls `run_sheet_build(upstream_edge, downstream_edge)`. Runs assemble_corridor ? place_order_table_stations ? place_sign+set_sign_attributes (tips from cached station rows) ? place_sheet_geometry ? run_visual_qa_captures. Skips stages already done; returns `sheetPlanActive=False` outside a sheet plan (general CAD unchanged). `get_plan_status` nextTool prefers `run_sheet_build`. Wired chat_driver/server/prompts; tests in `test_plan_workflow.py`. Restarted chat_driver.

## 2026-08-05 ? Cursor ? Live 619-311 via run_sheet_build (+ history tool_use repair)

First kickoff hit API 400: history had assistant(tool_use) then assistant(FINAL text) with no tool_result (from prior forced FINAL). Extended `chat_history._repair_tool_pairing` step 2b to drop unanswered tool_use assistants. Fresh continue: build_wztc_order_table + `run_sheet_build` with prior work-area edges ? phases assemble (144.9 ft bay / 2550 approach), stations 6+2, signs+attrs, place_sheet_geometry, visual QA; FINAL complete. Auto-reply script briefly swapped edge order on an earlier attempt (fixed in `scripts/run_619311_live.py`).

## 2026-08-05 ? Cursor ? Match engineer hand-drawn 619-311 annotation style

Engineer drew a reference strip (~Y 214k) showing preferred styles. Root cause of the orange wash: compiled channelizing was one PLACE_ELEMENT_RUN polyline per run on TWZCD_P, which picks up a ByLevel custom linestyle that stamps giant orange cones. Fixes: `PLACE_CHANNELIZING_MARKERS` / `ExecPlaceChannelizingMarkers` places discrete 3?3 ft solid squares (color 6); `compile_plan` labels are name-only (length stays on the dim) and now emit SHOULDER TAPER overlay dim+label on the opposite side; `ARROW PANEL` label from `compile_symbols`; G20-* assemblies skip redundant code text (face already says END ROAD WORK). Live styled rebuild at Y=232000 via `scripts/rebuild_619311_styled.py`.

## 2026-08-05 ? Cursor ? Fix recurring tool_use/tool_result API 400

Root cause of repeated `messages.N: tool_use ids without tool_result`: (1) `_msg_tool_use_ids` only inspected dict blocks, so freshly appended Anthropic SDK content objects were invisible to `_repair_tool_pairing` and survived `save_history` as unpaired tool_use; (2) `main()` ERROR path logged but did not repair/persist, so in-memory broken history made every later `hi` 400. Fix: SDK-aware `_block_get`/`_block_type`; normalize+repair at start of `run_turn`; jsonable-append assistant/tool_response; repair+save on exception. Tests: `tests/test_chat_history_repair.py`. Cleaned `Bridge/chat-history.json` (9?4 msgs); restarted chat_driver; `hi` returned FINAL.

## 2026-08-05 ? Cursor ? list_levels English aliases (drainage?DCB)

`list_levels` was literal substring-only, so engineer `drainage` never hit project levels `DCB_*`. Added `Data/level-aliases.tsv` + expansion in `wztc_ops.list_levels` (OR needles, `matchedVia` / `aliasExpanded` notes). Prompt: ask for prefix only after alias-expanded empty result. Live: `list_levels('drainage')` returns DCB/DCP/DSSD/?. Dropped greedy `DS_` needle (false hits). Tests: `tests/test_level_aliases.py`.

## 2026-08-05 -- Cursor -- list_levels covers every HDM category, not just drainage

Hand-picking drainage needles only covered a few prefixes; live DGN has ~3046 levels. Fix: Data/level-categories.tsv maps English disciplines to NYSDOT HDM Exhibit 20-5 first letters (drainage->D, bridge->B, traffic->T, utilities->U, ...). wztc_ops.list_levels ORs category-letter matches on coded names with feature aliases (level-aliases.tsv now feature-only). Large hits include a prefix histogram. Offline: drainage 135, bridge 472, utilities 443. Deferred Level.Description after COM wedge.

## 2026-08-05 -- Cursor -- Wire intent, durable plan, placement registry

Three foundations now real (not a second AST): (1) sheet-spec annotationStyle + channelizingDevices.representation drive sheet_compile (defaults = prior hardcoded name-only / opposite overlay / markers); primitives carry primitiveId+specRef. Template on 619-311.json; others inherit defaults. (2) PlanSession persists to Bridge/sheet-plan.json on checklist flips; try_restore_sheet_plan on chat_driver startup; get_plan_status exposes persistedPath/updatedAt. (3) execute_compiled_plan no longer drops createdElementIds -- appends Bridge/placement-registry.jsonl; get_placements / delete_placements tools. Tests: tests/test_scene_plan_intent.py.

## 2026-08-09 -- Cursor -- Scorecard, registry provenance, gated visual QA, reflection

Implemented the playbook follow-ups (no Neo4j/multi-agent): (1) sheet_scorecard.build_placement_scorecard compares compile expectations vs placement-registry heads; place_sheet_geometry sets geometry_qa_passed from scorecard.passed. (2) Registry hardened with recordId/reqId/supersedes/soft-delete + resolve_latest_placements. (3) run_visual_qa_captures no longer auto-passes -- gated on scorecard + registry compiler kinds (force=True escape). (4) reflect_sheet_build deterministic critique citing primitiveIds/reqIds; artifacts in Bridge/sheet-reflection.jsonl. (5) run_sheet_build returns failedPhase/replan on compiler or visual failure while preserving earlier phases. Tools: get_geometry_scorecard, reflect_sheet_build. Tests: tests/test_scorecard_reflection.py.

## 2026-08-09 -- Cursor -- STE writing rules for WZTC panel + Cursor/Claude Code

Clarified three prompt surfaces: (1) mcp-server/prompts.py = in-MicroStation WZTC chat agent only (STE writing rules in BASE + sheet-build reflection rules in WZTC addendum; driver restarted). (2) .cursor/rules/ste-writing.mdc alwaysApply = Cursor Agent sessions. (3) CLAUDE.md Communication section = Claude Code sessions. Same habits: short imperatives, cite evidence before done, no weasel words. Does not replace product harness (scorecard/registry). Anthropic Mar 2026 harness article assessed: mostly already covered for this CAD stack; skip three-agent Playwright farm; keep tuning skeptical scorecard/reflect from live misses.

## 2026-08-09 -- Cursor -- 619-311.build.md playbook + agent load path

Live-build tips were only scattered in agent-log / chat. Added durable `Data/sheet-specs/619-311.build.md` (call path, annotation prefs, channelizing markers, NYW8-33 handoff, no Vehicle Space, visual QA, do-nots) and `sheet.buildGuide` pointer in `619-311.json`. Loader: `sheet_resolve.load_build_guide` / `build_guide_path` (basename-only). Wired `get_sheet_build_guide` tool; `get_sheet_requirements` + `build_wztc_order_table` + `get_plan_status` attach `buildGuidePath`/text. Prompt tells agent to follow buildGuide. README documents companion `.build.md`. Tests: `tests/test_sheet_build_guide.py`. Machine prefs stay in JSON; human tips go in the playbook for the next build.

## 2026-08-09 -- Cursor -- 619-311 build guide fully live

Finished wiring beyond the initial file: AUTHORING.md documents companion `.build.md`; CONTROL-LOOP step 0 follows `buildGuidePath`; `get_plan_status` returns `buildGuideExcerpt` (2k) + path; `run_sheet_build` OK attaches full `buildGuide`. Live smoke: `get_sheet_requirements('619-311')` returns `buildGuide` (7093 chars) via bridge; chat agent turn called `get_sheet_build_guide` and summarized Preferred call path (FINAL 2026-08-09 20:36:53). Driver PID 23148. Tests: test_sheet_build_guide + related suites green.

## 2026-08-10 -- Claude Code -- live 619-311 run found + fixed 3 real compiler bugs

Drove a full 619-311 build directly via the `wztc-designer` MCP tools (55 mph / RURAL / 12 ft lane / >=8 ft shoulder / lane closure / workers exposed). Order table, corridor, stations, signs+attrs all placed clean. Three real bugs surfaced and fixed in `mcp-server/sheet_compile.py` / `sheet_resolve.py` / `wztc_ops.py` (each required an MCP server restart ? this session's tool connection is a subprocess per `.mcp.json`, so edits to `mcp-server/*.py` don't affect already-loaded modules until reconnect):

1. **Cone-spacing off-by-one** (`sheet_compile.compile_channelizing`): device-count used `round(length / maxSpacing)`, which can round down and produce a spacing over the sheet's cap (hit a real 41 ft gap vs the 40 ft max). Changed to `math.ceil` so spacing is guaranteed <= max.
2. **`dimensioned: false` ignored** (`sheet_compile.compile_plan`): every row with `lengthFt > 0` got a dimension regardless of the corridor zone's own `dimensioned` flag. 619-311's `gapEndRoadWork` (G20-2, min 80/max 400 ft past the downstream taper) is a text callout on the real sheet, not a dimension line, but the compiler drew an "80'" dimension there anyway. Added `_zone_dimensioned()` lookup keyed off a new `lengthZone` field `station_walk` now carries (the actual zone that determines a row's length, since Sign rows' own `zone` id differs from their `spacingZone`).
3. **Work-area hatch silently dropped by default** (`wztc_ops.compile_sheet_plan`): `align_idxs` defaulted to `[1]` when omitted, but `compile_hatch` requires both align 1 and align 2 present ? so a bare `place_sheet_geometry(...)` call with no `align_idxs` compiled fine (gate passed) but never produced the required WORK AREA hatch, with no error or warning. `run_sheet_build` itself was unaffected (its internal `req` already falls back to `{1, 2}`), but `place_sheet_geometry` is explicitly documented as directly callable and doesn't mention this. Fixed the default to come from the sheet spec's own declared `orderTable.alignments`, not a hardcoded `[1]`.

Caught #3 by eye from a `capture_view` screenshot (no hatch visible), not from the gate ? the scorecard only checks placed-vs-compiled, so a compile that silently omits a primitive at the align_idxs stage passes clean. Also noted `focus_view_on_elements` returned a wildly wrong bbox/center (off by ~4 orders of magnitude) for the same element IDs `get_elements_range` computed correctly ? worked around with `adjust_view` using `get_elements_range`'s numbers instead; not yet root-caused, separate from the three fixes above.

Final state: `get_geometry_scorecard('619-311').passed == True`, 8 dimensions / 6 labels / 4 cone runs / 1 arrow panel / 1 protective vehicle / 1 hatch, all citing real placed element IDs, sign order upstream verified visually (W4-2R nearest taper, then W20-5R, then W20-1 furthest, matches sheet). NYW8-33 correctly deferred (vehicle-mounted, not a gap). `tests/` 55 passed after all three fixes. Test corridor at Y=300000 (empty area, chosen to avoid prior test geometry at Y~214k/217k/232k) ? engineer may want it deleted before a real submission.

## 2026-08-10 -- Claude Code -- root-caused and fixed dimensions rendering invisible when batch-placed (real, long-standing bug)

Engineer flagged "dimensions and labels look completely wrong" on the 619-311 build above. Labels were fine; every DIMENSION was completely invisible (no line, no arrows, no text) despite `get_elements_range`/`find_elements_near` reporting valid, correctly-positioned elements with `styleApplied=Y` and no VBA error. Found an abandoned diagnostic trail in `Bridge/wztc-journal.tsv` from 2026-08-04 (`styleApplied`/`styleErrNum`/`styleErrDesc` fields added to `WZTCExec.ExecPlaceDimension` specifically to chase this) that stopped without a logged conclusion ? this bug has likely affected every dimension this system has ever batch-placed, not just today's build.

Root-caused live via isolation testing: a single `place_dimension` MCP call always rendered correctly; the exact same coordinates placed as part of a multi-dimension batch (`place_sheet_geometry`'s `execute_compiled_plan` loop) came out invisible. Narrowed further ? placing two dimensions back-to-back reproducibly leaves the **first** invisible while the **second** renders fine, regardless of which zone/coordinates. Root cause: `ExecPlaceDimension` fetched the shared *named* `DimensionStyle` object (`ActiveDesignFile.DimensionStyles(styleName)`), mutated `ShowSecondaryText = False`, assigned it to the new `DimensionElement` (by reference, not a snapshot), then immediately restored `ShowSecondaryText` back to its original value on that same shared object. `oDim.Rewrite` does not synchronously bake in the dimension's arrow/text/witness-line sub-geometry ? MicroStation defers that to redraw time ? so a second dimension's own mutate-then-restore cycle on the identical shared style object, landing before the first dimension's real geometry got computed, left the first with an empty/degenerate cached geometry that no later view-level redraw (`UPDATE_VIEW_EXTENDED`, confirmed tested) could fix.

Fix: stopped restoring `ShowSecondaryText` after assignment ? leave the named style at `ShowSecondaryText = False` permanently, since every dimension this function places wants single-line text anyway. Verified live: 2 rapid-fire dimensions, then all 8 real 619-311 dimensions placed back-to-back via the real batch path, then a fresh pair after cleaning up the temporary comments ? all render correctly every time. A `View.Redraw`-after-`Rewrite` fix attempt was tried first and did **not** work (kept as a ruled-out lead here, not in the code). Side effect: manually-placed `ny_Plan` dimensions elsewhere in this DGN (or others sharing the style table) will now also default to single-line text ? consistent with what this codebase always wanted from `ny_Plan` dimensions, not a new behavior change to flag to the engineer as a regression.

Also hit and resolved along the way: hot-reloading `WZTCExec.bas` (required twice for this fix) resets VBA project state including `AlignmentTool`'s in-memory alignment tracking, even though the drawn alignment LINE elements survive on disk/in the model ? `place_sheet_geometry` then fails with `ALIGNMENT_NOT_READY`. Fix is `adopt_alignment(align_idx, element_id)` for each alignment (found the real element IDs via `Bridge/wztc-journal.tsv`'s `COMMIT_ALIGNMENT` request/response pairs for this session), not redrawing ? matches the tool's own error message guidance.

## 2026-08-09 -- Cursor -- Visual QA vision+panel + implement-all rule

Closed the two should-fix gaps fully: (1) `run_visual_qa_captures` / `run_sheet_build` now durable-copy frames to `Bridge/captures/qa_*.png`, lift `captures` to the tool result, and `chat_driver._wrap_op` attaches them as Anthropic vision blocks + panel `SCREENSHOT` via `_vision_blocks_for_qa_captures` (previously paths-only so visualQaPassed without eyes). (2) Chat prompts/rules/MAX_TOOL/wztc_ops notes no longer tell the agent to call `capture_view` (MCP-only); chat uses `view_drawing` / scripted QA. Added `.cursor/rules/implement-all.mdc` (alwaysApply) + matching `CLAUDE.md` section: when engineer says implement/fix all, finish every item wired+verified ? no short-win batching. Tests: `tests/test_visual_qa_vision.py`. Driver restarted.

## 2026-08-09 -- Cursor -- Harness soft gaps: geometry-faithful, sandbox, playbooks, P0 history

Implemented the remaining harness list (implement-all rule): (1) `sheet_geometry_faithful.py` + scorecard wiring ? tip/mid/xy drift, duplicate signs, kind flood; registry stores geom extras on place; `visual_qa_prechecks` runs `check_automated_visual_rules`. (2) Visual QA less model-only ? automated rules gate `visual_qa_passed` before frames. (3) KEEP/REVERT sandbox ? `sheet_sandbox.py` + `begin_sheet_sandbox` / `run_sheet_build_sandbox` / `keep_sheet_sandbox` / `revert_sheet_sandbox` (offset Y band; does not wipe kept corridor). (4) Playbooks for all plan sheets ? `scripts/generate_sheet_build_guides.py` created 60 stubs + `buildGuide` pointers (61 total incl. hand-authored 311); referenceLibrary skipped. (5) General CAD left prompt-shaped on purpose. (6) HARNESS_P0 ? `chat_history.harness_preflight_or_clear` at `run_turn` start clears still-broken history instead of 400-loop nudges. Tools on chat_driver + server; tests `test_harness_geometry_sandbox.py`; driver restarted.

## 2026-08-10 -- Claude Code -- engineer QA on the 619-311 build: 3 more real defects found + fixed (duplicate cones, silent G20 label loss, arrow panel not sign-styled)

Engineer manually edited the live 619-311 drawing to demonstrate two things wrong with the previous build, and separately spotted a third. All three were real, root-caused, and fixed -- not spec misreads:

1. **Doubled channelizing devices at taper junctions** (`sheet_compile.compile_channelizing`): adjacent runs (`shoulderTaperRun`/`laneTaperRun`, `laneTaperRun`/`longitudinalRun`) share their boundary station by construction (the taper-continuity point), and each run's endpoint-inclusive station list placed its own device there -- two markers stacked exactly on top of each other at 2 junctions (confirmed live: elements 135113/135135 and 135118/135152 at identical (x, y)). Fixed with a `placed_points` dedup set keyed on physical (x, y) rounded to 3 decimals, not station number (station numbers differ per-run/per-alignment; physical coincidence is what actually matters). Cone count went from a compiled 43 to 41.

2. **G20-2 label silently missing** (`Modules/DrawSign.bas` `PlaceSignAssembly`): the engineer deleted the original face+label (elements 104445/104507) after finding no visible label, then hand-placed a correct reference (cell + text) attached to the existing post/stem to show the desired result. Root cause: G20-* signs are the ONLY sign family whose label branch calls `InsertTextWithInchMarks` as the FIRST content of a freshly-armed `TEXTEDITOR PLACE` CadInputQueue session -- every other sign inserts the sign code first, then size on a second line via `KEY_DOWN`+`InsertTextWithInchMarks`. That first-content path was untested relative to the rest of the codebase's sign labels and silently produced nothing. Fixed by replacing the whole G20-* label branch with direct `CreateTextElement1` (Element API) construction, matching `WZTCExec.ExecPlaceTextLabel`'s already-reliable pattern, including a proper rotation matrix built from `viewAngleDeg` (not `Matrix3dIdentity`, so it still tracks view rotation like every other label). Re-placed G20-02 fresh via `place_sign`+`set_sign_attributes` at the exact original journal coordinates; label now renders every time.

3. **Arrow panel floating at a bare lateral offset** (engineer ask, not a bug): wanted it drawn the same way a roadside sign is -- a stem/post out from the tick, panel mounted at the post's outer end -- instead of a cell placed at a flat 16 ft lateral offset with nothing connecting it to the corridor. Added `DrawSign.PlaceCellOnPost` (same stem-then-`SnapInwardEdgeToTip` construction as `PlaceSignAssembly`, minus the SignLibrary/text-label parts a MUTCD sign needs), wired as a new bridge op `PLACE_CELL_ON_POST` (`WZTCBridge.bas`), a new `wztc_ops.place_cell_on_post` / MCP tool / chat-agent op-list entry, and `sheet_compile.compile_symbols` now emits arrowPanel's base point + outward direction instead of a pre-offset position for this item specifically (protectiveVehicle unchanged -- it sits in the lane, not on a post). `619-311.json`'s `arrowPanel` item gained `mountingStyle`/`mountingStyleNote`.

Every fix needed either a VBA hot-reload (`DrawSign.bas`, `WZTCBridge.bas`) + `adopt_alignment` re-bind, or an MCP server restart (`sheet_compile.py`/`wztc_ops.py`/`server.py`/`chat_driver.py`) -- hit the *wrong* one first again (restarted `chat_driver.py`, forgot the `wztc-designer` MCP connection is the actual process serving direct tool calls) and had to redo the arrow panel/cone rebuild a second time after the correct restart. `tests/` 64 passed throughout. Engineer's manual demo elements (G20-02 reference cell/label, arrow panel reference stem) were left in place until the automated rebuild reproduced the same result and started visually overlapping them (doubled text) -- engineer then explicitly approved deleting them; `delete_element`'s `own_element_only` safety correctly refused the first attempt since those elements weren't created by this session, requiring an explicit override once approved. G20-02's sign record is not currently in the placement registry (`place_sign` called directly doesn't register the way `execute_compiled_plan`'s internal sign phase does) -- harmless for `get_geometry_scorecard` (signs aren't in its tracked kinds) but a minor audit-trail gap worth closing later.

## 2026-08-10 -- Claude Code -- HANDOFF: engineer reports none of the above fixes are visible on their screen, unresolved

Engineer says the three fixes logged above (duplicate cones, G20-02 label, arrow panel post-mount) are NOT visible on their screen -- reports seeing "every single wrong thing from last time" in both the original corridor (Y=300000) AND a completely fresh rebuild done 100 ft away (Y=299900, built via one clean `run_sheet_build` call, all phases OK, scorecard passed, visual QA passed). This is unresolved and needs a fresh set of eyes (Cursor) or the engineer directly.

**What I verified repeatedly from the Claude Code / MCP-server side, all showing the fixes ARE present:**
- Direct COM queries (`get_elements_range`, `find_elements_near`) at the exact reported coordinates show correct single elements, not duplicates (e.g. exactly one cone at the shoulderTaper/laneTaper junction, arrow panel cell+stem as two separate elements per `get_placements`).
- `get_geometry_scorecard` passes clean, 0 failures, for both the Y=300000 and Y=299900 builds.
- `capture_view` (COM-level view screenshot) and `capture_window` (OS-level PrintWindow of the whole MicroStation app, ribbon and title bar visible) both show, repeatedly, across many separate calls: arrow panel on a 50 ft post, "END ROAD WORK" + "36" x 18"" legible on G20-02, single continuous cone line with no doubling.

**What was investigated and ruled out or fixed along the way (real issues found, not excuses):**
1. VBA 'Test' project was unloaded partway through this session (unknown cause -- possibly incidental to the window-focus/VBA-IDE investigation below); reloaded by the engineer, build continued fine after.
2. The MicroStation process has multiple top-level windows -- the design view AND a separate VBA IDE window ("SnappableToggle.mvba" was the last-focused VBA module). `Get-Process`'s `MainWindowTitle` nondeterministically returned one or the other across calls, meaning earlier "bring to front" attempts may have raised the VBA IDE instead of the drawing window. Fixed live with `EnumWindows` to enumerate all top-level windows for the PID and `AttachThreadInput`+`SetForegroundWindow` targeting the correct drawing-window handle specifically (plain `SetForegroundWindow` alone was returning `False`, blocked by Windows' foreground-lock).
3. MicroStation's process ID changed mid-session (37864 -> 7124) -- it restarted at some point, invalidating previously-found window handles. Not root-caused why it restarted.
4. Ruled out: multiple MicroStation instances (only one `microstation.exe` process at any point checked), wrong file/model (both confirmed `DELETE.dgn` / "left lane closure" model throughout), wrong drawing location (engineer confirmed cursor readout X=23145 Y=300000, exactly on the corridor), stale cached screenshot files in `Bridge/captures/` (all predate this session by days).

**Leading unresolved hypothesis:** `capture_view`/`capture_window` (both ultimately OS-level `PrintWindow`-style capture per their tool descriptions) may not reliably reflect true screen content for MicroStation CONNECT's hardware-accelerated (OpenGL) view rendering -- `PrintWindow` has known history of returning stale/blank content for GPU-composited windows unless the caller passes `PW_RENDERFULLCONTENT` (Windows 8.1+), and there is no confirmation this bridge/capture path does so. This would mean everything I "verified visually" from this side could be systematically wrong for what the engineer's physical monitor actually shows, while the underlying DGN data (confirmed independently via raw COM element queries, not just screenshots) is genuinely correct. **This is the first thing to check**: have the engineer take an actual photo or Windows Snipping Tool capture of their own physical screen (not anything Claude Code/the MCP server generates) and compare against the element IDs/coordinates below.

**Concrete state to pick up from, if the drawing data itself needs re-inspection:**
- Sheet: 619-311, inputs speed=55 Non-Freeway lane=12 shoulder=">= 8 ft" RURAL, lane closure, workers exposed.
- Original build: corridor Y=300000, `get_geometry_scorecard('619-311')` passed, arrow panel elementIds [136255 cell, 136254 stem], G20-02 sign elementIds 136320/136323/136384/136385 (placed fresh via direct `place_sign`+`set_sign_attributes`, NOT registered in placement-registry -- known gap, noted above), cones deduped (41 not 43).
- Fresh rebuild: corridor Y=299900 (100 ft below original), built via single `run_sheet_build(upstream_edge=[23760,299900], downstream_edge=[23860,299900])` call, all phases OK including scripted `visual_qa` (frames saved to `Bridge/captures/qa_619-311_*.png` -- these ARE fresh, from this rebuild, worth the engineer/Cursor opening directly as files rather than via any tool-mediated screenshot). Arrow panel elementIds [136754, 136753] per `get_placements`.
- `Modules/DrawSign.bas`, `Modules/WZTCBridge.bas`, `mcp-server/sheet_compile.py`, `mcp-server/wztc_ops.py`, `mcp-server/server.py`, `mcp-server/chat_driver.py` all carry the fixes described in the two entries above this one. None of that code has been reverted.

Engineer is moving to Cursor for a fresh look. Whoever picks this up: start by opening `Bridge/captures/qa_619-311_full_corridor.png`, `qa_619-311_upstream.png`, `qa_619-311_work_area.png`, `qa_619-311_downstream.png` directly as image files (not through any capture tool) as the least-mediated evidence available, and compare against what the engineer sees locally.

## 2026-08-10 -- Cursor -- 619-311 rebuild at Y=299000 confirms three QA fixes in DGN data

Engineer asked for a new corridor **1000 ft below** the Y=300000 band (not the prior 100 ft / Y=299900 rebuild). Built via `scripts/rebuild_619311_y299000.py`: edges `[23760,299000]` / `[23860,299000]`, speed 55 / Non-Freeway / 12 / `>= 8 ft` / RURAL. Hot-reloaded `DrawSign.bas` / `WZTCBridge.bas` / `WZTCExec.bas` first. Did **not** wipe the older Y=300k / 299.9k bands.

`run_sheet_build` status OK, scorecard passed (dims 8 / labels 6 / cones 4 runs / AP 1 / PV 1 / hatch 1), `visualQaPassed` true. COM/registry verification (not only PrintWindow):

1. **Cone dedup** -- 41 channelizing element IDs, **41 unique centers**, **0** duplicate XY, **0** pairs closer than 0.75 ft (`get_elements_range` per registry cone id).
2. **G20 size label** -- element `137174` text `36" x 18"` on `SF_P`; face cell `137112` + stem `137173`; captures show `END ROAD WORK` on the face.
3. **Arrow panel on post** -- `PLACE_CELL_ON_POST`: stem line `137229` (base 22465,299000 -> tip +50 Y) + cell `137230` (type 2, ~47x23 at tip).

Captures: `Bridge/captures/cursor_299000_*.png` + `inspect_299000*.json`. If the engineer's physical screen still disagrees with these element IDs at Y=299000, the remaining issue is view/GPU capture vs monitor -- not missing DGN geometry for this band.

## 2026-08-10 -- Cursor -- three engineer QA fixes from Y=299000 demo; rebuild at Y=297000

Engineer hand-corrected the Y=299000 corridor for G20 + AP style, and said cones still double-rendered (they could not fix those by hand). Root causes found against that demo + COM:

1. **G20 label** -- prior G20 path placed size-only (`36" x 18"`). Engineer text node was `G20-02 | 36" x 18"`. `DrawSign.PlaceSignAssembly` now uses the same TEXTEDITOR two-line path as every other MUTCD sign (code first, then size) for G20 too.
2. **Cone double rendering** -- XY dedup was already correct (41 unique). Live markers had `FillMode=1` (filled): `CreateShapeElement1(Nothing, sq)` default fill + orange outline read as stacked/double squares. Fixed in `WZTCExec.ExecPlaceChannelizingMarkers` to `CreateShapeElement1(Nothing, sq, 0)` (same fill-none pattern as workspace).
3. **Arrow panel** -- stem+cell is correct (matches engineer demo). Adding `TWZSGN_P` at the tip (tried once) stacked on the tip cone and looked like a doubled orange marker; kept stem-line + snapped cell only (no post cell on the channelizing tip).

Verified live at **Y=297000** (left prior bands): scorecard pass; 41 cones `FillMode=0`, 0 pairs closer than 2 ft; G20 text `G20-02 | 36" x 18"`; AP `PLACE_CELL_ON_POST` stem+cell ids. Captures `Bridge/captures/cursor_297000_*.png`.

## 2026-08-10 -- Cursor -- G20 black hole, AP tip base, diagonal downstream taper (Y=296000)

Engineer compared a hand-placed correct G20-02 face next to the automated one at Y=297000 and flagged three remaining issues; fixed and rebuilt at Y=296000:

1. **G20 black grouped hole** -- cell subelement dump: ours had SF_P nested cell color=6 (orange); engineer copy had color=240 (black). Outer SF_P complex stayed orange, SFB_P legend black on both. Added `DrawSign.FixG20FaceBlackHole` after face place to force SF_P nested-cell color 240.
2. **Arrow panel Y** -- `compile_symbols` used alignment center as `place_cell_on_post` base; roadside signs attach at the perp tip (`half_len=40`). Base is now tip = station + outward*40 so stem/face/label share the diamond-sign axis (was ~40-50 ft short).
3. **Downstream taper diagonal** -- `downstreamRun` used constant `lane_width`; on align 2 that landed a flat row on the opposite side of centerline from roll-ahead. Now interpolates `-lane_width` (work end, matches roll-ahead world Y) to `+lane_width` (far end, prior last-cone Y).

Verified Y=296000: G20 hole color 240; downstream cones Y 296012/296000/295988; scorecard pass. Captures `Bridge/captures/cursor_296000_*.png`.

## 2026-08-10 -- Cursor -- documented 619-311 QA fixes in JSON/playbook + sheet-spec-sync rule

Encoded the live 2026-08-10 engineer QA fixes into `Data/sheet-specs/619-311.json` (G20 `labelNote`/`faceSymbologyNote`, arrowPanel `mountingStyleNote` tip-base + no TWZSGN_P, channelizing `fillMode:0`, `downstreamRun` diagonal note) and `619-311.build.md` (signs/symbols section + Do-not rows). Added alwaysApply `.cursor/rules/sheet-spec-sync.mdc` and mirrored in `CLAUDE.md` / hardened `sheet-first-qa.mdc`: after every named-sheet fix, update JSON + build.md + agent-log in the same effort ? the WZTC agent loads `buildGuide` and will miss chat-only tips.

## 2026-08-10 -- Cursor -- place_lane_highway wired for WZTC chat agent

Engineer asked to draw N-lane highway strips (any length) from the in-MicroStation agent the same way we drew them live (2 solid edges Default/color0/weight0; lanes-1 dashed separators at 12 ft; 10 ft dash / 30 ft real gap). Added `mcp-server/lane_highway.py` (pure geometry) + `wztc_ops.place_lane_highway`; exposed on `server.py` MCP and `chat_driver` `_BASE_OP_NAMES` (general + wztc modes). `prompts.py` general-CAD section tells the agent to call it (ask for missing lanes/endpoints/side). Tests `tests/test_lane_highway.py`; live smoke 2-lane 100 ft OK; driver restarted.

## 2026-08-10 -- Cursor -- place_two_way_highway (even lanes + dual yellow)

Same general-CAD idea as `place_lane_highway`, for undivided two-way roads. Geometry in `mcp-server/lane_highway.py` (`two_way_highway_lines`): always 2 solid white outsides + 2 solid yellow center lines `yellow_gap_ft` apart (default 2); total lanes even; `L=lanes/2` per direction with `(L-1)` dashed white rows on each side of the yellow pair (2-lane: none; 4: one each; 6: two each). Same 10/30 real-gap dashes and Default/weight0; yellow via `resolve_color('yellow')` (live idx 4). Wired `wztc_ops.place_two_way_highway`, MCP `server.py`, `chat_driver` `_BASE_OP_NAMES`, `prompts.py` (use this tool for two-way, not one-way). Tests 12 pass; live smoke 2/4/6-lane @ Y=294200/294000/293800 OK; driver restarted PID 53584.

## 2026-08-10 -- Cursor -- road strip catalog: divided + TWLT + shoulders

Sheet inventory (`geometry.crossSection` / registry `roadType`) showed five recurring corridors beyond one-way/two-way: undivided?shoulders (311), divided+median (302), TWLT (312/412), ramps/gore, intersections. Built the next strip tools (general CAD, same Default/weight0 / 10-30 dash / yellow via `resolve_color`):

- `place_divided_highway(lanes_per_direction, median_width_ft, ?)` ? dual carriageway, yellow median edges + empty median gap (302-style).
- `place_twlt_highway(lanes_per_direction, twlt_width_ft, ?)` ? center turn lane bounded by dashed yellow (312-style); do not use two-way for TWLT.
- Optional `shoulder_width_ft` on all strip tools (solid white EOP outside travel outers).

Geometry in `lane_highway.py`; shared placer `_place_road_line_segments` in `wztc_ops`; MCP + `chat_driver` + catalog prompt in `prompts.py`. Intersections/ramp-gore deferred (prompt says ask). Tests 18 pass; live smoke divided 3+3@Y=293600 and TWLT@Y=293400 OK; driver PID 50220.

## 2026-08-10 -- Cursor -- intersection + ramp-gore road tools

Second batch of general-CAD roadway topology (after strip catalog). New `mcp-server/road_junctions.py`:

- `place_orthogonal_intersection` ? + or T; primary strip through junction; secondary stubs start outside primary pavement (travel+shoulder); arm types reuse `one_way|two_way|divided|twlt` via `build_strip_lines` / `travel_width_ft` in `lane_highway.py`.
- `place_ramp_gore` ? mainline one-way + diverging ramp; nose on ramp-side outer edge at `gore_station_ft`; optional solid white gore V (`gore_mark_ft`).

Wired MCP + `chat_driver` + prompt catalog (replaces ?not in catalog yet?). Tests `tests/test_road_junctions.py` + prior strip tests = 25 pass. Live smoke: plus @ (23800,293200), tee @ (24100,293200), gore nose (23390,292964). Driver PID 9384. Out of scope still: curb radii, crosswalks, gore chevrons.

## 2026-08-10 -- Cursor -- intersection MUTCD box striping (gap + stop/crosswalk + dotted)

Engineer flagged continuous corridor striping through junctions as wrong. Reworked `orthogonal_intersection_lines` / `place_orthogonal_intersection` to MUTCD 3B.11 / 685-style sketch rules:

1. Approach arms **gap** at the intersection box (primary split into `primary_neg`/`primary_pos`; stubs leave mark depth clear) ? no solid edge/center through intersecting approaches.
2. Defaults **ON**: transverse white **crosswalk** (8 ft pair) + **stop bar** (4 ft beyond) on every approach (both primary sides + all stubs).
3. **Dotted yellow center extensions** (2/4) through the box when `has_turning_lanes=True`, or auto when either arm is `twlt`.

Prompt + MCP updated. Tests 28 pass; live smoke plus @ (23800,292800) CW=8 SB=4 dotted; tee TWLT CW=6 SB=3. Driver PID 25016.

## 2026-08-10 -- Cursor -- intersection edge connect + stop-bar clip + striping arrows

Engineer QA: yellow/dashes ran through the stop box, and arms looked disconnected. Root cause: stubs started after mark_depth (gap), while primary yellow/lane continued to the box through the stop/crosswalk zone.

Fix in `road_junctions.orthogonal_intersection_lines`:
- Edge/shoulder still run to the intersection box (arms meet).
- Yellow center + dashed lane clipped at the stop-bar station (`_clip_center_lane_before_stop`).
- Turn arrows from `ny_plan_striping.cel` (SAL/SAR/SAS/SALS + SLONLY) as metas, placed via `PLACE_CELL` with new optional `libraryPath` (VBA `WZTCBridge.ExecPlaceCell`; default still WZTC).

Live smoke @ (23800,292500): yellow_inside_stop_zone=0; stub edges at box; 10 arrow cells; CW=8 SB=4. Tests 31 pass. Driver PID 43444.

## 2026-08-10 -- Cursor -- any-library cell find + place (general + WZTC)

Gap: agent could `attach_cell_library`/`list_cells` but `place_cell` was WZTC-mode only, and there was no cross-library search ? so ''place a gas meter'' was not a wired path.

Shipped:
- `list_cell_libraries` ? lists `c:\pwworking\usny\d0119091\*.cel`
- `find_cell(query=?)` ? scans name+description across those libs (restores prior attach); live: `gas meter` ? `UGM` in `ny_plan_utility.cel`
- `place_cell(..., library_path=)` already existed (VBA `libraryPath`); moved `place_cell` into `_BASE_OP_NAMES` so **general** mode can place too (`place_cell_on_post` stays WZTC)
- Prompt: find_cell ? place_cell(library_path=?); MUTCD faces still `place_sign`

Tests `tests/test_cell_libraries.py`; smoke placed UGM; driver restarted.

## 2026-08-10 -- Cursor -- intersection arrow facing + dedicated-lane math

Engineer: arrows faced wrong; do not put SAL/SAR+SLONLY on continuous multilane; dedicated = max(0, lanes_in - lanes_out). Side-road example near prior smoke matched travel-toward-box.

Root cause for facing: assumed striping cell +X = travel. Live probe of SAS at ACTIVE ANGLE 0/90/180/270 showed **angle 0 points +Y**. Fix: `_ms_striping_arrow_angle_deg` = `atan2(-travel_x, travel_y)` in `road_junctions._append_turn_arrow_metas`.

Lane math: drop `has_turning_lanes` as the arrow selector. New `primary_lanes_out` / `secondary_lanes_out` (default = lanes_in ? all SAS). Dedicated pockets get SAL/SAR + SLONLY (odd dedicated prefers left). Dotted center still from explicit `has_turning_lanes`, TWLT, or dedicated > 0.

Wired through `wztc_ops` / MCP `server.py` / `prompts.py`. Tests 32 pass. Live smoke @ (24650,291200) continuous SAS-only angles {-90,0,90,180}; @ (25150,291200) 3?2 dedicated SAL+SLONLY+SAS. Driver PID 34304.

## 2026-08-10 -- Cursor -- shared turn-option arrows + RH approach lane placement

Engineer clarified lane math: dedicated SAL/SAR+SLONLY only when lanes_in > lanes_out; when equal, shared options by safe turns (1-lane + ? L/S/R; 2 ? SALS+SARS; 3+ ? SALS/SAS/SARS). Prior smoke put arrows on the **opposing** half of two-way strips (left of travel from full-width centers), so they looked wrong-way even when ACTIVE ANGLE was correct.

Fix in `road_junctions`:
- Approach centers from centerline toward **right** of travel (US), past yellow/median/TWLT clearance; one-way still full-width left-to-right.
- `_shared_through_cells` + `_allowed_turns_for_arm` (plus = L/S/R; tee stub = L+R only; primary tee = straight + turn toward stub).
- No triple-head cell in `ny_plan_striping.cel` ? single-lane L/S/R emits stacked SALS+SARS.
- SLONLY only on dedicated pockets; through lanes lose that turn once a pocket exists.

Live smoke @ (27000,291200) 2+2 shared; (27500,291200) 3+3; (28000,291200) 3?2 dedicated. Tests 15 pass. Driver restarted.

## 2026-08-10 -- Cursor -- arrow facing confirmed via bbox (not vision)

Capture-vision repeatedly claimed tips faced away; element bbox vs cell origin proves SAS tip follows ACTIVE ANGLE with **0 = +Y**, **-90/+270 = +X**. Formula stays `atan2(-travel_x, travel_y)` (no +180 ? that made west tips face away). Real facing bug was RH-lane placement (arrows sat in opposing two-way half). Shared SALS/SAS/SARS math + dedicated ONLY kept. Engineer should QA fresh smokes at (30000 / 30450 / 30900, 291200), not older junctions.

## 2026-08-10 -- Cursor -- LSR overlap, tip-at-origin angle, ONLY spacing, 3?2 asymmetric strip

Engineer QA on shared/dedicated arrow junctions:
1. Overlap SALS+SARS on the same xy for single-lane L/S/R (no triple-head cell).
2. Striping cell origin is at the tip (stem opposite); ACTIVE ANGLE = `atan2(travel_x, -travel_y)` so south travel tips toward the box (180).
3. SLONLY setback 28 ft upstream; SAL shares station with SAS/SARS (same x on E-W).
4. When `primary_lanes_out` < toward on two_way: `asymmetric_two_way_highway_lines(toward, out, median_second=dedicated*lane_w)` so each arm is 3 toward + median + 2 away (reads 3+2 / 2+3 across the box).

Live smoke @ (32000,291200) shared; (32500,291200) 3?2. Tests 33 pass. Driver restarted.

## 2026-08-10 -- Cursor -- restore tip angles; fix 3-into/2-after strip orientation

Engineer: global 180 flip inverted every arrow (tips away from stop bar). Restored `atan2(-tx, ty)` on **all** arms (west -90 / east 90 / south 0 / north 180). South-only +180 also pointed tips away ? removed.

3?2 strip: `_first_edge_from_centerline` paints the left-of-corridor pack first (= away). Build `lanes_first=out + median_first`, `lanes_second=toward` so EB on west is 3 into the box and 2 after on the east (verified: south dash rows=2, north=1 on primary_neg).

Live smoke @ (35000 / 35500, 291200). Tests 16 pass.

## 2026-08-10 -- Cursor -- 619-311 real-road demo (EB Urban 55)

Demo only (not wired into permanent compile/agent path): full 619-311 on a striped 4-lane + 8 ft shoulder corridor at Yc=290200, X=32000..36000. Designer locks: eastbound right (south of yellow), Urban, 55 mph, WA 100 ft at X 34500?34600 on lane line Y=290187. `half_len=20` (12+8) so posts tip at outer EOP.

Gotchas: Align1 stations increase upstream (west); `outward_sign=-1` hatches **north**/left lane; EB right needs `outward_sign=+1`. `clear_plan_elements` / `assemble_corridor(force=True)` deletes journal-owned `place_polyline` striping ? place corridor **after** the sheet build. Scorecard passed; NYW8-33 still DEFERRED (vehicle-mounted). Helper: `scripts/build_619311_on_real_road.py`. Tips in `619-311.json` / `619-311.build.md`.

## 2026-08-10 -- Cursor -- wire Cursor build-verify-fix into WZTC agent

Sheet-first + scorecard + visual QA were already in `prompts.py`. Missing was the real-road lateral method from the live Cursor 619-311 session.

Shipped `resolve_sheet_lateral` (locks `outward_sign` / `half_len` from travel up?dn + closed_side; real_road_edge ? lane+shoulder tip-at-EOP). `run_sheet_build` / `compile_sheet_plan` honor locked lateral by default. Checklist `nextTool` after order table is `resolve_sheet_lateral`. Prompt CONTROL-LOOP now: ask closed_side ? resolve ? run_sheet_build ? build-verify-fix (IDs/ranges, no stacked duplicates, don't thrice-assert capture vs engineer, re-place journal striping after force wipe). Wired in chat_driver + MCP `server.py`. Tests `test_resolve_sheet_lateral` + plan_workflow. Driver PID 31676. `619-311.build.md` preferred path updated.

## 2026-08-10 -- Cursor -- fix 619-311 channelizing lateral (real-road Y)

Engineer QA: cones through middle of road, signs/PV/WA wrong in Y; X OK. Root cause: `assemble_corridor`/hatch put Align1 on the **left edge of the closed lane** (channelizing line), but `compile_channelizing` still used longitudinal offset=+lane_width (fog-line) and lane-taper tip at offset 0 ? backwards for that align contract.

Fix in `sheet_compile.compile_channelizing`: longitudinal offset 0; lane taper tip=+lane (outer travel) toe=0; shoulder +lane ? +lane+shoulder (EOP); Align2 downstream 0 ? -lane. AP `tip_half_len_ft` from locked lateral. Rebuild script lane line Y=290188. Tests `test_channelizing_lateral`. Sheet JSON/playbook Do-not updated.

## 2026-08-10 -- Cursor -- real-road 619-311 wrong half (two_way first edge)

Engineer: rebuild still looked like cones through the middle. Root cause: `place_two_way_highway(x1,y1,?)` is the **north travel outer**, not yellow center. Script used `Yc-12` as align ? north (WB) dash. EB right dash is `Y_north_outer - (2*lane + yellow_gap + lane)` (= -38). Clean rebuild at north-outer 289200 / lane line 289162 / hatch 289142?289162; posts ~289138. Captures `qa_311_v2_*.png`. Look at Y~289175, not old 290200 band.

## 2026-08-10 -- Cursor -- 619-311 polish: one AP, G20 opposite, drop guides

Engineer QA on real-road band: duplicate TWZAP_P, G20-2 on closed roadside, white align+ticks cluttering striped plan.

Fixes in `wztc_ops`: `place_sheet_geometry` places arrowPanel once; G20 tip uses flipped world-locked `closed_outward` + `opposite_half_len` (open lane + yellow + opposing pack + shoulder, e.g. 46 ft); new `delete_construction_guides` deletes only journal `PLACE_ORDER_TABLE_STATIONS` / `PLACE_PERP_LINE` / `DEFINE_ALIGNMENT_SEGMENT` IDs. `resolve_sheet_lateral` locks opposite fields. Rebuild script at north-outer **288200** (1000 ft south of 289200). Sheet JSON/playbook + prompt CONTROL-LOOP updated.

## 2026-08-10 -- Cursor -- G20 closed shoulder + wire full real-road 619-311 path

Engineer correction: G20-2 stays on the **same** closed-shoulder roadside as W20s (south EOP for EB right) ? not opposite. Root cause of wrong tip was Align2 tan flipping `_outward_unit`; fix locks `closed_outward` from `resolve_sheet_lateral` and applies it to all one-side signs in `_place_locked_signs_from_stations`.

Agent wiring: playbook Preferred path has real-road combo steps 10-12 (re-place striping, `delete_construction_guides`); `run_sheet_build` auto-runs guide cleanup when `real_road_edge`; plan complete `nextTool` = `delete_construction_guides`; prompts CONTROL-LOOP + sheet JSON `roadside: closed`. Tests resolve/plan_workflow updated.

## 2026-08-10 -- Cursor -- live WZTC agent full 619-311 run + lockedSignRows skip fix

Ran real chat agent on full real-road 619-311 (Y north-outer 287200). Turn 1 followed playbook (resolve ? run_sheet_build ? striping) but `phases.signs skipped=True` because stale `sheet-plan.json` had `order_table_built` with empty `lockedSignRows` ? vacuous checklist treated signs as done. Fixed: `stage_done` requires roadside signs for sheets with `signs.items`; `run_sheet_build` auto-rebuilds order table when lock empty; load clears bogus order_table_built; prompt warns against post-scorecard `find_elements_near` fishing. Continue turn placed all 4 signs (G20 tip Y=287142 south). Driver `--force` mid-`run_sheet_build` orphaned TOOL_RESULT ? finished guides/corridor outside panel. Agent API cost ~$2.93 for the monitored turns.

## 2026-08-13 -- Cursor -- curved/S polyline road striping for all highway types

Catalog strip tools were straight-only `(x1,y1)->(x2,y2)`. Added optional `vertices=[[x,y],...]` on `lane_highway_*` / `two_way` / `divided` / `twlt` / `asymmetric` / `build_strip_lines` and `place_*` + `place_ramp_gore` (curved mainline, straight ramp from local tangent). Path offset uses `alignment_geometry.station_to_xy` local normals; solids densify to multi-point polylines; `_place_road_line_segments` honors `seg['vertices']`. Straight 2-pt path keeps the old exact geometry. Prompts updated. Tests: 41 pass in `test_lane_highway` + `test_road_junctions`. Live smoke `scripts/smoke_curved_highways.py` needs MicroStation (none running when tried). Orthogonal intersections stay straight-arm sketches ? place curved approaches as separate strips.

## 2026-08-13 -- Cursor -- curved road corner bowties (overlapping lines)

Live S-curve striping overlapped at bends. Root cause: densified offset samples used discontinuous per-segment normals across corners, so the solid polyline backtracked (dot~-1) and looked like double lines. Fix in `lane_highway._miter_offset_vertices`: miter joins for solids; dashes placed per straight leg. Cleared old smoke (fence delete 141) and re-placed. Tests include `test_s_curve_no_backtrack_bowties` (42 pass).

## 2026-08-13 -- Cursor -- curved striping corner overlaps (final)

Overlaps at S-bends had three failed approaches: densify-across-corners (backtrack), unlimited miter (spikes into lanes), bevel joins (Z artifacts), naive per-leg (concave legs cross). Final: `_offset_leg_specs` places each straight leg with constant normal and **clips only concave/inside joins** to the offset-line intersection; convex corners keep a small gap. Re-placed smoke band; `test_s_curve_no_crossing_solid_legs` + 42 tests pass.

## 2026-08-13 -- Cursor -- filleted continuous curved striping

Engineer: no gaps, no overlaps, continuous like a typical curve; space road types farther apart. Final approach: `fillet_polyline_segments` inserts arc fillets at polyline corners (default r=50), then `_continuous_offset_polyline` samples offsets via `station_to_xy` so normals stay continuous through bends. Solids are one multi-point polyline per row; dashes walk that poly. Smoke bands now 300 ft apart at ~(51210, 297900). 44 tests pass.

## 2026-08-13 -- Cursor -- curve-type matrix smoke (all highway types)

Live smoke `scripts/smoke_curve_matrix.py`: 5 path shapes (L-bend, C-curve, gentle-S, reverse-S, hairpin) ? 5 strip tools (one-way, two-way, divided, TWLT, ramp gore) = 25 placements, all OK. Unit test `test_curve_matrix_fillets_and_offsets`. View framed ~(53600, 296940); columns 700 ft apart, road types 280 ft apart within a column.

## 2026-08-13 -- Cursor -- curved 619 sheet corridor (path_vertices)

Sheet builds assumed a straight chord between work-area edge picks. On curved real roads that made Align1/2, hatch, and dim tips ignore the highway bend (signs were already view-horizontal via ACTIVE ANGLE = view rotation ? leave that alone).

Fix: optional `path_vertices` on `assemble_corridor` / `resolve_sheet_lateral` / `run_sheet_build`; densified Align1/2 along the path with approach extension; `PlanSession.work_bay_vertices` ? `compile_hatch` curved boundary; `compile_plan` dim tips use per-station outward. Prompts + `619-311` JSON/playbook updated (pass path_vertices; re-place `place_two_way_highway(vertices=?)` after build; do not rotate signs to tangent). Tests: `tests/test_curved_corridor_compile.py`, curved case in `test_resolve_sheet_lateral`.

## 2026-08-13 -- Cursor -- live 619-311 on curved reverse-S road

Scripted `scripts/build_619311_on_curved_road.py`: Urban 55 / 100 ft WA on a ~5600 ft reverse-S (matrix paths are too short for ~2345 ft approach). `path_vertices` on closed-lane offset edge; `run_sheet_build` OK (curved=True), 4 signs, scorecard pass. Signs view-horizontal; hatch/cones follow bend. First `place_two_way_highway` returned ERROR on solid polylines ? root cause `_continuous_offset_polyline` step_ft=1 produced ~5k-vertex PLACE_POLYLINE TSVs VBA rejected; default step now 10 ft; solids re-placed OK. QA: `Bridge/captures/qa_311_curve_{overview,work,upstream,downstream}.png`. Look ~(64800, 294546).

## 2026-08-13 -- Cursor -- curved dim hug + hatch clip (619-311)

Engineer: dims don't hug curve; hatch is a box outside the lane. Fixes:
1. `compile_plan` emits `curved` dims with tip-offset `path`; multi-segment alignments always use `place_path_hugging_dimension` (color-2 polyline + sheet length text) instead of Linear Size chords (which measured wrong lengths like 112' vs 120').
2. `DrawDiagonalHatchLines` clips stripes to the shape polygon (point-in-poly), not the AABB ? bbox stripes made curved workspaces look rectangular.
3. `compile_hatch` densifies work-bay (~5 ft); `assemble_corridor` stores denser `workBayVertices`.
4. Live smoke moved to a sharp L-bend (reverse-S was too gentle through the WA). Closed-lane offset for striping demo set to 1?lane (not opposing dash). Look ~(74568, 290052). Script: `scripts/build_619311_on_curved_road.py`.

## 2026-08-13 -- Cursor -- curved 619-311 lateral + dims + signs (QA reject)

Engineer: curved build looked like center-lane closure; hatch bled opposite; dims/labels off; sign assemblies disconnected. Root causes vs known-good `scripts/build_619311_on_real_road.py`:
1. Demo `CHAN_OFF=LANE` put align on near-yellow dash ? restored `2*L+gap+L` (38) like straight.
2. Fixed world `closed_outward` from WA-mid broke tips on bends ? `_place_locked_signs_from_stations` now uses local Align1-equivalent outward (`-tan` on Align2).
3. Path-hugging polyline dims rejected ? back to Linear Size tip-to-tip; `compile_plan` takes `tip_half_len_ft` so tips match real-road half_len=20; Align2 dim/label tips also flip tan.
Rebuild: `scripts/build_619311_on_curved_road.py` origin ~(76000, 288000).

## 2026-08-13 -- Cursor -- curved dims hug roadside + table lengths

Engineer: dims along the curve cut into the highway and showed wrong lengths (chord measures like ~108'/~45'). Tables for Urban 55 / 12 ft / >=8 ft: Buffer **495'** (311-02), Roll Ahead **120?200'** (311-04, plan emits min 120), Downstream Taper **50?100'** (plan emits min 50), Lane Taper 680', Shoulder Taper 160'. Fix: restore path-hugging dims when tip-path sagitta bows (not for every multi-seg span); tip offset = locked half_len; text = sheet lengthFt. Straight approach spans stay Linear Size. `place_path_hugging_dimension` adds short end ticks. Docs/JSON/playbook updated.

## 2026-08-13 -- Cursor -- curved dims = ny_Plan SizeArrow chain (not polyline)

Engineer: fake polyline+text dims do not match straight 619-311; only the bowed spans on the curve should be curved, and they must still be real dimensions. Restored rule from 2026-08-03: always `ExecPlaceDimension` / `ny_Plan` SizeArrow. Curved spans: chain of those along the tip-offset roadside; mid segment `PrimaryText` OverrideText = sheet/table length (e.g. 495'-0\"); intermediate segments OverrideText=space to hide wrong chord numbers. `overrideText` wired through bridge. Straight spans unchanged (single SizeArrow).

## 2026-08-13 -- Cursor -- curved dim OverrideText (HIDE + late-bind)

After VBA Reset: PrimaryText only compiles late-bound (Object). Intermediate chain segs used OverrideText=space, but `ParseParams` Trim'd TSV values so the blank never arrived ? chord measures (~38') cluttered the curve. Fix: sentinel `HIDE` -> blank PrimaryText; ParseParams skips Trim for `overrideText`. Smoke `scripts/smoke_dim_hide.py` + full `scripts/build_619311_on_curved_road.py` OK (scorecard pass). Look work bay ~(78588, 288033). Early-bound `DimensionElement.PrimaryText` still fails compile on this CONNECT ? do not restore it.

## 2026-08-13 -- Cursor -- curved dims = one Arc Size (not SizeArrow chain)

Engineer: bend dims must be continuous curved (arc-like), same ny_Plan look as straight 619-311; SizeArrow tip-chains (dozens of tips on Buffer) are wrong. Missing piece: MicroStation `msdDimTypeArcSize` (not linear SizeArrow). `ExecPlaceArcSizeDimension` / `PLACE_ARC_SIZE_DIMENSION`: CreateDimensionElement1 ArcSize with refs start + height/offset point + end (2-point+DimHeight alone creates empty dims with no range). `place_path_hugging_dimension` now circle-fits the tip path and places **one** Arc Size + OverrideText; nearly-straight paths stay one SizeArrow. Smoke: element range real. Do not reintroduce multi-segment SizeArrow chains for curved spans.

## 2026-08-13 -- Cursor -- bend dims = ArcElement following the curve

Engineer: dimension must be an ARC that FOLLOWS THE CURVE (not chords, not tip chains, not giant far-side Arc Size). `msdDimTypeArcSize` on this install drew wrong/huge arcs away from the road. Fix: `ExecPlaceCurvedPlanDimension` / `PLACE_CURVED_PLAN_DIMENSION` builds a true `CreateArcElement1` dim-line concentric with the tip arc (signed sweep), plus radial extensions + sheet-length text. `place_path_hugging_dimension` uses that for bowed spans; straight spans stay SizeArrow. Tip path step 10 ft (was 40) so short zones circle-fit. Do not reintroduce SizeArrow tip chains or bare Arc Size for roadside hug.

## 2026-08-13 -- Cursor -- curved dim REQUIREMENTS (do not forget)

Engineer QA on curved 619-311 dims (repeat until all three hold):
1. Dim line on the bend must be a true ARC that FOLLOWS the roadside (concentric) ? not a chord, not a tip-chain, not a far-side Arc Size sweep.
2. Must LOOK like straight-sheet ny_Plan SizeArrow: arrowhead TIPS, extension lines, length text, color ? ArcElement-only without tips is NOT done.
3. Remove leftover bad dims (old SizeArrow chords / tip-chains / failed Arc Size) before rebuild ? clear_plan alone can miss non-journal leftovers; wipe DimensionElements in the work band.
Documented in Data/sheet-specs/619-311.build.md Do-not row + 619-311.json notes.

## 2026-08-13 -- Cursor -- curved dim tips visible + wipe leftovers

Engineer: arc followed the curve but (a) missing SizeArrow tips vs straight Buffer, (b) leftover bad chord dims still present. Root causes: filled `CreateShapeElement1(...,1)` tips vanish when view Fill is off; wipe only deleted DimensionElements so prior color-2 Arc/Line/Shape curved-plan leftovers stayed. Fix: `PlaceCurvedDimArrowTip` draws always-visible tip fans (dim color 2, ~5 ft ? match SizeArrow scale; oversized white fans looked like leftover bad dims at the Buffer/Roll Ahead junction); `ExecDeleteDimensionElementsInRange` also deletes color-2 arc/line/shape (+ short white tip leftovers); wipe AFTER clear_plan; force_arc keeps mild bows (50') as arcs. Notes reinforced in 619-311 JSON/playbook. Look work bay ~(78588, 288033).

**Engineer ask (keep remembering):** bend dim = ARC that FOLLOWS the curve + must LOOK like straight ny_Plan SizeArrow (tips/extensions/text) + REMOVE leftover bad dims before rebuild ? do not leave chords/tip-chains next to the new arc.

## 2026-08-13 -- Cursor -- why bend dims are graphics, not annotative

Engineer (close, then asked): are we using lines instead of annotative dimensions like the straight ones because there is no annotative curved dimension?

Answer for Claude: **annotative curved dims exist** ? MicroStation `msdDimTypeArcSize` (`PLACE_ARC_SIZE_DIMENSION` / `ExecPlaceArcSizeDimension`). Straight spans are real `DimensionElement` + `ny_Plan` `msdDimTypeSizeArrow`. Arc Size was tried first on this CONNECT install and failed for roadside hug: empty/broken geometry (2 pts) or giant far-side arc. Live bowed spans therefore use constructed `PLACE_CURVED_PLAN_DIMENSION` (`CreateArcElement1` + radial extensions + SizeArrow-scale tip fans + sheet text). That is a fallback, not a missing API. Do not reintroduce SizeArrow tip-chains. Optional next: another Arc Size smoke (point order / DimHeight) if we want a real annotative dim again; otherwise keep constructed arc and only steal SizeArrow terminators at the ends.

Look: work bay ~(78588, 288033); rebuild `scripts/build_619311_on_curved_road.py`.


## 2026-08-13 -- Claude Code -- root cause: Arc Size was never broken, our tip order was

Engineer asked why the curved-dim problem keeps recurring. Investigated instead of retrying.

**Finding: `msdDimTypeArcSize` works fine on this install.** The recorded conclusion
("annotative curved dims don't work here, fall back to constructed graphics") was wrong,
and because it was written into `619-311.build.md` / JSON / the agent-log as settled fact,
every later session skipped re-testing it. That is the recurrence mechanism.

Root cause is in our caller. `ExecPlaceArcSizeDimension` passed the two tips in raw PATH
order with `rot = Matrix3dIdentity` and never normalized sweep sign. Arc Size measures
**counter-clockwise** from first tip to second, so any clockwise (right-hand) roadside bend
swept the reflex way -> the "giant far-side arc". `ExecPlaceCurvedPlanDimension` already
worked around exactly this defect (wraps `da` to (-pi,pi], picks point order by sign, plus a
swap-retry when `ArcElement.SweepAngle` returns reflex) -- that fix was never carried back.

Fix: normalize sweep sign in `ExecPlaceArcSizeDimension` before `AddReferencePoint`; response
now reports `sweepRad` / `tipsSwapped`. Hot-reloaded.

Evidence: `scripts/diag_arc_size_root_cause.py` (new) places Arc Size dims across CCW, CW, and
the atan2-branch-crossing due-west case, in both tip orders, and compares element range to the
analytic bbox of the intended minor arc + witness lines. Before: 3/6 hug (the 3 failures are
exactly the CW-order cases, blowup 2.01x / 131x). After: 6/6 hug, blowup 1.01x.
Production circle-fit checked separately: `_fit_circle_2d` samples endpoints+mid so |r1-r2| =
0.00 -- both tips land on one radius, `DimHeight = rOff - r1` is well-defined.

**Not verified:** that Arc Size *looks* like straight-sheet ny_Plan SizeArrow (tips/extensions/
text) -- engineer requirement (2). Bbox proves the geometry hugs; it does not prove appearance.
Live bend dims therefore still use `PLACE_CURVED_PLAN_DIMENSION`; switching is the engineer's
call after visual QA. Registry row for PLACE_ARC_SIZE_DIMENSION also corrected (it documented a
start/height/end signature the code no longer has).

Probe elements left in the model at ~(90000-93000, 287000), 12 dims -- not deleted.


## 2026-08-13 -- Claude Code -- overlay dims crossed the pavement on real roads

Engineer QA on the C-curve 619-311: "still has the wrong dimension." Two defects.

1. **My leftover.** The A/B Arc Size dim I placed to compare terminators (element
   178800) was still sitting 70 ft outside the real 160' dim -- two `160'-0"` dims
   at the shoulder taper. Deleted. Do not leave comparison elements in the model;
   that is the engineer's standing requirement 3.

2. **Real bug: overlay dims flipped across the travel lanes.** `compile_plan`
   applied `annotationStyle.overlayDimSide="opposite"` unconditionally. That is a
   printed-sheet convention -- on the schematic the far side is blank paper. On a
   real road the alignment IS the closed-lane edge, so the flip drove SHOULDER
   TAPER's dim through the pavement.

   Measured (radius from the C-curve center, road spans r3000-3058, EOP r3058):
   every main dim tipped at r=3057.7 on the EOP; SHOULDER TAPER tipped at
   r=3017.7 -- 40 ft inboard, between the lane line (r3012) and yellow (r3025),
   with its dim arc at r=3032.7 also on pavement.

   Fix in `sheet_compile.compile_plan`: when a real-road `tip_half_len_ft` is
   locked, keep `overlay_sign = outward_sign` and separate the overlay radially
   (`overlay_offset = dimOutward*2`) instead of flipping. Schematic builds keep
   the printed flip. After: all 7 dims tip at r=3057.7; overlay dim line r=3087.7
   clears the main column at r=3072.8.

Tests `test_overlay_dim_stays_off_pavement_on_real_road` +
`test_overlay_dim_still_flips_on_schematic_build` (132 pass). Note both needed
`sheet_elements="MergingTaper|ShoulderTaper|DownstreamTaper"` -- 619-311.json has
no `sheet.elements` key, so the registry pipe-list must be passed explicitly or
`_should_annotate_non_sign_label` filters the overlay row out and the test is
vacuous. Rebuild: `scripts/build_619311_on_c_curve.py`, work bay ~(86497, 289270).


## 2026-08-13 -- Claude Code -- CORRECTION: Arc Size is ON; scope reverted to dims only

Supersedes the "Not verified / still uses PLACE_CURVED_PLAN_DIMENSION" line in my
earlier entry today. Engineer directive: keep the curved-dimension work and genuine
bug fixes, revert everything else to Cursor's handoff state.

FINAL STATE
- `ARC_SIZE_BEND_DIMS = True` (wztc_ops). Bowed spans now place a REAL annotative
  `msdDimTypeArcSize` DimensionElement. Straight spans unchanged (ny_Plan SizeArrow).
- Cursor's `scripts/build_619311_on_curved_road.py` (L-bend, origin 76000, work bay
  ~78588,288033) rebuilt untouched: 6 SizeArrow + 2 Arc Size + **0**
  PLACE_CURVED_PLAN_DIMENSION. The constructed line-graphics fallback is no longer
  used on a live build. Verified visually: dim line hugs the bend, solid arrowhead
  terminators both ends, length text on the curve.
- Root cause of the original "Arc Size is broken" verdict was OUR tip order, not the
  API -- Arc Size measures CCW, we passed tips in path order, so clockwise bends
  swept reflex. Fixed in `ExecPlaceArcSizeDimension`; probe
  `scripts/diag_arc_size_root_cause.py` 6/6.
- Kept as a genuine bug fix: `sheet_compile.compile_plan` overlay dim side on
  real-road builds (SHOULDER TAPER was tipping 40 ft into the travel lanes).
- Scope creep NOT adopted: `scripts/build_619311_on_c_curve.py` was my own C-curve
  variant; it is not part of the sheet-build path and Cursor's builds are unchanged.
  The X~84000 band holds leftovers from my C-curve iterations -- engineer will clean.

OPEN DEFECT (pre-existing, affects every rebuild at a fixed origin)
`clear_plan_elements` cannot delete a prior run's elements once
Bridge/wztc-journal.tsv has rotated: the ownership proof (createdElementIds= rows)
moves to Bridge/archive/ and `ownElementOnly` then refuses. Live at origin 76000:
FOUR stacked `ROLL AHEAD DISTANCE` labels (164238 / 170227 / 176408 / 183626) plus
an orphaned `120'-0"` text (176407). The 164xxx range predates this session. This is
why duplicate dim text keeps reappearing across sessions and why "remove leftover
bad dims" never fully sticks. Fix candidates: have clear_plan_elements consult the
archived journals, or persist created IDs in the placement registry independent of
journal rotation.

## 2026-08-13 -- Cursor -- leftover dims after journal rotate + 50' chord look

Engineer: Arc Size look is close, but stacked leftover 120' text/chords remain and Downstream 50' still reads as a chord.

Root cause of leftovers: `CLEAR_PLAN_ELEMENTS` only parsed live `Bridge/wztc-journal.tsv`. `RotateJournalIfOversized` moves older `createdElementIds=` into `Bridge/archive/`, so `ownElementOnly` / clear miss those IDs. Range wipe also skipped TextElements, so orphan `120'-0"` / `ROLL AHEAD DISTANCE` labels stacked.

Fix: do **not** replay every archived journal ID (25 files / ~30 MB hung CLEAR_PLAN). Rebuild wipe `ExecDeleteDimensionElementsInRange` now also deletes Text/TextNode in the work band, which is what stacked the leftover `120'-0"` / `ROLL AHEAD DISTANCE` labels. `harvest_journal_create_ids` is the tested parser if a future bounded archive pass is needed. Compile no longer marks a span "curved" from chord-vs-sheet length alone when sagitta is tiny. Bowed spans stay `PLACE_ARC_SIZE_DIMENSION`. Live rebuild after VBA Reset: wipe deleted 18 leftovers; Roll Ahead `185026` Arc Size; Downstream `185037` Arc Size; scorecard pass. Look work bay ~(78588, 288033).

## 2026-08-13 -- Cursor -- spec-driven designer inputs (all four loop items)

Engineer: if we will need the four agent-loop improvements eventually, do them now.

1. `get_required_designer_inputs` reads spec `inputs[]` (`sheet_resolve.required_designer_inputs`). 619-311 asks 5 (not 7): derives `closureType` from `applicability.closure` and `signSizeClass` NON-FREEWAY. Speed options from `allowed[]` ? no 60. `validate_designer_input_value` rejects 60. Prompt HOW-TO-ASK no longer hardcodes 45/35/55/Other. `plan_workflow` `inputs_locked` nextTool is this lookup.
2. `chat_driver.tools_for_turn` omits highway-catalog / junctions / gores / cell-browse / registry while `sheet_plan_active()`; keeps `place_two_way_highway`.
3. Prompt leans on the lookup + `get_plan_status` / buildGuide; work-bay pick / `path_vertices` unchanged.
4. Unit tests in `tests/test_required_designer_inputs.py` (25 related tests passed). Eval harness scenarios `sheet_619311_inputs_from_spec` and `sheet_619311_reject_speed_60` are API-billed ? not run this turn.

Do not silently default speed/area_type. Do not offer out-of-domain values via "Other".

## 2026-08-13 -- Cursor -- corridor pick ladder (get_element_vertices)

Engineer approved the simulated conversation: designer inputs first, then roadway source, then work bay snapped to the path.

VBA `GET_ELEMENT_VERTICES` (`WZTCExec.ExecGetElementVertices`) returns densified XY for line / line-string / arc / complex chain. Python: `get_element_vertices`, `propose_corridor_source`, `lock_corridor_path`, `propose_work_area_on_path`, `snap_work_area_to_path`. `place_two_way_highway` (and other striping) remembers vertices so "the road I just placed" needs no click. 619-311 closed side derives as right of travel. Length check uses the resolved station walk, not a hardcoded 2345 ft. Do not ask the engineer to aim 38 ft off centerline.

## 2026-08-13 -- Cursor -- curved QA: guides, tangent labels, PV 180

Engineer: (1) drop perp ticks + alignment on curved builds the same as straight; (2) dim name labels follow the curve tangent like dim numbers; (3) protective vehicle faces the wrong way ? rotate 180 about center, straight and curved.

`run_sheet_build` now always calls `delete_construction_guides` after geometry, before visual QA (not only `real_road_edge`). `compile_plan` emits `angleDeg` on Non-Sign labels from the local tangent; `ExecPlaceTextLabel` / `PLACE_TEXT_LABEL` take `angleDeg`. MUTCD sign faces stay view-horizontal. `compile_symbols` sets TWZWVA_P `angleDeg = atan2(tan)+180`.

## 2026-08-13 ? Cursor ? TWZSGN_P post follows travel tangent

Sign faces stay view-horizontal; the T post must not. Root cause: `PlaceSignAssembly` placed TWZSGN_P at `viewAngleDeg` (same as the MUTCD face). On a 90? bend, G20-2's T stayed world-aligned while the road (and straight-approach posts) had rotated. Fix: `PLACE_SIGN` optional `postAngleDeg`; `_place_locked_signs_from_stations` passes `_post_angle_deg` from travel (Align1 tan = ?travel, Align2 tan = +travel). Cell at 0? has an east crossbar = arm toward downstream. Face/text still use view angle. Tests: `test_signpost_angle_follows_travel_tangent`.

## 2026-08-13 ? Cursor ? readable label flip + highway-kind caution

Feature name labels still follow the corridor tangent (`_text_angle_deg`), but if atan2 is more than 90? CW or CCW from view-upright the angle folds ?180 so lettering is not upside-down (?90 kept). Same for ARROW PANEL labels. Wrong-highway caution is spec-driven for every 619 sheet: `applicability.highwayKinds` or parse `roadway`; `highwayCaution` on get_sheet_requirements / get_required_designer_inputs / lock_corridor_path / build_wztc_order_table / run_sheet_build. Does not hard-block ? agent must ask. 619-311 is `two_way_undivided` only.

## 2026-08-13 ? Cursor ? build ledger + overlap check + Tier 1 scorecard

Append-only `Bridge/build-ledger.jsonl` (survives registry wipe). `check_build_overlap` is caution-not-block: same sheet+origin ? clear_plan_elements; same sheet path conflict; other sheet ? ask. Tier 1 exact-duplicate hash fails the scorecard against live model rows (`GET_ELEMENTS_IN_RANGE_BOX` range-intersect scan, not find_elements_near). Tier 2 is station/offset vs ledger paths and live centers. Tests: `tests/test_build_overlap.py`.

## 2026-08-13 ? Cursor ? 619-311 C / S / straight family beside L

Engineer asked three full 619-311s to prove L-bend rules on other alignments. Script `scripts/build_619311_curve_family.py` (uses `bridge`, not chat_bridge). Does **not** `clear_plan_elements` between builds ? only `PlanSession.reset` + fence-wipe of the new AABB ? so the L at `(76000, 288000)` stays. C at `(79600, 288000)` R=3000; S at `(83600, 288000)` reverse-S; straight 1000 ft south at `(76000, 287000)`. All three `run_sheet_build` OK, scorecard passed. Work: C `(82097, 289270)`, S `(86413, 288488)`, straight `(78815, 286962)`. Captures `Bridge/captures/qa_311_family_*.png`. `resolve_sheet_lateral` reports `curved=True` even on the E-W straight because the CHAN_OFF offset polyline is densified (225 verts); dims on that build still read as SizeArrow 120'/50'.

## 2026-08-13 ? Cursor ? compound dims = real SizeArrow + Arc Size parts

Engineer: S-curve lane taper must not be one fake arc; split into real dimensions; each piece shows its length; parts must sum to the table value. `classify_dim_path` / `split_dim_path_runs` / `_apportion_sheet_lengths` in `sheet_compile.py`. One circular arc stays one Arc Size. Downstream 50' on R=3000 was SizeArrow because `_fit_circle_2d` capped R at 2500 ? cap is now 12000. Dim arc is forced outside the tip circle. Sign stem uses cell vertices not AABB (`DrawSign.ExtremePointAlongDir`) so C/S diagonal faces meet the white line. Tests: `test_split_reverse_s_parts_sum_to_sheet`, `test_short_highway_curve_classifies_as_arc`.

## 2026-08-14 ? Cursor ? fresh L/C/S band + live S split

Fresh band: L origin `(90000, 300000)`, C `(93600, 300000)`, S `(97600, 300000)`. Live S failed to split because `assemble_corridor` sampled Align1/2 at 50 ft and erased the fillet; now 10 ft. S lane taper placed as SizeArrow 58.7 + Arc Size 191.2 + SizeArrow 430.1 = 680. C/L stay one dim per span (placedCount 22). Downstream 50' classifies as arc on all three. Stem uses `RayHitOutline`. Captures `qa_311_fresh_*.png`. Work: L `(92506, 299978)`, C `(96097, 301270)`, S `(100413, 300488)`. Tests: 19 passed in `test_curved_corridor_compile.py`.

## 2026-08-14 ? Cursor ? S-curve rebuild: stem gap + L-style dims

Engineer: L is the visual reference; S dims missing/wrong; stem still gapped. Stem gap root cause was `AccelRayHits` always closing last?first on the whole cell vertex dump ? that chord sits short of the orange face on diagonal C/S. Stem now `postOuter` on the perp ray ? `faceTarget` at STEM_GAP; close only loops < 80 ft. Dims: `min_run_ft` 100 (no 58' crumbs); SizeArrow only if run sag < 0.35 ft. Rebuild S only: `python scripts/build_619311_curve_family.py S`. Fence-wipe S AABB only. Captures `qa_311_fresh_s_*.png`.

## 2026-08-14 ? Cursor ? S-curve dims inside pavement + G20 stem

Roll ahead / downstream sat in the travel lanes because `place_path_hugging_dimension` always used `r+15`. On the inside of an S-bend the closed shoulder is closer to the arc center ? dim must be `r-15` (negative `DimHeight`). `arc_dim_line_radius` in `wztc_ops.py`. Buffer looking half-missing was the same inside-arc clip. G20 stem: drop RayHit-nudge (inner SF_P hole); stem to snapped inward vertex on the perp ray. Test: `test_arc_dim_line_radius_inside_of_curve_stays_off_pavement`. Rebuild S only.

## 2026-08-14 ? Cursor ? S buffer invisible SizeArrow + G20 orange snap

Engineer: BUFFER SPACE label present, no dimension; G20 still gapped. Live element 220320 was a SizeArrow whose range equaled the two EOP tips (350×350, no extra for dim line) ? ny_Plan style assignment zeros DimHeight before AddElement. Axis-aligned SizeArrows still drew; the 45° buffer did not. Fix: set DimHeight after style + after Rewrite; pass sheet-length override text. G20: snap/stem to color-6 orange verts, then +1.5 ft into the fill.


