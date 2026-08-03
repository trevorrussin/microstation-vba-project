# Agent Work Log

Shared cross-tool log. This project gets worked on from both Claude Code and
Cursor — neither tool can see the other's session history or memory, so this
file is the bridge. **Whichever AI tool finishes a non-trivial, non-obvious
piece of work appends an entry here before finishing.**

"Non-obvious" is the bar — don't log routine edits that are already clear from
`git log`/`git diff`. Log the things a future session (in either tool) would
otherwise have to rediscover: a design decision and why, a bug whose root
cause wasn't where it looked, a gotcha specific to this codebase, a dead end
that shouldn't be retried.

## Entry format

```
## YYYY-MM-DD — <tool> — <one-line summary>
<2-6 sentences: what changed, why, anything surprising. Reference specific
files/functions. If it's a bug fix, say what the real root cause was, not
just the symptom.>
```

Newest entries at the bottom (append-only, chronological — matches `git log`
ordering conventions already used elsewhere in this repo).

## How each tool uses this file

- **Claude Code**: reads this at the start of a session when the user
  mentions work happened elsewhere ("I did some work in Cursor"), or
  periodically checks it — see the pointer in `CLAUDE.md`. Anything
  load-bearing gets folded into Claude Code's own persistent memory
  (`~/.claude/projects/.../memory/`) from here, same as if the user had
  described it directly in chat.
- **Cursor**: instructed via `.cursor/rules/agent-log.mdc` to append here
  after finishing meaningful work.

This is a manual bridge, not automatic sync — each tool only reads/writes it
when its own instructions or the user tells it to.

---

## 2026-08-02 — Claude Code — sign rotation now matches current view angle

`DrawSign.bas`/`WZTCExec.bas`: sign face cells were rotating to match the
alignment/perpendicular direction, which is mathematically guaranteed to
flip some signs upside-down and has nothing to do with the view. Fixed to
capture the view's rotation (`ViewRotationAngleDegrees`, via `v.Rotation`
before it gets reset to identity) and use that for `ACTIVE ANGLE` instead —
verified live via direct COM reads that a placed cell's rotation exactly
matches the view's rotation at placement time, both unrotated (0°) and
rotated (60°) cases. See Claude Code memory `feedback-sign-rotation-matches-view`
for the full writeup.

## 2026-08-02 — Claude Code — bounded reuse for Bridge/results_*.tsv

`WZTCBridge.bas`'s `WriteResultRows` used to write one new
`results_<reqId>.tsv` file per query call, forever (19+ leftover files
found after one session). Changed to a bounded pool of 8 reusable
filenames (`results_slot0.tsv`..`results_slot7.tsv`, chosen via
`reqId mod 8`) — each write already truncates/overwrites via `Open ...
For Output`, so no deletion is involved anywhere, file count just stays
bounded. Only safe because a single batch never has more than 8
multi-row-result ops today (`call_batch` in `bridge_client.py` is only
ever called with one op at a time) — if real multi-op batching gets
added later, bump `RESULT_POOL_SIZE` past the largest expected batch.

Hit a real VBA gotcha along the way: a new module-level `Const` placed
between two existing procedures (instead of in the top-of-file
declarations block) fails to compile with "Only comments may appear
after End Sub, End Function, or End Property" — moved
`RESULT_POOL_SIZE` up next to `BRIDGE_DIR` etc. to fix. See Claude Code
memory `feedback-vba-compile-error-recovery` for the full incident,
including a clean scriptable way to detect/dismiss/reset a blocking VBA
compile-error dialog via `VBE.CommandBars` (no manual clicking, no
guessing at dialog internals).

## 2026-08-02 � Cursor � keyin probe 3s hang timeout + wave8 promote

`scripts/keyin_batch.py`: live `SendKeyin` now runs in a child process with a hard **3s** timeout (`SENDKEYIN_TIMEOUT_SEC`); hangs are recorded as `HANG` ? `unsafe-blocked` and the batch continues. Also stopped executing `tool`/`datapoint` kinds (they activate-and-wait � `TITLEBLOCK PLACE` was the wave8 forever-hang). Wave8 sparse-category harvest probed/promoted **+398** registry rows; 4 PDF embed/layers keyins hit the new timeout and were added to the skip list. Registry ~2208 / ~1959 verified.

## 2026-08-02 � Cursor � drawing recipe probe (element-delta bar)

Added `scripts/recipe_batch.py` + `Data/recipe-candidates.tsv`. Unlike settings
`keyin_batch.py`, a drawing recipe only promotes when graphical element count
on `DELETE.dgn` increases (plus COM alive / 3s step timeout). First live results:
`PLACE_LINE` and `PLACE_SHAPE_CONSTRAINED` ? `verified-headless-safe`;
`HATCH_ICON` (Legacy two-identical-seed pattern) completed without hang but
`+0` elements ? stays `needs-testing`. Circle/block/arc/smartline not seeded
(no in-repo CadInputQueue sequence).

## 2026-08-02 � Cursor � Phase C edit direct_api + Phase B WZTC registry rows

Phase C: added `ExecCopy/Rotate/Scale/Mirror/ArrayElementByID` in
`WZTCExec.bas` using Element API patterns live-proven on `DELETE.dgn`
(Clone+Move, ScaleUniform, Matrix3d Z-rotate Transform, Mirror two-point) �
wired through `WZTCBridge` + MCP tools. No CadInputQueue inventing.
Phase B: catalogued existing place bridge ops as `direct_api`
`verified-headless-safe` rows (`PLACE_CELL` was flipped from bare
`unsafe-blocked` COMMAND). Re-import `WZTCExec.bas` and `WZTCBridge.bas`
in the VBA IDE before exercising the new edit ops.

## 2026-08-02 � Cursor � hatch Element API + PLACE_ARC / PLACE_TEXT_LABEL

CadInputQueue `HATCH ICON` (Legacy twin-seed) stays unreliable headlessly
(+0 elements on DELETE.dgn). Switched workspace hatch to
`CreateHatchPattern1` + `ClosedElement.SetPattern(..., Matrix3dIdentity)`
� live `HasPattern=True`. New bridge/MCP: `HATCH_ELEMENT`, `PLACE_ARC`
(placeArcModeEx=3), `PLACE_TEXT_LABEL` (TEXTEDITOR INSERT_TEXT). All three
plus updated `PLACE_WORKSPACE` verified OK via bridge on DELETE.dgn.
`HATCH_ICON` registry row ? `unsafe-blocked`. Hot-reloaded WZTCExec +
WZTCBridge. Note: PrintWindow captures often omit associative hatch lines
even when HasPattern is True � trust HasPattern / in-app view for hatch.

## 2026-08-02 � Cursor � Tier1-3 general geometry ops

Added Element-API geometry suite to `WZTCExec` / `WZTCBridge` / MCP:
Tier1 place (circle/ellipse/block/polyline/polygon) + symbology; Tier2
copy-parallel (lines), crosshatch/remove-hatch, break-line, extend-line
(recreate, not EndPoint � EndPoint assign hung VBA), fillet/complex
(needs-testing); Tier3 fence block + copy/move/delete contents, select/clear.
TRIM/CHAMFER left interactive-only (no COM ConstructTrim/Chamfer).
Live: Tier1 + symbology + copy-parallel OK on DELETE.dgn before
`LineElement.EndPoint` hang wedged VBA `[running]`. Reset/Ctrl+Break from
automation failed � user must interrupt VBA (or restart MicroStation), then
hot-reload and re-verify extend/fence/fillet/complex.

## 2026-08-02 � Cursor � Tier1-3 live-verified + complex-string fix

After VBA Reset, phased live verify on DELETE.dgn: extend (recreate-line path,
not EndPoint), break, crosshatch/remove, fence define/copy/undefine, fillet all
OK. `CREATE_COMPLEX_STRING` initially failed compile �
`CreateComplexStringElement1` needs `ChainableElement()`, not `Element()`;
fixed in `ExecCreateComplexString` via `el.AsChainableElement`, hot-reloaded,
live OK (`partCount=2`). All Tier1�3 geometry bridge ops now
`verified-headless-safe` except TRIM/CHAMFER (still interactive-only). Gotchas:
`scale` as a VBA local name conflicts with MicroStation `Scale` (use
`lenScale`); PrintWindow often omits hatch lines � trust `HasPattern`;
VBA `[running]`/`[break]` blocks hot-reload until Reset.

## 2026-08-02 — Cursor — manual-search robustness + sheet-registry Non-Freeway wave

`manual_search.py` / `ingest_manuals.py`: repo-relative index paths (no hardcoded
`c:\repos\...`); missing index returns `heading=INDEX_MISSING` instead of silent
`[]`; multi-token zero-hit queries retry with OR then quoted phrase (fixes
`lane closure` + `source=supplement` empties under FTS5 AND). Re-ingest embeds
`619-NNN` into stdsht page headings. `sheet-registry.tsv` grown 6→19 sheets:
fixed 619-311 (Book 3 p.81-82 tables 311-01..05); added Non-Freeway–first
080/201–203/308–309/312–314/317 plus Short Term 304–306. Helper
`scripts/extract_sheet_signs.py` dumps candidates from sheet-owned `.dgn` pages
only. Live bridge: `get_sheet_requirements('619-201')` found=True; unseeded
still found=False.

## 2026-08-02 — Cursor — sheet-registry batch 2 (Short Term remainder + Intermediate)

Grew `sheet-registry.tsv` 19→38. Added general notes/legend (010–011), Short
Duration freeway (205–207), remaining common Short Term (315–316, 318–319,
321–325), and Intermediate starters (401–403, 407, 410). Skipped 619-012
(full sign catalog, not a placement typical), 619-101/102 (in DesignerRef but
no Book 3 owned pages), and multi-sheet barrier 001 (no `.dgn` text ownership).
Signs from sheet-owned `.dgn` pages only. Live bridge OK for 205/322/407;
101 still found=false.

## 2026-08-02 — Cursor — sheet-registry batch 3 (fill most remaining Book 3)

Grew `sheet-registry.tsv` 38→80. Added detail/general (001–006), mowing/special
(021–023, 031–033, 041, 060, 090–091), Mobile (110–114), remaining Short Duration
(208–209, 211–212), remaining Intermediate (412, 414–418, 421–423), and Long Term
(501–504, 517–518, 520, 523–524). Still unseeded (~11): 012 catalog; 050–051 /
101–104 / 204 / 419–420 / 519 not present (or not owned) in Book 3 PDF text.
503 seeded with Detour element but empty signs (drawing text unreadable to
extractor). Live spot-check via bridge after write.

## 2026-08-02 — Cursor — sheet-registry last-row miss was off-by-one

Root cause of `619-524` not found was not Line Input dropping the last
line: `ReadAllLines` returned it fine (`last=619-524`), but
`GetSheetRequirements` / `ListRegisteredSheets` looped `For i = 2 To n`
and read `lines(i - 1)`, so the last 1-based data row `lines(n)` was never
examined (sentinel after 524 had made 524 become `lines(n-1)` and appear
to "fix" it). Corrected to `lines(i)`. Kept ADODB.Stream whole-file read.
Verified live: 524/523/001/201 found; 101 still correctly unseeded.

## 2026-08-02 — Cursor — sheet-registry complete for DesignerRef (91/91)

Seeded remaining 11 DesignerRef sheets: 012 (sign catalog, empty signs on
purpose), plus stubs for sheets absent from 2026 Book 3 PDF (050-051,
101-104, 204, 419-420) and 519 (TOC listed, p.149-150 blank). Stubs keep
title/roadType/duration but empty signs with an explicit confirm-against-
current-sheet note — no invented sign lists. Live: 012/101/204/519/524 all
found=True. Book 3 extras not in DesignerRef (024-026, 034, 042-046) still
unseeded.

## 2026-08-02 — Cursor — restore full signs + sheet NOTES into registry

Rebuilt all 80 complete `sheet-registry.tsv` rows without the old ~255-char
truncation: full PDF sign lists again (e.g. 619-502 now includes W20-1 /
W24-1bL/R). `notes` now append instructional NOTES text extracted from Book 3
pages (numbered note bodies + key SHALL/SHOULD lines) after the table/page
cite, separated by `||`. Cap ~1800 chars per notes field. Stubs/012 unchanged.
Live bridge OK for 502/307 with long notes.

## 2026-08-02 — Cursor — trim sheet-registry notes to cites only

Removed noisy PDF NOTES bodies from complete rows; `notes` are back to short
table/page cites (e.g. `Tables 502-01..502-05; Book 3 p.138-139`). Full sign
lists kept. Stub/catalog notes unchanged. Rationale: drawing-text NOTES were
fragmented and inflated tool-result tokens for little agent benefit.

## 2026-08-02 — Cursor — click-first + cost caps after arc session

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
(existing history scrubbed 1.2MB→229KB). Re-import `WZTCChatPanel.frm` for
the pick-button status caption; restart `chat_driver.py` for the Python
changes.

## 2026-08-02 — Cursor — chat-log false resync + list_levels cap

Conversation pane "repeated history" + stale Reference pick prompt: 
`WZTCChatTimer` treated any `ReadAllLines=0` (file locked mid-append) as
rotation (`n < mLastLineCount` → reset to 0) and replayed the entire
`chat-log.tsv` into an already-filled panel — including old
`ASK_USER_CHOICE` lines that re-showed "Use coordinates I already gave."
Fix: only resync when `n > 0 And n < mLastLineCount`; call new
`ResetTranscriptPanes` on real rotation; `FINAL` always `HideChoiceButtons`.
Also `list_levels(name_contains=...)` + hard cap — unfiltered returned 3046
rows and ~$0.50 of follow-on input on the orange-color turn. Prompt: don't
guess color indices. Re-import `WZTCChatTimer.bas` + update panel form code;
restart chat_driver.

## 2026-08-02 — Cursor — prefer hot_reload for VBA sync

Engineer preference (standing): after editing existing `.bas`/`.frm`/`.cls`
on disk, run `python mcp-server/hot_reload.py <files>` instead of asking for
manual delete+re-import. Cursor rule: `.cursor/rules/hot-reload-sync.mdc`.
Python (`chat_driver.py` / `wztc_ops.py`) still needs a process restart.
New IDE components and UserForm Designer controls still need manual Import.

## 2026-08-02 — Cursor — resolve_color / list_colors (COM) + level filter

Color indices are per-DGN (DELETE.dgn: 3=red, 6≈orange). `list_colors` +
`resolve_color(name|rgb)` live in `wztc_ops.py` via
`ActiveDesignFile.ExtractColorTable` / `FindClosestColor` (not the VBA
bridge) — confirmed orange→index 6. First attempt typed `ColorTable` in
`WZTCQuery.bas` and hot-reload left the VBA project unable to Accept any
further CodeModule writes (every hot_reload COM-exceptions); color ops
moved to Python COM and Query/Bridge reverted to pre-color text. If bridge
keyins still fail: VBA IDE Reset, then manual re-import WZTCQuery +
WZTCBridge from disk. `list_levels` refuses empty `name_contains`. Prompt
requires resolve_color before named-color symbology. Restart chat_driver.

## 2026-08-02 — Cursor — resolve_color / list_colors + require level filter

Color indices are per-DGN (DELETE.dgn: 3=red, 6≈orange via FindClosestColor).
Added `LIST_COLORS` / `RESOLVE_CLOSEST_COLOR` in `WZTCQuery`/`WZTCBridge`
(ExtractColorTable + GetColors / FindClosestColor, KB0039791), wrapped as
`list_colors` + `resolve_color(name|rgb)` in `wztc_ops` / chat tools /
`server.py`. `list_levels` now refuses empty `name_contains`. Prompt:
resolve_color before change_element_symbology for named colors. Hot-reloaded
Query+Bridge; restart chat_driver for Python.

## 2026-08-02 — Cursor — cells / line styles / fonts + registry false-OK scrub

Agent stumble items 1–3/7/8. New Python-COM tools in `wztc_ops.py` (same
pattern as resolve_color — not VBA bridge): `list_line_styles` (requires
filter; 471 styles live), `resolve_line_style` (aliases like dashed→
`( Dashed )`; **Name** is the lookup key, Number is not — `LineStyles(-104)`
fails), `cell_library_status` / `attach_cell_library` /
`list_cells` via `AttachCellLibrary` + `GetCellInformationEnumerator(False,
False)`, `list_fonts`/`resolve_font`, `list_text_styles`/`resolve_text_style`.
`change_element_symbology` gained `line_style_name`; VBA
`ExecChangeElementSymbology` prefers Name over collection index (hot-reloaded
Exec+Bridge). Annotation scale still from `describe_drawing_state` only.
Restart chat_driver for new tools.

## 2026-08-02 — Cursor — false-OK CommandName audit (item #8)

Replaced the precautionary family downgrade with a live audit:
`scripts/keyin_false_ok_audit.py`. Probe does SendKeyin **without** the
immediate SendReset that masked pending prompts in `keyin_batch._one_keyin`,
then classifies by `CommandState.CommandName` (ARMED if non-empty and ≠
baseline `Element Selection`). Ran `--scope all-gated` on DELETE.dgn:
1733 rows → CLEAN 1661, ARMED 67, HANG 3, SKIP 2. Applied: 67 →
`unsafe-blocked` (ZOOM_IN/OUT, WINDOW_CENTER/ORIGIN, many SET_* /
UPDATE_VIEW / REFERENCE_* that prompt "Select view|reference|point"),
11 CLEAN `needs-testing` restored to verified (incl. ZOOM_HALF/DOUBLE/
PREVIOUS, FIT_ELEMENT/SELECTION/FENCE, ACCUDRAW_ROTATE_VIEW), 3 HANGs →
`needs-testing`. Results: `Bridge/keyin-false-ok-audit.json`. Empty
CommandName after NULL/NOCOMMAND is CLEAN, not ARMED.

## 2026-08-02 — Cursor — ask_user_choice element pick + selection-tool reality

Added `allow_element_pick` alongside `allow_point_pick` on
`ask_user_choice` (`chat_driver.py` + `WZTCChatPanel.frm`). Element
identify is injected as a normal choice-button caption ("Identify an
element in the drawing") — no extra Designer control. Clicking that
caption runs GetInput → `LocateElement` and replies `elementId=… type=…
level=… [cell=…]` (does not merely echo the label). Point pick stays on
dedicated `btnPickPoint`. Prefer element-pick for "which existing thing";
point-pick for "put it here."

Selection-tool audit vs color-quality resolve helpers: agent-direct
`select_element(id)` / `clear_selection` are solid (need an ID first —
element-pick now supplies that). Registry `POWERSELECTOR MODE
NEW/ADD/SUBTRACT/INVERT`, `SINGLE`, `CHOOSE ALL`/`NONE`,
`SELECTION SET INVERT` are settings-mode toggles (CLEAN in false-OK
audit) — they change how the interactive tool behaves, they do **not**
select anything headlessly. `POWERSELECTOR BLOCK/SHAPE/LINE/CIRCLE` stay
on Element Selection CommandName but still expect engineer clicks for
the fence — not color-class resolve tools. No working headless "select by
level" op in the registry; use `list_levels` + find/classify, or
element-pick. `CHOOSE_ELEMENT` remains unsafe-blocked.

## 2026-08-02 — Claude Code — agent-driven 8-step wizard built + live-verified through a real chat conversation

Built and live-verified all 4 components of the plan at
`~/.claude/plans/polished-purring-reef.md` (full detail in Claude Code
memory `project_agent_driven_8step_wizard_status.md`): `BUILD_WZTC_ORDER_TABLE`
(`WZTCRules.bas`), `FIND_REFERENCE_LINEWORK` (`WZTCQuery.bas`, defaults to
active-model-only per the engineer's direction — reference-attachment
scanning is built and opt-in via `includeReferences=True` but not
exercised by default), `DEFINE_ALIGNMENT_SEGMENT`/`COMMIT_ALIGNMENT`
(`AlignmentTool.bas`/`WZTCExec.bas`), and `PLACE_ORDER_TABLE_STATIONS`
(`PerpPlacement.bas`, batches what used to be one `place_perp_line` call
per order-table item into one call per alignment). All 5 wired into
`chat_driver.py`'s `_WZTC_OP_NAMES` and confirmed callable by the actual
chat agent, not just direct bridge test scripts — full plan (work space →
alignment → spacing → sign → tick line) drawn successfully end-to-end
through the real `WZTCChatPanel.frm` conversation.

**Real incident, same session**: an uncapped O(n³) chaining algorithm in
`FindReferenceLinework` (rebuilding a connected path from scanned reference
elements) hung MicroStation badly enough that Ctrl+Break didn't recover it
— required a full restart. Root-caused to two things: unbounded chaining
complexity (fixed with a hard 80-segment cap before the expensive pass
runs) and reference-attachment COM calls (`.AsVertexList`, `att.Name`)
throwing instead of returning `Nothing` the way `.AsLineElement` does
elsewhere in this codebase (fixed with whole-block error guards, not
line-by-line). Full writeup: Claude Code memory
`feedback_reference_scan_hang_and_com_errors.md`.

**Two real `chat_driver.py` bugs found via the live conversation, both
fixed**: (1) `enter_mode`'s effect is deferred to the next turn by design,
but nothing told the agent that, so it retried WZTC tools in the same turn
repeatedly and concluded tooling was broken — fixed via explicit
`GENERAL_MODE_HINT` guidance. (2) `_SESSION_MODE` didn't survive a
`chat_driver.py` restart even though conversation history does, causing
the same confusion a second way (agent's own history showed it "already"
in wztc mode when the fresh process was actually back in general) — fixed
by persisting the mode to `Bridge/chat-session-mode.txt`, loaded at
startup with a safe fallback to `"general"` on anything unrecognized.

**New tool**: `mcp-server/restart_chat_driver.py` — safe automated restart
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
agent chose the less-efficient per-item path once live — not yet re-verified
in a second live run. `find_reference_linework`'s `includeReferences=True`
path is still unverified against real reference geometry.
