# WZTC Designer MCP Server

Exposes the WZTC agent tool surface (plan Layers 5–6, milestones M5–M7) as a
real, connectable MCP server: query the model, get deterministic MUTCD spacing
numbers, draw signs/elements/work space, edit agent-created elements via a
safety-gated MicroStation command registry, and search the NYSDOT/MUTCD
reference manuals — without clicking through the 8-step wizard.

This is the Python side of the bridge. It does not replace `WZTCBridge.bas` —
it drives it, over the file + COM-keyin transport proved in M1/M2.

**M7** split each op's actual implementation out into `wztc_ops.py`, shared
between this stdio MCP server and `chat_driver.py` — the agent loop behind the
in-MicroStation chat panel (`UserForms/WZTCChatPanel.frm`). Both processes can
run against the same live MicroStation session at once; `chat_driver.py` uses
its own request/response file pair (`bridge_client.chat_bridge` /
`WZTCBridge.RunChatToolRequest`) so the two never race on the same files — see
`Bridge/README.md`. `server.py` is the tool surface for Claude Code / any MCP
client; `chat_driver.py` is the standalone driver for the in-app panel. Run
`python chat_driver.py` yourself to use the panel — it isn't auto-launched.

## What this talks to

```
this server  ──►  win32com (reads: attach to live MicroStation session)
             ──►  Bridge/request.tsv + a COM-sent keyin (writes)
                        │
              WZTCBridge.RunRequest  (VBA, in MicroStation)
                        │
              WZTCExec / WZTCQuery / WZTCRules / WZTCSheetRegistry
              / WZTCCommandRegistry
```

Every tool call here writes `Bridge/request.tsv`, sends
`VBA RUN [Test]WZTCBridge.RunRequest` via `CadInputQueue.SendKeyin`, and reads
`Bridge/response.tsv` back. This has been confirmed synchronous against this
install (~0.4s round trip) — see `bridge_client.py` for the retry-loop
fallback kept in case that ever isn't true under load.

## Prerequisites

1. **MicroStation 2023 CONNECT must be open**, with a design file active and
   the `Test` VBA project loaded (containing the current
   `Modules/WZTCBridge.bas`, `WZTCExec.bas`, `WZTCQuery.bas`, `WZTCRules.bas`,
   `WZTCSheetRegistry.bas`, `WZTCCommandRegistry.bas`, plus their existing
   dependencies — `PerpPlacement.bas`, `DrawSign.bas`, `SignLibrary.bas`,
   `DrawElements.bas`, `SharedState.bas`). If you've just pulled changes to any
   of those files, re-import them in the VBA IDE first (delete/Remove the old
   module, File → Import File — a plain re-import over an existing module of
   the same name does not reliably take effect in this IDE, confirmed live)
   — see the repo's `CLAUDE.md` File Sync Protocol. For the in-MicroStation
   chat panel specifically, also import `UserForms/WZTCChatPanel.frm` and
   `Modules/WZTCChatTimer.bas` (controls must be added to the form manually —
   see that file's header comment), and add the manually-added controls once.
2. **Python 3.10+** with the packages in `requirements.txt`:
   ```
   pip install -r requirements.txt
   ```

## Register with Claude Code

```
claude mcp add wztc-designer -- python "c:\repos\microstation-vba-project\mcp-server\server.py"
```

To remove: `claude mcp remove wztc-designer`.

## Tool surface (M5 + M6 + M7)

| Group | Tools |
|---|---|
| Query | `find_elements_near`, `station_to_point`, `get_alignment_stationing`, `list_levels`, `classify_site_features` |
| Compute | `compute_spacing`, `get_sheet_requirements` |
| Draw | `place_perp_line`, `place_sign`, `place_workspace`, `place_element_run`, `place_cell`, `set_sign_attributes`, `handoff` |
| Registry / Edit | `list_registry_commands`, `describe_registry_command`, `run_registry_command`, `move_element`, `change_element_level`, `edit_text_element`, `delete_element` |
| Session | `undo_last_op`, `get_journal`, `list_deferred_handoffs` |
| Reference | `search_reference_manual` |

`search_reference_manual` doesn't touch `WZTCBridge` at all — it's a local
FTS5 full-text search (`Data/manual-index.sqlite`) over the three NYSDOT/MUTCD
reference PDFs in `Project Documentation/`, built by `ingest_manuals.py`
(run it once, and again any time the source PDFs change). Both the PDFs and
the sqlite index are gitignored — a fresh clone needs:

```bash
# place part6.pdf, B-2011Supplement-adopted.pdf, 2026_1_stdsht_usc_book_3.pdf
# under Project Documentation/, then:
python mcp-server/ingest_manuals.py
```

If the index is missing, the tool returns one hit with `heading=INDEX_MISSING`
instead of a silent empty list. Multi-word queries that miss under FTS5 AND
are retried with OR / phrase matching. Ground engineer-facing questions about
MUTCD/NYSDOT requirements in the returned excerpt + page citation, not recollection.

`manual_search.render_page_image(source, page_num, out_path)` renders the
actual PDF page (via PyMuPDF) rather than just returning text — the chat
panel (below) uses this to show the real manual/sheet page a search hit
came from, alongside its text excerpt.

**Not exposed — by design:** `test_registry_command`. That VBA op bypasses the
`needs-testing` gate for exactly one manual IDE promotion run. Reachable only
by hand-editing `Bridge/request.tsv` and sending the keyin yourself. Arbitrary
execution stays impossible for the agent even during testing.

**Not yet exposed — never built in M2:** `get_model_context`. Compose equivalent
context today from `list_levels` + `get_alignment_stationing` instead.

## The engineering-judgment boundary

This server never computes a spacing, taper length, or sign size itself.
`compute_spacing` and `get_sheet_requirements` wrap `WZTCRules.bas` /
`WZTCSheetRegistry.bas` so those numbers stay deterministic and traceable to
a PE-reviewed table — the calling agent decides *what* to place and *how to
respond to a site condition*, never invents a number that belongs in one of
those two tools.

Every `place_*` / edit tool takes an optional `reason`. It rides through
untouched into `Bridge/wztc-journal.tsv`.

## Dimensions, callouts, and other interactive-only commands

Authoritative list: `Data/command-registry.tsv` rows where `safetyStatus` is
`interactive-only-use-handoff` or `unsafe-blocked` (currently
`DIMENSION_SIZE_WITH_LINES`, `DIMENSION_LINEAR_SIZE_ARROW`,
`TEXTEDITOR_PLACENOTE`, `CHANGE_ATTRIBUTES_INTERACTIVE`). Do not invent a
second "red list" elsewhere — the registry is the single source.

`handoff(kind="dimension" | "callout", ...)` queues these and returns
`DEFERRED`. After drawing everything else, call `list_deferred_handoffs()` and
tell the engineer what still needs a few clicks through the existing forms.

`run_registry_command` on any of those rows returns a clear `ERROR` pointing
at `HANDOFF` — never a silent no-op, and never a fake `OK`.

## Edit vocabulary (M6)

`move_element` / `change_element_level` / `edit_text_element` / `delete_element`
default to `own_element_only=True`: the target must appear as a
`createdElementIds=` / `elementId=` value in this session's journal. That
matches "elements the agent itself created" rather than silently expanding to
edit anything in the DGN.

Responses for move/level/text embed prior state (`priorDeltaX/Y`, `priorLevel`,
`priorText`) so `undo_last_op` can reverse them. `delete_element` always
returns `notUndoable=Y` — honesty over fake success, same spirit as `HANDOFF`.

## Undo

`undo_last_op()` walks `Bridge/wztc-journal.tsv` backward for the most recent
undoable op:

- Draw ops → delete `createdElementIds` / `elementId`
- Move / level / text → re-apply the embedded prior-state fields
- `DELETE_ELEMENT` rows (`notUndoable=Y`) are skipped

It does **not** use MicroStation's own undo stack. Independently tested by
`Debug/DebugAgentLoopTest.bas` and `Debug/DebugCommandRegistryTest.bas`.

## Command registry growth

See `Data/README.md` (command-registry.tsv section) for the promotion process
(`needs-testing` → `verified-headless-safe`) and the structural close-out
guard that refuses any `COMMAND:` recipe lacking `DATAPOINT:` + `RESET`.

## In-MicroStation chat panel (M7)

`chat_driver.py` is a persistent process holding the actual Claude Opus 5
agent loop (`client.beta.messages.tool_runner`, adaptive thinking, most
`wztc_ops` functions plus `ask_user`/`view_drawing`/`adjust_view`/
`web_search` as always-on "base" tools — see Session modes below for what's
base vs. mode-gated) behind `UserForms/WZTCChatPanel.frm` — a modeless
in-MicroStation dialog. "Python owns the brain and the hands; VBA owns only
the face": the panel only polls `Bridge/chat-log.tsv` (via
`WZTCChatTimer.bas`, a Win32 `SetTimer` loop — the first polling mechanism
in this codebase) and appends to `Bridge/chat-input.tsv` on Send; every
actual tool call goes through the identical `WZTCBridge.ExecuteOp` dispatch
every other op in this repo already uses, so watching it work is the same
visible MicroStation activity as always.

Requires `ANTHROPIC_API_KEY` set (or an `ant auth login` profile) — picked up
automatically by `anthropic.Anthropic()`, no code for it here. Run
`python chat_driver.py`; it isn't auto-launched by VBA (see `Bridge/README.md`
for why). Full protocol/file-schema docs: `Bridge/README.md` → "M7 — chat
panel protocol".

### Session modes (2026-08-02)

The agent boots in **general mode** — broad MicroStation drawing/query
capability, no domain-specific tools or rules loaded — and switches into
**wztc mode** only once the engineer clearly starts a WZTC task, via a tool
call the model makes itself (`enter_mode("wztc")`; `exit_mode()` returns to
general). This is deliberately *not* a hidden classifier: the switch is
visible in the panel transcript ("— Switched to wztc mode —", a new
`MODE_CHANGED` chat-log line type) and the engineer can always tell which
mode is active.

Base tools (`BASE_TOOLS`, always loaded) cover general drawing/edit/query,
the view/registry tools, journal/undo, `ask_user`/`ask_user_choice`,
`view_drawing`, `web_search`, and `enter_mode`/`exit_mode` themselves. WZTC
tools (`WZTC_TOOLS`, loaded only in wztc mode) are `compute_spacing`,
`get_sheet_requirements`, `resolve_sign_code`, `place_perp_line`,
`place_sign`, `place_workspace`, `place_element_run`, `place_cell`,
`set_sign_attributes`, and `search_reference_manual` — plus the WZTC-only
system-prompt rules (the engineering-judgment boundary, `road_type`
handling, sign-code resolution). Modes **stack**, not replace: wztc mode is
base + WZTC, so the agent never loses general drawing capability while
doing WZTC work. `_MODE_TOOLS`/`_MODE_SYSTEM_PROMPT` are plain dicts keyed
by mode name, and `run_turn()` reads `_SESSION_MODE` fresh on every call —
adding a future mode (drainage, terrain, etc.) is a matter of defining its
own tool list + prompt addendum and adding a dict entry, not restructuring.

`_SESSION_MODE` is a module-level global, not persisted across a
`chat_driver.py` restart — every fresh process starts in general mode.
Switching modes changes both `tools` and `system` on the next
`run_turn()` call, which invalidates the prompt-cache prefix for that one
turn (expected and cheap for a deliberate, infrequent switch — see
`prompt-caching.md`'s invalidation hierarchy). Verified live before
building this: the Messages API tolerates conversation history containing
`tool_use`/`tool_result` blocks for a tool no longer declared in the
current request's `tools` list, so a mode switch never needs special
handling for stale tool_use blocks from a different mode.

`ask_user` is a tool, not a special stop condition — it writes an `ASK_USER`
line and blocks (a plain polling wait, safe here since this process has no UI
thread to keep responsive, unlike the VBA panel) until the engineer replies
in the panel, matching the documented pattern for promoting question-asking
to a tool call in agentic loops.

`ask_user_choice` (2026-08-02) is the same pattern with real clickable
buttons — analogous to Claude Code's own `AskUserQuestion`, used only for a
genuine decision with 2-4 concrete options rather than every question. One
option can be "click a point in the drawing" (`allow_point_pick=True`),
which runs a `GetInput` click-capture from the panel's own native button-
Click event and returns formatted `(x, y, z)` coordinates as the answer.
Deliberately NOT built by routing a `GetInput` wait through `WZTCBridge.
ExecuteOp` — MicroStation's COM interface has no non-blocking way to detect
a click, and the bridge's `SendKeyin` call has no timeout of its own, so
that would hang this whole process with no way to cancel (the same failure
mode that already hung `WZTCViewCapture.bas` once). The engineer can always
ignore the buttons and type a free-form reply instead — every answer path
(button click, point pick, typing) converges on the same
`Bridge/chat-input.tsv` append `ask_user`'s reply already uses.

After every `search_reference_manual` call, the panel also shows the actual
matched PDF page (top hit only) in `imgScreenshot`, alongside a citation
line in the activity trace — same display mechanism the post-turn drawing
screenshot uses (last-shown-wins if both fire in one turn), driven by
`manual_search.render_page_image` + a new `REFERENCE_IMAGE` chat-log line
type. Best-effort: a missing/gitignored PDF or an out-of-range page just
means no image that turn, never a failed tool call — the text excerpt the
model already has is unaffected either way.

`view_drawing` (2026-08-02) is the tool that closes the gap between "the
panel shows a screenshot" and "the agent actually sees it" — every other
image feature (the post-turn auto-focus screenshot, the reference-manual
page render) only ever displayed a picture to the human; the model itself
never received one as vision input. `view_drawing` returns its screenshot
as a real `image` content block inside the tool result (base64 PNG, same
capture path/resize as the auto-focus screenshot), so the *next* turn the
model actually looks at it. Deliberately agent-triggered, not automatic
after every tool call or every turn — the system prompt tells the model to
call it selectively (a substantial design change, or suspected errors),
since each call costs real image tokens (~1500-2000, at
`view_capture.MAX_LONG_EDGE`'s existing 1568px resize) and calling it after
every small edit would pile that up fast on a long multi-step turn.

`web_search` (2026-08-02) is a separate, narrowly-scoped escape valve for
when the agent gets stuck on MicroStation's own VBA/COM API layer — not
general internet access. `allowed_domains` hard-restricts it to
`docs.bentley.com`, `communities.bentley.com`, and
`bentleysystems.service-now.com` (Bentley's own documentation/KB/
programming forum); the system prompt tells the agent it's a last resort
for API/automation troubleshooting only, never for spacing/sign/MUTCD
content, which stays on `compute_spacing`/`get_sheet_requirements`/
`search_reference_manual`. `stackoverflow.com` was in the original
allowlist but had to come out — confirmed live, Anthropic's crawler is
blocked from it (`400 ... not accessible to our user agent`), so it would
never have returned anything anyway. Deliberately the basic
`web_search_20250305` tool, not the dynamic-filtering `web_search_20260209`
variant — measured live against the same query, dynamic filtering's
code-execution pass more than doubled real cost ($0.15 vs $0.065) for no
quality difference on an already-narrow 3-domain allowlist. `max_uses: 3`
caps the worst case within a single turn.

`adjust_view` (2026-08-02) is the reliable replacement for the entire
`ZOOM_*`/`PAN_VIEW_*` command-registry family, all of which are now
`needs-testing` (disabled) — a live "zoom out 40%" request exposed that
several of them (`ZOOM_OUT`, `ZOOM_OUT_CENTERED`, `ZOOM_HALF`) silently
activate a MicroStation tool and leave the view waiting on a manual
"select point" click that never arrives headlessly, despite returning
`OK`. Root cause (confirmed by reading the code, not guessed): `scripts/
keyin_batch.py`'s probe calls `SendKeyin` then immediately `SendReset`
with no check for a pending prompt in between, so a fast/no-error return
was mistaken for completion; separately, `WZTCCommandRegistry.bas`'s
`CheckCloseOutGuard` (whose whole job is catching exactly this
activate-and-abandon pattern) only inspects `COMMAND:` steps, not bare
`KEYIN:` steps, so this class slipped through both the harvest-time probe
and the runtime safety gate. `adjust_view` sidesteps the whole class by
setting `View.Center`/`Extents` directly via COM (`view_capture.
navigate_view`, already proven live for the auto-focus screenshot) —
never touches `CadInputQueue`, can't leave a pending click, and supports
an exact percentage no registry key-in could ever offer. `view_capture.
get_view_state` is the new read-side counterpart (`View.Extents.X/Y` =
current width/height, `View.Center` = current center) that makes the
percentage math possible. `scripts/keyin_batch.py`'s `promote()` no
longer auto-qualifies `kind="view"` candidates for `verified-headless-
safe` on a clean probe return alone — future view-kind harvests default
to `needs-testing` until a human confirms in the IDE that the action
completes with no click.

## Eval harness (`eval_harness.py`)

A small fixed set of scenarios that run real turns through `chat_driver.
run_turn` and check the agent's actual tool-call trace / final answer —
first step toward answering "did this prompt/tool/model change break
anything" with something other than anecdote. Not a drawing-correctness
checker (it never looks at the DGN) — that's the natural next layer, not
this one.

```bash
python eval_harness.py                 # run every scenario
python eval_harness.py --list           # print scenarios, no API calls
python eval_harness.py --only sign_code_translation
python eval_harness.py --report ../Bridge/eval-results.json
```

Real API calls (billed) and real `WZTCBridge` calls against whatever model
is currently open — run it against `DELETE.dgn`, not a real project file,
same as any manual panel test. Swaps in isolated stand-ins for
`chat_driver.LOG` (writes to `Bridge/eval-log.tsv`, not the live panel's
`chat-log.tsv`) and `chat_driver.INPUT` (a canned non-answer for `ask_user`,
so a scenario can't hang waiting for a human to type into the real panel)
before running — see the module docstring. First run caught a real
methodology trap worth remembering: an early stub `ask_user` reply that said
"use standard defaults" led the agent to actually place a sign from a
fabricated station/alignment; rewording the stub to a neutral non-answer
made it correctly stop and ask instead — the fix was the test's leading
wording, not the agent.

## Known gaps / not verified yet

- **Multi-session MicroStation**: RESOLVED 2026-08-02. `ms_connect.
  get_microstation_app()` (used by `bridge_client.py`, `hot_reload.py`,
  `view_capture.py`, and the two `scripts/*_batch.py` dev tools) replaces
  the old ambiguous `GetObject(Class=...)` attach — it enumerates every
  Running Object Table entry matching MicroStation's CLSID and requires
  exactly one to have the target VBA project loaded, raising a clear error
  on 0 or 2+ matches instead of silently picking one. This closes the gap
  that was the confirmed likely trigger for a real MicroStation crash the
  same day (two instances running, an ambiguous attach landed on one mid-
  test) — see `ms_connect.py`'s docstring for the incident and the ROT
  mechanics discovered while fixing it.
- **`get_model_context`** isn't implemented (see above).
- **Recipe DSL interpreter** for `keyin_recipe` rows is new code — settings-
  only seeds reuse proven `SendKeyin` strings, but each promoted row should
  still be checked against its known-good call site before trusting it in
  production sheets.
- **`search_reference_manual` recall** is uneven on some reasonable queries —
  not chased down yet, tracked as a known gap rather than fixed blind.
- **Sheet-registry sign codes vs. SignLibrary.bas keys**: most codes as
  printed on a 619 sheet (`get_sheet_requirements`) aren't a 1:1 string match
  to a `SignLibrary.bas` key — `resolve_sign_code` bridges the padding/suffix
  gap and surfaces every ambiguous variant rather than guessing one, but a
  real chunk of sheet-registry codes (mostly R- and W1/W3/W4/W5/W7/W8/W9-
  series and NY-custom signs) aren't in `SignLibrary.bas` at all yet — that's
  a content gap (new sign definitions needed), not something a lookup table
  can paper over.
