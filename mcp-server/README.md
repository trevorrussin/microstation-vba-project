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
(run it once, and again any time the source PDFs change). Ground
engineer-facing questions about MUTCD/NYSDOT requirements in the returned
excerpt + page citation, not recollection.

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
agent loop (`client.beta.messages.tool_runner`, adaptive thinking, every
`wztc_ops` function + `search_reference_manual` + `ask_user` as tools) behind
`UserForms/WZTCChatPanel.frm` — a modeless in-MicroStation dialog. "Python
owns the brain and the hands; VBA owns only the face": the panel only polls
`Bridge/chat-log.tsv` (via `WZTCChatTimer.bas`, a Win32 `SetTimer` loop — the
first polling mechanism in this codebase) and appends to `Bridge/
chat-input.tsv` on Send; every actual tool call goes through the identical
`WZTCBridge.ExecuteOp` dispatch every other op in this repo already uses, so
watching it work is the same visible MicroStation activity as always.

Requires `ANTHROPIC_API_KEY` set (or an `ant auth login` profile) — picked up
automatically by `anthropic.Anthropic()`, no code for it here. Run
`python chat_driver.py`; it isn't auto-launched by VBA (see `Bridge/README.md`
for why). Full protocol/file-schema docs: `Bridge/README.md` → "M7 — chat
panel protocol".

`ask_user` is a tool, not a special stop condition — it writes an `ASK_USER`
line and blocks (a plain polling wait, safe here since this process has no UI
thread to keep responsive, unlike the VBA panel) until the engineer replies
in the panel, matching the documented pattern for promoting question-asking
to a tool call in agentic loops.

## Known gaps / not verified yet

- **Multi-session MicroStation**: `win32com.client.GetObject` attaches to
  *a* running MicroStation instance; behavior with two sessions open is
  unconfirmed.
- **`get_model_context`** isn't implemented (see above).
- **Recipe DSL interpreter** for `keyin_recipe` rows is new code — settings-
  only seeds reuse proven `SendKeyin` strings, but each promoted row should
  still be checked against its known-good call site before trusting it in
  production sheets.
- **`chat_driver.py`'s live agent loop is untested** — everything testable
  without a live Anthropic API key has been verified directly (tool schema
  generation for every wrapped function, the CRLF-safe chat-log writer, the
  input-file polling/cursor logic), but the actual `run_turn` loop — the
  `tool_runner` call, `generate_tool_call_response()`-based history
  mirroring, multi-turn context across separate turns — has not yet been
  exercised against a real API key. First live test is pending.
