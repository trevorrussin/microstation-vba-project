# WZTC Designer MCP Server

Exposes the WZTC agent tool surface (plan `Layer 5`, milestone M5) as a real,
connectable MCP server: query the model, get deterministic MUTCD spacing
numbers, and draw signs/elements/work space in MicroStation — without
clicking through the 8-step wizard.

This is the Python side of the bridge. It does not replace `WZTCBridge.bas` —
it drives it, over the file + COM-keyin transport proved in M1/M2.

## What this talks to

```
this server  ──►  win32com (reads: attach to live MicroStation session)
             ──►  Bridge/request.tsv + a COM-sent keyin (writes)
                        │
              WZTCBridge.RunRequest  (VBA, in MicroStation)
                        │
              WZTCExec / WZTCQuery / WZTCRules / WZTCSheetRegistry
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
   `WZTCSheetRegistry.bas`, plus their existing dependencies —
   `PerpPlacement.bas`, `DrawSign.bas`, `SignLibrary.bas`, `DrawElements.bas`,
   `SharedState.bas`). If you've just pulled changes to any of those files,
   re-import them in the VBA IDE first (delete the old module, File → Import
   File) — see the repo's `CLAUDE.md` File Sync Protocol.
2. **Python 3.10+** with the packages in `requirements.txt`:
   ```
   pip install -r requirements.txt
   ```
   (`mcp` is the official MCP Python SDK; `pywin32` is the COM bridge to
   MicroStation. Both install cleanly on this machine as of this writing —
   `mcp` 2.0.0, `pywin32` 312.)

## Register with Claude Code

```
claude mcp add wztc-designer -- python "c:\repos\microstation-vba-project\mcp-server\server.py"
```

This registers it as a stdio-transport server (the default — see
`server.py`'s `mcp.run()` call). Once registered, restart/reload Claude Code
and the 17 tools below become directly callable in a session.

To remove: `claude mcp remove wztc-designer`.

## Tool surface (M5)

| Group | Tools |
|---|---|
| Query | `find_elements_near`, `station_to_point`, `get_alignment_stationing`, `list_levels`, `classify_site_features` |
| Compute | `compute_spacing`, `get_sheet_requirements` |
| Draw | `place_perp_line`, `place_sign`, `place_workspace`, `place_element_run`, `place_cell`, `set_sign_attributes`, `handoff` |
| Session | `undo_last_op`, `get_journal`, `list_deferred_handoffs` |

**Not yet exposed — M6 scope per the plan:** `move_element`, `change_level`,
`edit_text`, standalone `delete_element`. `WZTCBridge.bas` has no dispatch
case for these yet; `undo_last_op` uses the equivalent delete-by-ID VBA
primitive (`WZTCExec.ExecDeleteElementsByID`) internally, but it isn't wired
up as its own callable op.

**Not yet exposed — never built in M2:** `get_model_context` (listed in the
plan's Layer 5 table but no `GET_MODEL_CONTEXT` op exists in
`WZTCBridge.bas`). Compose equivalent context today from `list_levels` +
`get_alignment_stationing` instead.

## The engineering-judgment boundary

This server never computes a spacing, taper length, or sign size itself.
`compute_spacing` and `get_sheet_requirements` wrap `WZTCRules.bas` /
`WZTCSheetRegistry.bas` so those numbers stay deterministic and traceable to
a PE-reviewed table — the calling agent decides *what* to place and *how to
respond to a site condition*, never invents a number that belongs in one of
those two tools.

Every `place_*` tool takes an optional `reason`. It's not decoration — it
rides through untouched into `Bridge/wztc-journal.tsv`, which is the audit
trail a PE would need to answer "why is that sign there." Pass it whenever a
placement is adjusted from the default (an obstruction dodge, a non-standard
station), per the plan's worked example:

```
station_to_point(1, 320)          → x, y, tangent
find_elements_near(x, y, 25)      → utility pole cell 4.2 ft off proposed post location
place_sign(..., reason="shifted 4 ft off perp — utility pole at 3+20")
```

## Dimensions and callouts — the "red list"

`DIMENSION SIZE WITH LINES` and `TEXTEDITOR PLACENOTE` have no programmatic
`CadInputQueue` precedent anywhere in this repo (see `WZTCExec.bas`'s
header). `handoff(kind="dimension" | "callout", ...)` queues these instead of
faking success — it returns `DEFERRED`, not `OK`. After drawing everything
else, call `list_deferred_handoffs()` and tell the engineer what still needs
a few clicks through the existing `PlaceElements`/`PlaceCells` forms.

## Undo

`undo_last_op()` does **not** use MicroStation's own undo stack — its exact
grouping behavior across a multi-element op (e.g. a sign's post + face +
text + arc, placed in one `place_sign` call) has not been verified in the
IDE, and the plan explicitly flags the `MARK` keyin/API as unconfirmed.
Instead it walks `Bridge/wztc-journal.tsv` backward for the most recent op
that created elements and isn't already undone, and deletes exactly those
element IDs (`WZTCExec.ExecDeleteElementsByID`) — deterministic, and
independently tested by `Debug/DebugAgentLoopTest.bas` without needing to
guess at undo-stack semantics.

## Known gaps / not verified yet

- **Multi-session MicroStation**: `win32com.client.GetObject` attaches to
  *a* running MicroStation instance; behavior with two sessions open is
  unconfirmed (flagged in the plan, still true here).
- **`get_model_context`** isn't implemented (see above).
- **Edit vocabulary** (`move_element`, `change_level`, `edit_text`) is M6.
