# Bridge

File-based transport for `Modules/WZTCBridge.bas`. An external process (or a
human with a text editor, for manual testing) writes `request.tsv`, triggers
`VBA RUN [ProjectName]WZTCBridge.RunRequest` in MicroStation, and reads
`response.tsv` for the result. Every executed op is also appended to
`wztc-journal.tsv`, regardless of which entry point (this file pair, the
chat-tool pair below, or a direct `WZTCBridge.ExecuteOp` call) issued it.

These `.tsv`/`.json` runtime files are git-ignored — this README is what
keeps the folder itself in the repo.

## Protocol

One op per line, tab-separated:

```
<reqId>\t<OP_TYPE>\tkey1=val1\tkey2=val2...
```

Response mirrors it, one line per request, in the same order:

```
<reqId>\t<OK|ERROR>\tkey1=val1...
```

`WZTCBridge.ExecuteOpInner` dispatches ~25 op types over this pair (M1's
`PLACE_CELL` proof, M2-M5's query/compute/draw/session ops, M6's command
registry and edit vocabulary) — see `mcp-server/wztc_ops.py` for the full,
current op list with params.

## M1 op: PLACE_CELL (the original transport proof, still the simplest example)

```
0001	PLACE_CELL	cellName=TWZAP_P	ptX=1000	ptY=1000	ptZ=0	angleDeg=0
```

`cellName` must be a valid entry from `CellPlacer.GetCellCatalogue()` (e.g.
`TWZAP_P` — Arrow Panel). `ptZ` and `angleDeg` are optional, default `0`.

Success response:

```
0001	OK	elementId=88213	note=placed TWZAP_P at 1000,1000
```

## Manual test (no Python needed)

1. Open a design file with the WZTC VBA project loaded.
2. Create `request.tsv` in this folder with the line above.
3. In MicroStation's Key-in bar, type:
   `VBA RUN [ProjectName]WZTCBridge.RunRequest`
   (replace `[ProjectName]` with the actual project name shown in the VBA IDE)
4. Confirm a `TWZAP_P` cell appeared at design coordinates (1000, 1000).
5. Check `response.tsv` — should show `OK` and an `elementId`.
6. Check `wztc-journal.tsv` — should show the request and response lines appended.

This proves the VBA-side half of the bridge works before wiring up an
external Python client to trigger it over COM.

## M7 — chat panel protocol

A second, separate file pair + VBA entry point for the in-MicroStation chat
panel (`UserForms/WZTCChatPanel.frm` + `mcp-server/chat_driver.py`), kept
independent of `request.tsv`/`response.tsv` so the chat driver and an
external MCP client (Claude Code) can both be attached to the same live
MicroStation session at once without racing on the same files — each
process's `reqId` counter independently starts at `P1` (see
`bridge_client.py`), so sharing files would risk two different ops
colliding under the same `reqId`.

| File | Written by | Read by |
|---|---|---|
| `chat-tool-request.tsv` / `chat-tool-response.tsv` | `chat_driver.py` (`bridge_client.chat_bridge`) | `WZTCBridge.RunChatToolRequest` |
| `chat-input.tsv` | `WZTCChatPanel.frm` (on Send) | `chat_driver.py` (`InputWatcher`) |
| `chat-log.tsv` | `chat_driver.py` (`ChatLog`) | `WZTCChatPanel.frm` (polled via `WZTCChatTimer.bas`) |
| `chat-history.json` | `chat_driver.py`, after every turn | `chat_driver.py`, on startup |

`chat-tool-request.tsv`/`chat-tool-response.tsv` follow the exact same
protocol as `request.tsv`/`response.tsv` above — `WZTCBridge.
RunChatToolRequest` is a byte-for-byte copy of `RunRequest`, just reading
and writing this second file pair, dispatching through the same
`ExecuteOp`.

`chat-input.tsv`: one line per user message, `<timestamp>\t<message>`,
appended by `WZTCChatPanel.btnSend_Click`.

`chat-log.tsv`: one line per event, `<timestamp>\t<TYPE>\tkey=val...`,
appended by `chat_driver.py`'s `ChatLog` class:

| TYPE | Fields | Meaning |
|---|---|---|
| `THINKING` | `text` | Model's in-progress reasoning/commentary before a tool call or the final answer |
| `TOOL_CALL` | `name`, `input` (JSON) | A tool is about to run — `input` is for display only, never parsed structurally on the VBA side |
| `TOOL_RESULT` | `name`, `status` (`OK`/`ERROR`), `summary` | That tool call finished |
| `ASK_USER` | `question` | Model is blocked waiting for the engineer's reply (see `chat_driver.ask_user`) |
| `FINAL` | `text` | Turn complete |
| `ERROR` | `note` | The turn failed outright (API error, etc.) — the driver keeps running, ready for the next message |

**Must be CRLF, not bare LF** — same requirement as `Data/sheet-registry.tsv`
/ `Data/command-registry.tsv` (see `Data/README.md`). `WZTCChatTimer.bas`
reads this file with `Line Input#`, which silently reads an all-bare-LF file
as a single giant line instead of one line per event — confirmed live during
M7 Stage 1 bring-up, not theoretical. `chat_driver.py`'s `ChatLog` opens the
file with `newline="\r\n"` specifically to guarantee this regardless of
platform defaults.

Any text value that could contain a real newline or tab (model output,
tool-call JSON) is flattened to a single physical line before being written
(`chat_driver._flatten`) — an embedded newline would otherwise look like a
second, malformed log entry to the line-count-based reader.

`chat-history.json`: the full conversation as a plain JSON array of
`{"role": ..., "content": ...}` messages (Anthropic wire format), reloaded
on `chat_driver.py` startup so closing/reopening the panel — or restarting
the driver process — doesn't lose context. No session/multi-conversation
concept: one file, one ongoing conversation; delete it to start fresh.

**Lifecycle**: `chat_driver.py` is user-started (`python chat_driver.py`),
not auto-launched by VBA — see the M7 plan's Stage 6 decision for why
(`Shell()` has no precedent in this repo and no proven pattern to lean on).
If nothing responds in the panel, check whether the driver process is
actually running before assuming something's broken.
