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
