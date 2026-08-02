# Agent Work Log

Shared cross-tool log. This project gets worked on from both Claude Code and
Cursor â€” neither tool can see the other's session history or memory, so this
file is the bridge. **Whichever AI tool finishes a non-trivial, non-obvious
piece of work appends an entry here before finishing.**

"Non-obvious" is the bar â€” don't log routine edits that are already clear from
`git log`/`git diff`. Log the things a future session (in either tool) would
otherwise have to rediscover: a design decision and why, a bug whose root
cause wasn't where it looked, a gotcha specific to this codebase, a dead end
that shouldn't be retried.

## Entry format

```
## YYYY-MM-DD â€” <tool> â€” <one-line summary>
<2-6 sentences: what changed, why, anything surprising. Reference specific
files/functions. If it's a bug fix, say what the real root cause was, not
just the symptom.>
```

Newest entries at the bottom (append-only, chronological â€” matches `git log`
ordering conventions already used elsewhere in this repo).

## How each tool uses this file

- **Claude Code**: reads this at the start of a session when the user
  mentions work happened elsewhere ("I did some work in Cursor"), or
  periodically checks it â€” see the pointer in `CLAUDE.md`. Anything
  load-bearing gets folded into Claude Code's own persistent memory
  (`~/.claude/projects/.../memory/`) from here, same as if the user had
  described it directly in chat.
- **Cursor**: instructed via `.cursor/rules/agent-log.mdc` to append here
  after finishing meaningful work.

This is a manual bridge, not automatic sync â€” each tool only reads/writes it
when its own instructions or the user tells it to.

---

## 2026-08-02 â€” Claude Code â€” sign rotation now matches current view angle

`DrawSign.bas`/`WZTCExec.bas`: sign face cells were rotating to match the
alignment/perpendicular direction, which is mathematically guaranteed to
flip some signs upside-down and has nothing to do with the view. Fixed to
capture the view's rotation (`ViewRotationAngleDegrees`, via `v.Rotation`
before it gets reset to identity) and use that for `ACTIVE ANGLE` instead â€”
verified live via direct COM reads that a placed cell's rotation exactly
matches the view's rotation at placement time, both unrotated (0Â°) and
rotated (60Â°) cases. See Claude Code memory `feedback-sign-rotation-matches-view`
for the full writeup.

## 2026-08-02 â€” Claude Code â€” bounded reuse for Bridge/results_*.tsv

`WZTCBridge.bas`'s `WriteResultRows` used to write one new
`results_<reqId>.tsv` file per query call, forever (19+ leftover files
found after one session). Changed to a bounded pool of 8 reusable
filenames (`results_slot0.tsv`..`results_slot7.tsv`, chosen via
`reqId mod 8`) â€” each write already truncates/overwrites via `Open ...
For Output`, so no deletion is involved anywhere, file count just stays
bounded. Only safe because a single batch never has more than 8
multi-row-result ops today (`call_batch` in `bridge_client.py` is only
ever called with one op at a time) â€” if real multi-op batching gets
added later, bump `RESULT_POOL_SIZE` past the largest expected batch.

Hit a real VBA gotcha along the way: a new module-level `Const` placed
between two existing procedures (instead of in the top-of-file
declarations block) fails to compile with "Only comments may appear
after End Sub, End Function, or End Property" â€” moved
`RESULT_POOL_SIZE` up next to `BRIDGE_DIR` etc. to fix. See Claude Code
memory `feedback-vba-compile-error-recovery` for the full incident,
including a clean scriptable way to detect/dismiss/reset a blocking VBA
compile-error dialog via `VBE.CommandBars` (no manual clicking, no
guessing at dialog internals).

## 2026-08-02 — Cursor — keyin probe 3s hang timeout + wave8 promote

`scripts/keyin_batch.py`: live `SendKeyin` now runs in a child process with a hard **3s** timeout (`SENDKEYIN_TIMEOUT_SEC`); hangs are recorded as `HANG` ? `unsafe-blocked` and the batch continues. Also stopped executing `tool`/`datapoint` kinds (they activate-and-wait — `TITLEBLOCK PLACE` was the wave8 forever-hang). Wave8 sparse-category harvest probed/promoted **+398** registry rows; 4 PDF embed/layers keyins hit the new timeout and were added to the skip list. Registry ~2208 / ~1959 verified.

## 2026-08-02 — Cursor — drawing recipe probe (element-delta bar)

Added `scripts/recipe_batch.py` + `Data/recipe-candidates.tsv`. Unlike settings
`keyin_batch.py`, a drawing recipe only promotes when graphical element count
on `DELETE.dgn` increases (plus COM alive / 3s step timeout). First live results:
`PLACE_LINE` and `PLACE_SHAPE_CONSTRAINED` ? `verified-headless-safe`;
`HATCH_ICON` (Legacy two-identical-seed pattern) completed without hang but
`+0` elements ? stays `needs-testing`. Circle/block/arc/smartline not seeded
(no in-repo CadInputQueue sequence).

## 2026-08-02 — Cursor — Phase C edit direct_api + Phase B WZTC registry rows

Phase C: added `ExecCopy/Rotate/Scale/Mirror/ArrayElementByID` in
`WZTCExec.bas` using Element API patterns live-proven on `DELETE.dgn`
(Clone+Move, ScaleUniform, Matrix3d Z-rotate Transform, Mirror two-point) —
wired through `WZTCBridge` + MCP tools. No CadInputQueue inventing.
Phase B: catalogued existing place bridge ops as `direct_api`
`verified-headless-safe` rows (`PLACE_CELL` was flipped from bare
`unsafe-blocked` COMMAND). Re-import `WZTCExec.bas` and `WZTCBridge.bas`
in the VBA IDE before exercising the new edit ops.

## 2026-08-02 — Cursor — hatch Element API + PLACE_ARC / PLACE_TEXT_LABEL

CadInputQueue `HATCH ICON` (Legacy twin-seed) stays unreliable headlessly
(+0 elements on DELETE.dgn). Switched workspace hatch to
`CreateHatchPattern1` + `ClosedElement.SetPattern(..., Matrix3dIdentity)`
— live `HasPattern=True`. New bridge/MCP: `HATCH_ELEMENT`, `PLACE_ARC`
(placeArcModeEx=3), `PLACE_TEXT_LABEL` (TEXTEDITOR INSERT_TEXT). All three
plus updated `PLACE_WORKSPACE` verified OK via bridge on DELETE.dgn.
`HATCH_ICON` registry row ? `unsafe-blocked`. Hot-reloaded WZTCExec +
WZTCBridge. Note: PrintWindow captures often omit associative hatch lines
even when HasPattern is True — trust HasPattern / in-app view for hatch.

## 2026-08-02 — Cursor — Tier1-3 general geometry ops

Added Element-API geometry suite to `WZTCExec` / `WZTCBridge` / MCP:
Tier1 place (circle/ellipse/block/polyline/polygon) + symbology; Tier2
copy-parallel (lines), crosshatch/remove-hatch, break-line, extend-line
(recreate, not EndPoint — EndPoint assign hung VBA), fillet/complex
(needs-testing); Tier3 fence block + copy/move/delete contents, select/clear.
TRIM/CHAMFER left interactive-only (no COM ConstructTrim/Chamfer).
Live: Tier1 + symbology + copy-parallel OK on DELETE.dgn before
`LineElement.EndPoint` hang wedged VBA `[running]`. Reset/Ctrl+Break from
automation failed — user must interrupt VBA (or restart MicroStation), then
hot-reload and re-verify extend/fence/fillet/complex.

## 2026-08-02 — Cursor — Tier1-3 live-verified + complex-string fix

After VBA Reset, phased live verify on DELETE.dgn: extend (recreate-line path,
not EndPoint), break, crosshatch/remove, fence define/copy/undefine, fillet all
OK. `CREATE_COMPLEX_STRING` initially failed compile —
`CreateComplexStringElement1` needs `ChainableElement()`, not `Element()`;
fixed in `ExecCreateComplexString` via `el.AsChainableElement`, hot-reloaded,
live OK (`partCount=2`). All Tier1–3 geometry bridge ops now
`verified-headless-safe` except TRIM/CHAMFER (still interactive-only). Gotchas:
`scale` as a VBA local name conflicts with MicroStation `Scale` (use
`lenScale`); PrintWindow often omits hatch lines — trust `HasPattern`;
VBA `[running]`/`[break]` blocks hot-reload until Reset.
