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
