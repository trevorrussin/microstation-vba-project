# Data

External data files read by both VBA (`Modules/WZTCSheetRegistry.bas`) and any future
Python MCP server — deliberately kept as plain data files rather than VBA modules, since
both sides need to read the same source (per the plan's M4 design).

## sheet-registry.tsv

Maps a 619-series sheet number to the signs, elements, and applicability notes shown on
that sheet. **Seeded incrementally, not exhaustive** — as of this writing it covers 6 of
the 91 current 619 sheets (the most common Short Term scenarios: 301, 302, 303, 307, 310,
311). An unpopulated sheet simply isn't in this file; the tool falls back to manual entry
for those, same as before this registry existed.

**Provenance:** extracted from `2026_1_stdsht_usc_book_3.pdf` (the actual January 2026
NYSDOT Standard Sheets book), via a research agent reading the plain-text PDF dump, not
hand-typed from memory. 619-311's row is flagged incomplete — its spacing-table page
wasn't fully extracted and needs re-verification against the source before being trusted
for real spacing values.

**Sign codes are as printed on the sheet** (MUTCD/NY style, e.g. `W20-1`, `NYW8-33`), not
yet mapped to `SignLibrary.bas`'s internal zero-padded entry names (e.g. `W20-01RA`).
That mapping is a known follow-up — don't assume a code in this file resolves 1:1 to a
`SignLibrary.AddSign` call without checking.

### Format

Tab-separated, one row per sheet:

```
sheetNum  title  roadType  duration  signs  elements  notes
```

- `signs` / `elements`: pipe-separated lists (`|`), matching the same list-encoding
  convention `WZTCExec.bas` uses for vertex lists.
- `elements` uses short internal names (`ProtectiveVehicle`, `MergingTaper`,
  `DownstreamTaper`, `ChannelizingDevices`, `Flagger`, `ArrowPanel`, `CenterlineCones`) —
  these are descriptive labels for now, not yet wired to `DrawElements.GetElementLevel`'s
  numeric element-type indices.
- `notes`: free text — applicability conditions, which Book 3 spacing tables apply (by
  table number, not transcribed values — the actual numbers still need per-sheet
  verification before being trusted as spacing inputs), and known extraction gaps.

### Extending this file

To add another sheet: read its actual page range in `2026_1_stdsht_usc_book_3.pdf`, don't
guess from the sheet title alone — several titles look similar (e.g. every "Right Lane
Closure" variant has different sign sets depending on road type and duration).

**Line endings must be CRLF, not bare LF.** VBA's `Line Input #` (used by
`WZTCSheetRegistry.ReadAllLines`) reads a bare-LF file as a single giant line instead of
one line per row — confirmed by testing, not theoretical. Every file VBA itself writes
(e.g. `Bridge/*.tsv` via `Print #`) is CRLF for exactly this reason. If you edit this file
with a tool that saves LF-only (common on non-Windows editors), convert it back to CRLF
before it'll read correctly.
