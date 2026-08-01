# Data

External data files read by both VBA (`Modules/WZTCSheetRegistry.bas`,
`Modules/WZTCCommandRegistry.bas`) and the Python MCP server — deliberately
kept as plain data files rather than VBA modules, since both sides need to
read the same source.

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
`WZTCSheetRegistry.ReadAllLines` / `WZTCCommandRegistry.ReadAllLines`) reads a bare-LF
file as a single giant line instead of one line per row — confirmed by testing, not
theoretical. Every file VBA itself writes (e.g. `Bridge/*.tsv` via `Print #`) is CRLF for
exactly this reason. If you edit this file with a tool that saves LF-only (common on
non-Windows editors), convert it back to CRLF before it'll read correctly.

## command-registry.tsv

Catalog of named MicroStation command **recipes** the agent may (or may not) run.
Safety is a property of the whole call sequence, never the bare command string alone —
the same token (e.g. `PLACE CELL ICON`) appears at both headless-safe and
`GetInput`-dependent call sites elsewhere in this repo.

Read by `Modules/WZTCCommandRegistry.bas`. The MCP tools `list_registry_commands` /
`describe_registry_command` / `run_registry_command` surface it; `TEST_REGISTRY_COMMAND`
exists only on the VBA side (manual IDE promotion) and is **not** exposed to the agent.

### Columns

```
opName  category  safetyStatus  recipeLines  vbaFunction  requiredParams
optionalParams  createsElements  ownElementOnly  sourceRefs  addedDate
promotedDate  notes
```

- `category`: `keyin_recipe` (interpreted from `recipeLines`) or `direct_api`
  (bespoke `WZTCExec` function; the row is bookkeeping/gating only — call the
  dedicated bridge op, not `RUN_REGISTRY_COMMAND`).
- `safetyStatus` is the enforcement gate:
  - `verified-headless-safe` — only status `ExecuteRecipe` / `RUN_REGISTRY_COMMAND` will run
  - `needs-testing` — catalogued, refused at execution
  - `interactive-only-use-handoff` — known to need a live click; points at `HANDOFF`
  - `unsafe-blocked` — confirmed activate-and-abandon; not promotable without redesign
- `recipeLines` mini-DSL (pipe-separated): `KEYIN:text{param}` / `COMMAND:text{param}` /
  `SETCELL:{cellName}` / `DATAPOINT:{ptX},{ptY},{ptZ}` / `RESET` / `DEFAULTCOMMAND`
- `requiredParams` / `optionalParams`: pipe-separated param names (`level|color`)
- `createsElements` / `ownElementOnly`: `Y`/`N`

### Structural close-out guard

Any recipe with a `COMMAND:` step must also include `DATAPOINT:` and `RESET`,
regardless of `safetyStatus`. Second defense against the
`BBMarkupProcessor.ExecuteAddDimension` / `ExecuteAddCallout` activate-and-abandon
anti-pattern (tool armed, `Status = "Done"`, no points sent).

### Promotion process (`needs-testing` → `verified-headless-safe`)

1. Add a row with `safetyStatus=needs-testing` — inert by construction.
2. Test by hand in the MicroStation VBA IDE (type the keyin, or hand-edit
   `Bridge/request.tsv` with `TEST_REGISTRY_COMMAND` and send the keyin yourself).
3. On success, flip `safetyStatus` to `verified-headless-safe` and fill
   `promotedDate` / `sourceRefs`. No VBA or Python code change needed for
   `keyin_recipe` promotions.

### Seed / current inventory

Built by live COM probe against this install, sourced from **outside** this
repo as well as in-repo call sites: COD OT MicroStation Keyin Reference,
Axiom’s 80+ two-letter key-ins, WSDOT CAE function-key notes, plus CONNECT
long-form `ACTIVE`/`SET`/`LOCK`/`VIEW` variants. Probe results for the latest
external batch are in `Bridge/keyin-probe-results.json`.

`verified-headless-safe` = settings / view / lock / selection keyins that
returned without hanging. `unsafe-blocked` = bare tools, dialogs, and
precision datapoints that return fast but arm UI or need clicks.
`needs-testing` = higher-impact file/undo keyins not batch-probed.
`interactive-only-use-handoff` = dimensions/callouts.

M1–M5 draw ops are **not** catalogued here yet — they already work via
dedicated bridge ops.

Same CRLF requirement as `sheet-registry.tsv`.
