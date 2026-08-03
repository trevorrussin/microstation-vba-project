# Sheet specs

One JSON file per NYSDOT standard sheet, holding **everything on that sheet** in a form
the agent and the placement code can read directly: the tables verbatim, the station
sequence, the symbol anchors, which things get dimensioned, the printed notes, and the
machine-checkable rules those imply.

Currently complete: **619-311** (plan sheet) and **619-011** (shared table library). Every
other sheet still falls back to `Data/sheet-registry.tsv`, which is a six-token summary and
is not sufficient to draw from — see "Why this exists" below. See `Data/sheet-specs/STATUS.md`
for the batch plan and per-sheet gate status as more sheets are authored.

## Two sheet kinds

Most 619 sheets are **plan sheets** (`sheet.kind` unset or `"plan"`): a corridor, an order
table, signs placed on an alignment. `619-311.json` is the worked example.

A few sheets are **reference libraries** (`sheet.kind: "referenceLibrary"`), like 619-011
("Work Zone Traffic Control General Tables and Legend"): no corridor, no order table, no
signs — just the master `tables`, a `legend`, and `details`. Plan sheets excerpt from these
rather than reprinting every value; `619-011.json`'s `knownExcerpts` section documents which
of 619-311's tables are confirmed (by cell-for-cell comparison, not assumption) to be an
excerpt of a 619-011 master table. `scripts/validate_sheet_spec.py` branches on `sheet.kind`
and runs a lighter structural check for reference-library sheets (no corridor/signs to
cross-reference).

## Why this exists

`Data/sheet-registry.tsv` stores, for 619-311, a pipe list of sign codes and a pipe list
of element names. That is enough to answer "does this sheet have an arrow panel", and
nothing else. It cannot answer:

- In what order do the advance warning signs appear?
- How far apart? (`sheet-registry.tsv` has no numbers at all; `WZTCRules.ComputeSpacing`
  has numbers, but several of them match no table on the sheet.)
- Where exactly does the arrow panel go?
- Which zones get a dimension, and what does the sheet call them?
- What text goes on the W20-1 sign?

Because none of that was readable, the placement code hardcoded assumptions and each
rebuild drifted a different way. `619-311.json` is the fix: the sheet becomes data.

## Reading a spec

The file is plain JSON, read directly by Python. It deliberately does **not** go through
the VBA bridge — it is a static data file, and a MicroStation round trip to read a file
on disk buys nothing. VBA receives resolved numbers as bridge parameters, the same way
it already receives `sheetElements`.

```python
import json, pathlib
spec = json.loads(pathlib.Path("Data/sheet-specs/619-311.json").read_text())
```

## Sections

| Section | What it holds |
|---|---|
| `sheet` | Identity, revision, provenance, source URL, local PDF copy |
| `applicability` | Road type, closure, speed range, lane widths, shoulder bands, area types |
| `inputs` | The parameters this sheet needs, with allowed values and which tables consume them |
| `tables` | 311-01 … 311-05, verbatim, as structured lookups |
| `geometry` | Stationing convention, cross section, lateral offsets |
| `corridor` | Ordered zones in direction of travel — the layout backbone |
| `orderTable` | The same zones split into the two alignment walks the order-table engine uses |
| `signs` | Per-sign order, legend substitution, flags, sizes, mounting |
| `symbols` | Arrow panel, protective vehicles, channelizing devices, spotter, work area hatch |
| `annotations` | Exactly what is dimensioned and labeled, in the sheet's own wording |
| `details` | Detail 311A |
| `notes` | The five printed notes, verbatim, plus plan callouts |
| `rules` | Machine-checkable assertions with the failure each one guards against |
| `knownCodeDeviations` | Where the current VBA disagrees with the sheet, and which section settles it |
| `tableRoles` | Canonical role name -> this sheet's actual table id, so generic tooling never hardcodes `"311-NN"` |
| `legend` | (reference-library sheets only) Symbol/description rows, e.g. 619-011's 33-item WZTC legend |
| `knownExcerpts` | (reference-library sheets only) Which plan sheets are confirmed to reprint an exact excerpt of one of this sheet's tables |

### `corridor` is the backbone

`corridor.zones` is an ordered list in **direction of travel** — first thing a driver
meets to last. Each zone is one of:

- `kind: gap | taper | buffer | clearance` — has a length, so it becomes a station-table
  row and usually gets a dimension
- `kind: sign` — a sign position
- `kind: symbol | workArea` — placed at a station but contributes no length
  (`lengthSource: null`)

`lengthSource` is declarative, either a table lookup or a fixed range:

```json
{ "table": "311-02", "column": "laneTaper",
  "lookupBy": ["preconstructionPostedSpeedMph", "laneWidthFt"] }
```

```json
{ "fixedRange": { "minFt": 50, "maxFt": 100 } }
```

`orderTable.alignments` then projects those zones onto the upstream and downstream
walks, each with its own station 0 at the work area edge. `excludedRows` records the
rows that the generic default emits but this sheet does **not** have, with the reason —
that list is the direct antidote to `WZTCRules.GetDefaultUpstreamItems` returning the
same seven rows for every sheet.

### `confidence`

Every section carries one:

- `verbatim` — taken from the PDF's vector text layer via coordinate-grouped word
  extraction. Tables, notes, legend. Trust these as printed.
- `drawing` — measured off the drawing geometry (anchors, lateral extents, which end of
  a zone a symbol sits at). Correct as read, but if a value here ever contradicts the
  PDF, the PDF wins.

### `rules`

Each rule pairs an assertion with the `commonFailure` it guards against, so a QA pass can
report "violated arrow-panel-anchor" instead of "looks off". These are the checks worth
running before asking the engineer to eyeball anything.

## Validating

```bash
python scripts/validate_sheet_spec.py Data/sheet-specs/619-311.json
```

This script is sheet-generic (reads `tableRoles`/`corridor`/`inputs` from the spec, not
hardcoded to 619-311). Checks structural integrity (every zone/table/sign cross-reference
resolves), transcription invariants (a skip line is exactly 40 ft, and channelizing devices
are always skip lines + 1 — a typo in any taper triplet breaks this), and then resolves a
worked example end to end via `mcp-server/sheet_spec.py` and prints the resulting station
table. Pass `--speed`/`--lane-width`/`--shoulder`/`--area`/`--closure`/`--exposure` to
resolve a specific case; omitted ones default from the spec's own `inputs` declarations.

Round-trip against the source PDF (`Bridge/roundtrip/<sheet>.py`, using the shared
primitives in `scripts/pdf_table_extract.py`) and the live-build check
(`scripts/live_build_check.py`) are the other two gates — see `AUTHORING.md` Step 6 and
`STATUS.md` for the full four-gate bar every sheet must clear.

## Authoring another sheet

See `AUTHORING.md` for the full procedure. Summary:

1. Get the actual PDF from the NYSDOT standard sheets repository. Do not work from
   `2026_1_stdsht_usc_book_3.pdf` page text alone — it interleaves columns.
2. Pull the tables with coordinate-grouped word extraction (`scripts/pdf_table_extract.py`),
   not plain `get_text()`. Plain extraction scrambles multi-column table rows; grouping
   words by rounded `y` and sorting by `x` reconstructs them correctly.
3. Read the plan geometry from `scripts/extract_plan_geometry.py`'s vector output for
   anchors and lateral extents, not by eyeballing a rendered image — watch for DATUM
   SHARING (see AUTHORING.md's "The error this catches").
4. Fill `corridor` first; everything else references its zone ids. Declare `tableRoles`
   before writing `tables`.
5. Mark each section `verbatim` or `drawing` honestly.
6. Run all four gates (`validate_sheet_spec.py`, a new `Bridge/roundtrip/<sheet>.py`,
   `live_build_check.py`) before marking the sheet `done` in `STATUS.md`.

The schema is shared, so a second sheet should not need new section types. If it does,
bump `schemaVersion` and note what changed here.
