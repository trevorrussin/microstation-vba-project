# Authoring a sheet spec

How to turn a NYSDOT 619 standard sheet PDF into a `Data/sheet-specs/<sheet>.json`
that the placement code can execute. Follow this in order. `619-311.json` is the
worked reference — read it alongside this.

The governing principle: **nothing in a sheet spec should come from looking at a
picture.** These PDFs are vector, not scans. Every number and every station
boundary can be extracted as coordinates. Where this process was skipped during
the 619-311 draft, it produced a wrong answer that survived review and was only
caught later by the vector extractor (see "The error this catches" below).

---

## Step 0 — Get the right PDF

Per-sheet PDF from the NYSDOT repository:

```
https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository/<sheet>.pdf
```

Save under `Bridge/captures/`. Record the URL and retrieval date in `sheet.sourceUrl`
/ `sheet.retrieved`.

**Do not transcribe from `2026_1_stdsht_usc_book_3.pdf`.** That is how
`Data/sheet-registry.tsv` was seeded, and it is why `Data/README.md` warns that its
numbers "still need per-sheet verification" — book-page extraction interleaves
columns and yields plausible garbage.

Check the page count. Sheets marked "(2 Sheets)" in `Modules/DesignerRef.bas` have a
second page; do not assume plan-on-page-1 and tables-on-page-2 without looking.

## Step 1 — Map the sheet before extracting anything

```python
words = page.get_text("words")     # (x0, y0, x1, y1, text, block, line, word)
```

Search for `TABLE` to get the coordinates of every table title. That gives you the
x/y windows for every later step, so no window is a guess. For 619-311:

| Region | Window |
|---|---|
| Main plan | x < 400 |
| Notes column | x 405–660 |
| Tables | x > 740 |
| Detail 311A + legend key | x 430–680 |

## Step 2 — Extract the tables by coordinate

Never read a table from a rendered image. Image reads of these tables come back
**truncated and confidently wrong** — during the 619-311 pass, one produced a fluent,
well-formatted transcription of three rows out of seven and gave no sign anything was
missing.

```python
rows = defaultdict(list)
for w in words_in_window:
    rows[round(w[1] / 3.0)].append(w)          # group by rounded y
for k in sorted(rows):
    cells = sorted(rows[k], key=lambda w: w[0])  # order by x
```

Four failure modes, each of which hit the 619-311 pass:

- **A logical row split across two y-bands.** The 45 mph row of table 311-02 landed at
  y=463.0 and y=463.6. If a row looks half-empty, re-group with a coarser key
  (`round(y / 8)`).
- **The window silently eats the last row.** Filtering `w[3] <= y1` drops any row whose
  *bottom* crosses the boundary. This lost the RURAL row of 311-03 and the WARNING FLAG
  row of 311-05, and both extractions looked successful.
  **Always assert the row count against what the sheet shows.**
- **Too-wide x windows pull in the neighbouring column.** A loose notes window swallowed
  table 311-05 and the title block.
- **Multi-line headers** produce many y-bands above the data. Data rows are the ones whose
  first token is numeric.

## Step 3 — Notes and legend

Same technique, constraining **both** x and y to the notes column. The notes share
y-bands with plan callouts on their left and tables on their right; a flat word join
interleaves all three into text like
`RURAL 500 500 500 NOTES: 45 - 50 80/2 * PRECONSTRUCTION 1. SHORT-TERM STATIONARY IS...`.

Transcribe notes verbatim. They carry real constraints — Note 4 on 619-311 ("no workers,
equipment or other vehicles in the buffer space or the roll ahead distance") is a
placement rule, not prose.

## Step 4 — Plan geometry: run the extractor, don't eyeball

```bash
python scripts/extract_plan_geometry.py Bridge/captures/619-311.pdf
```

This is the step that makes the whole thing reliable. It works because:

- Dimension lines are long vertical strokes in narrow x bands. **Their endpoints are the
  segment boundaries**, exactly.
- Each dimension's text label sits at the **midpoint** of the segment it dimensions, so
  matching label centres to segments labels every segment deterministically.
- Symbols (arrow panel, vehicles, signs, hatch) are orange vector paths — colour
  `(1.0, 0.5, 0.0)` — whose bounding boxes give exact stations.
- On a plan with traffic moving up the page, **descending y is downstream**.

Its output gives you the corridor directly, as text:

```
y   182.6 ->   210.2  50'-100' DOWNSTREAM TAPER
y   258.0 ->   303.2  ROLL AHEAD DISTANCE (SEE TABLE 311-04)
y   364.8 ->   406.3  BUFFER SPACE (SEE TABLE 311-02)
y   406.3 ->   479.9  LANE TAPER (SEE TABLE 311-02)
y   479.9 ->   509.2  SHOULDER TAPER (SEE TABLE 311-02)   <- left column
y   479.9 ->   572.1  A (SEE TABLE 311-03)
y   572.1 ->   650.7  B (SEE TABLE 311-03)
y   650.7 ->   728.6  C (SEE TABLE 311-03)
```

Write `corridor.zones` from that list in y order. Gaps between consecutive segments are
real and meaningful — the 210.2→258.0 gap on 619-311 is the undimensioned WORK AREA.

### The error this catches

**Read the DATUM SHARING section before assigning order-table rows.** Two dimensions
starting at the same y are measured from the same point, which means the shorter lies
*inside* the longer rather than following it.

On 619-311, `SHOULDER TAPER` (479.9→509.2) and `A` (479.9→572.1) share datum 479.9. The
shoulder taper is therefore an **overlay inside gap A**, not a sequential station upstream
of it. The first draft of the spec — written by looking at the drawing — got this wrong
and placed every advance warning sign 120 ft too far upstream. On a not-to-scale drawing
the difference is invisible by eye and obvious in the coordinates.

Encode an overlay with `consumesStation: false`, `containedIn`, and a `stationAnchor`, and
list it under the alignment's `overlayZones` rather than `rows`.

### What still needs judgement

Anchors — *which end* of a zone a symbol attaches to — come from comparing symbol blob
y-extents to segment boundaries. That comparison is arithmetic, but deciding what it means
is not. Two worked examples from 619-311:

- Roll ahead ends at y=303.2; the vehicle body spans 304.7–348.5. So the dimension meets the
  vehicle's **downstream (front)** end — the roll ahead distance is the clear space *ahead*
  of the vehicle, toward the work area.
- The arrow panel spans 439.0–459.6, inside the lane taper segment 406.3–479.9 and hard
  against its upstream end. So it anchors at the lane-taper/shoulder-taper junction, not at
  the upstream end of the shoulder taper.

Record the coordinates you reasoned from in the field's `note`. That makes a reviewer's
spot-check cheap and makes a fabricated anchor visible.

## Step 5 — Encode

Write `corridor` first; every other section references its zone ids. Then `tables`,
`orderTable`, `signs`, `symbols`, `annotations`, `notes`, `rules`.

Mark every section `confidence: "verbatim"` (from the text layer) or `"drawing"` (derived
from vector geometry). If a `drawing` value ever contradicts the PDF text, the text wins.

Watch for sheet-driven values that look like code concerns:

- **Legend substitution.** Table 311-03's XX/YY columns are *sign legend text*, not
  distances — W20-1 reads "ROAD WORK XX". This is what resolves SignLibrary's
  Ahead-vs-1000-FT variant ambiguity. Check every sheet for placeholder letters on signs.
- **Vehicle-mounted signs.** NYW8-33 "LANE CLOSED" is mounted on the protective vehicle, not
  a roadside post. Set `postMounted: false` and `mountedOn`.
- **Ranges, not values.** Table 311-04 roll ahead is min/max. Don't collapse it.
- **Band collapse.** Table 311-02 has exactly three shoulder columns (≤4 / 5–7 / ≥8 ft).
  Never interpolate per-foot values the sheet doesn't print.
- **Applicability limits.** 619-311's tables stop at 55 mph. Record the real range; don't
  assume the app's dropdown.

## Step 5.5 — Declare `tableRoles`

Add a `tableRoles` section mapping canonical role names (`taperAndBuffer`,
`advanceWarningSpacing`, `rollAheadDistance`, `signSizes`, `protectiveVehicle`) to this
sheet's actual table ids. `mcp-server/sheet_spec.py` and `scripts/validate_sheet_spec.py`
key off these roles, not off literal `"311-NN"` strings — that is what makes the same
Python code work on a sheet whose analogous table happens to be numbered differently, or
that references a shared 619-011 table instead of reprinting its own. See 619-311.json for
the worked example.

## Step 6 — Validate (same tooling for every sheet, not a new script per sheet)

```bash
python scripts/validate_sheet_spec.py Data/sheet-specs/<sheet>.json
```

This script is sheet-generic — it reads `tableRoles`, `corridor`, and `inputs` from the
spec itself, so a second sheet needs no code changes here, only a correct spec. Four
gates, tracked per sheet in `Data/sheet-specs/STATUS.md`:

1. **Structural** — every zone/table/sign cross-reference resolves. Catches little on its
   own but is free.
2. **Domain invariants** — the relationships that hold on any 619 WZTC sheet with the
   relevant table role: a skip line is exactly 40 ft, channelizing devices are always skip
   lines + 1, taper length is monotonic in both speed and lateral shift. One wrong digit in
   any taper triplet breaks at least one. If your sheet's `tableRoles` don't include a
   given role, the corresponding checks are skipped with a `WARN`, not silently passed —
   read the warnings.
3. **Round-trip** — re-extract every cell from the PDF and diff against the JSON. Not a
   generic script: table layouts vary too much across the catalog for one blind differ to
   be trustworthy (it would either miss real errors or need endless special cases). Instead,
   write `Bridge/roundtrip/<sheet>.py` using the shared primitives in
   `scripts/pdf_table_extract.py` (`words_in_window`, `group_rows`, `squash`,
   `assert_row_count`) — see `Bridge/roundtrip/619-311.py` for the worked pattern. This is
   the only unfakeable gate; a sheet is not done until it diffs clean.
4. **Live build** — `python scripts/live_build_check.py <sheet> --speed ... --lane-width ...
   --shoulder ... --area ... --road-type ... --category ...` calls `BUILD_WZTC_ORDER_TABLE`
   through the bridge with a real MicroStation session open and diffs the returned rows
   against the spec's own `orderTable` (present labels + `excludedRows` + overlay zones as
   forbidden labels) — sheet-generic, no per-sheet script needed here. Requires MicroStation
   open with WZTCBridge polling.

Update `Data/sheet-specs/STATUS.md` with the result of each gate. A sheet is `done` only
once all four pass.

## Step 7 — Hand back

State explicitly:

- Which sections are `verbatim` vs `drawing`
- Round-trip result (cell counts compared, failures)
- Any DATUM SHARING relationships found, and how you encoded them
- Anything the sheet shows that the current code contradicts → `knownCodeDeviations`

---

## Reusing work across sheets

The 619-3xx family shares a skeleton: advance signs → taper(s) → buffer → protective
vehicle → roll ahead → work area → downstream taper → END ROAD WORK. 619-302 / 311 / 313 /
317 / 325 are lane-closure variants with near-identical corridors.

Do one sheet per family carefully, then author siblings as diffs against it. An unexpected
structural difference from the family reference is a flag to slow down, not to paper over.

**Do not assume the per-sheet tables repeat.** Every sheet reprints its own numbered tables
(302-02, 311-02, 317-02…). They may well carry identical values, but that must be *verified
by extraction*, not assumed — assuming it yields silently wrong numbers that pass every
validator.

---

## Live-build playbook (`<sheet>.build.md`)

After the JSON is gate-green, add a companion markdown next to it:

```
Data/sheet-specs/619-311.json
Data/sheet-specs/619-311.build.md
```

Point at it from the JSON:

```json
"sheet": {
  "number": "619-311",
  "buildGuide": "619-311.build.md",
  ...
}
```

**Split of concerns:**

| Lives in | Content |
|---|---|
| `<sheet>.json` | Machine prefs the compiler must obey (`annotationStyle`, channelizing `representation`, `annotations`, `rules`) |
| `<sheet>.build.md` | Human/agent tips: preferred call path, visual QA checklist, NYW8-33 handoff, do-nots, script pointers |

The agent loads the playbook via `get_sheet_requirements` / `get_sheet_build_guide` /
`get_plan_status.buildGuidePath`. When a live build surfaces a new preference, update
the JSON (if code must obey) and/or append the tip to `.build.md` — do not leave it only
in `dev-notes/agent-log.md`. See `619-311.build.md` as the worked example.

