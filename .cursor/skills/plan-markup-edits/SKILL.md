---
name: plan-markup-edits
description: >-
  Apply engineer PDF markups to MicroStation plan sheets: estimate/pay-item
  tables, NOTES, plan callouts with leaders, and short tags without leaders.
  Use for markups, redlines, estimate-table edits, note renumbering, or
  place/update callouts on any project. Always run the bounding-box text
  overlap gate before claiming place/edit done.
---

# Plan-sheet markup edits (table + notes + callouts)

Follow this process **every time**. Do not invent layout. The markup PDF (and any SAMPLE table on it) is authority for format.

**When extending this skill after a new lesson:** write the **general process** (how to classify markup, what to place, what to verify). Do **not** bake in project-specific names, sheet numbers, pay-item IDs, or label strings from one job.

## 0. Before editing

1. Confirm the **active file/model** (host sheet vs activated reference). Callouts often live in an activated **ref** (`fea_*.dgn`); tables/notes may live in a sheet model (`XS Stacked Views` / `tsp_*.dgn`).
   - **Never open/switch DGN files** to get there — if the wrong file is active, stop and ask the engineer to switch (see `.cursor/rules/never-switch-dgn.mdc`).
2. **MCP namespace:** use **`user-wztc-designer`** (MicroStation) for all live place/edit/view/capture. Do **not** call `user-autocad` unless the engineer explicitly asked for AutoCAD/DWG. Same tool names exist on both products — namespace is the gate.
3. Open the markup PDF. Classify each FreeText/cloud:
   - **Table / quantities / item numbers**
   - **NOTES** (wording + renumber after deletions)
   - **Plan callouts** (green highlight + red font + **leader** = place/update callout)
   - **Ignore** yellow instruction notes (linetype, hexagon, x-walk, freeze, move pole, etc.) unless the user explicitly expands scope
3. If notes are renumbered (delete a note, shift later numbers), plan **SEE NOTE n** updates on existing callouts in the same effort.

## 1. Estimate / LIST OF PAY ITEMS table

### Authority for format

- Prefer a **SAMPLE** table on the markup PDF (e.g. page with `LIST OF PAY ITEMS`).
- Match: column headers, **center** description column, wrap only at natural phrase breaks.
- Short descriptions stay **one line** when they fit the DESCRIPTION column (markup `SHOW AS 1 LINE` is mandatory).
- Do **not** invent narrow wraps.

### Hard rule: TextNode vs TextElement

- `TCSH_P_T` **TextNodes** in this environment auto-split at ~22 characters even when `LineLength` is raised. That is wrong for pay-item tables.
- Place / rebuild body **descriptions** as **`CreateTextElement1`** (single line per visual line), centered on the DESCRIPTION column X.
- Item / Unit / Quantity can stay as short TextNodes if they already do not wrap.
- Multi-line SAMPLE wraps = **two (or more) TextElements** stacked with the same pitch as existing 2-line rows (~2.5 master units here), not one auto-wrapping TextNode.

### Content workflow

1. Read current table body (level, X columns for ITEM / DESCRIPTION / UNIT / QUANTITY, row Y pitch from border cells).
2. Build the row list from markups (add/rename/qty/order).
3. Rewrite item/unit/qty TextNodes with `DeleteAllTextLines` + `AddTextLine` + `Rewrite` **only** for those short cells.
4. Remove old description TextNodes; place SAMPLE-matching TextElements.
5. Verify: one-line rows that the SAMPLE/markup require are truly 1 line; scan TextLinesCount / element type.

### Layout rules (cell / row / column)

- **Cell padding:** Give text breathing room. Do **not** let description (or other cell) text touch horizontal or vertical blue borders — keep a clear gap (target ≥ ~0.4 master units from Range to band edge here).
- **Row height:** Just tall enough to comfortably contain the text (1-line vs 2-line). Avoid huge empty rows; grow a band only when 2-line pitch + padding does not fit.
- **Column widths:** Wide enough that text is not awkwardly wrapped, but not dramatically wider than the SAMPLE / existing grid. Prefer SAMPLE one-line descriptions over inventing wraps; do not widen columns unless the markup/SAMPLE requires it.
- 2-line description stack pitch: must exceed TextElement **Range** height
  (here ~5.0) or the two line boxes overlap and look cramped. Target pitch
  ~5.8 with clear air between ink; do not use the old ~2.5–3.4 values.
  Prefer collapsing to one line when the DESCRIPTION column already holds
  similar-length singles.

### Adding rows / extending the grid (hard rule)

- Prefer **text-only** edits when the border grid already has enough rows.
- New row bands: **Clone** an existing full-width horizontal band + `Move(0, dy, 0)`. Do **not** invent geometry.
- Tall column fills that must grow: delete the broken shape and place a new **axis-aligned** `CreateShapeElement1` rectangle. **Never** `ModifyVertex` on live table borders — it scrambles corners into **diagonal** lines (live incident 2026-09-17 on `tsp_02` ESTIMATE OF QUANTITIES).
- Working text helper for this sheet family: `scripts/_fix_bts_table_format.py` (descriptions as TextElements).

## 2. NOTES block

### Authority for format

- Match **how notes looked on the sheet before** (left-justified under `NOTES`, wide lines).
- Markup adds wording (e.g. `AND DETAILS.`) and renumbers after strikeouts.

### Hard rule

- Use the **`NOTES` title** TextNode/component as the template (Justification **0** / left).
- Do **not** clone description-column center-justified text for notes.
- Prefer **TextElements** with wide wraps (note 1–2 often one full line). Avoid narrow TextNode auto-wrap.
- Persist with `SAVE DESIGN` keyin when editing an activated ref.

## 3. Plan callouts + leaders

### What to place (classify from the markup)

Look at each plan FreeText and whether it has a **leader** to geometry:

| Markup | Place as |
|--------|----------|
| Annotation text **with a leader** to a feature | Full callout cell + **visible** leader (§ below) |
| Short **tag / ID label** next to a feature (**no** leader in the markup) | Plain text only — match existing callout/annotation text style (usually black, no highlight fill). **No** leader and **no** red construction line. Put the text at the markup box (PDF→DGN). |
| Explicit user ask to change an existing callout | Update that callout (§ Updating) |

Do not invent leaders for tags that have none on the markup. Do not invent tags the markup does not show.

### Existing callout anatomy (typical annotation callout)

Working stock callouts are one unit (cell **or** shared GraphicGroup) containing:

1. TextNode (right-justified) — callout wording  
2. Short horizontal **stub** line at the text’s right  
3. **Note dimension** (type 33, `DimensionType` 53, style e.g. `ny_Plan`) — stub end → tip on feature, with **terminator/arrowhead** from the dim style  

That note dimension **is** the leader. Text, stub, leader, and arrowhead stay **connected** in every case — only the move rule changes (below).

### How leaders move (snapped vs free)

In both modes the callout stays one connected annotation (text ↔ stub ↔ leader ↔ arrow). They differ only in what translates:

| Tip state | Move text / callout | Arrowhead | Leader |
|-----------|---------------------|-----------|--------|
| **Snapped / associated** to a drawing element | Text can move (including to the other side of the feature); text orientation stays | **Stays fixed** on the snap element (does not translate across the drawing) | Stays attached; **reflows / flips** to the side the text is on |
| **Not snapped** (free tip) | Whole annotation translates together | Moves with the text | Moves with the text |

- Do **not** treat GraphicGroup “move everything” as correct for a **snapped** tip — that would drag the arrow off the element.
- Do **not** leave text, leader, and arrow as unrelated elements — they must stay connected in both modes.
- **Look vs drag behavior (important):** Stock-style callouts are **TextNode + short stub + note-dim leader** (`PrimaryText` usually `*`). That look matches the sheet. True “drag the wording, arrow stays / leader flips” only works when the wording is **on the associative note itself**. Separate TextNode + GraphicGroup will move as a block when Graphic Group lock is on. Prefer stock look unless the engineer explicitly wants drag-flip; then use Place Note text and match style as closely as possible.
- **Associative tip (arrow on feature):** Association Lock on → `PLACE NOTE` → tentative + `SendDataPointForLocate(target, tip)` → `SendDataPoint(textEnd)`. Then set `PrimaryText='*'`, color/level like stock, add TextNode + ~5-unit stub if keeping stock look. Verify tip with `MODIFY ELEMENT` on the leader’s text end (not Move+GG).
- CONNECT VBA `AddReferencePoint` accepts **`Point3d` only** (AssociationPoint = ByRef compile error).
- When editing: preserve existing snap associations; only use whole-unit move for callouts that are actually free.

### Updating existing callouts (SAFE)

- **Do use:** `SetTextLine(n, …)` + `Rewrite` on the TextNode inside the cell, then `CadInputQueue.SendKeyin('SAVE DESIGN')`.
- **Do not use:** `DeleteAllTextLines` on callout TextNodes — it can destroy the cell / fail to persist on attachment-backed elements.
- When notes renumber: update `(SEE NOTE n)` on every affected callout in the same pass.
- Full wording replacements: `SetTextLine` each line; `AddTextLine` only if you need extra lines.

### Adding new callouts with leaders

1. Prefer **clone a working stock callout** (text + stub + note dim already linked). Move the **whole** unit so the dim tip lands on the feature. Set text with `SetTextLine` only.
2. If you must build new: create a **note dimension** (`CreateDimensionElement1` + `DimensionType` 53 / note, style `ny_Plan`) with reference points **tip** then **stub end**. Keep `IsHidden=False` on new dims so the leader and arrowhead display. Match stub/text level and color to stock.
3. **Never** add a separate plain line as a fake leader. Separate lines have no arrowhead and disconnect when the text moves.
4. Link text + stub + note dim (cell when possible; otherwise shared GraphicGroup). Use whole-unit move only when the tip is **not** snapped (§ How leaders move).
5. If the tip should be snapped: associate the tip to the target element (do not rely on free XYZ + GraphicGroup drag).
6. Keep text where it landed. **Move only if it overlaps** other text (§4).
7. Verify: arrowhead at tip; text↔leader connected; then move text once — **snapped** ⇒ arrow stays, leader flips; **free** ⇒ whole annotation translates. Redraw if needed.

### Attachment / save gotchas

- Activated ref: `ActiveDesignFile` may be the ref, but `IsFromAttachment` can still be true on elements; `DesignFile.Save()` / `OpenDesignFileForProgram` may hit **Write inhibited**. Prefer **`SAVE DESIGN`** keyin after successful in-session `SetTextLine` edits.
- If a cell breaks, **`UNDO`** immediately; do not keep editing a type-7 orphan that used to be a type-2 callout cell.
- Confirm you are editing the file that **owns** the elements (search sibling DGNs if needed). Tables in sheet models vs callouts in `fea_*` refs is a common split.

## 4. Text position + overlap (**blocking gate**)

Applies to NOTES, callouts, circled labels, short tags, and any other placed/edited text.

**You may not claim done until this gate passes.** A center-distance-only check is not enough — long right-justified callouts can cover a tag while centers look fine.

### Keep original positions by default

- Editing existing text: change wording/numbers only; leave origin alone.
- Do not nudge “to look nicer” or to match markup FreeText box offsets.
- **Move only when overlap is real.** Prefer moving the element you just added (or whose longer wording caused the clash).

### How to detect overlap (required method)

1. For every element you placed or whose text you changed, get its **Range** box (Low/High X/Y).
2. Scan other text in the same plan area (plain text, text nodes, text inside cells). Use each item’s **Range**, not only origin.
3. Flag a clash when boxes intersect (small pad ~0.25 master units is OK) **or** centers are under ~1× text height *and* the strings would read on top of each other.
4. Leaders crossing leaders is OK. **Text-through-text is not.**

### If clash

1. Nudge **one** side (usually the new tag/callout) the smallest step that clears.
2. Re-scan ranges.
3. When reporting: name both texts, element IDs, and the nudge `(dx, dy)`.

### Done evidence (required in the reply)

Before saying the place/edit is finished, state explicitly:

- overlap gate: **pass** / **fail**
- what was scanned (e.g. new tags + nearby callouts/cells)
- clashes found (or “none”)
- any moves (id + dx/dy)

If you cannot run the range scan, say so and do not claim done.

## 5. Verify before claiming done

- Table / notes / callouts content checks as before (SAMPLE wraps, SEE NOTE sync, visible leaders where required).
- **Overlap gate (§4) must pass** with the done-evidence lines in your reply.
- Cite element IDs / levels / before-after text when reporting.

## Anti-patterns (forbidden)

- Inventing table wraps without opening the SAMPLE/markup page
- Rebuilding pay-item descriptions as auto-wrapping TextNodes
- Center-justifying NOTES like the estimate DESCRIPTION column
- `DeleteAllTextLines` inside callout cells
- Assuming cloned type-33 note dimensions will show — always add a real leader line
- Declaring callouts “done” without a visual check that leaders connect text → feature **with arrowhead**, and that moving the text moves leader + arrow with it
- Faking leaders with plain `CreateLineElement1` lines (no arrowhead; disconnects from text on move)
- Declaring text place/edit “done” without a **Range/bbox** overlap scan (§4)
- Using **only** origin/center distance for overlap (misses long callout strings)
- Moving notes/callouts from their original position when there is no overlap
