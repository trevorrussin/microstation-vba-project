---
name: loc01-curb-classify
description: >-
  Classify Loc01 (HWP24QC1) INNER steel-faced curb layers: Straight /
  Depressed / Corner. Use when coloring curb, reclassify Loc01 islands,
  flare-to-flare depressed spans, or fixing wrong Corner/Straight mixes.
  NW island (3dpoly 13D986) is the visual authority. Layers only unless
  the engineer explicitly asks to draw flares. Agent control on FAIL:
  engineering-failure-protocol (not domain-specific).
---

# Loc01 curb classify (Straight / Depressed / Corner)

## Agent control (not curb-specific)

On any FAIL / UNCERTAIN verify: follow
`.cursor/rules/engineering-failure-protocol.mdc` first.
That protocol is **universal** — FL/ramp/CAD problems are only instances.

This skill answers **how to classify curb correctly** when the plan is working.
It does **not** replace the failure-recovery control loop.

**Engineering harness (preferred entrypoint):**
`python scripts/engineering_harness.py …` — Screenshots-only sandbox,
manifest states (`PASSED|BLOCKED|ESCALATED`), sample-first verify, steel
restore from backup, then classify. Manifest:
`outputs/engineering-harness-manifest.json`. Do not advance while BLOCKED.

Curb batch instance (called by harness): `scripts/_acad_batch_classify_locations.py`
exits on first `verify_fail`/`error` and leaves the DWG open (`STATUS=BLOCKED`).

Per-sheet done gate: `scripts/_acad_verify_location_sheet.py` (xref under
Screenshots, no `pwworking`/`d017*`, steel LF, Depressed>0 if ramps).

### Hard gates for this domain (what PASS requires)

Blocking if unmet (treat as FAIL → protocol, not “continue batch”):

- Islands classified match ramp-cluster count (PDF/markup authority)
- Unique INNER FL per island (no template-handle reuse across nearby corners)
- INNER not OUTER (never largest-diag fallback)
- If flare/ramp blocks exist → Depressed LF > ~0 (no Straight in flare-to-flare)
- No hide/delete of ramps, source FLs, or outers

Known-good compare for this domain: Loc10 PASS (PRO-S `*U255`–`*U260` flares,
unique `14DD*` inners) vs Loc11+ no-INNER / `E52F` reuse.

---

## Authority

1. Engineer visual QA (especially **NW**)
2. Geometry rules below
3. Scripts under `scripts/_acad_reclassify_loc01_*.py` / `_acad_reclassify_others_nw_process.py`

Do **not** invent a new classification scheme mid-session.

## Where each layer goes (always)

| Layer | Color | Where |
|-------|-------|--------|
| **Depressed** | 3 (green) | INNER curb **flare-to-flare** on one ped ramp (includes flares + DWS between). **Split vertices** at those stations so the layer change lands on the flare edge. |
| **Corner** | 5 (blue) | **Outside** depressed, and the curb is a **curve of short tessellation chords** (larger arc: path/chord + heading in a ~5 ft window of **short** segs only). |
| **Straight** | 1 (red) | **Outside** depressed, and **essentially colinear** — including **long** elementary segments (≥ ~2.5 ft). Long segs are **never** treated as arc chords (NW’s ~11–14 ft red runs). |

## Hard constraints

- **INNER only** for classify/replace. Never modify OUTER FLs: `13D80A`, `13D25C`, `13D7DC`, `13CEC7`, `13D2EA`, `13D7BC`.
- **Do not Delete** inner sources after classify — **hide** (`Visible=False`) so colored segs show. Undo can restore visibility.
- **No flare drawing** unless the engineer explicitly asks. Use existing PROP-D solids / DWS **only to measure** Depressed spans.
- Do **not** offset-rebuild the inner curb. Same path as the green topo / FL.
- AutoCAD product: `user-autocad` / COM — not MicroStation MCP — for this DWG work.

## NW reference (correct process)

- Source: `AcDb3dPolyline` **`13D986`** (sparse verts ≈ 2 ft mean) — not a densified FL dump.
- Ramps: `*U24` (`13CDD7`), `*U28` (`13D60F`).
- Example Depressed LF: U24 ≈ 11.07 · U28 ≈ 8.85.
- Feature-line islands must be reduced to **NW-like vertex density** before Corner/Straight (RDP / long-seg rule). Dense 0.3 ft sampling → false “all Corner”. **Forbidden.**

## Island sources (INNER)

| Island | Primary | PROP ramps (measure) |
|--------|---------|----------------------|
| NW | `13D986` | U24, U28 |
| N | `13D25E` | U25 |
| SW | `13D7E8` | U27, U29 |
| E | `13CED2` | none → Corner/Straight only |
| MidN | `13D2EE` | none until flares exist |
| MidS | `13D7BF` | none until flares exist |

## Loc 05–98 (Screenshots batch) — structure vs Loc01

Sheets under `Screenshots/HWP24QC1_Location_*_SGD.dwg` (and a few without `_SGD`) are **unclassified** Civil DWGs:

- **INNER only** — never classify the OUTER loop. INNER = FL closer to PROP-D ramp cluster (paired centers within ~10 ft). **Never** fall back to largest-diag FL (that is OUTER — Loc09 failure 2026-09-30).
- Inner/outer may be `AeccDbFeatureLine` on **`C-TOPO-FEAT` and/or `PRO-S-CURB-LINE` / `PROP-S-CURB-LINE`** — inventory layers per sheet; do not assume TOPO only.
- Always **RDP 0.1** before Corner/Straight. Dense 0.3 ft sampling forbidden.
- **Ped-ramp / flare measure (Depressed):** Loc01-style `PROP-D-ROAD-RAMPS` **or** block refs with **≥2 AcDbSolid flares** on `PRO-S-CURB-LINE` / `PROP-S-CURB-LINE` / `S-CURB-DROP` (Loc10: `*U255`–`*U260` on `PRO-S-CURB-LINE`). **Never** classify Corner/Straight only when flare blocks exist — that puts Straight in depressed zones.
- **Hard verify:** if ramp/flare blocks found and Depressed LF ≈ 0 → **FAIL** (do not accept the sheet).
- **Do not hide or delete** PROP-D ramps, source FLs, or outers. Color-coded content must stay visible. Only delete our own Straight/Depressed/Corner `AcDbLine` segs when re-clearing.
- **Engineering loop every sheet:** inventory → classify INNER → hard gates →
  **PASS only then** save and next sheet. **FAIL → STOP batch** (Failure
  Recovery Loop at top of this skill) — diagnose vs known-good Loc; do not
  advance while BLOCKED.
- Helper: `scripts/_acad_batch_classify_locations.py` (`--force` to redo). Inventory: `scripts/_acad_inv_loc05_08_dwg.py`.

## Scripts

| Script | Role |
|--------|------|
| `scripts/_acad_reclassify_loc01_v7_keep.py` | Full INNER classify; hide sources; no delete |
| `scripts/_acad_reclassify_others_nw_process.py` | Reclass N/SW/E/MidN/MidS; **keep NW**; RDP + LONG_SEG |
| `scripts/_acad_verify_loc01_curb.py` | **Done gate** — must pass before claiming complete |
| `scripts/_acad_inv_loc05_08_dwg.py` | Read-only Loc05–08 (and path-arg) FL/ramp inventory |
| `scripts/_acad_batch_classify_locations.py` | Loc 05–98 batch; multi-layer ramp scan |

## Live quantities → Pre-Final Cost Estimate Working

**Target file (only):**  
`2026.09.28_HWP24QC1 - Pre-Final Cost Estimate_Working.xlsx`  
Updates Length **in place**. Never delete that workbook. Never write a second estimate file.

**Where numbers go:** Length (col D) on **`4.09 AD` / `4.09 BD` / `4.09 CD`**, matched by **Corner ID** (col C).  
**Do not** write `HW-200` or `C.E. (Main)` — they VLOOKUP/ref the 4.09 tabs.

**All locations (not just 01–02 / 50–52):**  
- Master Corner ID list = rows already on `4.09 AD` in the Working estimate.  
- CAD capture regions = `Data/corner-cad-regions.json` (add `center` + `radius_ft` per Corner ID as you classify each site).  
- Optional override: XData app `HWP24QC1` / `CORNER_ID` on steel lines.

```text
python scripts/_acad_live_curb_qty.py           # one shot into Working xlsx
python scripts/_acad_live_curb_qty.py --watch   # ~1 s while editing in AutoCAD
python scripts/_acad_live_curb_qty.py --dry-run
```

Dashboard CSV: `outputs/curb-live-qty.csv`.  
PDF Corner ID catalog: `Data/hwp24qc1-corner-ids-from-pdfs.json`.

## Done means (Loc01 curb)

Do **not** claim done until `python scripts/_acad_verify_loc01_curb.py` exits 0 (or engineer waives a named check).

Verify script checks:

1. All six OUTER handles still exist  
2. NW keep-zone still has Straight + Depressed + Corner steel  
3. Non-NW islands with PROP ramps have Depressed LF in expected ballpark  
4. No script claimed success after Delete of OUTER  
5. Prints per-island LF summary for engineer eyeball  

If verify fails: **STOP, re-plan**, do not densify/threshold-twiddle without stating the new plan.

## Site zone + template contamination (2026-10-01, Claude Code)

SGD hosts were seeded from one template and carry **foreign content**:
- A copied **PROP-D ramp kit** (`*U226`–`*U231`, `*U273`, `*U274`) + PRO-S polylines
  **`E52F` / `FC58`** (same handles, same 66.5/69.7 ft offset in every sheet), moved
  near or off each site. Loc20: kit 560 ft from site, 0 survey ents nearby.
  Batch was picking `E52F`/`FC58` as INNER → the Loc09/11+/20–24 systemic fail.
- Foreign survey drop/TC blocks (~1012000,207000 and ~1002000,210000) in many hosts —
  verify gate was counting them as "ramps".
- **Fix:** `scripts/_acad_site_zone.py` — keep a ramp/FL only if the sheet's survey
  xref (`T4318*` / `Survey*`) has ≥100 ents in the 3×3 100-ft cells around it.
  No loaded survey xref → FAIL (UNCERTAIN). Wired into batch + `_acad_verify_location_sheet.py`.
- **Flare read bug:** `ramp_flare_data` used early-bound `e.Coordinate` → silently 0
  flares → Depressed measured from DWS pad only (~2–3 ft short per ramp). Fixed with
  late-bound `DynDispatch(e).Coordinate`. Any steel saved before 2026-10-01 is suspect.
- **Corner ID:** assign island → nearest plan-xref `CORNER ID … 10xxxxx` MLeader arrow
  (≤40 ft, sheet's IDs only). Sort-order queue mis-assigned 3/5 on Loc03-04.
  Master Location→Corner list = Working `4.09 AD` (PDF catalog: 38 locs, truncated-ID errors).
- **Loc10_SGD in Screenshots (and its `_sgd_backup` copy) = Loc03-04 content**
  (Survey R = `T4318AG25712_3-4,53-55`). ESCALATED — need the real Loc10 SGD.
- Zero on-site islands = FAIL (UNCERTAIN), not PASS. `--all` = `_SGD` hosts only.
- Loc01 truth file: `OneDrive…\AutoCAD Sheet Sets\HWP24QC1_Location_01-02_50-52_SGD_RECOVERED_20260929_1705.dwg`.
  Use **Screenshots/** only for batch work — never pwworking.
- Visual QA: `scripts/_acad_curb_qa_capture.py` (survey/host clutter layers Off, zoom, PNG).


### Batch rules added 2026-10-01 (Claude Code, during full Loc batch)

- **INNER must touch a flare.** If the picked FL has 0 flare vertices within CURB_NEAR,
  reject it and use the tan outline (Loc13: zigzag `14592` picked, flares on line+arc).
- **Tan PRO-S outline = INNER curb** where no Loc01-style FL exists (engineer confirmed
  on Loc11-12). Chain touching `PRO-S-CURB-LINE` / `PROP-S-CURB-LINE` pieces —
  LWPolyline (bulges densified), **AcDbLine, AcDbArc** — seed = piece most flares touch.
- **Flare contact = shortest triangle edge (base)**; apex ignored (it fell inside
  CURB_NEAR on tight outlines and stretched Depressed up the side edge). DWS bbox
  corners used only when a ramp has no flares.
- **Island scope:** template kit removed by (a) `E52F` or `FC58` on PRO-S (either) →
  drop ramps ≤150 ft; (b) when plan CORNER ID labels exist, drop islands >60 ft from
  any label (Loc21/37-38/44/95 kit landed inside survey zone).
- **Merge islands that touch the same curb line** (Loc58-60: one 70 ft outline with a
  ramp at each end clustered as 2 islands).
- Corner ID: nearest label ≤40 ft; else 1 island + 1 ID → direct; else unassigned
  (map from PDF). COM busy: ensure_layers + flare reads wrapped in retry.
- **Open engineer question:** outline mode colors the *whole* tan box (back/side edges
  Straight). Confirm curb-face-only vs whole outline before trusting Straight LF on
  outline sheets (13, 15-18, 27-29, 35, 45-47, 58-60, 81-85 …).
- Escalated (design gaps, not classifier): Loc10 (file = Loc03-04 content), Loc33
  (two nested C-TOPO paths), Loc61 (no ramps at real corner 1061199), Loc62-67 island2
  (ramps, no curb line). Never save a partial sheet without engineer OK.
- QA without stealing focus: `scripts/_acad_curb_plot.py` (matplotlib from DWG geometry).

## Common mistakes (do not repeat)

- Angular-wedge Depressed instead of flare-to-flare  
- Deleting green INNER before colored segs exist on all islands  
- Touching OUTER as “source” / classifying OUTER instead of INNER  
- Falling back to **largest** FL when pairing fails (OUTER)  
- Assuming inners are only on `C-TOPO-FEAT`  
- Assuming ped ramps are only on `PROP-D-ROAD-RAMPS` (Loc10 flares on `PRO-S-CURB-LINE`)  
- Classifying with **0 Depressed** while flare blocks exist → Straight in drop areas  
- Hiding PROP-D ramps or source FLs so color-coded content disappears  
- Classifying only one island when the sheet/PDF shows many ramps  
- Drawing yellow flares when the task is layers-only  
- `sample_fl_vlax(step=0.3)` then same curvature thresholds as NW  
- Clearing **all** steel including correct NW  
- Logging tips only in chat — update this skill + agent-log  
- Skipping per-sheet verify / engineering loop on batch work  
- Continuing to the next Loc after a hard-gate FAIL (fake batch progress)  
- Treating repeated identical fails (Loc 11/12/13…) as independent problems  
  instead of one systemic root-cause investigation  
- Loosening thresholds or inventing a new scheme without a tested hypothesis
- Picking template `E52F`/`FC58` or the copied PROP-D kit as an island (use site zone)
- Trusting a Depressed LF without confirming flares were read (n_flares>0)
- Assigning Corner IDs by sort order instead of plan label proximity  
