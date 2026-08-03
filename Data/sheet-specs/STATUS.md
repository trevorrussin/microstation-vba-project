# Sheet spec batch status

Tracks every 619 sheet through the four gates a spec must clear before the
placement code should trust it. Same bar for sheet 74 as for sheet 2 — enforced
by these gates, not by memory. See `AUTHORING.md` for the authoring procedure
and `README.md` for the schema.

## Gates

1. **structural** — `python scripts/validate_sheet_spec.py Data/sheet-specs/<sheet>.json` passes with 0 errors
2. **invariants** — same command, domain invariants (skip line = 40 ft, devices = skips+1, monotonicity) pass
3. **round-trip** — `python Bridge/roundtrip/<sheet>.py` diffs 0 against the source PDF
4. **live-build** — `python scripts/live_build_check.py <sheet> ...` confirms `BUILD_WZTC_ORDER_TABLE` actually produces the spec's expected rows in a live MicroStation session

A sheet is **done** only when all four are checked. `drafted` means the JSON
exists but hasn't cleared every gate yet.

## Legend

| Status | Meaning |
|---|---|
| `not-started` | No spec file |
| `drafted` | Spec file exists, gates incomplete |
| `done` | All 4 gates passed |
| `blocked` | Cannot author — missing source PDF or other hard blocker (reason in Notes) |

## Phase 1 — shared table library

| Sheet | Status | structural | invariants | round-trip | live-build | Notes |
|---|---|---|---|---|---|---|
| 619-011 | done | pass | pass (1 documented sheet anomaly, not our error — see `knownAnomalies` on table 011-02) | pass (0 fails) | n/a (reference-library sheet, no order table to build) | Shared table library. `knownExcerpts` documents which of 619-311's tables are confirmed identical excerpts. Sign-size master table (needed for a 311-05-style role) was NOT found on this sheet — lives on 619-012 (now authored as referenceLibrary). |

### 619-011 findings that affect every family below
- Table 011-01 (protective vehicle) covers all 5 durations (Mobile/Short Duration/Short Term/Intermediate/Long Term) × Freeway/Non-Freeway — 619-311's 311-01 is confirmed to be exactly its SHORT_TERM/NON-FREEWAY slice. Every Intermediate/Long Term sibling sheet in Family 1 (412/414/423/523) should pull its protective-vehicle row from 011-01's INTERMEDIATE_TERM/LONG_TERM columns rather than re-transcribing.
- Table 011-02 (taper) covers 25-65 mph and lateral shift 4-12 ft — wider than 311-02's 25-55/10-12. Family reference sheets for Freeway (302, which reaches 65 mph) should cross-check against 011-02 first.
- Table 011-06 (advance warning spacing) adds a FREEWAY row (A=1000,B=1500,C=2640, "1 MILE"/"½ MILE") that 311-03 (non-freeway) doesn't have — needed for Family 2/3/5 (freeway) sheets.
- Table 011-07 gives the general taper-length formulas (Merging=L, Shifting=L/2, Shoulder=L/3, downstream/one-lane-two-way=50-100ft) — cite this instead of re-deriving per sheet.
- Table 011-05 (flare rates for positive barrier) is new, needed for Family 3 (freeway shoulder closure, which has a barrier option) and Family with long-term barrier sheets.
- Detail 011A confirms the 40 ft skip-line constant directly (30' skip + 10' line).

## Family 1 — Non-Freeway multilane lane closure (reference: 619-311)

| Sheet | Status | structural | invariants | round-trip | live-build | Notes |
|---|---|---|---|---|---|---|
| 619-311 | done | pass | pass | pass (0 fails) | pass (confirmed 2026-08-03) | Reference sheet for this family. Tooling in Phase 0 generalized against this sheet. |
| 619-202 | done | pass | pass (1 expected: no taperAndBuffer) | pass (0 fails) | pass (live 2026-08-03: ROLL + W04-02L/W20-01RF; no taper/buffer) | Short-duration left. No taper table; A/B only; operator in PV. Fresh PDF re-fetched (prior capture path-only text). |
| 619-203 | done | pass | pass (1 expected: no taperAndBuffer) | pass (0 fails) | pass (live 2026-08-03: ROLL + W04-02R/W20-01RF) | Short-duration right. Mirror of 202 with W4-2R. PV=011 SHORT_DURATION. |
| 619-312 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03: L + L/2 overlay + W04-02L/W20-05LA) | TWLT. AW A/B only; taper has no shoulder cols; L/2 overlay; W9-3/R4-7. |
| 619-317 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03: MERGING + W20-05RA) | Single lane. Tables == 311; MERGING TAPER label; channelizing 317-05; W20-5 (no R/L). Roles by content. |
| 619-325 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03) | Double interior short-term. Tables == 311; roles 01=PV/02=roll/03=taper/04=AW. |
| 619-412 | done | pass | pass | pass (sibling identity + phrases; rotation=270) | pass (live 2026-08-03) | Intermediate TWLT. Diff vs 312: PV=011 INTERMEDIATE; NYR9-11; 20' spacing; rot=270. |
| 619-414 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03) | Intermediate single lane. Taper/AW/roll == 311; PV=011 INTERMEDIATE; NYR9-11; 20' spacing. |
| 619-423 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03: W04-02L) | Intermediate double interior. PV=011 INTERMEDIATE; 20' spacing; W4-2L. |
| 619-523 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03) | Long-term double interior. PV=011 LONG_TERM; NYR9-11; 20' spacing. |

### 619-311 family siblings findings (authored 2026-08-03, all 4 gates)
- **Taper/AW/roll identity**: 317/325/414/423/523 taper+buffer cells match 311-02 exactly (7 speeds × lane+shoulder). 312/412 match buffer+lane only (no shoulder columns). Roll-ahead and URBAN/RURAL A/B/(C) match 311 on overlapping sheets.
- **PV duration slices from 619-011**: Short Duration (202/203), Intermediate (412/414/423), Long Term (523) pulled from 011-01 NON-FREEWAY columns rather than re-keyed from sheet glyphs — verified against PDF PV cells.
- **Short duration (202/203)**: no taper/buffer/G20-2/downstream; A/B only; operator stays in PV (Note 2). Prior `Bridge/captures/619-202.pdf` / `203.pdf` had path-only text (2 words); re-fetched from NYSDOT repo.
- **TWLT (312/412)**: L/2 shifting taper is an overlay (`consumesStation: false`) like 311's shoulder taper; AW prints A/B only; extra R4-7/W9-3.
- **tableRoles by content**: 325 numbers PV=01/roll=02/taper=03/AW=04 (reverse-ish of 317's 01=AW/02=taper/03=roll/04=PV). Do not key off suffix digit.
- **Tooling**: `sheet_spec.resolve` now tolerates AW rows without C and taper rows without `shoulderTaper`; `live_build_check` tolerates missing lane/shoulder taper in overrides.


## Family 2 — Freeway/divided lane closure (reference: 619-302)

| Sheet | Status | structural | invariants | round-trip | live-build | Notes |
|---|---|---|---|---|---|---|
| 619-302 | done | pass | pass (1 documented cross-sheet discrepancy vs 619-011, not our error — see `knownAnomalies` on table 302-02) | pass (0 fails) | pass (confirmed live against a real MicroStation session 2026-08-03, both signs and non-sign rows) | Family 2 reference sheet. |
| 619-303 | done | pass | pass (1 expected FREEWAY A≠B≠C warning) | pass (0 fails) | pass (live 2026-08-03: dual MERGING TAPER + 2L + W20-05aRM) | Two-lane Family 2 sibling. Structural variant of 302 (not a blind clone): two MERGING TAPER L with dimensioned 2L between; W20-5aR; table roles 02=roll/04=taper/05=sizes; 9 notes; two arrow panels. `zone_length` gained `scale` for 2L. |
| 619-402 | done | pass | pass (1 expected FREEWAY A≠B≠C warning) | pass (0 fails, E3 2-page) | pass (live 2026-08-03) | Intermediate one-lane sibling. Tables 402-03/04/02 == 302 taper/spacing/roll; **402-01 = PVH/PVL+TMIA**; new **402-05** channelizing matrix; Note 4 = **20'** spacing (not 40'); Note 8 regulatory R2-1/NYR2-*; 2 pages (tables on p2). Canonical PDF = `619-402_E3.pdf`. |
| 619-403 | done | pass | pass (1 expected FREEWAY warning) | pass (sibling-identity + phrase checks; rotation=270) | pass (live 2026-08-03: dual taper+2L+W20-05aRM) | Intermediate two-lane hybrid of 303+402. Table roles: 04=channelizing, 05=sizes, 06=spacing. W20-5a; 20' spacing; PVH/PVL. FREEWAY.C=2640 inferred (no 2640 token in text layer) — knownAnomaly. Pages rotation=270. |
| 619-504 | done | pass | pass (2 expected warnings: no rollAhead role; FREEWAY A≠B≠C) | pass (sibling identity + barrier phrases) | pass (live 2026-08-03: MERGING TAPER only upstream — no roll/buffer) | Long-term barrier sibling. No PV/roll-ahead tables; 504-03 flare rates; Notes 4-7 barrier rules; buffer column in 504-02 but not on plan. sheet_spec.resolve/zone_length now tolerate missing rollAhead/PV roles. |

### 619-302 findings (authored + verified 2026-08-03, all 4 gates passed)
- **Table numbering trap, confirmed real**: on 302, table 04 = REQUIRED SIGN SIZES and table 05 = ROLL AHEAD DISTANCE — the reverse of 311's numbering (311-04=roll ahead, 311-05=sign sizes). `tableRoles` is assigned by content in the spec, not by suffix digit.
- 302-01/302-03/302-04/302-05 confirmed byte-identical (cell-for-cell, two independent extraction passes) to 619-011's master tables and 619-311's overlapping values. 302-02 is 7-of-8 rows identical to 619-011's 011-02/011-03 — see the genuine discrepancy below.
- **Genuine differences from 311, all confirmed**: 302-02 covers 25-55+65 mph (**8 rows, NOT 9** — an initial recon guess of a 60mph row was wrong and had to be corrected; 60mph is genuinely absent, same gap as 619-011's buffer table 011-03); 302-03 (advance warning) is keyed by ROAD TYPE with a FREEWAY row (A=1000,B=1500,C=2640,"1 MILE"/"½ MILE") confirmed identical to 619-011's 011-06; 302-01 splits FREEWAY vs NON-FREEWAY (311 was non-freeway only); 302-04 has 6 rows including WARNING FLAG (an initial pass mistakenly read this as 5 rows/no flag — corrected); plan wording is "MERGING TAPER" not "LANE TAPER"; **three** protective vehicles (VEH #1/#2/#3, with VEH #2 conditional on shoulder >= 8ft per Note 8) vs 311's two; **8 printed notes** not 5, in a different order than 311's (content-matched, not assumed by number); **no** "10'-0\" (MIN.)" lateral dimension (confirmed absent, not carried over from 311).
- **Cross-sheet discrepancy, not a transcription error**: 302-02's 65mph/12ft lane-taper cell prints `800/20/21` (internally consistent: 20×40=800), but 619-011's 011-02 prints `800/19/20` for the same lookup (already flagged there as an internal sheet-print anomaly). Two different official NYSDOT sheets disagree on this one cell. 619-302.json uses its own directly-verified value.
- **Same DATUM SHARING trap as 311** confirmed via `scripts/extract_plan_geometry.py` (shoulder-taper-inside-gap-A) — encoded the same way.
- **Likely registry error confirmed**: `Data/sheet-registry.tsv`'s 619-302 row lists sign codes R2-1/NYR2-2/NYR2-6 and none appear anywhere in the actual per-sheet PDF — book-PDF transcription noise, not a real sheet gap. The "(EITHER PVH OR PVL)" fragment from the book PDF also doesn't exist on the real sheet.
- **Tooling generalization gap found and fixed**: `mcp-server/sheet_spec.py`'s `resolve()` hardcoded the advance-warning-spacing table's row key as `"areaType"` (619-311's convention); 619-302/619-011 key the same table role as `"roadType"` instead. Fixed to support both. Also added `applicability.speedRangeMph.allowed` as an explicit-list alternative to `{min,max,increment}` for sheets with genuine gaps (like 302-02's missing 60mph), in both `sheet_spec.py` and `validate_sheet_spec.py`.
- **`scripts/live_build_check.py` bug found and fixed via the actual live run**: its generic `expected_labels()` compared against the sheet's bare `signCode` ("W4-2R") instead of the resolved SignLibrary key the bridge actually returns ("W04-02R") — different digit padding means the bare code isn't even a substring of the resolved key, so every sign check silently failed until this was caught by actually running gate 4 live rather than trusting the dry-run alone.

## Family 3 — Freeway shoulder closure (reference: 619-301)

| Sheet | Status | structural | invariants | round-trip | live-build | Notes |
|---|---|---|---|---|---|---|
| 619-301 | done | pass | pass | pass (0 fails; rotation=270 token/phrase checks) | pass (live 2026-08-03: SHOULDER TAPER sequential + W21-05aR/W21-05bRM/W20-01RM; no MERGING TAPER) | Family 3 reference. Shoulder closure: no AW spacing table (plan 1000/1500/1320); roll-ahead by PV **GVW**; PVH+TMIA; speeds 45–65 only; W21-5aR/W21-5bR. `sheet_spec.resolve` now optional AW table + GVW roll-ahead + optional laneTaper. |
| 619-205 | done | pass | pass (1 expected: no taperAndBuffer role) | pass (0 fails) | pass (live 2026-08-03: roll + W21-05 + W20-01RA; no taper/buffer) | Short-duration sibling. No taper/buffer table; order = ROLL AHEAD + W21-5 + W20-1 only; P+TMIA; speed-keyed roll-ahead; ~1000′ gaps. |
| 619-315 | done | pass | pass | pass (0 fails; rotation=270; gapC=2640) | pass (live 2026-08-03: SHOULDER TAPER + W21-05aR/W21-05bRM/W20-01RM @2640′) | Ramp-approach short-term. GVW roll-ahead like 301; **7 shoulder bands** + lateralShiftTaper; plan C=**2640′** (not 301’s 1320); W3-7a on plan. |
| 619-401 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03: W20-05aRM + W21-05bRM; no MERGING) | Intermediate. **W20-5aR** (not W21-5aR) + W21-5bR; speed roll-ahead; 401-03 adds laneTaper; **20′** device spacing; channelizing matrix. |
| 619-415 | done | pass | pass | pass (0 fails; role numbering trap) | pass (live 2026-08-03) | Intermediate ramp. Roles by content: **415-01=taper**, 02=roll, 03=PV; lateralShiftTaper; NYR9-11; 20′ spacing. |
| 619-501 | done | pass | pass (1 expected: no rollAhead role) | pass (0 fails) | pass (live 2026-08-03: SHOULDER TAPER only upstream — no roll/buffer) | Long-term barrier. No PV/roll-ahead; 501-03 flare; order starts at SHOULDER TAPER; positiveBarrier corridor zone. |

## Family 4 — Parkway

| Sheet | Status | structural | invariants | round-trip | live-build | Notes |
|---|---|---|---|---|---|---|
| 619-306 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03: MERGING+BUFFER+W04-02R/W20-05RM/W20-01RM @1000/1500/2640; no SHOULDER TAPER row) | Family 4 reference. Parkway shoulder<8 short-term. Hybrid: MERGING+DOWNSTREAM like 302, fixed gaps like 301 (no AW table). Table 306-03==302-02 on 45-65; NO NYW8-33; 3 notes only. |
| 619-212 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03: SHOULDER TAPER + W04-02R/W20-05RM/W20-01RM @500/1500/2640; no MERGING/BUFFER) | Short duration. Plan shoulder taper only; table still has lane cols. NYW8-33 on PV. Gaps 500/1500/½mi. Operator stays in vehicle. |
| 619-114 | done | pass | pass (1 expected: no taperAndBuffer role) | pass (0 fails) | pass (live 2026-08-03: ROLL AHEAD + W20-05RA @500′) | Mobile. 3 tables; moving roll-ahead; NYW8-33+W20-5R; 500′ min / 2 mi max; >15 min → 619-212. |
| 619-041 | done | pass | pass (1 expected: no taperAndBuffer role) | pass (0 fails) | pass (live 2026-08-03: ROLL AHEAD + W08-23) | Mowing / moving non-freeway. W8-23 only; NON-FREEWAY PV speed bands; roll-ahead incl ≤40; work area ≤40′; >5 min → 619-201. |

### 619-306 findings (Family 4 reference, all 4 gates passed 2026-08-03)
- **No shoulder-taper dimension on plan** despite L/3 columns in table 306-03 — parkway shoulder<8 sheet dimensions MERGING+DOWNSTREAM only. Do not emit SHOULDER TAPER as a sequential row (unlike 302's gap-A overlay).
- **No advance-warning spacing table** — fixed plan callouts 1000'/1500'/2640' (same FREEWAY A/B/C values as 302-03, but hardcoded like 301).
- **No NYW8-33** on size table / plan (unlike 302/212) — only G20-2/W4-2R/W20-1/W20-5R/FLAG.
- **Exactly 3 notes** (not 302's 8) — no left-lane symmetry / VEH#2 / transverse / 40' notes printed.
- Table 306-03 buffer+lane+shoulder cells identical to 302-02 on overlapping 45/50/55/65 speeds.

## Family 5 — Ramp-adjacent single lane closure (reference: 619-318)

| Sheet | Status | structural | invariants | round-trip | live-build | Notes |
|---|---|---|---|---|---|---|
| 619-318 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03: MERGING+dual W20-01RM @1320 + W04-02R/W20-05RM) | Family 5 reference. Fixed gaps 1000/1500/1320/1320 (two W20-1); NO AW spacing table. 318-01==302-02 @45-65. PVH+TMIA. Ramp W4-1R/R1-2/W3-2. |
| 619-316 | done | pass | pass | pass (0 fails; rotation=270) | pass (live 2026-08-03: SHOULDER TAPER + W21-05/W20-01RA) | Partial exit ramp. Shoulder walk (no MERGING row). Gaps 1000/1500. |
| 619-319 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03) | Near exit ramp. Same corridor as 318; adds E5-1/E5-2/W5-4. |
| 619-113 | done | pass | pass (1 expected: no taperAndBuffer) | pass (0 fails; rotation=270) | pass (live 2026-08-03: ROLL + W21-05aL @1000′) | Mobile-only. 2 tables; note-derived roll 80/160; no taper. |
| 619-211 | done | pass | pass (1 expected: no taperAndBuffer) | pass (0 fails) | pass (live 2026-08-03: ROLL + W21-05aL/W20-01RA) | Short-duration left shoulder on exit ramp. |
| 619-416 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03: SHOULDER TAPER + W21-05aR/W20-01RA @1000/500) | Intermediate partial exit. 7-band shoulder; W21-5aR. |
| 619-417 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03: MERGING=200 from 12ft band alias) | Intermediate entrance. 7-band grid; laneTaper aliases 10/11/12 ft cols. |
| 619-418 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03) | Intermediate exit-ramp channelizing. Lane+3band like 319. |
| 619-517 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03: MERGING + W20-01RM @2640′) | Long-term entrance. 7-band; gaps 1000/1500/2640. |
| 619-518 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03: MERGING + W20-01RM @2640′) | Long-term exit-ramp. Lane+3band; gaps 1000/1500/2640. |

### 619-318 findings (Family 5 reference, all 4 gates passed 2026-08-03)
- **No advance-warning spacing table** — plan callouts 1000'/1500'/1320'/1320' (two W20-1 signs). Do not apply 302-03 FREEWAY 1000/1500/2640.
- **Table 318-03** is ADVANCE PLACEMENT OF WARNING SIGN (NY2C-4: 930…1365) — advisory, not A/B/C.
- **Speeds 45/50/55/65 only**; 318-01 cell-identical to 302-02 on overlap. PV table FREEWAY-only PVH+TMIA.
- **Shoulder taper overlay inside gap A** (same datum-sharing pattern as 302).
- **Two structural sub-shapes in this family**: (A) lane+3band merging corridor (318/319/418/518); (B) 7-band shoulder grid where plan MERGING L aliases the 10/11/12 ft columns (417/517) or partial-exit uses SHOULDER TAPER walk (316/416); (C) mobile/short-duration minimal (113/211).

## Family 6 — Two-Lane Two-Way (reference: 619-307)

| Sheet | Status | structural | invariants | round-trip | live-build | Notes |
|---|---|---|---|---|---|---|
| 619-307 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03: ROLL+BUFFER+W20-07/W20-04*/W20-01* @URBAN45; no MERGING) | Family 6 reference. Buffer-only taperAndBuffer (no lane/shoulder cols). AW/buffer/PV/roll == 311 on overlap. Signs W20-7/W20-4/W20-1/W3-4(cond)/G20-2. |
| 619-308 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03) | Prior to intersection. Tables == 307; roles 01=AW/02=buf/03=sizes/04=PV/05=roll. |
| 619-309 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03) | AFAD. Primary 01-03; alt 04-06; PV/roll on 309B-*. R10-6 STOP HERE ON RED. |
| 619-314 | done | pass | pass | pass (0 fails; rotation=270) | pass (live 2026-08-03: ROLL+BUFFER+W20-07/04/01 @500′) | Moving flaggers. Fixed 500′ gaps (no AW table). Roll 2-band (no ≤40). |
| 619-321 | done | pass | pass (1 expected: no taperAndBuffer) | pass (0 fails; rotation=270) | n/a (pedestrian — no vehicle corridor walk; BUILD_WZTC_ORDER_TABLE errors on sign-only payload) | Sidewalk detour. Sign sizes + channelizing only. |
| 619-322 | done | pass | pass (1 expected: no taperAndBuffer) | pass (0 fails) | n/a (pedestrian/crosswalk — same as 321) | Crosswalk closure. + advance-placement guidelines table (not A/B/C). |
| 619-323 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03) | Intersection flagging. AW+sizes+channelizing; buf/PV/roll encoded from 311-identical NON-FREEWAY cells (sheet refs via notes/323-04). |
| 619-324 | done | pass | pass | pass (0 fails; rotation=270) | pass (live 2026-08-03: ROLL+BUFFER+W20-05/04/01; shoulder overlay) | TWLT shift. Buffer+shoulderTaper L/3; no merging L row. |
| 619-407 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03) | Intermediate flagger. Buffer speeds **45/50/55/65 only**; NYR9-11; 20′ channelizing. |
| 619-419 | blocked | | | | | **blocked: PDF missing/unreadable** — NYSDOT repo returns HTML 406/login page; not in captures/; DesignerRef-only in registry. |
| 619-420 | blocked | | | | | **blocked: PDF missing/unreadable** — same as 419. |
| 619-421 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03) | Intermediate intersection flagging. PV/roll on 421B-*; NYR9-11. |
| 619-422 | done | pass | pass | pass (0 fails) | pass (live 2026-08-03) | Intermediate TWLT shift. Sibling of 324 + NYR9-11. |
| 619-519 | done | pass | pass (1 expected: no taperAndBuffer) | pass (0 fails) | n/a (pedestrian — same as 321) | Long-term sidewalk. |
| 619-520 | blocked | | | | | **blocked: PDF missing/unreadable** — image-only capture (no text layer; OCR needed). Alt E1 same. |
| 619-524 | done | pass | pass (1 expected: no rollAhead) | pass (0 fails) | pass (live 2026-08-03: BUFFER+W03-03/W20-04*/W20-01*; shoulder overlay; no roll) | Long-term temp signal. R10-6L/R + W3-3; flare table 524-04. |
| 619-090 | done | pass | pass (1 expected: no rollAhead) | pass (0 fails; rotation=270) | pass (live 2026-08-03: BUFFER+W20-07/W03-04/W20-01*) | Temporary road closure. AW+buffer only; no PV/roll. |
| 619-091 | done | pass | pass (1 expected: no rollAhead) | pass (0 fails; rotation=270) | pass (live 2026-08-03) | Temporary intersection closure. Same schema as 090. |

### 619-307 findings (Family 6 reference, authored 2026-08-03)
- **No merging/lane taper** — flagger stop/slow control; `taperAndBuffer` is buffer-only (longitudinalBufferSpace rows, no laneTaper/shoulderTaper). Validate/resolve already tolerate optional taper columns (same path as 205/114).
- **AW/buffer/PV/roll cell-identical to 311** on overlapping speeds (25–55). Sign set differs: W20-7 + W20-4 (+ conditional W3-4) instead of W4-2R + W20-5R.
- **W3-4 is conditional** (Note 5: add if queue past W20-4) — in `signs.items` but not a default order-table row.
- **PV optional** (Note 6) — order table still includes ROLL+BUFFER for the “with PV” configuration used by live-build.
- **407 buffer trap**: intermediate sheet drops 25–40 mph rows; only 45/50/55/65 (incl. 645/16). Do not clone 307’s 7-row buffer blindly.
- **314**: no AW table — fixed 500′ plan gaps; roll is 2-band (≥55, 45–50) only.
- **Pedestrian (321/322/519)**: live-build n/a — `BUILD_WZTC_ORDER_TABLE` rejects sign-only payloads; no vehicle corridor to check.
- **419/420/520 blocked** on PDF availability (missing / image-only).

## Family 7 — Mobile operations (reference: 619-111)

| Sheet | Status | structural | invariants | round-trip | live-build | Notes |
|---|---|---|---|---|---|---|
| 619-111 | done | pass | pass | pass (`619-family7.py`) | pass (live 2026-08-03: ROLL AHEAD + W20-05RA; no tapers) | Family reference. 2 sheets; primary = Sheet 2 (≥8' shoulder) with W20-5R @1500′–½ mi. Sheet 1 (<8') has NYW8-33+W4-2R only. Speed-keyed moving roll-ahead (like 114). Fallback 619-206. |
| 619-110 | done | pass | pass | pass | pass (live 2026-08-03: ROLL AHEAD + W08-23) | Shoulder mobile; W8-23; **PVH+TMIA**; roll-ahead **GVW-keyed moving** (200/5–240/6 light; 160/4–200/5 heavy). Fallback 619-205. rot=270. |
| 619-112 | done | pass | pass | pass | pass (live 2026-08-03: ROLL AHEAD + W20-05aRA) | 2-sheet right two-lane; primary Sheet 2 adds W4-2R; W20-5AR advance; same GVW moving roll-ahead as 110. Fallback 619-207/209. rot=270. |
| 619-113 | done (see Family 5) | pass | pass | pass | pass | Authored under Family 5 (ramp mobile). Keep here as cross-ref only. |

### 619-111 findings (Family 7 reference, all 4 gates passed 2026-08-03)

- **Two structural sub-shapes**: (A) 111/114-style **speed-keyed** moving roll-ahead + P,TMIA PV matrix; (B) 110/112 **GVW-keyed** moving roll-ahead + PVH+TMIA (header `45-60 / w 55` is context like 301, not row keys).
- **2-page sheets** (111/112): Sheet 1 = narrow shoulder (<8′); Sheet 2 = ≥8′ with longer advance gap. Spec `tableRoles` / corridor / orderTable model **Sheet 2**; Sheet 1 tables kept for completeness.
- **No tapers / no AW table** on any F7 sheet — minimal mobile order (ROLL AHEAD + one advance sign, or W8-23 at PV for 110).

## Family 8 — Stop-and-go (reference: 619-101)

| Sheet | Status | structural | invariants | round-trip | live-build | Notes |
|---|---|---|---|---|---|---|
| 619-101 | blocked | | | | | No PDF: NYSDOT US/metric/transportation-systems repos return HTML error stubs (~10 KB) for 619-101[_E1\|_E3\|_20250501]. DesignerRef-only; not in 2026 Book 3. Historical master list had 101–104; current stop-and-go on NYSDOT index appears as 619-045/046 (also undownloadable from same repos). |
| 619-102 | blocked | | | | | Same — no real PDF in public repos. |
| 619-103 | blocked | | | | | Same — no real PDF in public repos. |
| 619-104 | blocked | | | | | Same — no real PDF in public repos. |

## Family 9 — Mowing/mulching/marking (reference: 619-023)

| Sheet | Status | structural | invariants | round-trip | live-build | Notes |
|---|---|---|---|---|---|---|
| 619-023 | done | pass | pass (1 expected: no taperAndBuffer) | pass (`619-family9.py`) | pass (live 2026-08-03: ROLL AHEAD + W21-08) | Family 9 reference. Shoulder <8' lane encroachment mowing/mulching. PVH/PVL×speed; GVW×speed roll (min=heavy/max=light); W21-8 primary. Fallback 619-022. rot=270. |
| 619-021 | done | pass | pass (2 expected: no taper/roll roles) | pass | n/a (sign-only — no roll row; BUILD_WZTC_ORDER_TABLE sign-only path) | Work beyond shoulder mowing. AW spacing + W21-8; plan gap 500′–2 mi. |
| 619-022 | done | pass | pass (1 expected: no taperAndBuffer) | pass | pass (live 2026-08-03: ROLL AHEAD + W21-08) | Non-freeway mowing shoulder/lane. PVH/PVL both closures; W21-8 + W8-23. |
| 619-031 | done | pass | pass (1 expected: no taperAndBuffer) | pass | pass (live 2026-08-03: ROLL AHEAD + W20-01RA) | Two-lane mulching/herbicide. P,TMIA like 041; speed-keyed moving roll 3-band; W20-1 + W8-23. Registry title "Freeway Mowing" wrong — PDF is two-lane mulching. |
| 619-032 | done | pass | pass (1 expected: no taperAndBuffer) | pass | pass (live 2026-08-03: ROLL AHEAD + W23-01) | Herbicide lane encroachment shoulder <8'. Sibling of 023; fallback 619-031. |
| 619-033 | done | pass | pass (1 expected: no taperAndBuffer) | pass | pass (live 2026-08-03: ROLL AHEAD + W20-01RA @FREEWAY) | Freeway shoulder mulching. FREEWAY-only PVH+TMIA; roll bands ≥60 / 45-55. Plan typo SEE TABLE 034-02 → 033-02. |
| 619-050 | blocked | | | | | **blocked: PDF missing** — NYSDOT repos return HTML error stubs; DesignerRef-only; not in 2026 Book 3. |
| 619-051 | blocked | | | | | **blocked: PDF missing** — same as 050. |
| 619-060 | done | pass | pass (1 expected: no taperAndBuffer) | pass | pass (live 2026-08-03: ROLL AHEAD + W23-01) | 2-sheet pavement marking. Tables on sheet 2; PVH/PVL + GVW×speed roll; W23-1/NYW8-*/W3-4. |

### 619-023 findings (Family 9 reference, authored 2026-08-03)
- **Two PV shapes in this family**: (A) PVH/PVL×NON-FREEWAY speed bands (022/023/032/060); (B) classic P,TMIA like 041 (031); (C) FREEWAY-only PVH+TMIA (033).
- **GVW×speed roll matrix** (022/023/032/060/033): printed as light/heavy GVW columns × speed rows. Encoded speed-keyed with `min=heavy` / `max=light` plus verbatim `lightGvw`/`heavyGvw` for round-trip. ≤40 band has equal min=max=120 (both GVWs) — validate now allows min≤max.
- **No tapers / no AW spacing role** on corridor mowing sheets (021 has AW table but plan uses fixed 500′–2 mi).
- **050/051 blocked** on PDF availability (same HTML-stub pattern as Family 8).

## Misc detail/reference sheets — not corridor-driven, lower complexity per sheet

| Sheet | Status | structural | invariants | round-trip | live-build | Notes |
|---|---|---|---|---|---|---|
| 619-001 | done | pass | pass (ref-lib) | pass (phrase) | n/a (referenceLibrary — barrier details, 7 pages) | Temporary Positive Barrier detail library. |
| 619-002 | blocked | | | | | **blocked: PDF missing** — NYSDOT repos return HTML error stubs. |
| 619-004 | done | pass | pass (ref-lib) | pass (phrase) | n/a (referenceLibrary) | Portable temporary wooden sign support details + dim limits. |
| 619-005 | done | pass | pass (ref-lib) | pass (phrase) | n/a (referenceLibrary) | PTRS spacing + sign sizes. |
| 619-006 | done | pass | pass (ref-lib) | pass (phrase) | n/a (referenceLibrary) | Speed-feedback PVMS placement/messaging — supplements corridor sheets. |
| 619-010 | done | pass | pass (ref-lib) | pass (phrase) | n/a (referenceLibrary) | General Notes — duration definitions table. |
| 619-012 | done | pass | pass (ref-lib) | pass (phrase) | n/a (referenceLibrary) | Sign Table catalog (3 pages) + color code legend. Size cells largely path-only; codes harvested from text layer. |
| 619-080 | done | pass | pass (2 expected: no taper/roll) | pass | n/a (sign-only work-beyond-shoulder) | Advance-placement distances + W20-1/G20-1/G20-2. |
