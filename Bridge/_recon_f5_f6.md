# Family 5 & 6 PDF Recon

Generated 2026-08-03 via PyMuPDF on `Bridge/captures/*.pdf`.

**Family 5** (ref 619-318): ramp-adjacent single lane closure — 10/10 PDFs available.
**Family 6** (ref 619-307): two-lane two-way / flagger / pedestrian — 16/18 PDFs available.

**Missing F6 PDFs:** 619-419, 619-420
**Image-only F6 PDFs (no text layer):** 619-520

## Reference confirmation

### 619-318 (Family 5 reference)
Confirm as F5 anchor: **YES**. 2 pg, tables ['318-05', '318-03', '318-04', '318-01', '318-02', '318-06'], MERGING=True, RAMP=True, DOWNSTREAM=True, PV/AP implied via keywords PV=True AP=True.
Registry expects Tables 318-01..06 with merging/downstream taper + channelizing — PDF matches.

### 619-307 (Family 6 reference)
Confirm as F6 anchor: **YES**. 2 pg, tables ['307-05', '307-04', '307-03', '307-01', '307-02'], FLAGGER=True, AFAD=False, no merging taper expected for base flagger sheet.
Registry: Two-Lane Two-Way flagger operation, Tables 307-01..03 — PDF aligns.

## Schema classification summary

| Sheet | Family | Schema | Closest |
|---|---|---|---|
| 619-318 | F5 | corridor lane-closure clone (ramp-adjacent) | 619-302 (score 10) |
| 619-316 | F5 | corridor lane-closure clone (ramp-adjacent) | 619-301 (score 6) |
| 619-319 | F5 | corridor lane-closure clone (ramp-adjacent) | 619-302 (score 10) |
| 619-113 | F5 | corridor clone + mobile-only | 619-302 (score 3) |
| 619-211 | F5 | corridor lane-closure clone (ramp-adjacent) | 619-302 (score 3) |
| 619-416 | F5 | corridor clone + road/intersection closure, signal | 619-302 (score 5) |
| 619-417 | F5 | corridor clone + road/intersection closure, signal | 619-302 (score 12) |
| 619-418 | F5 | corridor clone + road/intersection closure, signal | 619-302 (score 12) |
| 619-517 | F5 | corridor clone + road/intersection closure, signal | 619-302 (score 10) |
| 619-518 | F5 | corridor clone + flagger, road/intersection closure, signal | 619-302 (score 10) |
| 619-307 | F6 | NOVEL: flagger | 619-311 (score 4) |
| 619-308 | F6 | NOVEL: flagger | 619-311 (score 4) |
| 619-309 | F6 | NOVEL: AFAD | 619-311 (score 4) |
| 619-314 | F6 | NOVEL: flagger | 619-302 (score 6) |
| 619-321 | F6 | NOVEL: sidewalk detour | 619-302 (score 3) |
| 619-322 | F6 | NOVEL: sidewalk detour, crosswalk, signal | 619-311 (score 0) |
| 619-323 | F6 | NOVEL: flagger | 619-311 (score 4) |
| 619-324 | F6 | NOVEL: signal | 619-302 (score 6) |
| 619-407 | F6 | NOVEL: flagger, road/intersection closure, signal | 619-311 (score 4) |
| 619-419 | F6 | **PDF missing** | — |
| 619-420 | F6 | **PDF missing** | — |
| 619-421 | F6 | NOVEL: flagger, road/intersection closure, signal | 619-311 (score 2) |
| 619-422 | F6 | NOVEL: flagger, road/intersection closure, signal | 619-302 (score 11) |
| 619-519 | F6 | NOVEL: flagger, sidewalk detour, road/intersection closure, signal | 619-302 (score 3) |
| 619-520 | F6 | NOVEL: crosswalk (image-only PDF) | 619-311 (score 0) |
| 619-524 | F6 | NOVEL: flagger, road/intersection closure, signal | 619-311 (score 5) |
| 619-090 | F6 | NOVEL: flagger, road/intersection closure | 619-301 (score 2) |
| 619-091 | F6 | NOVEL: flagger, road/intersection closure | 619-301 (score 2) |

## Family 5 — per-sheet briefs

### 619-318
- **Pages/rotation:** 2 / 0°
- **Title keywords:** SHOULDER; RAMP; SHOULDER TAPER
- **TABLE titles:** 318-05: CHANNELIZING DEVICE; 318-03: W20-1; 318-04: HEAVY; 318-01: LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS; 318-02: ROLL AHEAD DISTANCE; 318-06: REQUIRED SIGN SIZES
- **Key signs:** G20-2, NY2C-4, NYR2-2, NYR2-6, NYS, NYW8-33, R1-2, R2-1, W20-1, W20-5, W3-2, W4-1R (+1)
- **Language flags:** MERGING, SHOULDER, DOWNSTREAM, RAMP
- **Spacing tokens:** 500', 1000', 1320', 1500'
- **Closest family:** 619-302 (score 10)
- **Schema:** corridor lane-closure clone (ramp-adjacent)


### 619-316
- **Pages/rotation:** 2 / 270°
- **Title keywords:** SHOULDER; RAMP; PARTIAL EXIT RAMP CLOSURE
- **TABLE titles:** 316-01: G20-2; 316-04: CHANNELIZING DEVICE; 316-03: HEAVY; 316-05: REQUIRED SIGN SIZES; 316-02: ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES
- **Key signs:** G20-2, W20-1, W21-5, W21-5AR
- **Language flags:** SHOULDER, DOWNSTREAM, RAMP
- **Spacing tokens:** 500', 1000', 1500'
- **Closest family:** 619-301 (score 6)
- **Schema:** corridor lane-closure clone (ramp-adjacent)


### 619-319
- **Pages/rotation:** 2 / 0°
- **Title keywords:** SHOULDER; RAMP; RAMP AREA CHANNELIZING
- **TABLE titles:** 319-04: CHANNELIZING DEVICE; 319-03: HEAVY; 319-01: LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS; 319-02: ROLL AHEAD DISTANCE; 319-05: REQUIRED SIGN SIZES*
- **Key signs:** G20-2, NYR2-2, NYR2-6, NYW8-33, R2-1, W20-1, W20-5, W4-2R
- **Language flags:** MERGING, SHOULDER, DOWNSTREAM, RAMP
- **Spacing tokens:** 500', 1000', 1320', 1500'
- **Closest family:** 619-302 (score 10)
- **Schema:** corridor lane-closure clone (ramp-adjacent)


### 619-113
- **Pages/rotation:** 1 / 270°
- **Title keywords:** ON THE SHOULDER WHERE THE WORK IS BEING; REACHES THE RAMP, IT SHALL FOLLOW THE WORK; SHOULDER AS POSSIBLE. ONCE THIS VEHICLE
- **TABLE titles:** 113-02: REQUIRED SIGN SIZES; 113-01: PROTECTIVE VEHICLE REQUIREMENTS
- **Key signs:** W21-5AL
- **Language flags:** MOBILE, RAMP
- **Spacing tokens:** 1000'
- **Closest family:** 619-302 (score 3)
- **Schema:** corridor clone + mobile-only


### 619-211
- **Pages/rotation:** 1 / 0°
- **Title keywords:** RAMP; LEFT SHOULDER CLOSURE ON EXIT RAMP; WORK ZONE TRAFFIC CONTROL
- **TABLE titles:** 211-01: VEHICLE WITH TMIA; 211-02: REQUIRED SIGN SIZES
- **Key signs:** W20-1, W21-5AL
- **Language flags:** RAMP
- **Spacing tokens:** 1000'
- **Closest family:** 619-302 (score 3)
- **Schema:** corridor lane-closure clone (ramp-adjacent)


### 619-416
- **Pages/rotation:** 2 / 0°
- **Title keywords:** SHOULDER; RAMP; PARTIAL EXIT RAMP CLOSURE
- **TABLE titles:** 416-03: VEHICLE HEAVY; 416-05: REQUIRED SIGN SIZES; 416-04: CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIO; 416-02: ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES; 416-01: LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS
- **Key signs:** G20-2, NYR2-2, NYR2-6, NYR9-11, R2-1, W20-1, W21-5AR
- **Language flags:** SHOULDER, DOWNSTREAM, RAMP, CLOSURE, SIGNAL
- **Spacing tokens:** 500', 1000'
- **Closest family:** 619-302 (score 5)
- **Schema:** corridor clone + road/intersection closure, signal


### 619-417
- **Pages/rotation:** 2 / 0°
- **Title keywords:** SHOULDER; RAMP; SINGLE LANE CLOSURE NEAR ENTRANCE RAMP
- **TABLE titles:** 417-03: TAPER; 417-04: HEAVY; 417-06: REQUIRED SIGN SIZES; 417-05: CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIO; 417-02: ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES; 417-01: LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS
- **Key signs:** G20-2, NY2C-4, NYR2-2, NYR2-6, NYR9-11, NYS, NYW8-33, R1-2, R2-1, W20-1, W20-5, W3-2 (+2)
- **Language flags:** MERGING, SHOULDER, DOWNSTREAM, RAMP, CLOSURE, SIGNAL
- **Spacing tokens:** 500', 1000', 1320', 1500'
- **Closest family:** 619-302 (score 12)
- **Schema:** corridor clone + road/intersection closure, signal


### 619-418
- **Pages/rotation:** 2 / 0°
- **Title keywords:** SHOULDER; RAMP; RAMP AREA CHANNELIZING
- **TABLE titles:** 418-01: 418-01); 418-03: HEAVY; 418-02: ROLL AHEAD DISTANCE; 418-05: REQUIRED SIGN SIZES; 418-04: CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIO
- **Key signs:** G20-2, NYR2-2, NYR2-6, NYR9-11, NYW8-33, R2-1, W20-1, W20-5, W4-2R
- **Language flags:** MERGING, SHOULDER, DOWNSTREAM, RAMP, CLOSURE, SIGNAL
- **Spacing tokens:** 500', 1000', 1320', 1500'
- **Closest family:** 619-302 (score 12)
- **Schema:** corridor clone + road/intersection closure, signal


### 619-517
- **Pages/rotation:** 2 / 0°
- **Title keywords:** WORK ZONE; SHOULDER; RAMP
- **TABLE titles:** 517-03: (SEE NOTE 5); 517-04: PROTECTIVE VEHICLE HEAVY; 517-06: REQUIRED SIGN SIZES; 517-05: CHANNELIZING DEVICE APPLICATION FOR LONG-TERM STATIONARY WOR; 517-02: ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES; 517-01: LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS
- **Key signs:** G20-2, NY2C-4, NYR9-11, NYS, NYW8-33, R1-2, W20-1, W20-5, W3-2, W4-1R, W4-2R
- **Language flags:** MERGING, SHOULDER, DOWNSTREAM, RAMP, CLOSURE, SIGNAL
- **Spacing tokens:** 500', 1000', 1500', 2640'
- **Closest family:** 619-302 (score 10)
- **Schema:** corridor clone + road/intersection closure, signal


### 619-518
- **Pages/rotation:** 2 / 0°
- **Title keywords:** SHOULDER; RAMP; RAMP AREA CHANNELIZING
- **TABLE titles:** 518-03: VEHICLE WITH TMIA; 518-01: LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS; 518-02: ROLL AHEAD DISTANCE; 518-05: REQUIRED SIGN SIZES; 518-04: CHANNELIZING DEVICE APPLICATION FOR LONG-TERM STATIONARY WOR
- **Key signs:** G20-2, NYR9-11, NYW8-33, W20-1, W20-5, W4-2R
- **Language flags:** MERGING, SHOULDER, DOWNSTREAM, RAMP, CLOSURE, SIGNAL, FLAGGER
- **Spacing tokens:** 500', 1000', 1500', 2640'
- **Closest family:** 619-302 (score 10)
- **Schema:** corridor clone + flagger, road/intersection closure, signal


## Family 6 — per-sheet briefs

### 619-307
- **Pages/rotation:** 2 / 0°
- **Title keywords:** FLAGGER; LANE CLOSURE WITH FLAGGERS; TWO-LANE TWO-WAY ROADWAY
- **TABLE titles:** 307-05: REFER TO SHEET 2 OF 2 FOR ALL TABLES; 307-04: VEHICLE WITH TMIA; 307-03: REQUIRED SIGN SIZES*; 307-01: ADVANCE WARNING SIGN SPACING; 307-02: LONGITUDINAL BUFFER SPACE
- **Key signs:** G20-2, W20-1, W20-4, W20-7, W3-4
- **Language flags:** DOWNSTREAM, FLAGGER
- **Spacing tokens:** 500', 1000', 1500'
- **Closest family:** 619-311 (score 4)
- **Schema:** NOVEL: flagger


### 619-308
- **Pages/rotation:** 1 / 0°
- **Title keywords:** FLAGGER; CLOSED PAVED SHOULDER 8' OR WIDER, IF THE WORK SPACE MOVES; TO INTERSECTION
- **TABLE titles:** 308-04: VEHICLE WITH TMIA; 308-01: ADVANCE WARNING SIGN SPACING; 308-03: REQUIRED SIGN SIZES*; 308-05: ROLL AHEAD DISTANCE; 308-02: LONGITUDINAL BUFFER SPACE
- **Key signs:** G20-2, W20-1, W20-4, W20-7, W3-4
- **Language flags:** DOWNSTREAM, FLAGGER
- **Spacing tokens:** 500', 1000', 1500'
- **Closest family:** 619-311 (score 4)
- **Schema:** NOVEL: flagger


### 619-309
- **Pages/rotation:** 3 / 0°
- **Title keywords:** FLAGGER; SHOULDER; ASSISTANCE DEVICE AND FLAGGER
- **TABLE titles:** 309-01: ADVANCE WARNING SIGN SPACING; 309-03: REQUIRED SIGN SIZES*; 309-02: LONGITUDINAL BUFFER SPACE /; 309-04: ADVANCE WARNING SIGN SPACING; 309-06: REQUIRED SIGN SIZES*; 309-05: LONGITUDINAL BUFFER SPACE /
- **Key signs:** G20-2, R10-6, W20-1, W20-4, W20-7, W3-4
- **Language flags:** DOWNSTREAM, AFAD, FLAGGER
- **Spacing tokens:** 500', 1000', 1500'
- **Closest family:** 619-311 (score 4)
- **Schema:** NOVEL: AFAD


### 619-314
- **Pages/rotation:** 1 / 270°
- **Title keywords:** LANE CLOSURE WITH MOVING FLAGGERS; TWO-LANE TWO-WAY ROADWAY; WORK ZONE TRAFFIC CONTROL
- **TABLE titles:** 314-04: REQUIRED SIGN SIZES*; 314-03: LONGITUDINAL BUFFER SPACE; 314-01: (EITHER PVH OR PVL); 314-02: ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES
- **Key signs:** G20-1, G20-2, NYW8-33, W20-1, W20-4, W20-7, W3-4, W7-3A
- **Language flags:** DOWNSTREAM, FLAGGER
- **Spacing tokens:** 500'
- **Closest family:** 619-302 (score 6)
- **Schema:** NOVEL: flagger


### 619-321
- **Pages/rotation:** 1 / 270°
- **Title keywords:** SIDEWALK; SIDEWALK CLOSED; SIDEWALK DETOUR
- **TABLE titles:** 321-01: REQUIRED SIGN SIZES*; 321-02: CHANNELIZING DEVICE APPLICATION FOR SHORT-TERM STATIONARY WO
- **Key signs:** G20-2, R205, R9-11L, R9-11R, R9-9, W20-1
- **Language flags:** DOWNSTREAM, SIDEWALK, RAMP
- **Spacing tokens:** —
- **Closest family:** 619-302 (score 3)
- **Schema:** NOVEL: sidewalk detour


### 619-322
- **Pages/rotation:** 2 / [0, 270]
- **Title keywords:** WORK ZONE PROVISIONS; TABLE 322-02: CHANNELIZING DEVICE APPLICATION FOR SHORT-TERM STATIONARY WORK ZONES; STATIONARY WORK ZONES
- **TABLE titles:** 322-01: REQUIRED SIGN SIZES*; 322-02: CHANNELIZING DEVICE APPLICATION FOR SHORT-TERM STATIONARY WO; 322-03: GUIDELINES FOR ADVANCE PLACEMENT OF WARNING SIGNS
- **Key signs:** G20-2, R11-2, R8-3, R9-10, R9-11L, R9-11R, R9-9, W20-1
- **Language flags:** SIDEWALK, CROSSWALK, SIGNAL
- **Spacing tokens:** —
- **Closest family:** 619-311 (score 0)
- **Schema:** NOVEL: sidewalk detour, crosswalk, signal


### 619-323
- **Pages/rotation:** 3 / 0°
- **Title keywords:** FLAGGER; FLAGGING OPERATION AT INTERSECTION; TWO-LANE TWO-WAY ROADWAY
- **TABLE titles:** 323-04: TO ENSURE ADEQUATE STOPPING; 323-01: ADVANCE WARNING SIGN SPACING; 323-02: REQUIRED SIGN SIZES*; 323-03: CHANNELIZING DEVICE APPLICATION FOR SHORT-TERM STATIONARY WO
- **Key signs:** G20-2, W20-1, W20-4, W20-7, W20-7A, W3-4
- **Language flags:** DOWNSTREAM, FLAGGER
- **Spacing tokens:** 500', 1000', 1500'
- **Closest family:** 619-311 (score 4)
- **Schema:** NOVEL: flagger


### 619-324
- **Pages/rotation:** 2 / 270°
- **Title keywords:** SHOULDER; L/3 SHOULDER TAPER; SURFACE, LOW OR NO SHOULDER, DROP-OFFS, GEOMETRIC CONSTRIANTS, AND/OR POOR SIGHT CONDITIONS).
- **TABLE titles:** 324-03: BUFFER SPACE; 324-01: (EITHER PVH OR PVL); 324-04: GUIDELINES FOR ADVANCE PLACEMENT OF WARNING SIGNS; 324-05: ADVANCE WARNING SIGN SPACING; 324-07: REQUIRED SIGN SIZES*; 324-06: CHANNELIZING DEVICE APPLICATION FOR SHORT-TERM STATIONARY WO; 324-02: ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES
- **Key signs:** G20-2, NYW8-33, R4-7, W20-1, W20-4, W20-5
- **Language flags:** SHOULDER, DOWNSTREAM, SIGNAL
- **Spacing tokens:** 500', 1000', 1500'
- **Closest family:** 619-302 (score 6)
- **Schema:** NOVEL: signal


### 619-407
- **Pages/rotation:** 2 / 0°
- **Title keywords:** WORK ZONE; FLAGGER; LANE CLOSURE WITH FLAGGERS
- **TABLE titles:** 407-05: VEHICLE WITH TMIA; 407-01: (1.5-2.5 SKIP LINES); 407-04: REQUIRED SIGN SIZES*; 407-03: CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIO; 407-02: LONGITUDINAL BUFFER SPACE; 407-06: ROLL AHEAD DISTANCE
- **Key signs:** G20-2, NY9-11, NYR9-11, W20-1, W20-4, W20-7A, W3-4
- **Language flags:** DOWNSTREAM, CLOSURE, SIGNAL, FLAGGER
- **Spacing tokens:** 500', 1000', 1500'
- **Closest family:** 619-311 (score 4)
- **Schema:** NOVEL: flagger, road/intersection closure, signal


### 619-419
- **Status:** PDF not in captures/


### 619-420
- **Status:** PDF not in captures/


### 619-421
- **Pages/rotation:** 3 / 0°
- **Title keywords:** FLAGGING OPERATION AT INTERSECTION; TWO-LANE TWO-WAY ROADWAY; WORK ZONE TRAFFIC CONTROL
- **TABLE titles:** 421-03: REQUIRED SIGN SIZES*; 421-02: CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIO; 421-01: ADVANCE WARNING SIGN SPACING; 421-04: LONGITUDINAL BUFFER SPACE /
- **Key signs:** G20-2, NYR9-11, W20-1, W20-4, W20-7, W20-7A, W3-4
- **Language flags:** DOWNSTREAM, CLOSURE, SIGNAL, FLAGGER
- **Spacing tokens:** 500', 1000', 1500'
- **Closest family:** 619-311 (score 2)
- **Schema:** NOVEL: flagger, road/intersection closure, signal


### 619-422
- **Pages/rotation:** 2 / 0°
- **Title keywords:** WORK ZONE; SHOULDER; L/3 SHOULDER TAPER
- **TABLE titles:** 422-04: VEHICLE WITH TMIA; 422-02: BUFFER SPACE; 422-03: ROLL AHEAD DISTANCE; 422-07: REQUIRED SIGN SIZES*; 422-05: CHANNELIZING DEVICE APPLICATION FOR INTERMEDIATE-TERM STATIO; 422-01: ADVANCE WARNING SIGN SPACING; 422-06: GUIDELINES FOR ADVANCE PLACEMENT OF WARNING SIGNS
- **Key signs:** G20-2, NYR9-11, NYW8-33, R4-7, W20-1, W20-4, W20-5
- **Language flags:** MERGING, SHOULDER, DOWNSTREAM, CLOSURE, SIGNAL, FLAGGER
- **Spacing tokens:** 500', 1000', 1500'
- **Closest family:** 619-302 (score 11)
- **Schema:** NOVEL: flagger, road/intersection closure, signal


### 619-519
- **Pages/rotation:** 2 / 0°
- **Title keywords:** SIDEWALK; SIDEWALK CLOSED; SIDEWALK DETOUR
- **TABLE titles:** 519-01: REQUIRED SIGN SIZES*; 519-02: CHANNELIZING DEVICE APPLICATION FOR LONG-TERM STATIONARY WOR; 519-03: REQUIRED SIGN SIZES*
- **Key signs:** G20-2, R205, R9-11L, R9-11R, R9-9, W20-1
- **Language flags:** DOWNSTREAM, SIDEWALK, RAMP, CLOSURE, SIGNAL, FLAGGER
- **Spacing tokens:** —
- **Closest family:** 619-302 (score 3)
- **Schema:** NOVEL: flagger, sidewalk detour, road/intersection closure, signal


### 619-520
- **Pages/rotation:** 2 / 0°
- **Status:** image-only PDF (6268 KB, 34 extractable chars) — needs OCR or visual review
- **Schema:** expect crosswalk (321 sibling) — **PDF unreadable by text extract**


### 619-524
- **Pages/rotation:** 2 / 0°
- **Title keywords:** L/3 SHOULDER TAPER; SHOULDER; TWO-LANE TWO-WAY ROADWAY
- **TABLE titles:** 524-02: FOR L AND NOTE 3); 524-04: FOR FLARE RATES AND NOTE 3); 524-05: REQUIRED SIGN SIZES*; 524-03: CHANNELIZING DEVICE APPLICATION FOR LONG-TERM STATIONARY WOR; 524-01: ADVANCE WARNING SIGN SPACING
- **Key signs:** G20-2, NYR9-11, R10-6L, R10-6R, W20-1, W20-4, W3-3
- **Language flags:** SHOULDER, DOWNSTREAM, CLOSURE, SIGNAL, FLAGGER
- **Spacing tokens:** 500', 1000', 1500'
- **Closest family:** 619-311 (score 5)
- **Schema:** NOVEL: flagger, road/intersection closure, signal


### 619-090
- **Pages/rotation:** 1 / 270°
- **Title keywords:** FLAGGER; TEMPORARY ROAD CLOSURE; TWO-LANE TWO-WAY ROADWAY
- **TABLE titles:** 090-01: ADVANCE WARNING SIGN SPACING; 090-02: LONGITUDINAL BUFFER SPACE; 090-03: REQUIRED SIGN SIZE*
- **Key signs:** W20-1, W20-7, W3-4
- **Language flags:** CLOSURE, FLAGGER
- **Spacing tokens:** 500', 1000', 1500'
- **Closest family:** 619-301 (score 2)
- **Schema:** NOVEL: flagger, road/intersection closure


### 619-091
- **Pages/rotation:** 1 / 270°
- **Title keywords:** 4. FOR INTERSECTIONS WITH MULTIPLE LANE APPROACHES, A SITE; 3. FLAGGER SYMBOL SIGN (W20-7) AND "BE PREPARED TO STOP"; FLAGGER
- **TABLE titles:** 091-02: LONGITUDINAL BUFFER SPACE; 091-01: ADVANCE WARNING SIGN SPACING; 091-03: REQUIRED SIGN SIZES*
- **Key signs:** W20-1, W20-7, W3-4
- **Language flags:** CLOSURE, FLAGGER
- **Spacing tokens:** 500', 1000', 1500'
- **Closest family:** 619-301 (score 2)
- **Schema:** NOVEL: flagger, road/intersection closure


## Recommendations

1. **619-318** — Use as F5 structural reference if tables 318-01..06 + merging/downstream taper corridor zones match 619-302 shape with ramp callouts; clone 302 corridor schema then add ramp labels.
2. **619-307** — Use as F6 base for flagger tables/sign spacing; do **not** reuse 302 merging taper schema — flagger sheets need `Flagger`/`CenterlineCones` zones and optional PV.
3. **Novel schema required (F6):**
   - 619-307: NOVEL: flagger
   - 619-308: NOVEL: flagger
   - 619-309: NOVEL: AFAD
   - 619-314: NOVEL: flagger
   - 619-321: NOVEL: sidewalk detour
   - 619-322: NOVEL: sidewalk detour, crosswalk, signal
   - 619-323: NOVEL: flagger
   - 619-324: NOVEL: signal
   - 619-407: NOVEL: flagger, road/intersection closure, signal
   - 619-421: NOVEL: flagger, road/intersection closure, signal
   - 619-422: NOVEL: flagger, road/intersection closure, signal
   - 619-519: NOVEL: flagger, sidewalk detour, road/intersection closure, signal
   - 619-524: NOVEL: flagger, road/intersection closure, signal
   - 619-090: NOVEL: flagger, road/intersection closure
   - 619-091: NOVEL: flagger, road/intersection closure
   - 619-419: **PDF missing** — expect NOVEL (intermediate sidewalk)
   - 619-420: **PDF missing** — expect NOVEL (intermediate crosswalk)
4. **Corridor clones (F5):** 316, 319, 416–418, 517–518 likely share 318 table stack with ramp/shoulder variants.
5. **Corridor clones (F6):** 308, 407, 421, 422 may share 307 flagger base; 314 moving flaggers may need mobile variant.
6. **Download still needed:** 619-419, 619-420
7. **OCR/visual review needed:** 619-520 (large image PDFs with no extractable text).