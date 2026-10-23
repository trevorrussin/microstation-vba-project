---
title: Plan sheets - named boundaries, Create Drawing, plan annotation
when: Cutting NYSDOT plan sheets (single, double plan, plan and profile), annotating plan Drawing models, deleting or re-cutting sheets.
keywords: plan sheet, plan sheets, named boundary, civil plan, plan border, north arrow, scale bar, match line, delete sheets, sheet index
---

# Plan sheets (ORD, NYSDOT seeds)

Model chain: Named Boundary -> Drawing model (live reference of that slice, where plan annotation is added) -> Sheet
model (references the Drawing model). Annotate in the Drawing model; the sheet only references it.

```yaml playbook
playbook:
  topic: plan-sheets-named-boundaries
  title: Plan sheets - named boundaries, Create Drawing, plan annotation
  when: Cutting NYSDOT plan sheets, annotating plan Drawing models, deleting or re-cutting sheets
  keywords: [plan sheet, plan sheets, named boundary, civil plan, plan border, north arrow, scale bar, match line, delete sheets, sheet index]
  sources:
    - "ORD - Cutting Plan Sheets - NYSDOT-20240130_123938-Meeting Recording.mp4 @ 00:08"
    - "000_ORD_Tutorial.pdf page 465"
    - "MicrostationTipsAndTricks.pptx slide 18"
  preflight:
    - {id: PF1, check: level, name: NBS_P, why: "plan border (sheet) proposed"}
    - {id: PF2, check: level, name: NBSP_P, why: "plan border, sheet plotting"}
    - {id: PF3, check: level, name: NBC_P, why: "plan border clipping"}
    - {id: PF4, check: level, name: NNR_P_T, why: "plan names - roadway, common highway and street (text)"}
    - {id: PF5, check: cell, name: NNAG, lib: ny_sheet.cel, why: "sheet north arrow (grid)"}
    - {id: PF6, check: cell, name: "240_1", lib: ny_barscale.cel, why: "scale bar 1in=20ft (plan scale 1:240)"}
    - {id: PF7, check: text_style, name: ny_Prop Normal, why: "plan notes"}
    - {id: PF8, check: text_style, name: ny_Prop Title, why: "plan titles"}
  steps:
    - id: S1
      title: Container / master reference file set up
      who: engineer
      manual_reason: "Reference attachment and project file setup are GUI/ProjectWise actions"
      how: |
        Plans that run for miles use a master reference file (Mstr_ref) with everything the plans show, levels, line
        styles and scales set. The sheet file attaches ONLY the master with Live Nesting, Nesting Depth 1, Display
        Overrides never. Give every attachment a Logical Name, Orientation Coincident - World. Depth 1 stops needless
        nested references from displaying.
      cite: "MicrostationTipsAndTricks.pptx slide 16"
    - id: S2
      title: Place Named Boundary - Civil Plan
      who: engineer
      manual_reason: "Named Boundary dialog (custom Bentley controls)"
      how: |
        Drawing Production > Named Boundaries > Place Named Boundary, mode Civil Plan. Pick the Drawing Seed FIRST
        (NYSDOT recording: `Plan Border`, nyu_Plan_Only.dgnlib) and select the alignment before changing names (the Name
        follows the geometry). Detail Scale CUSTOM 1:240 (1in=20ft) - do not change the scale the seed provides unless
        the sheet scale really changes. Name = `<road>-1` (end the top name in 1 so numbering starts uniformly),
        Group (New) + group name, Start/Stop stations, Length 250, Left Offset -60, Right Offset 60, Overlap 0,
        Boundary Chords. Clear the Start/Stop lock checkboxes if you want to pick graphically. Create Drawing on and
        Show Dialog on creates sheets immediately; for plan-and-profile leave Create Drawing OFF until the profile
        boundaries exist. Status bar: "5(5) Named Boundaries Created".
      cite: "ORD - Cutting Plan Sheets - NYSDOT-20240130_123938-Meeting Recording.mp4 @ 00:08"
    - id: S3
      title: Create Drawing dialog (plan)
      who: engineer
      manual_reason: "Create Drawing dialog fields are custom controls"
      how: |
        View Type Civil Plan, Discipline Civil, Purpose Plan View; Drawing Seed `Plan Border`; Seed Model nyu_Plan_Only.dgnlib;
        Annotation Group `Plan Annotation`; Sheet Model Detail Scale `1:240 (By Named Boundary)`; Add To Sheet Index on.
        Message center: "1(5) Sheets Created"; then "Updating Title Blocks...". Drawing model annotation scale and the sheet
        Detail Scale must both match the plan scale (1in=50ft -> 1:600).
      cite: "ORD - Cutting Plan Sheets - NYSDOT-20240130_123938-Meeting Recording.mp4 @ 00:46"
    - id: S4
      title: Review sheets - drawing model vs sheet model
      who: engineer
      manual_reason: "View Group / Models dialog change the engineer's active view"
      how: |
        View Group lists Drawing models and Sheet models. The Drawing model holds the Named Boundary shape, North Arrow and
        Match Line annotation; the Sheet model holds the border and references it. Saved Views are created with the sheets.
      cite: "000_ORD_Tutorial.pdf page 467"
    - id: S5
      title: Delete sheets completely (three parts)
      who: engineer
      manual_reason: "Destructive model/saved-view deletion is engineer-only"
      how: |
        A sheet is three parts: (1) Drawing + Sheet models (Home > Primary > Models > Delete Model(s)), (2) Saved Views
        (View > Saved Views > Delete), (3) the Named Boundary group (Named Boundaries dialog > Delete Selected Named
        Boundary or Group). Skipping one leaves orphans ("at least one reference model was not found").
      cite: "000_ORD_Tutorial.pdf page 472"
    - id: S6
      title: Plan annotation levels are usable in the Drawing model
      who: agent
      how: "Plan border/clipping/plotting and plan-name text levels come from the seed; use them as given, never retype names onto other levels."
      cite: "BRP Profiles-20240607_073333-Meeting Recording.mp4 @ 03:27"
      test: {action: levels, levels: [NBS_P, NBSP_P, NBC_P, NNR_P_T]}
    - id: S7
      title: North arrow from the sheet cell library
      who: agent
      how: "Grid north arrow is `NNAG` (inverted `NNAGI`, magnetic `NNAM`) in ny_sheet.cel; the cell carries its own levels - do not override them."
      cite: "MicrostationTipsAndTricks.pptx slide 18"
      test: {action: place_cell, name: NNAG, lib: ny_sheet.cel, expect_levels: [NNA_P], expect_size: [9.8, 50.6], size_tol: 1.5}
    - id: S8
      title: Graphic scale bar for the plan scale
      who: agent
      how: "Bar cells are named `<annotation scale>_1` in ny_barscale.cel: 240_1 = 1in=20ft, 480_1 = 1in=40ft, 600_1 = 1in=50ft, 1200_1 = 1in=100ft. Content is on NSB_P / NSB_P_T."
      cite: "BRP Profiles-20240607_073333-Meeting Recording.mp4 @ 10:33"
      test: {action: place_cell, name: "240_1", lib: ny_barscale.cel, expect_levels: [NSB_P, NSB_P_T], expect_size: [52.0, 8.5], size_tol: 1.5}
    - id: S9
      title: Plan notes and titles use the NYSDOT text styles
      who: agent
      how: "ny_Prop Normal for notes, ny_Prop Title for titles; the Drawing model's annotation scale supplies the plotted size."
      cite: "BRP Profiles-20240607_073333-Meeting Recording.mp4 @ 11:53"
      test: {action: text_style, style: ny_Prop Normal, text: "PLAN TEST NOTE"}
```
