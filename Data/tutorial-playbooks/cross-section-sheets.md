---
title: Cross section sheets - named boundaries, Create Drawing, annotation
when: Cutting or annotating cross sections (uniform or non-uniform), grid/axis annotation on a cross section drawing.
keywords: cross section, cross sections, xs, named boundary civil cross section, section navigator, grid, axis
---

# Cross section sheets (ORD, NYSDOT seeds)

Cross sections are live references: the 3D Design model is clipped into one Drawing model per station, annotated there,
and the Drawing models are placed on Sheet models. Annotation (grid, axis, labels) belongs in the Drawing model.

```yaml playbook
playbook:
  topic: cross-section-sheets
  title: Cross section sheets - named boundaries, Create Drawing, annotation
  when: Cutting or annotating cross sections
  keywords: [cross section, cross sections, xs, named boundary civil cross section, section navigator, grid, axis]
  sources:
    - "Cutting a Cross Section in ORD-20240125_142500-Meeting Recording.mp4 @ 01:33"
    - "000_ORD_Tutorial.pdf page 416"
    - "000_ORD_Tutorial.pdf page 426"
  preflight:
    - {id: PF1, check: level, name: XAX_P, why: "cross section axis line"}
    - {id: PF2, check: level, name: XAXTXT_P_T, why: "cross section axis text"}
    - {id: PF3, check: level, name: XGLH_P, why: "grid horizontal lines"}
    - {id: PF4, check: level, name: XGLV_P, why: "grid vertical lines"}
    - {id: PF5, check: cell, name: "120_1", lib: ny_barscale.cel, why: "graphic scale bar 1in=10ft (cross section scale)"}
    - {id: PF6, check: text_style, name: ny_Prop Normal, why: "cross section annotation text"}
  steps:
    - id: S1
      title: Both the 2D plan model and the 3D model must be open in views
      who: engineer
      manual_reason: "Opening views / model views is a GUI action; the agent must not change the engineer's view layout"
      how: |
        Cross section named boundaries are only created when the 3D model is visible: open the 2D and 3D models in
        views (Multi-Model Views). Attach the terrain, geometry and corridor files first (References dialog,
        Orientation Coincident - World); whatever shows in the 3D view is what appears on the sections.
      cite: "000_ORD_Tutorial.pdf page 416"
    - id: S2
      title: Place Named Boundary - Civil Cross Section
      who: engineer
      manual_reason: "Named Boundary dialog (custom Bentley controls)"
      how: |
        Drawing Production > Named Boundaries > Place Named Boundary, mode Civil Cross Section. Pick the Drawing Seed FIRST
        (it supplies defaults): NYSDOT recording used `XS 11 x 17 - 10 Scale Landscape`. Then: Detail Scale CUSTOM 1:60,
        Group (New), Name = road/alignment name, Start/Stop stations (e.g. 323+00.00 / 325+00.00), Left/Right Offset
        -75 / 75, Interval 50, Vertical Exaggeration 1.0, Top Clearance 10, Bottom Clearance 5, Elevation Datum Spacing 2,
        Event Point List (None), Create Drawing on, Show Dialog on. Click: accept start, accept stop, create boundaries
        (status bar prompts "Accept/Reject. Identify end point to place cross section boundaries").
      pitfall: "The boundaries do not exist until they appear in the 3D view."
      cite: "Cutting a Cross Section in ORD-20240125_142500-Meeting Recording.mp4 @ 01:33"
    - id: S3
      title: Create Drawing dialog
      who: engineer
      manual_reason: "Create Drawing dialog fields are custom controls"
      how: |
        Mode Cross Section; View Name `<name> - <station>`; Drawing Seed `XS 11 x 17 - 10 Scale Landscape`
        (nyu_XS_10 Scale.dgnlib); Drawing Model annotation scale CUSTOM 1:120 with Annotation Group `XS Grid w/ Annotation`;
        Sheet Model: Create Sheet Model on, Sheets (New), Sheet annotation CUSTOM 1:1, Drawing Boundary = the seed,
        Detail Scale `1:60 (By Named Boundary)`; Open Model on. "5(5) Named Boundaries Created" confirms the boundaries.
      cite: "Cutting a Cross Section in ORD-20240125_142500-Meeting Recording.mp4 @ 01:59"
    - id: S4
      title: Non-uniform sections need ONE group in station order
      who: engineer
      manual_reason: "Named Boundaries dialog"
      how: |
        A single set of sheets cannot be created from two groups. To widen part of the corridor (e.g. Bottom Clearance 10
        for utilities), place the second run into the SAME group (Group = existing name) in station order - sections
        land on sheets in the order the boundaries were defined. Delete a wrong group from the Named Boundaries dialog.
      cite: "000_ORD_Tutorial.pdf page 426"
    - id: S5
      title: Annotate in the Drawing model with the cross-section levels
      who: agent
      how: |
        Grid and axis come from the seed (levels XAX_P, XAXTXT_P_T, XGLH_P, XGLV_P) - do not redraw them. Add notes with the
        NYSDOT text styles. Scale bar for 1in=10ft is cell `120_1`.
      cite: "000_ORD_Tutorial.pdf page 420"
      test: {action: levels, levels: [XAX_P, XAXTXT_P_T, XGLH_P, XGLV_P]}
    - id: S6
      title: Cross section note text uses the Proposed Normal style
      who: agent
      how: "Place notes with ny_Prop Normal (annotation scale supplies the plotted height)."
      cite: "000_ORD_Tutorial.pdf page 420"
      test: {action: text_style, style: ny_Prop Normal, text: "XS TEST NOTE"}
    - id: S7
      title: Graphic scale bar for the cross section scale
      who: agent
      how: "1in=10ft is `120_1` in ny_barscale.cel; cell content is on NSB_P / NSB_P_T."
      cite: "BRP Profiles-20240607_073333-Meeting Recording.mp4 @ 10:33"
      test: {action: place_cell, name: "120_1", lib: ny_barscale.cel, expect_levels: [NSB_P, NSB_P_T], expect_size: [25.8, 4.3], size_tol: 1.5}
```
