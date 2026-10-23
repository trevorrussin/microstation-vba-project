---
title: NYSDOT profile sheets - cut, create drawing, annotate
when: Building or finishing a NYSDOT profile sheet (Named Boundary Civil Profile, Drawing/Sheet models, grade/bridge/scale annotation, vertical clearance marks).
keywords: profile, profile sheet, named boundary, vertical clearance, bridge deck, scale bar, bridge, grade, existing ground
---

# Profile sheets (NYSDOT / ORD)

Source recordings are silent screen captures; the steps below come from their on-screen dialogs (OCR + keyframes).
Where the recording used the AECOM BRP workspace names (`nyu_*`, `CP_Normal`), this playbook gives the name that
exists in THIS workspace (`ny_*`, `ny_Prop Normal`) - verified live by the runner.

```yaml playbook
playbook:
  topic: profile-sheet-annotation
  title: NYSDOT profile sheets - cut, create drawing, annotate
  when: Building or finishing a NYSDOT profile sheet
  keywords: [profile, profile sheet, named boundary, vertical clearance, bridge deck, scale bar, bridge, grade, existing ground]
  sources:
    - "Cutting a Profile in ORD-20240312_111345-Meeting Recording.mp4 @ 00:29"
    - "BRP Profiles-20240607_073333-Meeting Recording.mp4 @ 10:33"
    - "000_ORD_Tutorial.pdf page 468"
  preflight:
    - {id: PF1, check: level, name: XGLH_P, why: "profile grid horizontal lines"}
    - {id: PF2, check: level, name: XGLV_P, why: "profile grid vertical lines"}
    - {id: PF3, check: level, name: BDE, why: "bridge deck edge in profile"}
    - {id: PF4, check: level, name: NSB_P, why: "plan scale bars"}
    - {id: PF5, check: cell, name: 480_Vert, lib: ny_barscale.cel, why: "graphic scale bar 1in=40ft horiz/vert"}
    - {id: PF6, check: text_style, name: ny_Prop Normal, why: "annotation text"}
    - {id: PF7, check: text_style, name: ny_Prop Title, why: "profile title"}
    - {id: PF8, check: line_style, name: "( Dashed )", why: "existing surface profile line"}
    - {id: PF9, check: model, name: Profile, why: "a Profile Drawing model exists in the open DGN (annotate there, not in the Sheet model)"}
  steps:
    - id: S1
      title: Place Named Boundary Civil Profile (Station Limits)
      who: engineer
      manual_reason: "Drawing Production > Named Boundaries dialog; custom Bentley controls, not safely automatable"
      how: |
        Open the Profile Model view first (Create Drawing fails with "No view found for profile boundaries" otherwise).
        Drawing Production > Named Boundaries > Place Named Boundary, mode Civil Profile. Seed `Profile Only`,
        Detail Scale 1:480 (1in=40ft), Name `Profile 1`, Method Station Limits, Vertical Exaggeration 5.0,
        Available Profile Height 50, Top/Bottom Clearance 1, Elevation Datum Spacing 5, Station Datum Spacing 100,
        Profile Shifts Datum Stations, Use Terrains on, Create Drawing on, Show Dialog on. Click in the profile view to accept.
      cite: "Cutting a Profile in ORD-20240312_111345-Meeting Recording.mp4 @ 00:29"
    - id: S2
      title: Create Drawing dialog
      who: engineer
      manual_reason: "Create Drawing dialog fields are custom controls"
      how: |
        Mode Profile; View Name/Model Names without apostrophes; Drawing Seed `Profile Only` (nyu_Profile_Only.dgnlib);
        Annotation Group `Profile Grid`; Sheet Detail Scale `1:480 (By Named Boundary)`; Drawing Boundary `Profile Only`;
        Add To Sheet Index on; Open Model on. OK stays disabled while One Sheet Per Dgn is checked with an undefined path.
      pitfall: "Apostrophe in a group name (e.g. 1960's CL) makes the model name invalid - use 1960s."
      cite: "Cutting a Profile in ORD-20240312_111345-Meeting Recording.mp4 @ 01:17"
    - id: S3
      title: Work in the Drawing model
      who: agent
      how: "Annotation goes in the Drawing model; the Sheet model only references it. Before drawing, confirm the ACTIVE model is the Drawing (describe_drawing_state modelType 3); activate the Drawing model if it is not. Never switch DGN files."
      test: {action: model, name: Profile}
      cite: "000_ORD_Tutorial.pdf page 468"
    - id: S4
      title: Profile grid and axis levels are usable
      who: agent
      how: "Grid lines XGLH_P / XGLV_P and axis XAX_P come from the drawing seed; do not redraw them. Use them only to confirm alignment."
      test: {action: levels, levels: [XGLH_P, XGLV_P, XAX_P]}
      cite: "BRP Profiles-20240607_073333-Meeting Recording.mp4 @ 03:21"
    - id: S5
      title: Bridge deck edge on level BDE
      who: agent
      how: "Bridge deck / underside lines in the profile are on level BDE (Bridge, Deck, Edge), ByLevel color and weight."
      test: {action: levels, levels: [BDE]}
      cite: "BRP Profiles-20240607_073333-Meeting Recording.mp4 @ 06:39"
    - id: S6
      title: Graphic scale bar from the barscale cell library
      who: agent
      how: |
        Cell library ny_barscale.cel (nyu_barscale.cel in the recording). For a 1in=40ft profile place `480_Vert`
        (1"=40' Scale Horz.), which draws the HORIZONTAL SCALE and VERTICAL SCALE bars. Do NOT force a level: the
        cell's own content is already on NSB_P / NSB_P_T (Plan Scale Bars, Proposed) - verified live. Placed at scale 1.0
        in this workspace it measures 99.4 x 39.2 ft (an 80 ft bar plus labels), correct for the 1:480 annotation scale.
        (The recording typed X-scale 0.020833 because the BRP file's master units differ.)
      cite: "BRP Profiles-20240607_073333-Meeting Recording.mp4 @ 10:33"
      test: {action: place_cell, name: 480_Vert, lib: ny_barscale.cel, expect_levels: [NSB_P, NSB_P_T], expect_size: [99.4, 39.2], size_tol: 1.5}
    - id: S7
      title: Profile title text
      who: agent
      how: "Title text uses the Title text style (ny_Prop Title here; C_P_Title in the BRP file): PROPOSED <ROAD> PROFILE."
      cite: "BRP Profiles-20240607_073333-Meeting Recording.mp4 @ 11:53"
      test: {action: text_style, style: ny_Prop Title, text: "PROPOSED TEST PROFILE"}
    - id: S8
      title: Existing surface drawn dashed
      who: agent
      how: "Reviewer comment on the BRP profile sheets: dashed line for existing surfaces."
      cite: "BRP Profiles-20240607_073333-Meeting Recording.mp4 @ 09:47"
      test: {action: line_style, style: "( Dashed )"}
    - id: S9
      title: Update title blocks, position the drawing on the sheet
      who: engineer
      manual_reason: "Sheet Index / Move Reference / title-block update are GUI tools"
      how: "Move Reference to position the Drawing reference on the sheet border; title blocks update on open (\"Title Blocks updated successfully!\")."
      cite: "BRP Profiles-20240607_073333-Meeting Recording.mp4 @ 07:59"
    - id: S10
      title: Bridge vertical clearance marks - measured value, cited minimum
      who: agent
      how: |
        Draw a vertical double arrow from the profile up to the bridge underside (soffit) with the street name and "MIN. VC = x.xx'" beside it (engineer markup style). MIN. VC is MEASURED
        from the geometry: soffit line Y minus the profile Y at each end of the bridge, divided by the vertical scale (5 units/ft at VE 5), the smaller value wins. The minimum
        required clearance comes from the standard (HDM Chapter 2, section 2.6.9 Vertical Clearance, p.37), not from memory; the existing underclearance value is NBI Item 54. The QA gate's operation
        contracts and relation checks (`qa_gate expect --rel`) verify the measured numbers.
      cite: "BRP Profiles-20240607_073333-Meeting Recording.mp4 @ 03:57"
      test: {action: standard, doc: "nysdot/hdm/chapt_02.pdf", section: "2.6.9"}
```


## Notes
- Offsets: ORD civil offsets are negative LEFT of the baseline, positive RIGHT; our `copy_parallel` tool is + LEFT of start-to-end.
