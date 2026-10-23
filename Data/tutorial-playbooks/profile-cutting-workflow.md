---
title: Cutting a profile (and cross sections) from terrain along any line, named boundary to sheet
when: Creating a profile or cross-section view/sheet along an alignment or an ad-hoc line (geotech, bridge, abutment), re-cutting deeper, serialized profile sheets, or diagnosing "boundary will not fit" / "No view found".
keywords: cut a profile, cutting a profile, profile cut, quick profile from surface, set terrain active, open profile model, civil profile, profile only, boundary will not fit, no view found, serialize, geotech, abutment
---

# Cutting a profile in ORD (Bronx River Parkway recording, NYSDOT seeds)

The recording is narrated. Roles: the **engineer** drives every dialog in this workflow (they are custom Bentley
controls the agent must not drive); the **agent** knows the order, the settings, the naming rule and the failure modes,
tells the engineer exactly which step is next, and verifies the result afterwards.

```yaml playbook
playbook:
  topic: profile-cutting-workflow
  title: Cutting a profile (and cross sections) from terrain along any line, named boundary to sheet
  when: Creating a profile or cross-section view/sheet along an alignment or an ad-hoc line
  keywords: [cut a profile, cutting a profile, profile cut, quick profile from surface, set terrain active, open profile model, civil profile, profile only, boundary will not fit, no view found, serialize, geotech, abutment]
  sources:
    - "Cutting Profiles in ORD -  BRP-20241112_073049-Meeting Recording.mp4 @ 01:06"
    - "Cutting a Profile in ORD-20240312_111345-Meeting Recording.mp4 @ 00:29"
  preflight:
    - {id: PF1, check: level, name: AC_P, why: "the cut line is given the proposed alignment feature definition"}
    - {id: PF2, check: level, name: XGLH_P, why: "profile grid comes from the Profile Only seed"}
    - {id: PF3, check: level, name: XAX_P, why: "profile axis"}
    - {id: PF4, check: model, name: Profile, why: "profile Drawing/Sheet models appear once the drawing is created"}
  steps:
    - id: S1
      title: Draw the line to cut along and give it the proposed feature definition
      who: engineer
      manual_reason: "Civil Lines tool (Geometry > Horizontal > Lines); a civil feature, not plain graphics - the bridge tools cannot create one"
      how: |
        OpenRoads Modeling workflow: Geometry > Horizontal > Lines; click the start point and end point of the cut. Set the
        Feature Definition to AC_P (proposed alignment) so standards/style apply - an extra step, required only when this is an
        official alignment. A cross section is cut PERPENDICULAR to an alignment; with no plan alignment in the area this is
        a profile along an ad-hoc line.
      cite: "Cutting Profiles in ORD -  BRP-20241112_073049-Meeting Recording.mp4 @ 01:06"
    - id: S2
      title: Open the profile model in a second view
      who: engineer
      manual_reason: "Opening/arranging views changes the engineer's screen layout"
      how: |
        Geometry > Vertical > Open Profile Model; pick the horizontal alignment; the prompt says "Select or Open View" - open a
        second view and click it. The profile view MUST stay open for every later step (Create Drawing fails with "No view
        found" if it is closed; do not switch the active model afterwards if it is the only profile view).
      cite: "Cutting Profiles in ORD -  BRP-20241112_073049-Meeting Recording.mp4 @ 02:19"
    - id: S3
      title: Set the existing terrain ACTIVE
      who: engineer
      manual_reason: "Terrain > Active Terrain Model pick"
      how: |
        Terrain > Active terrain, pick the existing-ground terrain. Without an active terrain ORD cuts the line but shows no
        existing ground. You do not have to open the terrain or alignment file: reference them (References dialog; turn
        the survey/topo on only to locate yourself; the DTM/terrain file is what ORD reads).
      pitfall: "If existing ground is missing from the profile, the terrain is not active (the Set Active command is greyed when it already is)."
      cite: "Cutting Profiles in ORD -  BRP-20241112_073049-Meeting Recording.mp4 @ 04:17"
    - id: S4
      title: Project proposed surfaces onto the profile - Quick Profile From Surface
      who: engineer
      manual_reason: "Profile Creation tool prompts are interactive and need the profile view active"
      how: |
        With the PROFILE view active: Geometry > Profile Creation drop-down > Quick Profile From Surface (airplane icon). Click each
        reference surface to project (as many as needed); right-click (Reset) to end. It projects the surface as it is today.
      cite: "Cutting Profiles in ORD -  BRP-20241112_073049-Meeting Recording.mp4 @ 03:06"
    - id: S5
      title: Place Named Boundary - Civil Profile, size the window
      who: engineer
      manual_reason: "Named Boundary dialog (custom Bentley controls)"
      how: |
        Drawing Production > Named Boundaries > Place Named Boundary, mode Civil Profile, Drawing Seed `Profile Only` (NYS seeds),
        Detail Scale 1"=40' (1:480). The seed fills everything; change only what the profile needs: Stop Location / Length
        (e.g. 225), Vertical Exaggeration (omit for 1:1, or 5 for the NYSDOT 1:480 profile sheets), Available Profile Height =
        (highest - lowest elevation) rounded up with margin (lowest ~28.5, highest ~82 -> 54 -> use 60). Click the profile
        so the prompt knows which one; Start/Stop fill in. Use Terrains auto-shifts the elevation window to the terrains shown.
        Click in the view for the preview, then accept (status bar "Accept/Reject").
      pitfall: |
        "Boundary will not fit with given parameters": the profile height or Top/Bottom Clearance makes the boundary
        taller than the available height (here > 65). Toggle Use Terrains off to see it, reduce clearances to 1, then re-enable.
      cite: "Cutting Profiles in ORD -  BRP-20241112_073049-Meeting Recording.mp4 @ 07:46"
    - id: S6
      title: Create the Drawing model and Sheet model from the named boundary
      who: engineer
      manual_reason: "Named Boundaries dialog and Create Drawing dialog"
      how: |
        Placing the boundary only makes the viewport. Open the Named Boundaries dialog (small arrow), Profile Groups > your group, turn
        ON "Show Create Drawing dialog" (the pencil; otherwise ORD does what it wants), right-click the profile > Create Profile
        Drawing. Set View Name; Drawing Seed `Profile Only` (nyu_Profile_Only.dgnlib), Annotation Group `Profile Grid`, Sheets (New),
        Drawing Boundary and sheet Detail Scale `1"=40' (By Named Boundary)` (set it to 1:40 if the dialog shows a different value -
        NYSDOT preference files change), annotation scale 1:1, Open Model on. A progress bar appears bottom right; a long cut can look
        like a freeze - it is working.
      cite: "Cutting Profiles in ORD -  BRP-20241112_073049-Meeting Recording.mp4 @ 11:01"
    - id: S7
      title: Name for serialization - PIN_cph_pro-1
      who: agent
      how: |
        When cutting many NYSDOT profile sheets name the view/model `<PIN or D-number>_cph_pro-1` in all three places (view, drawing
        model, sheet model). ORD serializes automatically (-1, -2, -3 ...); type -1 each time. Category cph = contract plans,
        highway (PDM Appendix 14). Names must not contain an apostrophe (an apostrophe in a group name made Create Drawing report
        "Invalid name").
      cite: "Cutting Profiles in ORD -  BRP-20241112_073049-Meeting Recording.mp4 @ 12:23"
      test: {action: name, value: "X727.21_cph_pro-1", pattern: "^[A-Za-z0-9.]+_cph_pro-[0-9]+$"}
    - id: S8
      title: The Drawing model is where you draw; the Sheet model carries the border
      who: agent
      how: |
        Both a Drawing (model) view and a Sheet view are created. Lay out abutments / geotech / annotation in the DRAWING model
        (the live profile itself is hard to draw in). The sheet places the profile with a cell: delete that cell when you place the
        profile on a sheet with the BRP/NYSDOT border. Confirm a Profile Drawing model exists before annotating.
      cite: "Cutting Profiles in ORD -  BRP-20241112_073049-Meeting Recording.mp4 @ 14:27"
      test: {action: model, name: Profile}
    - id: S9
      title: Make the profile deeper / re-cut as a revision
      who: engineer
      manual_reason: "Profile view / Named Boundary edit and re-create"
      how: |
        To show more depth do not re-cut the boundary: change the depth in the profile view/boundary (20 ft deeper in the recording), then
        re-create the model and sheet with a new name (add "Rev") through Create Profile Drawing again; the new Drawing and Sheet
        models appear alongside the originals.
      cite: "Cutting Profiles in ORD -  BRP-20241112_073049-Meeting Recording.mp4 @ 15:43"
    - id: S10
      title: Cross sections along the same line
      who: engineer
      manual_reason: "Named Boundary / Create Drawing dialogs"
      how: |
        Named Boundary > Civil Cross Section, seed `XS 11 x 17 - 10 Scale Landscape`, Detail Scale 1"=5', select the alignment, click Start
        and Stop, Interval 25 ft, Event Point List (None); accept (two clicks): "10(10) Named Boundaries Created". Create the cross section DRAWING
        first (right-click group > Create Cross Section Drawing), then Create Cross Section Sheets from Drawings.
      pitfall: |
        "No view found for cross section boundaries": a seed/preference problem outside the drawing - log an IT ticket (Enterprise
        Solutions); workaround: still cut the profile along the line. Names auto-increment (AC1) when AC already exists.
      cite: "Cutting Profiles in ORD -  BRP-20241112_073049-Meeting Recording.mp4 @ 17:19"
```
