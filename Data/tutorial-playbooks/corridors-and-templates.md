---
title: Corridors - separate files, active profile, new corridor, template drops, point controls, 3D and cross sections
when: Building or reviewing an ORD corridor: file layout, active terrain/profile, New Corridor, template library, template drops (interval), point controls and parametric constraints, 3D model, dynamic cross sections, creating a terrain from a corridor.
keywords: corridor, corridors, template, template drop, drop interval, template library, itl, point control, parametric constraint, secondary alignment, final with mesh, dynamic cross section, superelevation, default-3d, corridor object, corridor model
---

# Corridors (ORD)

```yaml playbook
playbook:
  topic: corridors-and-templates
  title: Corridors - separate files, active profile, new corridor, template drops, point controls, 3D and cross sections
  when: Building or reviewing an ORD corridor
  keywords: [corridor, corridors, template, template drop, drop interval, template library, itl, point control, parametric constraint, secondary alignment, final with mesh, dynamic cross section, superelevation, default-3d, corridor object, corridor model]
  sources:
    - "BRP - Corridors.mp4 @ 00:12"
    - "Lunch and Learn - Corridors, Parametric Constraints & Point Controls-20230808_120243-Meeting Recording.mp4 @ 16:10"
    - "Depot Pl - Building a Corridor-20230609_102130-Meeting Recording.mp4 @ 01:05"
    - "ORD - Building a Corridor.pptx slide 5"
    - "Depot Pl - Setting a Point Control in a Corridor-20230609_125306-Meeting Recording.mp4 @ 06:02"
    - "000_ORD_Tutorial.pdf page 307"
  preflight:
    - {id: PF1, check: level, name: AC_P, why: "alignment feature definition for the corridor baseline"}
    - {id: PF2, check: level, name: PE_P, why: "pavement edge features the corridor creates"}
    - {id: PF3, check: state, key: is3D, value: "False", advisory: true, why: "corridor work starts in a 2D file; the Default-3D model is created and managed by ORD"}
  steps:
    - id: S1
      title: Keep alignments, corridors and terrains in separate files
      who: engineer
      manual_reason: "File organization and ProjectWise check-in"
      how: |
        Alignment file (ALG), corridor file and terrain file are separate: large combined files corrupt more easily and you would lose everything instead of one component.
        Check the alignment file IN before opening the corridor file. In the corridor file attach the alignment (ALG) and the existing terrain; turn off references you do not need. Corridor
        modeling always starts in a 2D file.
      cite: "BRP - Corridors.mp4 @ 00:31"
    - id: S2
      title: Terrain and profile must be ACTIVE
      who: engineer
      manual_reason: "Set Active is a hover-menu / ribbon pick on civil elements"
      how: |
        Import the existing terrain (.dtm; Feature Definition Terrain > Existing > Existing Contours; projection NY83-xx for the zone) and set it active: click the green boundary and hover, second icon
        "Set as Active Terrain Model", or Terrain > Edit > Active. Setting a terrain active creates the Default-3D model if none exists. Set the vertical profile active the same way (hover, Set As Active
        Profile) so ORD knows which profile the corridor uses; one active terrain and one active profile per alignment at a time. Always set the Feature Definition on anything modeling related.
      cite: "BRP - Corridors.mp4 @ 00:12"
    - id: S3
      title: Corridors > New Corridor - baseline, profile, name
      who: engineer
      manual_reason: "New Corridor tool prompts (lower-left status bar) are interactive"
      how: |
        Open the profile model of the alignment in a view, then Corridors > Create > New Corridor: "Locate corridor baseline" = the horizontal alignment, "Locate profile - reset for active
        profile" (right-click to accept the active profile), name the corridor, Feature Definition `Final` (use `Final with mesh` if you will create a terrain from the corridor). The corridor boundary
        drops into the view and the Create Template Drop tool starts automatically.
      cite: "BRP - Corridors.mp4 @ 01:42"
    - id: S4
      title: Template library and template drop
      who: engineer
      manual_reason: "Template library (.itl) and Create Template Drop dialogs lock the application while open"
      how: |
        Corridors > Create > New Template Drop (or the drop that starts after New Corridor), locate the corridor, pick the template (the three-dots / Alt+Down browser; expand Templates to the
        folder, e.g. Rural > Concrete Pavt w/ Concrete Shoulders > Undivided > 2 Lanes). Start Station: Alt = Lock To Start; End Station: Alt = Lock To End or a station; Drop Interval: 10 ft
        in training, the BRP recommendation is 1 ft (5 ft if less detail is needed; a smaller interval makes the file much larger). If the template library is empty the configuration did not load it: Create
        Template > File > Open the project .itl. While the template library window is open ORD is locked: close it before continuing. Start from a preloaded template and customize it.
      pitfall: "The template library window blocks the rest of the application; a prompt left open looks like a freeze."
      cite: "Lunch and Learn - Corridors, Parametric Constraints & Point Controls-20230808_120243-Meeting Recording.mp4 @ 16:10"
    - id: S5
      title: Corridor Object vs Corridor Model
      who: agent
      how: |
        The Corridor Object is a 2D closed shape with perpendicular handles in the 2D Default model (Construction class, can be turned off); it stores the corridor data. The Corridor Model is the set of 3D
        elements in the model named Default-3D. NEVER create or edit Default-3D by hand or the corridor stops working. The agent never edits corridor objects or the Default-3D model; flat 2D linework
        (Z=0) sits far below a 3D surface when viewed in 3D, which is expected.
      cite: "000_ORD_Tutorial.pdf page 311"
      test: {action: levels, levels: [PE_P, AC_P]}
    - id: S6
      title: Review - properties, 2 views Plan/3D, dynamic cross sections
      who: engineer
      manual_reason: "View layout and review tools change the engineer's screen"
      how: |
        F9 or right-hold in the view > View Control > 2 Views Plan/3D. Select the Corridor Object, hover, Corridor Properties (Use Active Profile True = it follows the active profile). Template drop
        boundary shows as a dashed closed shape; its Template Properties are editable. Corridors > Review > Dynamic Sections > Open Cross Section View, pick the corridor, click a view (View 7); set the
        view's Vertical Exaggeration (5 in training) and step through with the arrow keys. Level display is per view: use the "apply to all views" button to sync them.
      cite: "000_ORD_Tutorial.pdf page 316"
    - id: S7
      title: Point controls and secondary alignments
      who: engineer
      manual_reason: "Create Point Control tool is interactive civil modeling"
      how: |
        Corridors > Create > Create Point Control: Control Description (e.g. Ramp_CurbWest), Lock To Start / Lock To End (Alt) or a station range (R 90+24.30 to 90+32.08), Mode Both, Control Type Linear
        Geometry, pick the Plan Element (the feature, e.g. AC_P2 or a Ramp_Curb alignment), its Profile Element, the template Point (e.g. BN_AC, BN_PE_L, FG_AC, FG_PE_L) the control drives, Priority,
        "Use as Secondary Alignment" if it should also create dynamic sections, and the horizontal/vertical start/stop offsets. The Point Control list in the corridor properties shows each control with
        Enabled / Priority / Secondary Alignment so you can turn one off. Parametric constraints (Create Parametric Constraints) override template point values (pavement thickness, ditch width, slopes)
        for a station range.
        Workflow in the Depot Place recording: draw each curb/edge line plain first (Line with a direction key-in such as N25 50 08.6 E and a distance), trace it as a civil feature, give it its own
        profile (Profile Line Between Points, then Parabola Between Elements: vertical curve parameter 3.729, Trim/Extend Both), then Create Point Control (start 20+00.00, stop 21+37.16, description
        Ramp_CurbWest) picking the plan element (Ramp_CurbWest / Ramp_CurbEast / AC_P2) and its profile element (Sedgwick_CurbEast). The template points the control can drive are named in
        pairs: BN_AC, BN_PE_L, BN_PE_R (design/bottom) and FG_AC, FG_PE_L, FG_PE_R (finished grade). The control then appears under the corridor's Point Controls (Secondary Alignment, Curve Widening).
        In the template editor a point's properties hold its Feature Definition (e.g. Linear Pavement PET_P), Superelevation Flag and Use Feature Name Override; the corridor object tabs include
        End Condition Exception. The project template library is a .itl file kept with the models (DepotPI_ORD.itl); check it in like any other file.
      cite: "Depot Pl - Building a Corridor-20230609_102130-Meeting Recording.mp4 @ 01:09"
    - id: S8
      title: Superelevation, bulb-outs, tapers and widening are all part of the corridor
      who: engineer
      manual_reason: "Corridor design content and criteria belong to the engineer"
      how: |
        A complete corridor includes superelevation, project items (curb cuts, curb bulb-outs), shoulder tapers/widening and lane tapers/widening. Their values come from the governing standards
        (NYSDOT HDM chapters for superelevation and lane/shoulder widths), cited by section and page; never from memory. The corridor translates into a terrain (Final with mesh).
      cite: "BRP - Corridors.mp4 @ 04:45"
```
