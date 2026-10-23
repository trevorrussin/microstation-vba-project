---
title: OpenRoads modeling workflow overview - the order of work and which playbook covers each phase
when: Starting any OpenRoads/MicroStation task and you need to know where it sits in the modeling process, what the previous phase must have produced, and which playbook to open next.
keywords: workflow, process, order, phases, survey, terrain, alignment, profile, superelevation, corridor, template, quantities, labeler, sheets, drawing production, overview, resources, training
---

# OpenRoads modeling process (training workflow sheet)

The training spreadsheet lists the process in eleven phases with video/PDF resources for each. Use this playbook to place a task, check that the previous phase exists, and
open the matching playbook.

```yaml playbook
playbook:
  topic: ord-modeling-workflow-overview
  title: OpenRoads modeling workflow overview - the order of work and which playbook covers each phase
  when: Placing a task in the modeling process and choosing the next playbook
  keywords: [workflow, process, order, phases, survey, terrain, alignment, profile, superelevation, corridor, template, quantities, labeler, sheets, drawing production, overview, resources, training]
  sources:
    - "OpenRoads Training.xlsx sheet Sheet1"
    - "MicrostationTipsAndTricks.pptx slide 2"
    - "ProjectWise Overview, ORD H&V Alignments.pptx slide 2"
  preflight:
    - {id: PF1, check: level, name: AC_P, why: "phase 3-4 output: proposed alignments exist as AC_P features"}
  steps:
    - id: S1
      title: The eleven phases and the playbook for each
      who: agent
      how: |
        1 Import survey -> new-file-setup-and-coordinates, references-levels-models-measuring.  2 Create/edit existing terrain, label contours and spots -> terrain-rasters-imagery-elevations.
        3 Horizontal geometry (create, modify, annotate, report) -> alignments-horizontal-vertical, as-built-geometry-layout, offsets-tapers-geometry.  4 Vertical geometry -> alignments-horizontal-vertical.
        5 Superelevation -> corridors-and-templates.  6 Corridor modeling (templates, library, parametric constraints and point controls, create corridor, cross-section view, modify, assign superelevation) ->
        corridors-and-templates.  7 Create terrain (from elements, corridor, corridor top mesh; contours; element templates; label contours/spots; analyze points) -> terrain-rasters-imagery-elevations.  8 Earthwork quantities.
        9 Civil Labeler (labels). 10 Creation of sheets (plan, cross section, plan and profile, rectangular P&P, separate files, single plan roll, sheet index, deleting sheets and named boundaries) ->
        plan-sheets-named-boundaries, cross-section-sheets, profile-cutting-workflow, profile-sheet-annotation.  11 Drawing production (plan and digital deliverables).
        Annotation details: line-styles, hatching-and-patterns.
      cite: "OpenRoads Training.xlsx sheet Sheet1"
      test: {action: name, value: "terrain-rasters-imagery-elevations", pattern: "^[a-z0-9-]+$"}
    - id: S2
      title: Start every alignment station at 10+00.00
      who: agent
      how: "The training sheet's own note: always start alignment stationing from 10+00.00 (never 0+00, never negative stations)."
      cite: "OpenRoads Training.xlsx sheet Sheet1"
      test: {action: name, value: "10+00.00", pattern: "^[1-9][0-9]*\\+[0-9]{2}(\\.[0-9]+)?$"}
    - id: S3
      title: File formats the workflow expects
      who: agent
      how: |
        Survey data comes in as .csv / .dwg / .dgn / .dtm; terrain import reads .dtm; template libraries are .itl; imagery is .jp2. If a file arrives in another format, ask the engineer what it is; do not
        rename extensions.
      cite: "OpenRoads Training.xlsx sheet Sheet1"
      test: {action: name, value: "project.itl", pattern: "^[A-Za-z0-9_.-]+\\.(itl|dtm|dgn|dwg|csv|jp2)$"}
    - id: S4
      title: 'In prompts, "reset" means right-click'
      who: agent
      how: "In every ORD prompt \"reset\" is a right mouse click (Reset to complete, Reset for the active profile); Esc and Element Selection also leave a tool. Tell the engineer which, in those words."
      cite: "OpenRoads Training.xlsx sheet Sheet1"
      test: {action: name, value: "reset = right-click", pattern: "right-click"}
    - id: S6
      title: Know what the library files actually contain
      who: agent
      how: |
        The deck named "ORD - Creating Proposed Terrains" does NOT cover proposed terrains: its 14 slides are the ProjectWise overview, horizontal/vertical alignment and template-library slides
        (the same deck as "ProjectWise Overview, ORD H&V Alignments"; its file also has a damaged zip index and was recovered from the raw headers). For creating a terrain from a corridor use
        corridors-and-templates (Feature Definition Final with mesh) and the terrain import/set-active steps in terrain-rasters-imagery-elevations. Template library topics in those slides: types of components,
        feature definitions, point names, constraints, dynamic settings, active template, labels and parametric constraints, template testing, end condition priorities, drag and drop.
      cite: "ORD - Creating Proposed Terrains.pptx slide 2"
      test: {action: name, value: "corridors-and-templates", pattern: "^[a-z0-9-]+$"}
    - id: S5
      title: Where the learning material lives
      who: engineer
      manual_reason: "External resources (Bentley, DOT training sites, AECOM University, Teams/ProjectWise recordings)"
      how: |
        The workflow sheet links FHWA, Ohio DOT, Tennessee DOT ORD manuals, AECOM University, Bentley Connect learning and the upstate/downstate ORD recordings on Teams and ProjectWise. The recordings in this
        library (Standards/Tutorials) are those recordings.
      cite: "OpenRoads Training.xlsx sheet Sheet1"
```
