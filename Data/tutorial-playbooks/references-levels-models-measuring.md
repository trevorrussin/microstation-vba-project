---
title: References, levels, models and measuring widths - what to attach, what to turn on, where linework hides
when: Attaching or reading reference files, linework not showing after attaching, turning levels on/off, finding hidden linework in another model, measuring lane/shoulder/barrier widths and finding where a cross-section changes.
keywords: reference, references, attach reference, level display, levels, all on, model, model view, measure distance, measure, perpendicular, snaps, multi-snap, lane width, shoulder width, barrier width, widening, narrowing, transition, working folder, civil model
---

# References, levels, models and measuring

Training session (narrated). MicroStation analogy used by the instructor: a file is assembled from components (survey, roadway, striping, utilities)
kept in separate files so files stay small and organized; references put them back together; levels switch features on and off; models are folders inside a file.

```yaml playbook
playbook:
  topic: references-levels-models-measuring
  title: References, levels, models and measuring widths
  when: Attaching/reading references, linework not showing, level display, hidden linework in another model, measuring widths
  keywords: [reference, references, attach reference, level display, levels, all on, model, model view, measure distance, measure, perpendicular, snaps, multi-snap, lane width, shoulder width, barrier width, widening, narrowing, transition, working folder, civil model]
  sources:
    - "References, Levels, Model Views.mp4 @ 00:13"
    - "References, Levels, Model Views.mp4 @ 04:36"
    - "References, Levels, Model Views.mp4 @ 11:12"
    - "MicrostationTipsAndTricks.pptx slide 11"
    - "MicrostationTipsAndTricks.pptx slide 16"
  preflight:
    - {id: PF1, check: level, name: PE_P, why: "pavement edge, proposed (roadway linework)"}
    - {id: PF2, check: level, name: RGCB_P, why: "concrete barrier, proposed"}
    - {id: PF3, check: level, name: AC_P, why: "alignment, control, proposed"}
  steps:
    - id: S1
      title: Keep components in separate files, official work under CAD
      who: engineer
      manual_reason: "File organization is a ProjectWise folder decision"
      how: |
        Utilities with utilities, striping with striping, survey/map files as the design basis. Sketches, miscellaneous alternatives and anything
        unofficial live under Working; anything that goes into the plan set lives under CAD > AECOM > Civil > Model. File names follow the naming convention
        (feature ALG, HCL, SGN ...): read the convention before guessing what a file is.
      cite: "References, Levels, Model Views.mp4 @ 03:02"
    - id: S2
      title: Attach references and read the two columns that matter
      who: engineer
      manual_reason: "References dialog / Attach Reference is a GUI tool and changes the engineer's file"
      how: |
        Attach Tools > References > Attach. The File Name and Model columns are the important ones; keep the display toggle next to the file name. Surveys have
        `map` / `sur` in the name. Attach the roadway linework file (pavement edges, concrete barriers) and the striping file (sgn) with the defaults
        (Coincident - World) and OK the prompts.
      cite: "References, Levels, Model Views.mp4 @ 00:56"
    - id: S3
      title: Attached but nothing shows - the levels are off
      who: engineer
      manual_reason: "Level Display dialog changes the view the engineer is looking at"
      how: |
        A reference can be on while every level inside it is off. Open Level Display, right-click > All On (for the whole reference), then turn off
        what you do not want (signs, survey) so the work is readable. Levels separate features (gas vs water utility; striping vs utilities) so each plan
        sheet can show only what it needs; show the Description column to read what each level is (AC alignment, ACP alignment point, stationing tick ...).
      pitfall: "Linework that 'disappeared' is almost always a level that is off, a reference that is toggled off, or a different model."
      cite: "References, Levels, Model Views.mp4 @ 04:36"
    - id: S4
      title: Levels you will meet in roadway plans are usable
      who: agent
      how: "Pavement edge (PE_P), concrete barrier (RGCB_P) and control alignment (AC_P) are NYSDOT library levels in this workspace."
      cite: "References, Levels, Model Views.mp4 @ 06:17"
      test: {action: levels, levels: [PE_P, RGCB_P, AC_P]}
    - id: S5
      title: Models are folders inside a file - linework can hide in another model
      who: engineer
      manual_reason: "Opening/changing model views changes the engineer's screen"
      how: |
        A sheet file typically has a model for the references, a model for the border and a model that pulls them together. If you cannot find linework it may
        be in a different model; a reference's Settings (right-click) selects which model of the referenced file shows. Models can hold draft and alternative
        versions, which is powerful and also lets things hide. Use the active model deliberately and confirm it with describe_drawing_state.
      cite: "References, Levels, Model Views.mp4 @ 06:38"
    - id: S6
      title: Measure a width perpendicular to the line
      who: agent
      how: |
        Analyze > Measure Distance with Perpendicular snapping to the line; keep Multi-snap on (Intersection, Midpoint, Keypoint) plus Near when snapping to an element
        that has no key point. Typical readings in the recording: lane 12 ft, shared use path 14 ft narrowing to 12.52 ft. The agent measures from saved geometry the same
        way (perpendicular distance between two parallel lines) and reports the value to 0.01 ft.
      cite: "References, Levels, Model Views.mp4 @ 11:12"
      test: {action: measure, distance: 12}
    - id: S7
      title: Find where the cross-section changes
      who: agent
      how: |
        Walk the line from start to end measuring shoulder, lanes, barrier and shared-use path widths; look for places where they widen or narrow. Report the BEGIN and END
        of each transition (the contractor must be told exactly where a 10 ft shoulder becomes 1 ft or 0 ft, a 12 ft lane becomes 11 ft). Hovering a line shows where a segment
        begins and ends, which is usually the transition point. Stay on the main line and the shared use path unless told otherwise; the curve is where widths change. Mark each
        finding with a shape and a short text note (Annotate > Place Text); neatness does not matter, what you find does.
      cite: "References, Levels, Model Views.mp4 @ 20:20"
      test: {op: place_polyline, args: {vertices: [[4600, 960, 0], [4640, 960, 0], [4640, 970, 0], [4600, 970, 0], [4600, 960, 0]]}}
    - id: S8
      title: Annotation scale - small or huge text
      who: agent
      how: |
        Text placed at the wrong scale shows tiny or enormous ("it's coming in at a really small scale, but we'll fix that"); 48 was too big for the working scale.
        Place text with the sheet's text style and let the annotation scale supply the plotted size; the QA gate compares new text to the sheet's existing styles.
      cite: "References, Levels, Model Views.mp4 @ 10:04"
      test: {action: text_style, style: ny_Prop Normal, text: "EXIT RAMP GOES FROM ONE TO TWO LANES"}
```
