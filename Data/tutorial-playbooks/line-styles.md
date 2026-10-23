---
title: Line styles - custom NYSDOT styles, scale factor, copy parallel, changing attributes
when: Applying or fixing a line style (guide rail, fence, cut line, fill line, breaklines, dashed), a style that looks too small/large/continuous, copying an element parallel, or changing level/style/weight of existing elements.
keywords: line style, linestyle, scale factor, rgcb, lfnc, lcut, lfill, guide rail, fence, copy parallel, change attributes, match attributes, bylevel, dashed, line style scale
---

# Line styles

The recording is a silent screen capture of a roadway file (D900059_fea_rwy.dgn, 1"=40' = annotation scale 480). Every step
below is read from the on-screen dialogs and keyframes.

```yaml playbook
playbook:
  topic: line-styles
  title: Line styles - custom NYSDOT styles, scale factor, copy parallel, changing attributes
  when: Applying or fixing a line style, a style that looks too small/large, copying parallel, changing attributes
  keywords: [line style, linestyle, scale factor, rgcb, lfnc, lcut, lfill, guide rail, fence, copy parallel, change attributes, match attributes, bylevel, dashed, line style scale]
  sources:
    - "Manipulating Line Style Settings.mp4 @ 00:30"
    - "Manipulating Line Style Settings.mp4 @ 04:20"
    - "Manipulating Line Style Settings.mp4 @ 07:10"
  preflight:
    - {id: PF1, check: line_style, name: RGCB, why: "roadway guide rail, concrete barrier custom style"}
    - {id: PF2, check: line_style, name: LFNC, why: "landscape fence custom style"}
    - {id: PF3, check: line_style, name: LCUT_P, why: "landscape cut line custom style"}
    - {id: PF4, check: level, name: RGCB_P, why: "Roadway, Guide Rail, Concrete Barrier, Proposed"}
    - {id: PF5, check: level, name: LFNC_P, why: "Landscape, Fences, Proposed"}
  steps:
    - id: S1
      title: Use the NYSDOT line style by name, ByLevel color and weight
      who: agent
      how: |
        Custom line styles are named for their feature (RGCB concrete barrier, LFNC fence, LCUT_P cut line, LFILL_P fill line, O_BREAK breakline).
        A level carries its feature's style and symbology: place the element on the right level (RGCB_P, LFNC_P ...) and keep Color/Weight ByLevel
        ("Color: ByLevel (0), Weight: ByLevel" in Change Element Attributes). Resolve the exact style name with resolve_line_style before
        applying; the Number property is not a valid key.
      cite: "Manipulating Line Style Settings.mp4 @ 04:36"
      test: {action: line_style, style: RGCB}
    - id: S2
      title: Fence style applied and read back
      who: agent
      how: "LFNC draws the fence ticks (X marks). Apply with change_element_symbology(line_style_name=...) and confirm the element reports the style."
      cite: "Manipulating Line Style Settings.mp4 @ 06:56"
      test: {action: line_style, style: LFNC}
    - id: S3
      title: Style levels are usable
      who: agent
      how: "Guide rail, fence, cut line and fill line levels must accept elements before styles are applied."
      cite: "Manipulating Line Style Settings.mp4 @ 02:38"
      test: {action: levels, levels: [RGCB_P, LFNC_P, LCUT_P, LFILL_P]}
    - id: S4
      title: Line Style Settings - Scale factor = annotation scale
      who: engineer
      manual_reason: "Line Style Settings dialog (Settings > Line Style Settings): custom Bentley controls; the scale modifier is not readable through the object model, so the agent cannot verify it"
      how: |
        Select the style in the Names list. Tick `Scale factor` and set it to the annotation scale of the file (480 for 1"=40'); Origin/End/True Width
        control width; Shift None; "Click to Activate" makes it the active style for new elements. A custom style at scale 1.0 in a
        480-scale file draws far too small (looks continuous/invisible). Ask the engineer to confirm the scale factor by eye; the QA vision review
        is the only automated check.
      pitfall: "Dashes or ticks that look wrong on a sheet are almost always a style scale that does not match the model annotation scale."
      cite: "Manipulating Line Style Settings.mp4 @ 00:30"
    - id: S5
      title: Modify Line Style Attributes - width
      who: engineer
      manual_reason: "Modify Line Style Attributes tool settings (Width, Absolute, True Width)"
      how: |
        Modify > Modify Line Style Attributes: set Width (tick Absolute for a fixed width, True Width for plotted width), then click the
        element. Use Change Whole Element rather than a fence part when the element should keep one style.
      cite: "Manipulating Line Style Settings.mp4 @ 07:10"
    - id: S6
      title: Change Element Attributes / Match Attributes
      who: engineer
      manual_reason: "Change Element Attributes and Match Element Attributes are interactive tool-box tools"
      how: |
        Change Element Attributes: Active Attributes Level (e.g. RGCB_P), Color ByLevel, Style, Weight ByLevel, Class, Change Whole Element; pick the
        element. Match Element Attributes copies every setting (SmartMatch) from a good element onto others. The agent's equivalent is
        change_element_level + change_element_symbology (journal-gated); the QA gate checks the result.
      cite: "Manipulating Line Style Settings.mp4 @ 04:32"
    - id: S7
      title: Copy Parallel at a fixed distance
      who: agent
      how: |
        Move/Copy Parallel tool: Mode Miter, Distance (e.g. 0.75 or 1.5 ft), Use Active Attributes, Make Copy. "No Elements Found" means the pick was an element
        in a REFERENCE (only the active file's elements can be copied). The agent's copy_parallel is + LEFT of start-to-end (ORD offsets are
        the opposite); the contract measures the side and distance.
      cite: "Manipulating Line Style Settings.mp4 @ 04:20"
      test: {op: copy_parallel, args: {element_id: $line, distance: 0.75}}
    - id: S8
      title: Level Display to isolate what you are editing
      who: engineer
      manual_reason: "Level Display dialog changes the engineer's view"
      how: |
        Level Display (view settings): turn off all levels except the ones you are editing, so references and other features do not get picked. Restore
        the view afterwards; the agent must not change the engineer's level display.
      cite: "Manipulating Line Style Settings.mp4 @ 02:04"
```
