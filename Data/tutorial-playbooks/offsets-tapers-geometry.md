---
title: Offsets, tapers and parallel geometry - conventions and civil vs plain graphics
when: Placing pavement edges, lane lines, tapers, parallel copies, arrays or any geometry defined as an offset from a centerline/baseline.
keywords: offset, offsets, taper, tapers, parallel, copy parallel, pavement edge, edge of pavement, lane, array, centerline, baseline, curb return, driveway
---

# Offsets and tapers

ORD "civil geometry" features remember their relationships (design intent); plain CAD graphics do not. Our bridge
tools draw PLAIN graphics, so a later change to the baseline does NOT move geometry the agent offset from it.

```yaml playbook
playbook:
  topic: offsets-tapers-geometry
  title: Offsets, tapers and parallel geometry - conventions and civil vs plain graphics
  when: Placing pavement edges, lane lines, tapers, parallel copies, arrays or any geometry defined as an offset
  keywords: [offset, offsets, taper, tapers, parallel, copy parallel, pavement edge, edge of pavement, lane, array, centerline, baseline, curb return, driveway]
  sources:
    - "000_ORD_Tutorial.pdf page 174"
    - "000_ORD_Tutorial.pdf page 176"
    - "000_ORD_Tutorial.pdf page 180"
  preflight:
    - {id: PF1, check: level, name: PE_P, why: "pavement edge, proposed"}
    - {id: PF2, check: level, name: AC_P, why: "alignment, control, proposed (the baseline an offset hangs from)"}
  steps:
    - id: S1
      title: Know which geometry is civil and which is plain
      who: engineer
      manual_reason: "Choosing civil geometry tools and feature definitions is an engineering decision made in the ribbon"
      how: |
        Civil geometry (Geometry > Horizontal > Offsets and Tapers) keeps rules: edges stay parallel to the centerline, side
        roads stay perpendicular, arcs stay tangent. Plain graphics (what the agent draws) do nothing when the baseline moves.
        If the engineer expects the geometry to follow design changes, ask for civil Offsets and Tapers; otherwise tell the
        engineer the agent's offsets are static and must be redrawn after a baseline change.
      cite: "000_ORD_Tutorial.pdf page 168"
    - id: S2
      title: Sign convention - ORD vs the copy_parallel tool
      who: agent
      how: |
        ORD offsets: NEGATIVE = left of the baseline, POSITIVE = right (engineering standard; baseline direction is the
        direction it was drawn). Our `copy_parallel` tool is the opposite: distance > 0 = LEFT of start-to-end. To place an
        offset the engineer states in ORD terms use distance = -offset. The QA gate (OP-01) measures the side and fails a
        wrong one.
      cite: "000_ORD_Tutorial.pdf page 174"
      test: {op: copy_parallel, args: {element_id: $line, distance: 12}}
    - id: S3
      title: Constant offset keeps the edge parallel; changing one end makes a taper
      who: agent
      how: |
        ORD: the midpoint offset manipulator keeps an element parallel; changing an end manipulator makes a taper. With plain
        graphics build the taper explicitly: a polyline from offset d0 at the start station to d1 at the end station
        (taper length L, ratio = L / (d1 - d0)). Draw it as one polyline so it stays one element.
      cite: "000_ORD_Tutorial.pdf page 176"
      test: {op: place_polyline, args: {vertices: [[4600, 930, 0], [4700, 930, 0], [4750, 942, 0], [4850, 942, 0]]}}
    - id: S4
      title: Single Offset Entire Element with Mirror
      who: engineer
      manual_reason: "Civil Offsets and Tapers tool prompts (Locate Element, Enter Offset, Mirror) are interactive"
      how: |
        Geometry > Horizontal > Offsets and Tapers > Single Offset Entire Element. Click the CURVE of a complex centerline so the
        whole feature offsets (clicking a tangent offsets only the tangent). Type the offset and press Return to lock it,
        set Mirror = Yes to place both sides in one click (Mirror makes two ruled offsets, NOT a relationship between
        them). Use Persist Snaps and the Feature Definition toolbar (Linear > Pavement > Road_EdgeOfPavement) so standards
        are applied "set it and forget it". Disable Use Spiral Transitions unless the source has a spiral (a true parallel
        spiral is not mathematically possible).
      cite: "000_ORD_Tutorial.pdf page 180"
    - id: S5
      title: Repeated parallel copies (array)
      who: agent
      how: "array_element makes copies 1..N at i*(spacing_x, spacing_y); the original is NOT one of the N. Use it for evenly spaced parallels."
      cite: "MicrostationTipsAndTricks.pptx slide 19"
      test: {op: array_element, args: {element_id: $line, count: 3, spacing_x: 0, spacing_y: 5}}
    - id: S6
      title: Set a mark before an experiment
      who: engineer
      manual_reason: "Set Mark / Undo to Mark are Quick Access toolbar commands"
      how: "Quick Access > Set Mark (or key-in MARK) before trying a geometry change you may want to reverse; the agent's own undo is undo_last_op."
      cite: "000_ORD_Tutorial.pdf page 182"
```
