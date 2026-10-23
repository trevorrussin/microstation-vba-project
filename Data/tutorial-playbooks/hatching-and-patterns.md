---
title: Hatching, cross-hatching and pattern areas (NYSDOT detail patterns)
when: Hatching an area, cross-hatching, filling with a NYSDOT pattern cell (stone, concrete, embankment, topsoil...), or fixing a pattern that is too small / too large / missing.
keywords: hatch, hatching, cross hatch, crosshatch, pattern, pattern area, hatch area, associative, flood, fill, embankment, stone, concrete, drop pattern, true scale
---

# Hatching and patterns

```yaml playbook
playbook:
  topic: hatching-and-patterns
  title: Hatching, cross-hatching and pattern areas (NYSDOT detail patterns)
  when: Hatching an area, cross-hatching, filling with a NYSDOT pattern cell, or fixing a wrong-size pattern
  keywords: [hatch, hatching, cross hatch, crosshatch, pattern, pattern area, hatch area, associative, flood, fill, embankment, stone, concrete, drop pattern, true scale]
  sources:
    - "Hatching in Microstation-20250610_082510-Meeting Recording.mp4 @ 00:06"
    - "Hatching in Microstation-20250610_082510-Meeting Recording.mp4 @ 02:06"
  preflight:
    - {id: PF1, check: cell, name: OP620SM, lib: ny_detail_pattern.cel, why: "pattern cell: stone, medium (ny_detail_pattern.cel)"}
    - {id: PF2, check: cell, name: OP203FEM, lib: ny_detail_pattern.cel, why: "pattern cell: fill, embankment material"}
    - {id: PF3, check: level, name: O_Details_Breakline_P, why: "details breakline level for boundaries"}
  steps:
    - id: S1
      title: Draw the area as a shape first
      who: agent
      how: |
        Place a closed shape of the area to hatch before hatching: it is cleaner, and the shape becomes the associative boundary. The agent's
        hatch tool sets the pattern on an existing closed shape (HasPattern), it does not draw the area.
      cite: "Hatching in Microstation-20250610_082510-Meeting Recording.mp4 @ 00:06"
      test: {action: hatch, spacing: 5.0, angle: 45.0}
    - id: S2
      title: Hatch Area - associative boundary
      who: engineer
      manual_reason: "Annotate > Hatch Area tool settings (Drop Pattern off, Associative Boundary on, Snappable on) are interactive tool-box options"
      how: |
        Drawing workflow > Annotate > Hatch Area. Make sure Drop Pattern is NOT selected; tick Associative Boundary and Snappable; pick the
        boundary element. The hatch is then ONE item and updates when the boundary moves or a vertex is added. The Flood option hatches the enclosed
        area under the cursor, so no shape is needed. Spacing and angle are in the tool settings (e.g. spacing 1, angle 45).
      cite: "Hatching in Microstation-20250610_082510-Meeting Recording.mp4 @ 00:19"
    - id: S3
      title: Cross-hatch an existing hatch, and remove one
      who: agent
      how: |
        Cross-hatch = hatch first, then set the Cross hatch property true (Properties) and change spacing/angle there; the agent's
        crosshatch_element does this on an already-hatched closed element. remove_hatch removes the associative pattern. Both are read back
        (HasPattern) by the runner.
      cite: "Hatching in Microstation-20250610_082510-Meeting Recording.mp4 @ 01:26"
      test: {action: hatch, cross: true, remove: true, spacing: 5.0, angle: 45.0}
    - id: S4
      title: Pattern Area from a NYSDOT pattern cell
      who: engineer
      manual_reason: "Pattern Area tool with the cell-library picker is interactive"
      how: |
        Annotate > Pattern Area, Pattern Definition `From Cell`, Cell Library `ny_detail_pattern.cel` (New York detail pattern library),
        pick the cell (for example OP620SM stone medium, OP203FEM embankment material, OP203EIP embankment in place), tick Associative Boundary,
        then select the boundary element or flood. Patterns are annotative cells: the model annotation scale matters.
      cite: "Hatching in Microstation-20250610_082510-Meeting Recording.mp4 @ 02:06"
    - id: S5
      title: Pattern too small, too large or invisible
      who: engineer
      manual_reason: "Pattern properties (Properties dialog) on the placed element"
      how: |
        Alert "Do you really want to place 1721 x 1503 patterns?" means the scale is wrong - answer No, fix it, place again. Annotative patterns
        follow the model annotation scale: untick True Scale and/or set the model scale to the detail scale (the file in the recording was set
        to 10 scale; that made the pattern appear). Fix an existing pattern in Properties > Pattern: Scale, Angle, Origin (a good stone
        pattern showed PA=0, PS=120 at its scale).
      pitfall: "Selecting the boundary shows the pattern attached to it; moving/reshaping the boundary re-flows the pattern."
      cite: "Hatching in Microstation-20250610_082510-Meeting Recording.mp4 @ 03:05"
    - id: S6
      title: Level and class of the region
      who: agent
      how: |
        The associative region takes the ACTIVE level (the recording showed it on O_Details_Dimensions_Extension_Lines_P). Set the active level
        deliberately before hatching (patterns on the pattern/details level, not on a dimension level) and verify the level after.
      cite: "Hatching in Microstation-20250610_082510-Meeting Recording.mp4 @ 00:52"
      test: {action: levels, levels: [O_Details_Breakline_P]}
```
