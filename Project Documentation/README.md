# NYSDOT Workzone Traffic Control Designer

A MicroStation V8i VBA tool that automates the placement of NYSDOT workzone traffic control (WZTC) signs, elements, and cell symbols along a user-drawn alignment, in compliance with MUTCD New York State supplement standards.

---

## Problem Statement

NYSDOT traffic control designers must produce plan sheets showing workzone setups that comply with MUTCD NY spacing and sign requirements. Traditionally, this meant:

- Manually looking up spacing values from NYSDOT 619 standard sheets for the given road speed and category
- Manually calculating distances along the alignment for each sign, taper, and element
- Manually placing sign face cells, post cells, text labels, hatch areas, and device symbols one at a time

For a typical workzone, this process involves 30–100+ individual element placements, significant risk of arithmetic error, and no standard reference check during placement.

**This tool eliminates that manual process** by letting the designer draw the alignment once, configure the workzone parameters once, and then step through automated placement of all required elements.

---

## Assumptions

- The design file is in **feet** (master units = feet)
- The active MicroStation model is a 2D plan view
- Cell libraries are located at standard NYSDOT ProjectWise paths:
  - Sign faces: `c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel`
  - WZTC symbols: `c:\pwworking\usny\d0119091\ny_plan_wztc.cel`
- The alignment is drawn as a **continuous chain** — each segment starts at the endpoint of the previous one
- Workzone areas are approximately convex (for centroid-based auto-hatch to land inside)
- The user understands their workzone category and can identify the required sign numbers from the NYSDOT 619 standard sheets

---

## Architecture

The tool is structured as a 6-step sequential workflow, each step backed by a module + form pair:

```
LaunchWZTC (Module1)
    └─► WZTCDesigner
            ├─► SheetViewer (NYSDOT 619 reference viewer, modeless)
            └─► StartWZTCDrawing (Module6)
                    └─► AlignDraw  ← user draws lines/arcs
                            └─► GroupAndLaunchPlacement (Module6)
                                    └─► StartAlignmentPlacement (ModuleAlignmentPlacement)
                                            └─► PlacePerp  ← user places perp lines
                                                    └─► StartSignPlacement (ModuleSignPlacement)
                                                            └─► PlaceSign  ← user draws signs
                                                                    └─► StartWZTCElementsPlacement (ModuleWZTCElements)
                                                                            └─► PlaceElements  ← user draws elements
                                                                                    └─► StartWZTCCellPlacement (ModuleWZTCCells)
                                                                                            └─► PlaceCells  ← user places cells
```

### Module Roles

| File | Role |
|------|------|
| `Launcher.bas` | Entry points (`LaunchWZTC`, `LaunchNYSDOTViewer`) |
| `SignLibrary.bas` | Sign library: loads default signs, looks up spacing/size by sign number |
| `SignTypes.bas` | `Public Type signData` definition (shared across all modules) |
| `AlignmentTool.bas` | Alignment drawing: snapshot max element ID, show AlignDraw, group elements |
| `SharedState.bas` | **All public persistent state** — survives form unload/reload |
| `PerpPlacement.bas` | Path geometry engine: build arc/line chain, interpolate points, place perp lines |
| `SignPlacer.bas` | Sign step state machine: index tracking, accessor functions for `PlaceSign` |
| `DrawSign.bas` | Low-level sign drawing: click collection, projection onto perp line, post + face + text |
| `DrawElements.bas` | WZTC element drawing: level/color/weight setup, shape + hatch, line elements |
| `CellPlacer.bas` | Cell library: populate catalogue, attach library, interactive placement with count tracking |
| `PlaceButtons.cls` | `WithEvents` sink for dynamically created `Place Line` / `Skip` buttons |
| `SignNumBox.cls` | `WithEvents` sink for dynamically created sign number textboxes |

### State Persistence Model

Because MicroStation VBA resets all local variables when a form is unloaded, all designer configuration is stored in **public module-level variables** in `SharedState.bas`. This allows:

- Any form to read the configuration
- `WZTCDesigner` to restore its previous state when reopened via **Return to Designer**
- Sign geometry (perpendicular line midpoints and normal vectors) to pass from the alignment placement step to the sign drawing step without form dependencies

### Alignment Path Engine (`PerpPlacement.bas`)

The path engine stores each drawn segment as a `PathSeg` UDT containing:
- `IsArc` flag
- Start/end XYZ, center XYZ (arcs), radius, start angle, sweep angle
- Segment length (chord length for lines, arc length for arcs)

`BuildAlignmentPath` scans all elements newer than `wztcAlignmentStartMaxID`, sorts them into a connected chain by endpoint proximity, and resolves arc orientation by checking which of the arc's two geometric endpoints is closest to the current chain position.

`GetPointAndTangent(dist)` walks the chain using cumulative arc-length to return the XY position and unit tangent at any distance along the alignment.

`PlacePerpendicularLine` draws a tick-line of half-length 20 ft perpendicular to the tangent at each WZTC item location and stores the midpoint + normal vector for the sign drawing step.

---

## Example Output

For a 45 mph, 4-lane divided freeway workzone (Category F):

1. **Designer configures:**
   - Category: Freeway, Sheet 619-3, Speed: 45 mph, Freeway, Lane: 12 ft, Shoulder: 8 ft
   - Signs: W20-1 (48" x 48", 1000 ft spacing), W20-5 (48" x 48", 500 ft), R2-1 (36" x 48", 250 ft)
   - WZTC Order: Downstream Taper → Roll Ahead → W20-1 → Buffer → W20-5 → R2-1 → Work Area

2. **User draws alignment** (3 lines + 1 arc = ~2000 ft total)

3. **Perpendicular lines placed** at computed locations along the alignment:
   - Downstream Taper tick at 0 ft
   - Roll Ahead tick at 150 ft
   - W20-1 sign location at 650 ft (150 + 500)
   - Buffer Space tick at 900 ft
   - W20-5 at 1150 ft
   - R2-1 at 1400 ft
   - Work Area at 1650 ft

4. **Signs drawn:** User clicks post location on each tick — sign face cell, post cell (TWZSGN_P), text label (sign number + size) placed automatically at each location

5. **WZTC elements drawn:** Work Space polygon + auto-hatch, channelizing device lines on correct levels

6. **Cell symbols placed:** Arrow panels, flaggers, etc. from ny_plan_wztc.cel

**Result:** A complete NYSDOT-compliant WZTC plan layout drawn directly in the design file, with correct levels, colors, and element spacing.

---

## Design Decisions & Tradeoffs

### Decision: Public module variables instead of a class or database
**Why:** MicroStation VBA does not support persistent objects across form loads. Using public variables in a standard module (`SharedState.bas`) is the only reliable way to pass state between the sequential form steps. A class module would be reset when unloaded; a file-based store would add I/O complexity and a dependency on a writable path.

**Tradeoff:** State is lost when the VBA project is reset or MicroStation is closed. There is no save/load of a project file. For long multi-session projects, the user must redo the configuration step.

### Decision: Sequential form workflow (not a wizard)
**Why:** Each step requires active MicroStation interaction (drawing, clicking). A single form with tabs would need to hide/show sections and block interaction between steps, which is fragile in a modeless VBA environment.

**Tradeoff:** The user must proceed in order. Going backward requires navigating manually via Back buttons. There is no branching or non-linear workflow.

### Decision: `ae.Origin` for arc center instead of derivation from chain point
**Why:** MicroStation's `ArcElement` stores the geometric center in `.Origin`. Earlier code derived the center from `chainPt - r*(cos(sa), sin(sa))`, which assumes the path enters at the geometric start angle. MicroStation can store arcs in either direction relative to the drawn path, causing the center to be computed from the wrong endpoint.

**Tradeoff:** Using `.Origin` is correct and robust. It requires checking both geometric endpoints for proximity to the chain to determine travel direction.

### Decision: Centroid-based auto-hatch for Work Space
**Why:** The hatch command (`HATCH ICON`) requires a data point inside the closed shape. Computing the centroid of the user's clicked vertices gives a point near the center of the drawn area without requiring the user to make an extra click.

**Tradeoff:** For strongly non-convex shapes (e.g., an L-shape), the centroid may fall outside the boundary, causing hatch to fail silently. The typical workzone work-space area is approximately convex, so this is acceptable.

### Decision: DrawSign.bas name for the sign drawing module
**Why:** This module evolved from a testing scaffold and retained its name (`ModTest`) through iterative development. The name does not reflect its current role.

**Tradeoff:** Confusing name in a production codebase. Future refactoring should rename it to `ModuleSignDrawing` or similar. The module is fully functional; this is a cosmetic issue only.

### Decision: Form-based dynamic control creation instead of IDE-designed controls
**Why:** The sign table in `WZTCDesigner` and the WZTC order rows in `PlacePerp` are variable-length lists. Creating them dynamically in `UserForm_Initialize` allows any number of rows without requiring the IDE designer.

**Tradeoff:** Dynamic controls require `WithEvents` class modules (`PlaceButtons`, `SignNumBox`) since standard VBA event syntax only works with named IDE-designed controls. This adds two class files to the project. Dynamic controls also require the `ControlExists()` guard before every access, since they may not exist if the form fails to initialize.

### Decision: Sign size quote normalization at save time
**Why:** MUTCD sign sizes are always in inches (e.g., `48" x 48"`). Users sometimes type `'` (foot mark) instead of `"` (inch mark). The sign library stores correct `"` characters. Normalizing at `btnSubmit_Click` (`Replace(value, "'", Chr(34))`) converts user input silently without affecting library-filled values.

**Tradeoff:** A `'` typed intentionally in a sign size string is replaced. In practice, `'` in an inch context is always an error, so this is safe.

### Decision: Cell library path hardcoded to `c:\pwworking\...`
**Why:** NYSDOT ProjectWise WorkSpace maps all project files to a standard local path. Hardcoding avoids requiring the user to browse for libraries on every run.

**Tradeoff:** Tool will not work if the ProjectWise WorkSpace is not mounted or if a different project path is used. The path constant (`WZTC_CELL_LIB` in `CellPlacer.bas`, sign library paths in `SignLibrary.bas` and `DrawSign.bas`) would need to be updated for different ProjectWise environments.
