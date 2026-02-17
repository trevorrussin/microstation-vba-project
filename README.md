# NYSDOT Workzone Traffic Control Designer

A MicroStation VBA tool that automates the layout of NYSDOT workzone traffic control plans. Instead of manually placing every sign, taper, and element by hand, this tool lets you configure the workzone once and then walks you through placing everything in the correct order with the correct MUTCD NY spacings.

---

## What This Tool Does

Designing a workzone traffic control plan normally means:

- Looking up spacing values from the NYSDOT 619 standard sheets for the road speed, category, lane width, and shoulder width
- Calculating cumulative distances along the alignment for each sign, taper, buffer, and element
- Placing 30-100+ individual elements (sign faces, posts, text labels, hatched areas, channelizing devices) one at a time

This tool handles all of that automatically. You draw the alignment, fill in your workzone parameters, and the tool places everything at the right locations along the alignment with the correct levels, colors, and spacings.

---

## Requirements

- MicroStation with VBA support (V8i, CONNECT, or 2023)
- Design file units set to **feet**
- NYSDOT ProjectWise WorkSpace mounted (for cell libraries):
  - Sign faces: `ny_plan_nmutcd_signface.cel`
  - WZTC symbols: `ny_plan_wztc.cel`
- The alignment must be drawn as a **continuous chain** (each segment connects to the previous one)

---

## How It Works — Step by Step

The tool guides you through 6 steps in order:

### Step 1: Configure the Workzone (WZTCDesigner form)
Select the workzone category, 619 standard sheet number, road speed, road type (Freeway or Non-Freeway), lane width, and shoulder width. The form automatically calculates all MUTCD NY spacings (downstream taper, roll ahead, buffer space, merging taper, shoulder taper, etc.). Then add your required signs to the sign selection table — the built-in sign library auto-fills spacing and size when you type a sign number. You can view the 619 standard sheet for reference at any time. The WZTC Order panel shows the sequence items will be placed along the alignment, and you can reorder them.

### Step 2: Draw the Alignment (AlignDraw form)
Draw lines and arcs to trace the alignment path. Each segment connects automatically to the previous one. Click "Done" when finished — the tool groups all alignment elements together.

### Step 3: Place Perpendicular Lines (PlacePerp form)
The tool walks along the alignment and places an 80-ft perpendicular tick line at each item location (tapers, signs, work area, etc.), spaced according to the values from Step 1. For each item you can accept the suggested spacing or adjust it, and you can skip items you don't need.

### Step 4: Draw Signs (PlaceSign form)
For each sign that had a perpendicular line placed, you click where on the tick line to place the sign post. The tool automatically places the sign face cell, post cell, post line, and text label (sign number and size) at that location. For "Both Sides" signs, you click two points.

### Step 5: Draw WZTC Elements (PlaceElements form)
Draw the remaining workzone elements in sequence: Work Space polygon (with hatch), channelizing device lines, removal striping, temporary barrier, and barrier with warning lights. Each element is placed on its correct NYSDOT level. For the Work Space, you trace the boundary shape, then click inside it to apply the hatch pattern.

### Step 6: Place Cell Symbols (PlaceCells form)
Place any additional WZTC cell symbols (arrow panels, flaggers, etc.) from the ny_plan_wztc.cel library.

---

## File Descriptions

| File | What It Does |
|------|-------------|
| `Launcher.bas` | Starts the tool — run `LaunchWZTC` to begin |
| `SignLibrary.bas` | Contains 150+ MUTCD sign definitions with cell names, sizes (Freeway and Non-Freeway), and default spacings. When you type a sign number in the designer, this is where the auto-fill data comes from. |
| `AlignmentTool.bas` | Handles the alignment drawing step. Tracks what elements you draw so the tool knows which lines and arcs form your alignment. |
| `SharedState.bas` | Stores all your workzone configuration (speeds, spacings, sign selections, etc.) so it persists between the different steps. If you go back to the designer form, your previous selections are still there. |
| `PerpPlacement.bas` | The alignment geometry engine. It takes your drawn alignment, builds a connected path of lines and arcs, and calculates where each perpendicular tick line should go based on your configured spacings. |
| `DrawSign.bas` | Places the sign face cell, post cell, post line, and text label at each sign location along the perpendicular lines. |
| `DrawElements.bas` | Handles drawing the WZTC shape elements (work space, channelizing devices, barriers, etc.) on the correct NYSDOT levels. |
| `CellPlacer.bas` | Lets you browse and place additional WZTC cell symbols from the ny_plan_wztc.cel library. |
| `DesignerRef.bas` | Provides the WZTC order table logic and NYSDOT 619 standard sheet reference data. |
| `PlaceButtons.cls` | Handles the "Place Line" and "Skip" button clicks in the perpendicular line placement step. |
| `SignNumBox.cls` | Handles the sign number text box behavior — when you finish typing a sign number, it triggers the library lookup to auto-fill spacing and size. |

---

## Example

For a 45 mph freeway workzone (12 ft lanes, 8 ft shoulder):

1. **Configure:** Select Freeway, 45 mph, 12 ft lane, 8 ft shoulder. The tool calculates: Downstream Taper = 100 ft, Roll Ahead = 160 ft, Buffer Space = 360 ft, Merging Taper = 560 ft, Shoulder Taper = 120 ft. Add signs W20-05, W20-03, R02-01 to the sign table.

2. **Draw alignment:** Trace your alignment with lines and arcs (~2000 ft total).

3. **Place perpendicular lines:** The tool places 80-ft tick lines at each calculated location along the alignment.

4. **Draw signs:** Click on each tick line to place signs. The sign face, post, and label appear automatically.

5. **Draw elements:** Trace the work space boundary and click inside it to hatch. Draw channelizing device lines and barrier lines on the correct levels.

6. **Place cells:** Add any remaining symbols (arrow panels, flaggers, etc.).

**Result:** A complete NYSDOT-compliant WZTC plan layout in the design file, ready for review.

---

## Design Decisions & Tradeoffs

### Decision: Public module variables instead of a class or database
**Why:** MicroStation VBA does not support persistent objects across form loads. Using public variables in a standard module (`SharedState.bas`) is the only reliable way to pass state between the sequential form steps. A class module would be reset when unloaded; a file-based store would add I/O complexity and a dependency on a writable path.

**Tradeoff:** State is lost when the VBA project is reset or MicroStation is closed. There is no save/load of a project file. For long multi-session projects, the user must redo the configuration step.

### Decision: Sequential form workflow (not a wizard)
**Why:** Each step requires active MicroStation interaction (drawing, clicking). A single form with tabs would need to hide/show sections and block interaction between steps, which is fragile in a modeless VBA environment.

**Tradeoff:** The user must proceed in order. Going backward requires navigating manually via Back buttons. There is no branching or non-linear workflow.

### Decision: Dual-method arc center resolution instead of single derivation
**Why:** MicroStation can store arcs with either geometric endpoint as the start angle. Earlier code assumed the chain point was always at the geometric start angle and derived `center = chainPt - r*(cos(sa), sin(sa))`, which produced wrong centers when MicroStation stored the arc in reverse orientation. The fix first tries `ae.CenterPoint` (available in MicroStation 2023 / CONNECT edition), then falls back to computing both candidate centers and validating against `ae.Range` (works in all versions). Once the center is known, both geometric endpoints are compared to the chain point to determine travel direction.

**Tradeoff:** The Range-based fallback adds a small amount of computation per arc segment. In practice this is negligible since alignment chains rarely have more than a few dozen segments.

### Decision: User-click hatch for Work Space
**Why:** The hatch command (`HATCH ICON`) requires a data point inside the closed shape. Earlier code computed the centroid of the user's clicked vertices automatically, but this failed silently for non-convex shapes (e.g., L-shapes) where the centroid falls outside the boundary. The current approach asks the user to click inside the shape after drawing it, which works reliably for any shape.

**Tradeoff:** Requires one extra click from the user after drawing the work space boundary. In practice this is a minor step and gives the user full control over hatch placement.

### Decision: DrawSign.bas name for the sign drawing module
**Why:** This module evolved from a testing scaffold and was renamed from `ModTest` to `DrawSign` during refactoring to accurately reflect its role as the sign drawing engine.

**Tradeoff:** None remaining. The rename resolved the earlier naming confusion.

### Decision: Form-based dynamic control creation instead of IDE-designed controls
**Why:** The sign table in `WZTCDesigner` and the WZTC order rows in `PlacePerp` are variable-length lists. Creating them dynamically in `UserForm_Initialize` allows any number of rows without requiring the IDE designer.

**Tradeoff:** Dynamic controls require `WithEvents` class modules (`PlaceButtons`, `SignNumBox`) since standard VBA event syntax only works with named IDE-designed controls. This adds two class files to the project. Dynamic controls also require the `ControlExists()` guard before every access, since they may not exist if the form fails to initialize.

### Decision: Sign size quote normalization at save time
**Why:** MUTCD sign sizes are always in inches (e.g., `48" x 48"`). Users sometimes type `'` (foot mark) instead of `"` (inch mark). The sign library stores correct `"` characters. Normalizing at `btnSubmit_Click` (`Replace(value, "'", Chr(34))`) converts user input silently without affecting library-filled values.

**Tradeoff:** A `'` typed intentionally in a sign size string is replaced. In practice, `'` in an inch context is always an error, so this is safe.

### Decision: Cell library path hardcoded to `c:\pwworking\...`
**Why:** NYSDOT ProjectWise WorkSpace maps all project files to a standard local path. Hardcoding avoids requiring the user to browse for libraries on every run.

**Tradeoff:** Tool will not work if the ProjectWise WorkSpace is not mounted or if a different project path is used. The path constant (`WZTC_CELL_LIB` in `CellPlacer.bas`, sign library paths in `SignLibrary.bas` and `DrawSign.ba
