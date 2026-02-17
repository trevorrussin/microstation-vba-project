# WORKZONE TRAFFIC CONTROL DESIGNER — INSTALLATION GUIDE

## Quick Navigation

- [Step 1: Open VBA Editor](#step-1-open-vba-editor)
- [Step 2: Import All Source Files](#step-2-import-all-source-files)
- [Step 3: WZTCDesigner — Add Controls](#step-3-frmworkzonedesigner--add-controls)
- [Step 4: SheetViewer — NYSDOT Reference Viewer](#step-4-userform1--nysdot-reference-viewer)
- [Step 5: AlignDraw — Alignment Drawing Tool](#step-5-alignmentform--alignment-drawing-tool)
- [Step 6: PlacePerp — Perpendicular Line Placement](#step-6-frmalignmentplacement--perpendicular-line-placement)
- [Step 7: PlaceSign — Sign Drawing Tool](#step-7-frmsignplacement--sign-drawing-tool)
- [Step 8: PlaceElements — WZTC Elements Drawing](#step-8-frmwztcelements--wztc-elements-drawing)
- [Step 9: PlaceCells — Cell Library Placement](#step-9-frmwztccells--cell-library-placement)
- [Step 10: Run the Tool](#step-10-run-the-tool)
- [Troubleshooting](#troubleshooting)
- [File Reference](#file-reference)
- [Unused / Legacy Files](#unused--legacy-files)

---

## Overview

This guide walks you through installing the Workzone Traffic Control (WZTC) Designer tool in MicroStation VBA. The tool guides you step-by-step from configuring workzone parameters through drawing an alignment, placing perpendicular tick-lines, drawing signs, WZTC construction elements, and placing cell library symbols.

**Complete workflow:**
1. **WZTCDesigner** → Configure workzone (speed, category, signs, WZTC order)
2. **AlignDraw** → Draw alignment (lines and arcs) in MicroStation
3. **PlacePerp** → Place perpendicular lines at each WZTC item location
4. **PlaceSign** → Draw sign faces and posts at each perpendicular line
5. **PlaceElements** → Draw construction elements (work space, channelizing, barrier, etc.)
6. **PlaceCells** → Place WZTC cell library symbols from ny_plan_wztc.cel

---

## STEP 1: Open VBA Editor

1. Open **MicroStation**
2. Press **Alt + F11** to open the VBA Editor

---

## STEP 2: Import All Source Files

Import every `.bas`, `.frm`, and `.cls` file from the project folder into the VBA project:

1. In the VBA Editor menu: **File → Import File** (or right-click the project in Project Explorer → **Import File**)
2. Import these files **in order** (modules first, then forms, then classes):

### Standard Modules (`.bas`) — in `Modules/` subfolder
| File | Module Name | Purpose |
|------|-------------|---------|
| `Launcher.bas` | Launcher | Launch entry points |
| `SignLibrary.bas` | SignLibrary | Sign library: `signData` type, default signs, lookup by sign number |
| `AlignmentTool.bas` | AlignmentTool | Alignment drawing tool logic |
| `SharedState.bas` | SharedState | All public persistent state variables |
| `PerpPlacement.bas` | PerpPlacement | Path geometry + perpendicular lines |
| `DrawSign.bas` | DrawSign | Sign placement state + drawing (post, face, text) at perpendicular lines |
| `DrawElements.bas` | DrawElements | WZTC elements drawing |
| `CellPlacer.bas` | CellPlacer | Cell library placement |

### UserForms (`.frm`) — in `UserForms/` subfolder
| File | Form Name | Purpose |
|------|-----------|---------|
| `WZTCDesigner.frm` | WZTCDesigner | Main workzone designer |
| `SheetViewer.frm` | SheetViewer | NYSDOT 619 sheet reference viewer |
| `AlignDraw.frm` | AlignDraw | Alignment drawing tool |
| `PlacePerp.frm` | PlacePerp | Perpendicular line placement |
| `PlaceSign.frm` | PlaceSign | Sign drawing step |
| `PlaceElements.frm` | PlaceElements | WZTC elements drawing step |
| `PlaceCells.frm` | PlaceCells | Cell library placement step |

### Class Modules (`.cls`) — in `Class Modules/` subfolder
| File | Class Name | Purpose |
|------|-----------|---------|
| `PlaceButtons.cls` | PlaceButtons | `WithEvents` class for PlacePerp buttons |
| `SignNumBox.cls` | SignNumBox | `WithEvents` class for dynamic sign number textboxes |

> **Do NOT import** legacy files (see [Unused / Legacy Files](#unused--legacy-files) section).

---

## STEP 3: WZTCDesigner — Add Controls

After importing `WZTCDesigner.frm`, open it in the form designer and add these controls manually.

### 3A. Form Properties
| Property | Value |
|----------|-------|
| **(Name)** | `WZTCDesigner` |
| **Width** | `1220` |
| **Height** | `730` |
| **Caption** | `Workzone Design Tool` |
| **StartUpPosition** | `1 - CenterOwner` |

### 3B. Input Dropdowns (Top of Form)
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|---|---|---|---|---|---|---|---|
| Label | `lblCategory` | `Workzone Category:` | 10 | 20 | 120 | 18 | Bold |
| ComboBox | `cboCategory` | (empty) | 10 | 150 | 250 | 20 | Style: 2-Dropdown List |
| Label | `lblSheet` | `Standard Sheet:` | 40 | 20 | 120 | 18 | Bold |
| ComboBox | `cboSheet` | (empty) | 40 | 150 | 250 | 20 | Style: 2-Dropdown List |
| Label | `lblRoadSpeed` | `Road Speed (mph):` | 70 | 20 | 120 | 18 | Bold |
| ComboBox | `cboRoadSpeed` | (empty) | 70 | 150 | 250 | 20 | Style: 2-Dropdown List |
| Label | `lblRoadType` | `Road Type:` | 100 | 20 | 120 | 18 | Bold |
| ComboBox | `cboRoadType` | (empty) | 100 | 150 | 250 | 20 | Style: 2-Dropdown List |
| Label | `lblLaneWidth` | `Lane Width (ft):` | 130 | 20 | 120 | 18 | Bold |
| ComboBox | `cboLaneWidth` | (empty) | 130 | 150 | 250 | 20 | Style: 2-Dropdown List |
| Label | `lblShoulderWidth` | `Shoulder Width (ft):` | 160 | 20 | 120 | 18 | Bold |
| ComboBox | `cboShoulderWidth` | (empty) | 160 | 150 | 250 | 20 | Style: 2-Dropdown List |

### 3C. Spacing & Clearances Frame
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|---|---|---|---|---|---|---|---|
| Frame | `frameSpacingValues` | `Spacing & Clearances` | 195 | 10 | 560 | 220 | Code creates all labels + textboxes inside automatically |

> All textboxes inside `frameSpacingValues` are **created dynamically by code** — do not add them manually.

### 3D. Sign Table Section
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|---|---|---|---|---|---|---|---|
| Label | `lblSignTable` | `Required Signs & Details` | 425 | 10 | 500 | 18 | Bold |
| Frame | `frameSignTable` | `Sign Selection` | 445 | 10 | 560 | 190 | ScrollBars: Vertical |

### 3E. Action Buttons
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|---|---|---|---|---|---|---|---|
| CommandButton | `btnAddRow` | `+` | 645 | 20 | 40 | 25 | Bold, Size 12 — adds sign row |
| CommandButton | `btnRemoveRow` | `-` | 645 | 65 | 40 | 25 | Bold, Size 12 — removes last row |
| CommandButton | `btnReference` | `Reference (MUTCD)` | 645 | 220 | 130 | 25 | Opens NYSDOT 619 viewer |
| CommandButton | `btnSubmit` | `Submit & Draw` | 645 | 440 | 130 | 25 | Bold — validates and starts drawing |
| CommandButton | `btnClear` | `Clear All` | 645 | 580 | 90 | 25 | Resets all selections (confirm dialog) |
| Label | `lblStatus` | `Ready - Select options` | 680 | 20 | 550 | 20 | Status messages |

### 3F. WZTC Order Section
> **IMPORTANT:** All 6 controls below are placed **directly on the form** — do NOT nest them inside `frameWZTCOrder`.

| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|---|---|---|---|---|---|---|---|
| Frame | `frameWZTCOrder` | `WZTC Order` | 195 | 855 | 305 | 315 | Decorative border only |
| ListBox | `lstWZTCOrder` | (empty) | 215 | 865 | 220 | 255 | On form, NOT inside frame |
| CommandButton | `btnOrderUp` | `Up` | 225 | 1090 | 60 | 22 | Moves item up |
| CommandButton | `btnOrderDown` | `Down` | 253 | 1090 | 60 | 22 | Moves item down |
| CommandButton | `btnOrderDelete` | `X Del` | 285 | 1090 | 60 | 22 | Removes item |
| CommandButton | `btnRefreshOrder` | `Refresh Order` | 475 | 865 | 120 | 22 | Resyncs with sign table |

**CRITICAL:** Control names are case-sensitive and must match exactly.

---

## STEP 4: SheetViewer — NYSDOT Reference Viewer

`SheetViewer.frm` is imported as-is. No additional controls need to be added manually — the form layout is handled entirely in code.

**What it does:** Displays NYSDOT 619 Standard Sheet images as a reference while configuring the workzone. Opened by clicking **Reference (MUTCD)** in WZTCDesigner.

---

## STEP 5: AlignDraw — Alignment Drawing Tool

`AlignDraw.frm` is imported as-is. Add these controls manually in the form designer:

### Form Properties
| Property | Value |
|---|---|
| **(Name)** | `AlignDraw` |
| **Caption** | `Draw Alignment` |
| **Width** | `220` |
| **Height** | `185` |

### Controls
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|---|---|---|---|---|---|---|---|
| CommandButton | `cmdStartLine` | `Start Line Segment` | 20 | 20 | 160 | 28 | Draws a straight alignment segment |
| CommandButton | `cmdStartArc` | `Start Arc Segment` | 58 | 20 | 160 | 28 | Draws a curved alignment segment |
| CommandButton | `cmdDone` | `Done` | 130 | 20 | 160 | 30 | Bold — finalizes alignment and proceeds |
| Label | `lblStatus` | *(empty)* | 96 | 20 | 160 | 28 | Status messages |

**What it does:** Shown after **Submit & Draw** in WZTCDesigner. The user draws lines and arcs to define the workzone alignment. Clicking **Done** groups the drawn elements and advances to PlacePerp.

---

## STEP 6: PlacePerp — Perpendicular Line Placement

`PlacePerp.frm` is imported as-is. The form creates all dynamic controls (rows per WZTC order item) in code — no manual controls needed beyond the basics below.

### Form Properties
| Property | Value |
|---|---|
| **(Name)** | `PlacePerp` |
| **Caption** | `WZTC Alignment Placement` |
| **Width** | `500` |
| **Height** | `600` |

### Controls to Add Manually
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|---|---|---|---|---|---|---|---|
| Label | `lblTitle` | `Place perpendicular lines along alignment` | 8 | 10 | 470 | 16 | Bold |
| Frame | `frameItems` | *(empty)* | 30 | 10 | 470 | 480 | ScrollBars: Vertical — dynamic rows added here |
| CommandButton | `btnNext` | `Next: Draw Signs` | 520 | 10 | 145 | 25 | Bold — advances to sign placement |
| CommandButton | `btnBack` | `< Back` | 520 | 160 | 90 | 25 | Returns to AlignDraw |
| CommandButton | `btnReturnToDesigner` | `Return to Designer` | 520 | 258 | 145 | 25 | Returns to WZTCDesigner with state restored |

**What it does:** For each WZTC order item (spacing parameters and signs), shows the item name, spacing, and a **Place Line** button. Clicking Place Line draws a perpendicular tick-line on the alignment at the computed location and records sign geometry for the sign placement step.

---

## STEP 7: PlaceSign — Sign Drawing Tool

`PlaceSign.frm` is imported as-is. Add these controls manually:

### Form Properties
| Property | Value |
|---|---|
| **(Name)** | `PlaceSign` |
| **Caption** | `WZTC Sign Placement` |
| **Width** | `330` |
| **Height** | `275` |

### Controls
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|---|---|---|---|---|---|---|---|
| Label | `lblSignOf` | `Initialising...` | 8 | 10 | 300 | 16 | Sign counter (e.g. "Sign 1 of 3:") |
| Label | `lblSignName` | *(empty)* | 26 | 10 | 300 | 22 | Sign number, large bold blue text |
| Label | `lblSignSide` | *(empty)* | 52 | 10 | 300 | 16 | "One Side" or "Both Sides" description |
| Label | `lblInstruction` | *(empty)* | 76 | 10 | 300 | 34 | Click instructions, word-wrap |
| CommandButton | `btnDraw` | `Draw Sign` | 120 | 10 | 90 | 23 | Bold — click then click post in MicroStation |
| CommandButton | `btnNextSign` | `Next Sign` | 120 | 108 | 90 | 23 | Advances to next sign |
| CommandButton | `btnCancel` | `Cancel` | 120 | 206 | 75 | 23 | Cancel with confirm |
| CommandButton | `btnWZTCElements` | `Next: WZTC Elements` | 151 | 10 | 145 | 23 | Bold — proceed to WZTC elements step |
| Label | `lblStatus` | `Ready` | 182 | 10 | 300 | 42 | Status/error messages, word-wrap |
| CommandButton | `btnBack` | `< Back` | 232 | 10 | 90 | 23 | Returns to PlacePerp |
| CommandButton | `btnReturnToDesigner` | `Return to Designer` | 232 | 108 | 145 | 23 | Returns to WZTCDesigner with state restored |

**What it does:** Steps through each sign that had a perpendicular line placed. For each sign, the user clicks **Draw Sign** then clicks the post location(s) on the perpendicular line in MicroStation. The post line, post cell (TWZSGN_P), sign face cell, and text label are all placed automatically.

- **One Side** signs: 1 post click required
- **Both Sides** signs: 2 post clicks required; a connecting arc is drawn between the posts

---

## STEP 8: PlaceElements — WZTC Elements Drawing

`PlaceElements.frm` is imported as-is. Add these controls manually:

### Form Properties
| Property | Value |
|---|---|
| **(Name)** | `PlaceElements` |
| **Caption** | `WZTC Elements` |
| **Width** | `320` |
| **Height** | `320` |

### Controls
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|---|---|---|---|---|---|---|---|
| Label | `lblTitle` | `Draw WZTC Construction Elements` | 8 | 10 | 290 | 16 | Bold |
| Label | `lblElement` | *(empty)* | 28 | 10 | 290 | 18 | Current element name |
| Label | `lblInstruction` | *(empty)* | 50 | 10 | 290 | 34 | Instructions, word-wrap |
| CommandButton | `btnDrawElement` | `Draw Element` | 92 | 10 | 110 | 23 | Bold — starts drawing in MicroStation |
| CommandButton | `btnNextElement` | `Next Element` | 92 | 126 | 110 | 23 | Advances to next element type |
| CommandButton | `btnSkipElement` | `Skip` | 92 | 242 | 60 | 23 | Skips current element type |
| Label | `lblStatus` | `Ready` | 122 | 10 | 290 | 50 | Status messages, word-wrap |
| CommandButton | `btnGoCellLib` | `Next: Cell Library` | 178 | 10 | 130 | 23 | Bold — proceeds to cell library step |
| CommandButton | `btnBack` | `< Back` | 210 | 10 | 90 | 23 | Returns to PlaceSign |
| CommandButton | `btnReturnToDesigner` | `Return to Designer` | 210 | 108 | 145 | 23 | Returns to WZTCDesigner with state restored |

**What it does:** Steps through 5 WZTC construction element types:
1. **Work Space** (TWZWS2_P) — draw a closed polygon; workzone hatch applied automatically
2. **Channelizing Devices** (TWZCD_P) — draw line segments
3. **Pavement Marking Removal** (TWZPMRC_P) — draw line segments
4. **Temporary Barrier** (TWZBT_P) — draw line segments
5. **Temp. Barrier w/Warning Lights** (TWZBTWL_P) — draw line segments

All elements are drawn with **color 6, weight 2** on their respective levels. Right-click ends each element draw session.

---

## STEP 9: PlaceCells — Cell Library Placement

`PlaceCells.frm` is imported as-is. Add these controls manually:

### Form Properties
| Property | Value |
|---|---|
| **(Name)** | `PlaceCells` |
| **Caption** | `WZTC Cell Library` |
| **Width** | `320` |
| **Height** | `260` |

### Controls
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|---|---|---|---|---|---|---|---|
| Label | `lblTitle` | `Place WZTC Cell Library Symbols` | 8 | 10 | 290 | 16 | Bold |
| Label | `lblSelect` | `Select Symbol:` | 30 | 10 | 100 | 16 | — |
| ComboBox | `cmbCellSelect` | *(empty)* | 28 | 115 | 185 | 20 | Style: 2-Dropdown List — symbol list |
| CommandButton | `btnPlaceCell` | `Place Cell` | 58 | 10 | 100 | 23 | Bold — hides form, enter placement mode |
| CommandButton | `btnFinish` | `Finish` | 58 | 116 | 80 | 23 | Close form when done |
| Label | `lblStatus` | `Ready` | 90 | 10 | 290 | 50 | Status / placement count messages |
| CommandButton | `btnBack` | `< Back` | 148 | 10 | 90 | 23 | Returns to PlaceElements |
| CommandButton | `btnReturnToDesigner` | `Return to Designer` | 148 | 108 | 145 | 23 | Returns to WZTCDesigner with state restored |

**What it does:** Lets the user place any of 16 WZTC cell symbols from `ny_plan_wztc.cel` (Arrow Panel, Flagger, Impact Attenuator, etc.). The form hides during placement so MicroStation has full mouse focus. Right-click returns to the form. A running count of each symbol type placed is shown and saved.

Available symbols: Arrow Panel, Arrow Panel (Closed), Arrow Panel (Type), Barricade, Changeable Message Sign, Flagger, Flagger in Traffic, Impact Attenuator, Luminaire, Sequential Directional Taper, Seq. Dir. Taper (Dynamic), Sign Post (TWZSGN_P), Signal, Warning Lights, Warning Vehicle, Warning Vehicle w/Attenuator.

---

## STEP 10: Run the Tool

### From MicroStation
1. Open any MicroStation design file
2. Press **Alt + F11** to open the VBA Editor
3. In the Immediate Window (**Ctrl + G**), type:
   ```
   LaunchWZTC
   ```
4. Press **Enter** — the Workzone Designer form opens

### Alternatively
In the VBA Editor Immediate Window:
```
WZTCDesigner.Show vbModeless
```

---

## TROUBLESHOOTING

### "Sub or Function Not Defined"
- Verify all `.bas` files from the import table were imported (especially `DrawSign.bas`, `SignLibrary.bas`, `DrawElements.bas`, `CellPlacer.bas`)
- Check that class modules `PlaceButtons.cls` and `SignNumBox.cls` were imported

### "Control not found" Error
- Verify control names match exactly (case-sensitive)
- Check that `frameSpacingValues` and `frameSignTable` exist on `WZTCDesigner`
- Verify `frameItems` exists on `PlacePerp`
- Ensure `cmbCellSelect` exists on `PlaceCells`

### Form Opens Blank (no dropdowns)
- Code populates dropdowns on `UserForm_Initialize` — if blank, try closing and reopening with `LaunchWZTC`
- Check that `SharedState.bas` is imported (required for all public state variables)

### Signs Drawing in Wrong Location / Off the Arc
- Make sure `PerpPlacement.bas` is the current version (has `ae.Origin` arc fix)
- The alignment must be drawn as a continuous chain — each segment must start at the endpoint of the previous one

### Return to Designer Shows Blank Form
- Previous session state is restored automatically from `ModuleWZTCData` public vars
- If blank, no previous **Submit & Draw** was run in this VBA session (state resets when VBA project reloads)
- Use the **Clear All** button to explicitly reset if the old state looks wrong

### Cell Library Form Locks MicroStation
- This was fixed: `PlaceCells` now hides itself before entering `PLACE CELL` mode
- Verify `PlaceCells.frm` is the current version

### Workzone Hatch Not Appearing
- Hatch is applied automatically only when **3 or more** points are clicked before right-clicking
- The closed shape must be approximately convex for centroid-based auto-hatch to work

---

## FILE REFERENCE

### Active Source Files

| File | Type | Purpose |
|------|------|---------|
| `Launcher.bas` | Standard Module | Entry points: `LaunchWZTC`, `LaunchNYSDOTViewer` |
| `SignLibrary.bas` | Standard Module | Sign library: `signData` type, load/get signs, cell name and path for drawing |
| `AlignmentTool.bas` | Standard Module | Alignment drawing: `StartWZTCDrawing`, `GroupAndLaunchPlacement` |
| `SharedState.bas` | Standard Module | All public persistent state variables |
| `PerpPlacement.bas` | Standard Module | Path geometry engine + perpendicular line placement |
| `DrawSign.bas` | Standard Module | Sign placement state + `DrawSignAtPerpLine` — post, sign face, text (uses SignLibrary) |
| `DrawElements.bas` | Standard Module | WZTC element drawing (level/color/hatch) |
| `CellPlacer.bas` | Standard Module | Cell library placement + count tracking |
| `WZTCDesigner.frm` | UserForm | Main workzone configuration form |
| `SheetViewer.frm` | UserForm | NYSDOT 619 standard sheet reference viewer |
| `AlignDraw.frm` | UserForm | Alignment drawing (line/arc segments + Done) |
| `PlacePerp.frm` | UserForm | Perpendicular line placement per WZTC item |
| `PlaceSign.frm` | UserForm | Sign drawing step (post + face + text) |
| `PlaceElements.frm` | UserForm | WZTC elements drawing step |
| `PlaceCells.frm` | UserForm | Cell library placement step |
| `PlaceButtons.cls` | Class Module | `WithEvents` handler for dynamic placement buttons |
| `SignNumBox.cls` | Class Module | `WithEvents` handler for dynamic sign number textboxes |
| `INSTALLATION_GUIDE.md` | Documentation | This file |

---

## UNUSED / LEGACY FILES

The following files exist in the project folder but are **not part of the active workflow**. Do not import them into the VBA project.

| File | Reason Not Needed |
|------|-------------------|
| `LegacySignPlace.bas` | Intermediate sign placement approach — superseded by `DrawSign.bas` |
| `LegacyPrototype.bas` | Early hardcoded prototype with fixed project coordinates — not dynamic |
| `DebugTest.bas` | Debugging utility (control existence checker) — not part of workflow |
| `LegacyCells.bas` | Old hardcoded cell macro (`BmrWZTCOther`) — superseded by `CellPlacer.bas` |
| `LegacyElements.bas` | Reference/example file for HATCH ICON pattern — logic now in `DrawElements.bas` |
| `LegacyDesigner.frm` | Old form, not referenced anywhere in the workflow |
| `LegacyAlign.frm` | Old alignment drawing form — superseded by `AlignDraw.frm` |
| `DesignerRef.bas` | Code reference paste-from file for initial setup only — form code now lives in `WZTCDesigner.frm` |

---

## KEY FEATURES

### State Persistence Across Sessions
All workzone configuration (dropdowns, sign table, WZTC order, spacing values) is saved to public variables in `SharedState.bas` when **Submit & Draw** is clicked. Clicking **Return to Designer** on any form reopens `WZTCDesigner` with all previous selections restored. Use the **Clear All** button to explicitly start fresh.

### MUTCD NY Spacing Calculations
Automatically calculates MUTCD NY spacing values (Downstream Taper, Roll Ahead Distance, Vehicle Space, Buffer Space, Merging/Shifting Taper, Shoulder Taper, Advanced Warning Spacing) based on road speed. Supports 25–90 mph in 5 mph increments.

### Dynamic Sign Table
- Add rows with **+**, remove last row with **-**
- Each row: Sign Number, Spacing, Size, Side (One Side / Both Sides)
- Sign number auto-fills spacing and size from the sign library when you leave the field

### WZTC Order Panel
- Displays all workzone items in sequence (spacing labels + sign numbers)
- Reorder with Up/Down buttons, remove with X Del, refresh with Refresh Order
- Order determines the sequence of perpendicular line placement along the alignment

### Arc Alignment Support
Perpendicular lines are placed correctly along both straight and curved alignment segments using proper arc geometry (StartAngle, SweepAngle, origin).

### Auto-Hatch for Work Space
After drawing the Work Space closed polygon, workzone hatch is applied automatically at the centroid of the clicked points.

### Cell Placement Count Tracking
`PlaceCells` tracks how many times each cell type has been placed and displays a running count after each placement.
