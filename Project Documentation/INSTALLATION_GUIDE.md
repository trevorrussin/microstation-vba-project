# WORKZONE TRAFFIC CONTROL DESIGNER — INSTALLATION GUIDE

## Quick Navigation

- [Step 1: Open VBA Editor](#step-1-open-vba-editor)
- [Step 2: Import All Source Files](#step-2-import-all-source-files)
- [Step 3: frmWorkzoneDesigner — Add Controls](#step-3-frmworkzonedesigner--add-controls)
- [Step 4: UserForm1 — NYSDOT Reference Viewer](#step-4-userform1--nysdot-reference-viewer)
- [Step 5: AlignmentForm — Alignment Drawing Tool](#step-5-alignmentform--alignment-drawing-tool)
- [Step 6: frmAlignmentPlacement — Perpendicular Line Placement](#step-6-frmalignmentplacement--perpendicular-line-placement)
- [Step 7: frmSignPlacement — Sign Drawing Tool](#step-7-frmsignplacement--sign-drawing-tool)
- [Step 8: frmWZTCElements — WZTC Elements Drawing](#step-8-frmwztcelements--wztc-elements-drawing)
- [Step 9: frmWZTCCells — Cell Library Placement](#step-9-frmwztccells--cell-library-placement)
- [Step 10: Run the Tool](#step-10-run-the-tool)
- [Troubleshooting](#troubleshooting)
- [File Reference](#file-reference)
- [Unused / Legacy Files](#unused--legacy-files)

---

## Overview

This guide walks you through installing the Workzone Traffic Control (WZTC) Designer tool in MicroStation VBA. The tool guides you step-by-step from configuring workzone parameters through drawing an alignment, placing perpendicular tick-lines, drawing signs, WZTC construction elements, and placing cell library symbols.

**Complete workflow:**
1. **frmWorkzoneDesigner** → Configure workzone (speed, category, signs, WZTC order)
2. **AlignmentForm** → Draw alignment (lines and arcs) in MicroStation
3. **frmAlignmentPlacement** → Place perpendicular lines at each WZTC item location
4. **frmSignPlacement** → Draw sign faces and posts at each perpendicular line
5. **frmWZTCElements** → Draw construction elements (work space, channelizing, barrier, etc.)
6. **frmWZTCCells** → Place WZTC cell library symbols from ny_plan_wztc.cel

---

## STEP 1: Open VBA Editor

1. Open **MicroStation**
2. Press **Alt + F11** to open the VBA Editor

---

## STEP 2: Import All Source Files

Import every `.bas`, `.frm`, and `.cls` file from the project folder into the VBA project:

1. In the VBA Editor menu: **File → Import File** (or right-click the project in Project Explorer → **Import File**)
2. Import these files **in order** (modules first, then forms, then classes):

### Standard Modules (`.bas`)
| File | Module Name | Purpose |
|------|-------------|---------|
| `Module1.bas` | Module1 | Launch entry points |
| `Module3.bas` | Module3 | Sign library manager |
| `Module4.bas` | Module4 | `signData` type definition |
| `Module6.bas` | Module6 | Alignment drawing tool logic |
| `ModuleWZTCData.bas` | ModuleWZTCData | All public persistent state variables |
| `ModuleAlignmentPlacement.bas` | ModuleAlignmentPlacement | Path geometry + perpendicular lines |
| `ModuleSignPlacement.bas` | ModuleSignPlacement | Sign placement state machine |
| `ModuleWZTCElements.bas` | ModuleWZTCElements | WZTC elements drawing |
| `ModuleWZTCCells.bas` | ModuleWZTCCells | Cell library placement |
| `ModTest.bas` | ModTest | Sign drawing at perpendicular lines |

### UserForms (`.frm`)
| File | Form Name | Purpose |
|------|-----------|---------|
| `frmWorkzoneDesigner.frm` | frmWorkzoneDesigner | Main workzone designer |
| `UserForm1.frm` | UserForm1 | NYSDOT 619 sheet reference viewer |
| `AlignmentForm.frm` | AlignmentForm | Alignment drawing tool |
| `frmAlignmentPlacement.frm` | frmAlignmentPlacement | Perpendicular line placement |
| `frmSignPlacement.frm` | frmSignPlacement | Sign drawing step |
| `frmWZTCElements.frm` | frmWZTCElements | WZTC elements drawing step |
| `frmWZTCCells.frm` | frmWZTCCells | Cell library placement step |

### Class Modules (`.cls`)
| File | Class Name | Purpose |
|------|-----------|---------|
| `PlacementButtonHandler.cls` | PlacementButtonHandler | `WithEvents` class for frmAlignmentPlacement buttons |
| `SignNumberBoxHandler.cls` | SignNumberBoxHandler | `WithEvents` class for dynamic sign number textboxes |

> **Do NOT import** legacy files (see [Unused / Legacy Files](#unused--legacy-files) section).

---

## STEP 3: frmWorkzoneDesigner — Add Controls

After importing `frmWorkzoneDesigner.frm`, open it in the form designer and add these controls manually.

### 3A. Form Properties
| Property | Value |
|----------|-------|
| **(Name)** | `frmWorkzoneDesigner` |
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

## STEP 4: UserForm1 — NYSDOT Reference Viewer

`UserForm1.frm` is imported as-is. No additional controls need to be added manually — the form layout is handled entirely in code.

**What it does:** Displays NYSDOT 619 Standard Sheet images as a reference while configuring the workzone. Opened by clicking **Reference (MUTCD)** in frmWorkzoneDesigner.

---

## STEP 5: AlignmentForm — Alignment Drawing Tool

`AlignmentForm.frm` is imported as-is. Add these controls manually in the form designer:

### Form Properties
| Property | Value |
|---|---|
| **(Name)** | `AlignmentForm` |
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

**What it does:** Shown after **Submit & Draw** in frmWorkzoneDesigner. The user draws lines and arcs to define the workzone alignment. Clicking **Done** groups the drawn elements and advances to frmAlignmentPlacement.

---

## STEP 6: frmAlignmentPlacement — Perpendicular Line Placement

`frmAlignmentPlacement.frm` is imported as-is. The form creates all dynamic controls (rows per WZTC order item) in code — no manual controls needed beyond the basics below.

### Form Properties
| Property | Value |
|---|---|
| **(Name)** | `frmAlignmentPlacement` |
| **Caption** | `WZTC Alignment Placement` |
| **Width** | `500` |
| **Height** | `600` |

### Controls to Add Manually
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|---|---|---|---|---|---|---|---|
| Label | `lblTitle` | `Place perpendicular lines along alignment` | 8 | 10 | 470 | 16 | Bold |
| Frame | `frameItems` | *(empty)* | 30 | 10 | 470 | 480 | ScrollBars: Vertical — dynamic rows added here |
| CommandButton | `btnNext` | `Next: Draw Signs` | 520 | 10 | 145 | 25 | Bold — advances to sign placement |
| CommandButton | `btnBack` | `< Back` | 520 | 160 | 90 | 25 | Returns to AlignmentForm |
| CommandButton | `btnReturnToDesigner` | `Return to Designer` | 520 | 258 | 145 | 25 | Returns to frmWorkzoneDesigner with state restored |

**What it does:** For each WZTC order item (spacing parameters and signs), shows the item name, spacing, and a **Place Line** button. Clicking Place Line draws a perpendicular tick-line on the alignment at the computed location and records sign geometry for the sign placement step.

---

## STEP 7: frmSignPlacement — Sign Drawing Tool

`frmSignPlacement.frm` is imported as-is. Add these controls manually:

### Form Properties
| Property | Value |
|---|---|
| **(Name)** | `frmSignPlacement` |
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
| CommandButton | `btnBack` | `< Back` | 232 | 10 | 90 | 23 | Returns to frmAlignmentPlacement |
| CommandButton | `btnReturnToDesigner` | `Return to Designer` | 232 | 108 | 145 | 23 | Returns to frmWorkzoneDesigner with state restored |

**What it does:** Steps through each sign that had a perpendicular line placed. For each sign, the user clicks **Draw Sign** then clicks the post location(s) on the perpendicular line in MicroStation. The post line, post cell (TWZSGN_P), sign face cell, and text label are all placed automatically.

- **One Side** signs: 1 post click required
- **Both Sides** signs: 2 post clicks required; a connecting arc is drawn between the posts

---

## STEP 8: frmWZTCElements — WZTC Elements Drawing

`frmWZTCElements.frm` is imported as-is. Add these controls manually:

### Form Properties
| Property | Value |
|---|---|
| **(Name)** | `frmWZTCElements` |
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
| CommandButton | `btnBack` | `< Back` | 210 | 10 | 90 | 23 | Returns to frmSignPlacement |
| CommandButton | `btnReturnToDesigner` | `Return to Designer` | 210 | 108 | 145 | 23 | Returns to frmWorkzoneDesigner with state restored |

**What it does:** Steps through 5 WZTC construction element types:
1. **Work Space** (TWZWS2_P) — draw a closed polygon; workzone hatch applied automatically
2. **Channelizing Devices** (TWZCD_P) — draw line segments
3. **Pavement Marking Removal** (TWZPMRC_P) — draw line segments
4. **Temporary Barrier** (TWZBT_P) — draw line segments
5. **Temp. Barrier w/Warning Lights** (TWZBTWL_P) — draw line segments

All elements are drawn with **color 6, weight 2** on their respective levels. Right-click ends each element draw session.

---

## STEP 9: frmWZTCCells — Cell Library Placement

`frmWZTCCells.frm` is imported as-is. Add these controls manually:

### Form Properties
| Property | Value |
|---|---|
| **(Name)** | `frmWZTCCells` |
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
| CommandButton | `btnBack` | `< Back` | 148 | 10 | 90 | 23 | Returns to frmWZTCElements |
| CommandButton | `btnReturnToDesigner` | `Return to Designer` | 148 | 108 | 145 | 23 | Returns to frmWorkzoneDesigner with state restored |

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
frmWorkzoneDesigner.Show vbModeless
```

---

## TROUBLESHOOTING

### "Sub or Function Not Defined"
- Verify all `.bas` files from the import table were imported (especially `ModTest.bas`, `ModuleSignPlacement.bas`, `ModuleWZTCElements.bas`, `ModuleWZTCCells.bas`)
- Check that class modules `PlacementButtonHandler.cls` and `SignNumberBoxHandler.cls` were imported

### "Control not found" Error
- Verify control names match exactly (case-sensitive)
- Check that `frameSpacingValues` and `frameSignTable` exist on `frmWorkzoneDesigner`
- Verify `frameItems` exists on `frmAlignmentPlacement`
- Ensure `cmbCellSelect` exists on `frmWZTCCells`

### Form Opens Blank (no dropdowns)
- Code populates dropdowns on `UserForm_Initialize` — if blank, try closing and reopening with `LaunchWZTC`
- Check that `ModuleWZTCData.bas` is imported (required for all public state variables)

### Signs Drawing in Wrong Location / Off the Arc
- Make sure `ModuleAlignmentPlacement.bas` is the current version (has `ae.Origin` arc fix)
- The alignment must be drawn as a continuous chain — each segment must start at the endpoint of the previous one

### Return to Designer Shows Blank Form
- Previous session state is restored automatically from `ModuleWZTCData` public vars
- If blank, no previous **Submit & Draw** was run in this VBA session (state resets when VBA project reloads)
- Use the **Clear All** button to explicitly reset if the old state looks wrong

### Cell Library Form Locks MicroStation
- This was fixed: `frmWZTCCells` now hides itself before entering `PLACE CELL` mode
- Verify `frmWZTCCells.frm` is the current version

### Workzone Hatch Not Appearing
- Hatch is applied automatically only when **3 or more** points are clicked before right-clicking
- The closed shape must be approximately convex for centroid-based auto-hatch to work

---

## FILE REFERENCE

### Active Source Files

| File | Type | Purpose |
|------|------|---------|
| `Module1.bas` | Standard Module | Entry points: `LaunchWZTC`, `LaunchNYSDOTViewer` |
| `Module3.bas` | Standard Module | Sign library (load/get/match signs from cell library) |
| `Module4.bas` | Standard Module | `signData` Public Type definition |
| `Module6.bas` | Standard Module | Alignment drawing: `StartWZTCDrawing`, `GroupAndLaunchPlacement` |
| `ModuleWZTCData.bas` | Standard Module | All public persistent state variables |
| `ModuleAlignmentPlacement.bas` | Standard Module | Path geometry engine + perpendicular line placement |
| `ModuleSignPlacement.bas` | Standard Module | Sign placement state machine |
| `ModuleWZTCElements.bas` | Standard Module | WZTC element drawing (level/color/hatch) |
| `ModuleWZTCCells.bas` | Standard Module | Cell library placement + count tracking |
| `ModTest.bas` | Standard Module | `DrawSignAtPerpLine` — sign drawing with post/face/text |
| `frmWorkzoneDesigner.frm` | UserForm | Main workzone configuration form |
| `UserForm1.frm` | UserForm | NYSDOT 619 standard sheet reference viewer |
| `AlignmentForm.frm` | UserForm | Alignment drawing (line/arc segments + Done) |
| `frmAlignmentPlacement.frm` | UserForm | Perpendicular line placement per WZTC item |
| `frmSignPlacement.frm` | UserForm | Sign drawing step (post + face + text) |
| `frmWZTCElements.frm` | UserForm | WZTC elements drawing step |
| `frmWZTCCells.frm` | UserForm | Cell library placement step |
| `PlacementButtonHandler.cls` | Class Module | `WithEvents` handler for dynamic placement buttons |
| `SignNumberBoxHandler.cls` | Class Module | `WithEvents` handler for dynamic sign number textboxes |
| `INSTALLATION_GUIDE.md` | Documentation | This file |

---

## UNUSED / LEGACY FILES

The following files exist in the project folder but are **not part of the active workflow**. Do not import them into the VBA project.

| File | Reason Not Needed |
|------|-------------------|
| `Module2.bas` | Intermediate sign placement approach — superseded by `ModTest.bas` |
| `Module5.bas` | Early hardcoded prototype with fixed project coordinates — not dynamic |
| `TEST_MINIMAL.bas` | Debugging utility (control existence checker) — not part of workflow |
| `WZTCCellLibrary.bas` | Old hardcoded cell macro (`BmrWZTCOther`) — superseded by `ModuleWZTCCells.bas` |
| `WZTCDrawingElements.bas` | Reference/example file for HATCH ICON pattern — logic now in `ModuleWZTCElements.bas` |
| `WZTCUserForm.frm` | Old form, not referenced anywhere in the workflow |
| `UserForm2.frm` | Old alignment drawing form — superseded by `AlignmentForm.frm` |
| `WorkzoneDesigner.bas` | Code reference paste-from file for initial setup only — form code now lives in `frmWorkzoneDesigner.frm` |

---

## KEY FEATURES

### State Persistence Across Sessions
All workzone configuration (dropdowns, sign table, WZTC order, spacing values) is saved to public variables in `ModuleWZTCData.bas` when **Submit & Draw** is clicked. Clicking **Return to Designer** on any form reopens `frmWorkzoneDesigner` with all previous selections restored. Use the **Clear All** button to explicitly start fresh.

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
`frmWZTCCells` tracks how many times each cell type has been placed and displays a running count after each placement.
