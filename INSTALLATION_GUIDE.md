# WORKZONE TRAFFIC CONTROL DESIGNER - INSTALLATION GUIDE

## Overview
This guide walks you through installing the Workzone Traffic Control (WZTC) Designer tool in MicroStation VBA without needing external class modules.

---

## STEP 1: Open VBA Editor
1. Open **MicroStation**
2. Press **Alt + F11** to open the VBA Editor


## STEP 2: Create the UserForm

### 2A. Insert a New UserForm
1. In Project Explorer (left sidebar), right-click your project
2. Select **Insert → UserForm**
3. A new UserForm will be created

### 2B. Configure Form Properties
In the **Properties window** (bottom-left), set these values:

| Property | Value |
|----------|-------|
| **(Name)** | `frmWorkzoneDesigner` |
| **Width** | `600` |
| **Height** | `730` |
| **Caption** | `Workzone Design Tool` |
| **StartUpPosition** | `1 - CenterOwner` |


## STEP 3: Add Controls to the Form

### 3A. Open Toolbox

### 3B. Add Each Control (drag from Toolbox onto the form)

### Input Section (Top of Form)
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|--------------|--------|---------|-----|------|-------|--------|-------|
| **Label** | `lblCategory` | `Workzone Category:` | 10 | 20 | 120 | 18 | Bold font recommended |
| **ComboBox** | `cboCategory` | (empty) | 10 | 150 | 250 | 20 | Style: 2-Dropdown List |
| **Label** | `lblSheet` | `Standard Sheet:` | 40 | 20 | 120 | 18 | Bold font |
| **ComboBox** | `cboSheet` | (empty) | 40 | 150 | 250 | 20 | Style: 2-Dropdown List |
| **Label** | `lblRoadSpeed` | `Road Speed (mph):` | 70 | 20 | 120 | 18 | Bold font |
| **ComboBox** | `cboRoadSpeed` | (empty) | 70 | 150 | 250 | 20 | Style: 2-Dropdown List |
| **Label** | `lblRoadType` | `Road Type:` | 100 | 20 | 120 | 18 | Bold font |
| **ComboBox** | `cboRoadType` | (empty) | 100 | 150 | 250 | 20 | Style: 2-Dropdown List |
| **Label** | `lblLaneWidth` | `Lane Width (ft):` | 130 | 20 | 120 | 18 | Bold font |
| **ComboBox** | `cboLaneWidth` | (empty) | 130 | 150 | 250 | 20 | Style: 2-Dropdown List |
| **Label** | `lblShoulderWidth` | `Shoulder Width (ft):` | 160 | 20 | 120 | 18 | Bold font |
| **ComboBox** | `cboShoulderWidth` | (empty) | 160 | 150 | 250 | 20 | Style: 2-Dropdown List |

### Spacing Section
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|--------------|--------|---------|-----|------|-------|--------|-------|
| **Frame** | `frameSpacingValues` | `Spacing & Clearances` | 195 | 10 | 560 | 220 | Dynamically creates labels and textboxes |

### Sign Table Section
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|--------------|--------|---------|-----|------|-------|--------|-------|
| **Label** | `lblSignTable` | `Required Signs & Details` | 425 | 10 | 500 | 18 | Bold font |
| **Frame** | `frameSignTable` | `Sign Selection` | 445 | 10 | 560 | 190 | ScrollBars: Vertical |

### Action Buttons (Bottom of Form)
| Control Type | (Name) | Caption | Top | Left | Width | Height | Notes |
|--------------|--------|---------|-----|------|-------|--------|-------|
| **CommandButton** | `btnAddRow` | `+` | 645 | 20 | 40 | 25 | Font: Bold, Size 12 | Add row to sign table |
| **CommandButton** | `btnRemoveRow` | `-` | 645 | 65 | 40 | 25 | Font: Bold, Size 12 | Remove last row from table |
| **CommandButton** | `btnReference` | `Reference (MUTCD)` | 645 | 220 | 130 | 25 | — | View reference data |
| **CommandButton** | `btnSubmit` | `Submit & Draw` | 645 | 440 | 130 | 25 | Font: Bold | Process and draw configuration |
| **Label** | `lblStatus` | `Ready - Select options` | 680 | 20 | 550 | 20 | Status messages | Displays current state |

**CRITICAL:** Control names (in parentheses) must match EXACTLY - they are case-sensitive!

### 3C. Spacing and Clearances Frame

**No need to manually add textboxes inside frameSpacingValues!** The code automatically creates all labels and textboxes dynamically:

#### Auto-Generated Read-Only Textboxes (calculated from speed):
- Downstream Taper (ft)
- Vehicle Space (ft)
- Buffer Space (ft)
- Merging Taper (ft)
- Shifting Tapers (ft)
- Shoulder Tapers (ft)
- Advanced Warning Spacing (ft)

#### Auto-Generated User Input Textboxes:
- **# of Skip Lines** (editable - user enters number of skip lines)
- **# of Channelizing Devices** (editable - user enters number of devices)

**All controls are created in the UserForm_Initialize() event, so you only need to create the empty frame!**


## STEP 4: Copy & Paste the Code

### 4A. Open Code File
1. In your workspace, locate the file: **`WorkzoneDesigner.bas`**
2. Open it in Notepad or VS Code

### 4B. Copy All Code
1. Select all code: **Ctrl + A**
2. Copy: **Ctrl + C**

### 4C. Paste into UserForm
1. In VBA Editor, **double-click** your new UserForm (frmWorkzoneDesigner)
2. This opens the code window
3. **Delete the default code** (the empty event handlers)
4. Paste the code: **Ctrl + V**
5. Save: **Ctrl + S**


## STEP 5: Test the Installation

### Option A: Run with F5
1. Make sure the UserForm is selected
2. Press **F5** (or click Run button)
3. The form should appear

### Option B: Run from Immediate Window
1. Press **Ctrl + G** to open Immediate Window
2. Type: `frmWorkzoneDesigner.Show`
3. Press Enter

### Expected Behavior
- Dropdowns should populate with categories and speeds
- Road Type dropdown should show "Freeway" and "Non-Freeway" options
- Lane Width dropdown should show width options (10-15 ft)
- **Shoulder Width dropdown should show 7 options: ≤4 ft, 5-7 ft, 8 ft, 9 ft, 10 ft, 11 ft, 12 ft**
- Spacing & Clearances frame should display 9 textboxes:
  - 7 read-only calculated values
  - 2 editable user input fields (Skip Lines, Channelizing Devices)
- Sign table section creates dynamic rows with:
  - Sign Number field
  - Spacing field
  - Size field
  - Side selection dropdown ("One Side" or "Both Sides")
- **Plus (+) button adds new rows**
- **Minus (-) button removes the last row**


## TROUBLESHOOTING

### "Invalid Character" Error
- Make sure you copied the entire code without any special characters
- Re-copy from the WorkzoneDesigner.bas file

### "Control not found" Error
- Verify all control names match exactly (case-sensitive)
- Check that `cboCategory`, `cboSheet`, `cboRoadSpeed`, `cboRoadType` all exist
- Ensure `frameSpacingValues` and `frameSignTable` are created
- Verify textboxes are **inside** frameSpacingValues (not on the main form)

### Dropdowns Are Empty
- The code automatically populates dropdowns on form initialization
- If empty, check that the `UserForm_Initialize` event is firing
- Try pressing F5 to run the form again

### Some Controls Missing
- Double-check the control tables above
- Ensure all required controls are added to the form, including:
  - `lblRoadType` and `cboRoadType` for road type
  - `lblLaneWidth` and `cboLaneWidth` for lane width
  - **`lblShoulderWidth` and `cboShoulderWidth` for shoulder width (new)**
  - `btnRemoveRow` for removing rows from sign table (new)

### Compile Error: "Sub or Function Not Defined"
- This was fixed in the latest version
- Make sure you're using the updated WorkzoneDesigner.bas code
- All undefined references to `signWidthBoxes`, `signHeightBoxes`, etc. have been removed


## FILE REFERENCE

Your workspace contains:

| File | Purpose |
|------|---------|
| **WorkzoneDesigner.bas** | All the VBA code to paste into the UserForm |
| **INSTALLATION_GUIDE.md** | This file - step-by-step instructions |
| **WZTCUserForm.frm** | Example form (reference only) |


## KEY FEATURES

### Speed Range
- Supports road speeds from **25 mph to 90 mph** in 5 mph increments
- Automatically calculates MUTCD NY spacing values based on selected speed

### Road Type Selection
- **Freeway** vs **Non-Freeway** dropdown
- Allows customization based on roadway classification

### Lane Width Selection
- Dropdown with common lane widths from **10 ft to 15 ft**
- Helps determine appropriate spacing requirements

### Shoulder Width Selection
- **NEW:** Dropdown with 7 shoulder width options: **≤ 4 ft, 5-7 ft, 8 ft, 9 ft, 10 ft, 11 ft, 12 ft**
- Allows users to specify shoulder geometry for workzone design

### Spacing & Clearances Table
- Displays 7 automatically calculated values based on MUTCD NY formulas
- **NEW:** Two user input fields for:
  - **# of Skip Lines** - user-editable count of skip lines
  - **# of Channelizing Devices** - user-editable count of devices (cones, drums, etc.)

### Sign Placement
- Dynamic table for sign configuration
- **Dropdown** for selecting "One Side" or "Both Sides" for each sign
- **Add rows** using the "**+**" button
- **Remove rows** using the "**-**" button (new)
- Sign number, spacing, and size inputs for each row

### MUTCD Compliance
- Pre-populated with MUTCD NY standard sheets
- Automatic spacing calculations per MUTCD formulas
- Reference output shows complete configuration

## SUPPORT NOTES



**Installation Time:** ~10-15 minutes  
**Difficulty Level:** Beginner-Intermediate  
**MicroStation Version:** Compatible with VBA-enabled versions
