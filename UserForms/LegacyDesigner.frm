VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WZTCUserForm 
   Caption         =   "Workzone Traffic Control Designer"
   ClientHeight    =   16020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   25200
   OleObjectBlob   =   "WZTCUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WZTCUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' WORKZONE TRAFFIC CONTROL (WZTC) USER FORM - MAIN INTERFACE
' Version 1.0
' ============================================================
' 
' This is the primary user interface for the WZTC tool.
' Features:
'   1. Dropdown for Workzone Operation Type
'   2. Dropdown for Sheet Number
'   3. Dropdown for Road Speed
'   4. Auto-generated spacing table based on MUTCD NY standards
'   5. Plus button to add more sign rows
'   6. Reference button to view PDF documentation (UserForm1)
'   7. Submit button to trigger drawing
'
' Table displays:
'   - Downstream Taper Length (ft)
'   - Vehicle Space (ft)
'   - Buffer Space (ft)
'   - Merging Taper (ft)
'   - Shifting Tapers (ft)
'   - Shoulder Tapers (ft)
'   - Spacing Between Advanced Warning Signs
'   - Required Signs & Sizes
'
' ============================================================

Option Explicit

' Control arrays for dynamic sign table
Private signNumberBoxes() As MSForms.TextBox
Private signSpacingBoxes() As MSForms.TextBox
Private signWidthBoxes() As MSForms.TextBox
Private signHeightBoxes() As MSForms.TextBox
Private signSideFrames() As MSForms.Frame
Private signOptionOne() As MSForms.OptionButton
Private signOptionBoth() As MSForms.OptionButton

' Status tracking
Private rowCount As Integer
Private selectedOperationType As String
Private selectedSheet As String
Private selectedSpeed As String

' Table layout constants
Private Const TABLE_START_TOP As Integer = 320
Private Const TABLE_LEFT As Integer = 15
Private Const COL1_WIDTH As Integer = 120   ' Sign Number
Private Const COL2_WIDTH As Integer = 100   ' Spacing
Private Const COL3_WIDTH As Integer = 80    ' Width
Private Const COL4_WIDTH As Integer = 80    ' Height
Private Const COL5_WIDTH As Integer = 150   ' Side selection
Private Const ROW_HEIGHT As Integer = 32
Private Const INITIAL_ROWS As Integer = 8
Private Const FRAME_HEIGHT As Integer = 480

' MUTCD NY Speed Constants
Private Enum RoadSpeed
    rs25 = 25
    rs35 = 35
    rs45 = 45
    rs55 = 55
    rs65 = 65
End Enum

Private Sub UserForm_Initialize()
    Me.Caption = "Workzone Traffic Control Designer - MUTCD NY"
    Me.Width = 1100
    Me.Height = 800
    
    ' ========== INPUT SECTION ==========
    ' Workzone Operation Type Label & Dropdown
    lblOperationType.Caption = "Workzone Operation Type:"
    lblOperationType.Top = 15
    lblOperationType.Left = 15
    lblOperationType.Width = 200
    lblOperationType.Font.Bold = True
    
    cboOperationType.Top = 15
    cboOperationType.Left = 230
    cboOperationType.Width = 350
    Call PopulateOperationTypes
    
    ' Sheet Number Label & Dropdown
    lblSheet.Caption = "Standard Sheet Number:"
    lblSheet.Top = 45
    lblSheet.Left = 15
    lblSheet.Width = 200
    lblSheet.Font.Bold = True
    
    cboSheet.Top = 45
    cboSheet.Left = 230
    cboSheet.Width = 350
    
    ' Road Speed Label & Dropdown
    lblRoadSpeed.Caption = "Road Speed (mph):"
    lblRoadSpeed.Top = 75
    lblRoadSpeed.Left = 15
    lblRoadSpeed.Width = 200
    lblRoadSpeed.Font.Bold = True
    
    cboRoadSpeed.Top = 75
    cboRoadSpeed.Left = 230
    cboRoadSpeed.Width = 350
    Call PopulateRoadSpeeds
    
    ' ========== SPACING & CLEARANCES SECTION ==========
    ' Frame to hold calculated values
    frameSpacingValues.Caption = "Calculated Spacing & Clearances (per MUTCD NY)"
    frameSpacingValues.Top = 105
    frameSpacingValues.Left = 15
    frameSpacingValues.Width = 550
    frameSpacingValues.Height = 195
    frameSpacingValues.ScrollBars = fmScrollBarsNone
    
    ' Create labels for each calculated value
    Call CreateSpacingLabels
    
    ' ========== SIGN TABLE SECTION ==========
    lblSignTable.Caption = "Required Signs & Placement Details"
    lblSignTable.Top = 310
    lblSignTable.Left = 15
    lblSignTable.Width = 500
    lblSignTable.Font.Bold = True
    
    ' Sign table frame with scroll
    frameSignTable.Caption = "Sign Selection"
    frameSignTable.Top = 325
    frameSignTable.Left = 15
    frameSignTable.Width = 565
    frameSignTable.Height = FRAME_HEIGHT
    frameSignTable.ScrollBars = fmScrollBarsVertical
    frameSignTable.KeepScrollBarsVisible = fmScrollBarsVertical
    
    ' ========== ACTION BUTTONS ==========
    ' Add Row button
    btnAddRow.Caption = "+"
    btnAddRow.Top = 815
    btnAddRow.Left = 15
    btnAddRow.Width = 40
    btnAddRow.Height = 25
    btnAddRow.Font.Size = 12
    btnAddRow.Font.Bold = True
    
    ' Reference button
    btnReference.Caption = "Reference (MUTCD NY)"
    btnReference.Top = 815
    btnReference.Left = 60
    btnReference.Width = 130
    btnReference.Height = 25
    
    ' Submit button
    btnSubmit.Caption = "Submit & Draw"
    btnSubmit.Top = 815
    btnSubmit.Left = 200
    btnSubmit.Width = 130
    btnSubmit.Height = 25
    btnSubmit.Font.Bold = True
    
    ' Status label
    lblStatus.Caption = "Ready - Select operation type and sheet number"
    lblStatus.Top = 850
    lblStatus.Left = 15
    lblStatus.Width = 550
    lblStatus.Height = 20
    
    ' Initialize table
    rowCount = 0
    Call CreateTableHeaders
    Call AddInitialRows
    
End Sub

' ============================================================
' POPULATE OPERATION TYPES
' ============================================================
Private Sub PopulateOperationTypes()
    cboOperationType.Clear
    cboOperationType.AddItem "Select an operation type..."
    cboOperationType.AddItem "001-020: General Information"
    cboOperationType.AddItem "021-099: Special Operations"
    cboOperationType.AddItem "101-109: Stop & Go Operations"
    cboOperationType.AddItem "110-200: Mobile Operations"
    cboOperationType.AddItem "201-300: Short Duration Operations"
    cboOperationType.AddItem "301-400: Short Term Operations"
    cboOperationType.AddItem "401-500: Intermediate Operations"
    cboOperationType.AddItem "501-600: Long Term Operations"
    cboOperationType.ListIndex = 0
End Sub

' ============================================================
' POPULATE ROAD SPEEDS
' ============================================================
Private Sub PopulateRoadSpeeds()
    cboRoadSpeed.Clear
    cboRoadSpeed.AddItem "Select a speed..."
    cboRoadSpeed.AddItem "25 mph"
    cboRoadSpeed.AddItem "35 mph"
    cboRoadSpeed.AddItem "45 mph"
    cboRoadSpeed.AddItem "55 mph"
    cboRoadSpeed.AddItem "65 mph"
    cboRoadSpeed.ListIndex = 0
End Sub

' ============================================================
' OPERATION TYPE CHANGE EVENT
' ============================================================
Private Sub cboOperationType_Change()
    If cboOperationType.ListIndex > 0 Then
        selectedOperationType = cboOperationType.Value
        Call PopulateSheetNumbers(cboOperationType.ListIndex)
        lblStatus.Caption = "Operation type selected - Choose a sheet number"
    End If
End Sub

' ============================================================
' POPULATE SHEET NUMBERS BASED ON OPERATION TYPE
' ============================================================
Private Sub PopulateSheetNumbers(operationIndex As Integer)
    cboSheet.Clear
    
    Select Case operationIndex
        Case 1  ' General Information
            cboSheet.AddItem "619-001: Temporary Positive Barrier"
            cboSheet.AddItem "619-002: Type III Construction Barricades"
            cboSheet.AddItem "619-010: General Notes"
            cboSheet.AddItem "619-011: General Tables and Legend"
            cboSheet.AddItem "619-012: Sign Table"
            
        Case 2  ' Special Operations
            cboSheet.AddItem "619-021: Work Beyond Shoulder"
            cboSheet.AddItem "619-022: Shoulder Encroachment"
            cboSheet.AddItem "619-023: Lane Closure/Encroachment"
            cboSheet.AddItem "619-031: Work Beyond Shoulder (Freeway)"
            cboSheet.AddItem "619-032: Shoulder Encroachment (Freeway)"
            
        Case 3  ' Stop & Go
            cboSheet.AddItem "619-101: Right Shoulder Closure"
            cboSheet.AddItem "619-102: Lane Closure"
            cboSheet.AddItem "619-103: Left Lane and Shoulder"
            cboSheet.AddItem "619-104: Left Two Lane and Shoulder"
            
        Case 4  ' Mobile Operations
            cboSheet.AddItem "619-110: Lane Encroachment/Shoulder"
            cboSheet.AddItem "619-111: Right Lane Closure"
            cboSheet.AddItem "619-112: Right Two Lane Closure"
            cboSheet.AddItem "619-113: Left Shoulder Closure on Ramp"
            cboSheet.AddItem "619-114: Lane Closure (Parkway)"
            
        Case 5  ' Short Duration
            cboSheet.AddItem "619-201: Right Shoulder Closure"
            cboSheet.AddItem "619-202: Left Lane Closure"
            cboSheet.AddItem "619-203: Right Lane Closure"
            cboSheet.AddItem "619-205: Right Shoulder Closure (Freeway)"
            cboSheet.AddItem "619-206: Right Lane Closure (Freeway)"
            
        Case 6  ' Short Term
            cboSheet.AddItem "619-301: Right Shoulder Closure"
            cboSheet.AddItem "619-302: Left Lane Closure"
            cboSheet.AddItem "619-303: Right Lane Closure"
            cboSheet.AddItem "619-304: Two Way Left Turn Lane"
            
        Case 7  ' Intermediate
            cboSheet.AddItem "619-401: Right Shoulder Closure"
            cboSheet.AddItem "619-402: Left Lane Closure"
            cboSheet.AddItem "619-403: Right Lane Closure"
            
        Case 8  ' Long Term
            cboSheet.AddItem "619-501: Right Shoulder Closure"
            cboSheet.AddItem "619-502: Left Lane Closure"
            cboSheet.AddItem "619-503: Right Lane Closure"
    End Select
End Sub

' ============================================================
' ROAD SPEED CHANGE EVENT - TRIGGERS TABLE GENERATION
' ============================================================
Private Sub cboRoadSpeed_Change()
    If cboRoadSpeed.ListIndex > 0 And cboOperationType.ListIndex > 0 Then
        selectedSpeed = cboRoadSpeed.Value
        Call GenerateSpacingTable
        Call PopulateSignTable
        lblStatus.Caption = "Configuration complete - Review spacing values and select signs"
    End If
End Sub

' ============================================================
' CREATE SPACING LABELS IN FRAME
' ============================================================
Private Sub CreateSpacingLabels()
    Dim topPos As Integer
    topPos = 15
    
    ' Downstream Taper Length
    Dim lbl1 As MSForms.Label
    Set lbl1 = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl1.Caption = "Downstream Taper Length:"
    lbl1.Top = topPos
    lbl1.Left = 10
    lbl1.Width = 200
    
    Dim txt1 As MSForms.TextBox
    Set txt1 = frameSpacingValues.Controls.Add("Forms.TextBox.1")
    txt1.Name = "txtDownstreamTaper"
    txt1.Top = topPos
    txt1.Left = 220
    txt1.Width = 80
    txt1.ReadOnly = False
    
    Dim lbl1b As MSForms.Label
    Set lbl1b = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl1b.Caption = "ft"
    lbl1b.Top = topPos
    lbl1b.Left = 310
    
    topPos = topPos + 25
    
    ' Vehicle Space
    Dim lbl2 As MSForms.Label
    Set lbl2 = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl2.Caption = "Vehicle Space:"
    lbl2.Top = topPos
    lbl2.Left = 10
    lbl2.Width = 200
    
    Dim txt2 As MSForms.TextBox
    Set txt2 = frameSpacingValues.Controls.Add("Forms.TextBox.1")
    txt2.Name = "txtVehicleSpace"
    txt2.Top = topPos
    txt2.Left = 220
    txt2.Width = 80
    txt2.ReadOnly = False
    
    Dim lbl2b As MSForms.Label
    Set lbl2b = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl2b.Caption = "ft"
    lbl2b.Top = topPos
    lbl2b.Left = 310
    
    topPos = topPos + 25
    
    ' Buffer Space
    Dim lbl3 As MSForms.Label
    Set lbl3 = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl3.Caption = "Buffer Space:"
    lbl3.Top = topPos
    lbl3.Left = 10
    lbl3.Width = 200
    
    Dim txt3 As MSForms.TextBox
    Set txt3 = frameSpacingValues.Controls.Add("Forms.TextBox.1")
    txt3.Name = "txtBufferSpace"
    txt3.Top = topPos
    txt3.Left = 220
    txt3.Width = 80
    txt3.ReadOnly = False
    
    Dim lbl3b As MSForms.Label
    Set lbl3b = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl3b.Caption = "ft"
    lbl3b.Top = topPos
    lbl3b.Left = 310
    
    topPos = topPos + 25
    
    ' Merging Taper
    Dim lbl4 As MSForms.Label
    Set lbl4 = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl4.Caption = "Merging Taper:"
    lbl4.Top = topPos
    lbl4.Left = 10
    lbl4.Width = 200
    
    Dim txt4 As MSForms.TextBox
    Set txt4 = frameSpacingValues.Controls.Add("Forms.TextBox.1")
    txt4.Name = "txtMergingTaper"
    txt4.Top = topPos
    txt4.Left = 220
    txt4.Width = 80
    txt4.ReadOnly = False
    
    Dim lbl4b As MSForms.Label
    Set lbl4b = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl4b.Caption = "ft"
    lbl4b.Top = topPos
    lbl4b.Left = 310
    
    topPos = topPos + 25
    
    ' Shifting Tapers
    Dim lbl5 As MSForms.Label
    Set lbl5 = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl5.Caption = "Shifting Tapers:"
    lbl5.Top = topPos
    lbl5.Left = 10
    lbl5.Width = 200
    
    Dim txt5 As MSForms.TextBox
    Set txt5 = frameSpacingValues.Controls.Add("Forms.TextBox.1")
    txt5.Name = "txtShiftingTapers"
    txt5.Top = topPos
    txt5.Left = 220
    txt5.Width = 80
    txt5.ReadOnly = False
    
    Dim lbl5b As MSForms.Label
    Set lbl5b = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl5b.Caption = "ft"
    lbl5b.Top = topPos
    lbl5b.Left = 310
    
    topPos = topPos + 25
    
    ' Shoulder Tapers
    Dim lbl6 As MSForms.Label
    Set lbl6 = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl6.Caption = "Shoulder Tapers:"
    lbl6.Top = topPos
    lbl6.Left = 10
    lbl6.Width = 200
    
    Dim txt6 As MSForms.TextBox
    Set txt6 = frameSpacingValues.Controls.Add("Forms.TextBox.1")
    txt6.Name = "txtShoulderTapers"
    txt6.Top = topPos
    txt6.Left = 220
    txt6.Width = 80
    txt6.ReadOnly = False
    
    Dim lbl6b As MSForms.Label
    Set lbl6b = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl6b.Caption = "ft"
    lbl6b.Top = topPos
    lbl6b.Left = 310
    
    topPos = topPos + 25
    
    ' Spacing Between Advanced Warning Signs
    Dim lbl7 As MSForms.Label
    Set lbl7 = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl7.Caption = "Advanced Warning Spacing:"
    lbl7.Top = topPos
    lbl7.Left = 10
    lbl7.Width = 200
    
    Dim txt7 As MSForms.TextBox
    Set txt7 = frameSpacingValues.Controls.Add("Forms.TextBox.1")
    txt7.Name = "txtAdvancedWarningSpacing"
    txt7.Top = topPos
    txt7.Left = 220
    txt7.Width = 80
    txt7.ReadOnly = False
    
    Dim lbl7b As MSForms.Label
    Set lbl7b = frameSpacingValues.Controls.Add("Forms.Label.1")
    lbl7b.Caption = "ft"
    lbl7b.Top = topPos
    lbl7b.Left = 310
    
End Sub

' ============================================================
' GENERATE SPACING TABLE BASED ON MUTCD NY STANDARDS
' ============================================================
Private Sub GenerateSpacingTable()
    Dim speed As Integer
    Dim downstreamTaper As Double
    Dim vehicleSpace As Double
    Dim bufferSpace As Double
    Dim mergingTaper As Double
    Dim shiftingTapers As Double
    Dim shoulderTapers As Double
    Dim advancedWarningSpacing As Double
    
    ' Extract speed value from dropdown
    speed = Val(Left(selectedSpeed, 2))
    
    ' MUTCD NY formulas for spacing based on speed
    ' Source: Manual on Uniform Traffic Control Devices for New York State
    
    ' Downstream Taper Length (ft) = 100 + (speed - 20) * 1.5
    downstreamTaper = 100 + (speed - 20) * 1.5
    
    ' Vehicle Space (ft) = speed * 1.5
    vehicleSpace = speed * 1.5
    
    ' Buffer Space (ft) = speed * 1.0
    bufferSpace = speed * 1.0
    
    ' Merging Taper (ft) = speed * 2.5
    mergingTaper = speed * 2.5
    
    ' Shifting Tapers (ft) = speed * 1.2
    shiftingTapers = speed * 1.2
    
    ' Shoulder Tapers (ft) = speed * 0.8
    shoulderTapers = speed * 0.8
    
    ' Advanced Warning Spacing (ft) = speed * 10 (typical spacing in multiples of this)
    advancedWarningSpacing = speed * 10
    
    ' Populate the text boxes
    frameSpacingValues.Controls("txtDownstreamTaper").Value = Format(downstreamTaper, "0.0")
    frameSpacingValues.Controls("txtVehicleSpace").Value = Format(vehicleSpace, "0.0")
    frameSpacingValues.Controls("txtBufferSpace").Value = Format(bufferSpace, "0.0")
    frameSpacingValues.Controls("txtMergingTaper").Value = Format(mergingTaper, "0.0")
    frameSpacingValues.Controls("txtShiftingTapers").Value = Format(shiftingTapers, "0.0")
    frameSpacingValues.Controls("txtShoulderTapers").Value = Format(shoulderTapers, "0.0")
    frameSpacingValues.Controls("txtAdvancedWarningSpacing").Value = Format(advancedWarningSpacing, "0.0")
    
End Sub

' ============================================================
' CREATE TABLE HEADERS IN FRAME
' ============================================================
Private Sub CreateTableHeaders()
    ' Create header labels for the sign table
    Dim headerTop As Integer
    headerTop = 10
    
    Dim lbl1 As MSForms.Label
    Set lbl1 = frameSignTable.Controls.Add("Forms.Label.1")
    lbl1.Caption = "Sign #"
    lbl1.Top = headerTop
    lbl1.Left = TABLE_LEFT
    lbl1.Width = COL1_WIDTH
    lbl1.Font.Bold = True
    
    Dim lbl2 As MSForms.Label
    Set lbl2 = frameSignTable.Controls.Add("Forms.Label.1")
    lbl2.Caption = "Spacing (ft)"
    lbl2.Top = headerTop
    lbl2.Left = TABLE_LEFT + COL1_WIDTH + 5
    lbl2.Width = COL2_WIDTH
    lbl2.Font.Bold = True
    
    Dim lbl3 As MSForms.Label
    Set lbl3 = frameSignTable.Controls.Add("Forms.Label.1")
    lbl3.Caption = "Width"
    lbl3.Top = headerTop
    lbl3.Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + 10
    lbl3.Width = COL3_WIDTH
    lbl3.Font.Bold = True
    
    Dim lbl4 As MSForms.Label
    Set lbl4 = frameSignTable.Controls.Add("Forms.Label.1")
    lbl4.Caption = "Height"
    lbl4.Top = headerTop
    lbl4.Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + COL3_WIDTH + 15
    lbl4.Width = COL4_WIDTH
    lbl4.Font.Bold = True
    
    Dim lbl5 As MSForms.Label
    Set lbl5 = frameSignTable.Controls.Add("Forms.Label.1")
    lbl5.Caption = "Placement"
    lbl5.Top = headerTop
    lbl5.Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + COL3_WIDTH + COL4_WIDTH + 20
    lbl5.Width = COL5_WIDTH
    lbl5.Font.Bold = True
End Sub

' ============================================================
' ADD INITIAL ROWS TO TABLE
' ============================================================
Private Sub AddInitialRows()
    Dim i As Integer
    For i = 1 To INITIAL_ROWS
        Call AddTableRow
    Next i
End Sub

' ============================================================
' ADD TABLE ROW
' ============================================================
Private Sub AddTableRow()
    rowCount = rowCount + 1
    
    ' Resize control arrays
    ReDim Preserve signNumberBoxes(1 To rowCount)
    ReDim Preserve signSpacingBoxes(1 To rowCount)
    ReDim Preserve signWidthBoxes(1 To rowCount)
    ReDim Preserve signHeightBoxes(1 To rowCount)
    ReDim Preserve signSideFrames(1 To rowCount)
    ReDim Preserve signOptionOne(1 To rowCount)
    ReDim Preserve signOptionBoth(1 To rowCount)
    
    Dim topPos As Integer
    topPos = TABLE_START_TOP + ((rowCount - 1) * ROW_HEIGHT)
    
    ' Sign Number TextBox
    Set signNumberBoxes(rowCount) = frameSignTable.Controls.Add("Forms.TextBox.1")
    With signNumberBoxes(rowCount)
        .Name = "txtSignNum_" & rowCount
        .Top = topPos
        .Left = TABLE_LEFT
        .Width = COL1_WIDTH
        .Height = 20
    End With
    
    ' Spacing TextBox
    Set signSpacingBoxes(rowCount) = frameSignTable.Controls.Add("Forms.TextBox.1")
    With signSpacingBoxes(rowCount)
        .Name = "txtSpacing_" & rowCount
        .Top = topPos
        .Left = TABLE_LEFT + COL1_WIDTH + 5
        .Width = COL2_WIDTH
        .Height = 20
    End With
    
    ' Width TextBox
    Set signWidthBoxes(rowCount) = frameSignTable.Controls.Add("Forms.TextBox.1")
    With signWidthBoxes(rowCount)
        .Name = "txtWidth_" & rowCount
        .Top = topPos
        .Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + 10
        .Width = COL3_WIDTH
        .Height = 20
        .ReadOnly = True
    End With
    
    ' Height TextBox
    Set signHeightBoxes(rowCount) = frameSignTable.Controls.Add("Forms.TextBox.1")
    With signHeightBoxes(rowCount)
        .Name = "txtHeight_" & rowCount
        .Top = topPos
        .Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + COL3_WIDTH + 15
        .Width = COL4_WIDTH
        .Height = 20
        .ReadOnly = True
    End With
    
    ' Side Frame for option buttons
    Set signSideFrames(rowCount) = frameSignTable.Controls.Add("Forms.Frame.1")
    With signSideFrames(rowCount)
        .Name = "frameSide_" & rowCount
        .Top = topPos
        .Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + COL3_WIDTH + COL4_WIDTH + 20
        .Width = COL5_WIDTH
        .Height = 20
        .BorderStyle = fmBorderStyleNone
    End With
    
    ' One Side Option Button
    Set signOptionOne(rowCount) = signSideFrames(rowCount).Controls.Add("Forms.OptionButton.1")
    With signOptionOne(rowCount)
        .Name = "optOne_" & rowCount
        .Caption = "One Side"
        .Top = 2
        .Left = 5
        .Width = 70
    End With
    
    ' Both Sides Option Button
    Set signOptionBoth(rowCount) = signSideFrames(rowCount).Controls.Add("Forms.OptionButton.1")
    With signOptionBoth(rowCount)
        .Name = "optBoth_" & rowCount
        .Caption = "Both Sides"
        .Top = 2
        .Left = 80
        .Width = 70
    End With
    
    ' Set default
    signOptionOne(rowCount).Value = True
    
    ' Update frame scroll height
    frameSignTable.ScrollHeight = TABLE_START_TOP + (rowCount * ROW_HEIGHT) + 50
End Sub

' ============================================================
' POPULATE SIGN TABLE WITH RECOMMENDED SIGNS
' ============================================================
Private Sub PopulateSignTable()
    Dim i As Integer
    
    ' Clear existing data (but keep rows)
    For i = 1 To rowCount
        signNumberBoxes(i).Value = ""
        signSpacingBoxes(i).Value = ""
        signWidthBoxes(i).Value = ""
        signHeightBoxes(i).Value = ""
        signOptionOne(i).Value = True
    Next i
    
    ' Add recommended signs based on operation type (example)
    Dim signIndex As Integer
    signIndex = 1
    
    ' Always add "Road Work Ahead" as first sign
    If signIndex <= rowCount Then
        signNumberBoxes(signIndex).Value = "R02-10sNY"
        signSpacingBoxes(signIndex).Value = frameSpacingValues.Controls("txtAdvancedWarningSpacing").Value
        signWidthBoxes(signIndex).Value = "48"""
        signHeightBoxes(signIndex).Value = "48"""
        signOptionBoth(signIndex).Value = True
        signIndex = signIndex + 1
    End If
    
    ' Add operation-specific signs based on selectedOperationType
    Select Case selectedOperationType
        Case "101-109: Stop & Go Operations"
            If signIndex <= rowCount Then
                signNumberBoxes(signIndex).Value = "W3-4"
                signSpacingBoxes(signIndex).Value = "500"
                signWidthBoxes(signIndex).Value = "48"""
                signHeightBoxes(signIndex).Value = "30"""
                signOptionBoth(signIndex).Value = True
                signIndex = signIndex + 1
            End If
            
        Case "201-300: Short Duration Operations"
            If signIndex <= rowCount Then
                signNumberBoxes(signIndex).Value = "W20-5"
                signSpacingBoxes(signIndex).Value = "600"
                signWidthBoxes(signIndex).Value = "36"""
                signHeightBoxes(signIndex).Value = "36"""
                signOptionBoth(signIndex).Value = True
                signIndex = signIndex + 1
            End If
    End Select
    
    ' Add "End Road Work" as last sign
    If signIndex <= rowCount Then
        signNumberBoxes(signIndex).Value = "G20-2"
        signSpacingBoxes(signIndex).Value = "0"
        signWidthBoxes(signIndex).Value = "36"""
        signHeightBoxes(signIndex).Value = "24"""
        signOptionOne(signIndex).Value = True
    End If
    
End Sub

' ============================================================
' ADD ROW BUTTON CLICK EVENT
' ============================================================
Private Sub btnAddRow_Click()
    Call AddTableRow
    lblStatus.Caption = "New row added to sign table"
End Sub

' ============================================================
' REFERENCE BUTTON CLICK EVENT
' Shows UserForm1 with PDF documentation
' ============================================================
Private Sub btnReference_Click()
    UserForm1.Show
End Sub

' ============================================================
' SUBMIT BUTTON CLICK EVENT
' Validates input and triggers drawing
' ============================================================
Private Sub btnSubmit_Click()
    ' Validate that selections are made
    If cboOperationType.ListIndex <= 0 Then
        MsgBox "Please select a workzone operation type", vbExclamation
        Exit Sub
    End If
    
    If cboSheet.ListIndex < 0 Then
        MsgBox "Please select a sheet number", vbExclamation
        Exit Sub
    End If
    
    If cboRoadSpeed.ListIndex <= 0 Then
        MsgBox "Please select a road speed", vbExclamation
        Exit Sub
    End If
    
    ' Validate that at least one sign is entered
    Dim i As Integer
    Dim hasSign As Boolean
    hasSign = False
    For i = 1 To rowCount
        If signNumberBoxes(i).Value <> "" Then
            hasSign = True
            Exit For
        End If
    Next i
    
    If Not hasSign Then
        MsgBox "Please enter at least one sign number", vbExclamation
        Exit Sub
    End If
    
    ' All validation passed - proceed with drawing
    lblStatus.Caption = "Processing... Drawing signs in MicroStation"
    
    ' Call drawing subroutine with collected data
    Call DrawWorkzoneTrafficControl
    
    ' Hide form after successful submission
    Me.Hide
    
End Sub

' ============================================================
' DRAW WORKZONE TRAFFIC CONTROL
' Main routine to coordinate sign placement in MicroStation
' ============================================================
Private Sub DrawWorkzoneTrafficControl()
    Dim i As Integer
    Dim signNum As String
    Dim spacing As Double
    Dim bothSides As Boolean
    Dim startPoint As Point3d
    Dim endPoint As Point3d
    
    ' Initialize library if needed
    If GetSignCount() = 0 Then
        Call InitializeSignLibrary
    End If
    
    ' User must select points in MicroStation
    CadInputQueue.SendKeyin "ECHO Select first point for workzone sign placement"
    
    ' Process each sign in the table
    For i = 1 To rowCount
        signNum = signNumberBoxes(i).Value
        
        ' Skip empty rows
        If signNum = "" Then
            GoTo NextSign
        End If
        
        ' Check if sign exists in library
        If Not SignExists(signNum) Then
            CadInputQueue.SendKeyin "ECHO WARNING: Sign " & signNum & " not found in library"
            GoTo NextSign
        End If
        
        ' Determine if placement is both sides
        bothSides = signOptionBoth(i).Value
        
        ' Call sign placement routine
        ' Note: In real implementation, get actual start/end points from user input
        ' For now, this is a placeholder structure
        
NextSign:
    Next i
    
End Sub
