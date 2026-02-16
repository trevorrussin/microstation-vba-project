Option Explicit

' Control arrays for dynamic sign table
Private signNumberBoxes() As MSForms.TextBox
Private signSpacingBoxes() As MSForms.TextBox
Private signSizeBoxes() As MSForms.TextBox
Private signSideComboBoxes() As MSForms.ComboBox
Private rowCount As Integer

' Status tracking
Private selectedOperationType As String
Private selectedSheet As String
Private selectedSpeed As String
Private selectedRoadType As String

' Table layout constants
Private Const TABLE_START_TOP As Integer = 120
Private Const TABLE_LEFT As Integer = 20
Private Const COL1_WIDTH As Integer = 100   ' Sign Number
Private Const COL2_WIDTH As Integer = 100   ' Spacing
Private Const COL3_WIDTH As Integer = 80    ' Size
Private Const COL4_WIDTH As Integer = 120   ' Side selection
Private Const ROW_HEIGHT As Integer = 30
Private Const INITIAL_ROWS As Integer = 10

' WZTC Order table
Private wztcOrderTexts() As String
Private wztcOrderCount As Integer

' Handlers for sign number textbox Exit (auto-fill spacing/size from library)
Private signNumberHandlers As Collection

' ============================================================
' INITIALIZE FORM
' ============================================================
Private Sub UserForm_Initialize()
    On Error GoTo InitError

    Debug.Print "Starting UserForm_Initialize..."

    Set signNumberHandlers = New Collection

    Me.Caption = "Workzone Traffic Control Designer - MUTCD NY"

    ' Widen form to accommodate WZTC Order panel
    If Me.Width < 1820 Then Me.Width = 1820

    ' ========== INPUT SECTION ==========
    ' Workzone Category Label & Dropdown
    If ControlExists("lblCategory") Then
        lblCategory.Caption = "Workzone Category:"
        lblCategory.Top = 10
        lblCategory.Left = 20
        lblCategory.Width = 120
        lblCategory.Font.Bold = True
    End If
    
    If ControlExists("cboCategory") Then
        cboCategory.Top = 10
        cboCategory.Left = 150
        cboCategory.Width = 250
        Call PopulateCategories
    End If
    
    ' Sheet Number Label & Dropdown
    If ControlExists("lblSheet") Then
        lblSheet.Caption = "Standard Sheet Number:"
        lblSheet.Top = 40
        lblSheet.Left = 20
        lblSheet.Width = 120
        lblSheet.Font.Bold = True
    End If
    
    If ControlExists("cboSheet") Then
        cboSheet.Top = 40
        cboSheet.Left = 150
        cboSheet.Width = 250
    End If
    
    ' Road Speed Label & Dropdown
    If ControlExists("lblRoadSpeed") Then
        lblRoadSpeed.Caption = "Road Speed (mph):"
        lblRoadSpeed.Top = 70
        lblRoadSpeed.Left = 20
        lblRoadSpeed.Width = 120
        lblRoadSpeed.Font.Bold = True
    End If

    If ControlExists("cboRoadSpeed") Then
        cboRoadSpeed.Top = 70
        cboRoadSpeed.Left = 150
        cboRoadSpeed.Width = 250
        Call PopulateRoadSpeeds
    End If

    ' Road Type Label & Dropdown
    If ControlExists("lblRoadType") Then
        lblRoadType.Caption = "Road Type:"
        lblRoadType.Top = 100
        lblRoadType.Left = 20
        lblRoadType.Width = 120
        lblRoadType.Font.Bold = True
    End If

    If ControlExists("cboRoadType") Then
        cboRoadType.Top = 100
        cboRoadType.Left = 150
        cboRoadType.Width = 250
        Call PopulateRoadType
    End If

    ' Lane Width Label & Dropdown
    If ControlExists("lblLaneWidth") Then
        lblLaneWidth.Caption = "Lane Width (ft):"
        lblLaneWidth.Top = 130
        lblLaneWidth.Left = 20
        lblLaneWidth.Width = 120
        lblLaneWidth.Font.Bold = True
    End If

    If ControlExists("cboLaneWidth") Then
        cboLaneWidth.Top = 130
        cboLaneWidth.Left = 150
        cboLaneWidth.Width = 250
        Call PopulateLaneWidth
    End If
    
    ' Shoulder Width Label & Dropdown
    If ControlExists("lblShoulderWidth") Then
        lblShoulderWidth.Caption = "Shoulder Width (ft):"
        lblShoulderWidth.Top = 160
        lblShoulderWidth.Left = 20
        lblShoulderWidth.Width = 120
        lblShoulderWidth.Font.Bold = True
    End If

    If ControlExists("cboShoulderWidth") Then
        cboShoulderWidth.Top = 160
        cboShoulderWidth.Left = 150
        cboShoulderWidth.Width = 250
        Call PopulateShoulderWidth
    End If
    
    ' ========== SPACING & CLEARANCES SECTION ==========
    If ControlExists("frameSpacingValues") Then
        frameSpacingValues.Caption = "Calculated Spacing & Clearances (per MUTCD NY)"
        frameSpacingValues.Top = 195
        frameSpacingValues.Left = 10
        frameSpacingValues.Width = 280
        frameSpacingValues.Height = 260

        Call CreateSpacingLabels
    End If
    
    ' ========== SIGN TABLE SECTION ==========
    If ControlExists("lblSignTable") Then
        lblSignTable.Caption = "Required Signs & Placement Details"
        lblSignTable.Top = 205
        lblSignTable.Left = 300
        lblSignTable.Width = 270
        lblSignTable.Font.Bold = True
    End If

    If ControlExists("frameSignTable") Then
        frameSignTable.Caption = "Sign Selection"
        frameSignTable.Top = 205
        frameSignTable.Left = 300
        frameSignTable.Width = 540
        frameSignTable.Height = 260
        frameSignTable.ScrollBars = fmScrollBarsVertical
        frameSignTable.KeepScrollBarsVisible = fmScrollBarsVertical
    End If

    ' ========== WZTC ORDER SECTION ==========
    ' Controls to add manually in IDE (directly on form, not inside a frame):
    '   frameWZTCOrder  - Frame            "WZTC Order"
    '   lstWZTCOrder    - ListBox
    '   btnOrderUp      - CommandButton
    '   btnOrderDown    - CommandButton
    '   btnOrderDelete  - CommandButton
    '   btnRefreshOrder - CommandButton
    If ControlExists("frameWZTCOrder") Then
        frameWZTCOrder.Caption = "WZTC Order"
        frameWZTCOrder.Top = 195
        frameWZTCOrder.Left = 855
        frameWZTCOrder.Width = 305
        frameWZTCOrder.Height = 315
    End If

    If ControlExists("lstWZTCOrder") Then
        lstWZTCOrder.Top = 215
        lstWZTCOrder.Left = 865
        lstWZTCOrder.Width = 220
        lstWZTCOrder.Height = 255
    End If

    If ControlExists("btnOrderUp") Then
        btnOrderUp.Caption = "Up"
        btnOrderUp.Top = 225
        btnOrderUp.Left = 1290
        btnOrderUp.Width = 60
        btnOrderUp.Height = 22
    End If

    If ControlExists("btnOrderDown") Then
        btnOrderDown.Caption = "Down"
        btnOrderDown.Top = 253
        btnOrderDown.Left = 1290
        btnOrderDown.Width = 60
        btnOrderDown.Height = 22
    End If

    If ControlExists("btnOrderDelete") Then
        btnOrderDelete.Caption = "X Del"
        btnOrderDelete.Top = 285
        btnOrderDelete.Left = 1290
        btnOrderDelete.Width = 60
        btnOrderDelete.Height = 22
    End If

    If ControlExists("btnRefreshOrder") Then
        btnRefreshOrder.Caption = "Refresh Order"
        btnRefreshOrder.Top = 475
        btnRefreshOrder.Left = 1290
        btnRefreshOrder.Width = 120
        btnRefreshOrder.Height = 22
        btnRefreshOrder.Font.Size = 8
    End If

    ' ========== ACTION BUTTONS ==========
    ' Add Row button (moved below tables)
    If ControlExists("btnAddRow") Then
        btnAddRow.Caption = "Add Row +"
        btnAddRow.Top = 470
        btnAddRow.Left = 300
        btnAddRow.Width = 150
        btnAddRow.Height = 25
        btnAddRow.Font.Size = 8
        btnAddRow.Font.Bold = True
    End If

    ' Remove Row button (moved below tables)
    If ControlExists("btnRemoveRow") Then
        btnRemoveRow.Caption = "Remove Row --"
        btnRemoveRow.Top = 470
        btnRemoveRow.Left = 545
        btnRemoveRow.Width = 150
        btnRemoveRow.Height = 25
        btnRemoveRow.Font.Size = 8
        btnRemoveRow.Font.Bold = True
    End If

    If ControlExists("btnReference") Then
        btnReference.Caption = "Reference (MUTCD)"
        btnReference.Top = 50
        btnReference.Left = 580
        btnReference.Width = 130
        btnReference.Height = 25
    End If

    If ControlExists("btnSubmit") Then
        btnSubmit.Caption = "Submit & Draw"
        btnSubmit.Top = 85
        btnSubmit.Left = 580
        btnSubmit.Width = 130
        btnSubmit.Height = 25
        btnSubmit.Font.Bold = True
    End If

    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Welcome! This form helps you configure MUTCD NY workzone traffic control signs. Please select a workzone category."
        lblStatus.Top = 500
        lblStatus.Left = 20
        lblStatus.Width = 550
        lblStatus.Height = 20
    End If
    
    ' Initialize table
    Debug.Print "Initializing sign table..."
    rowCount = 0
    If ControlExists("frameSignTable") Then
        Debug.Print "frameSignTable exists, creating table..."
        On Error Resume Next
        Call CreateTableHeaders
        Debug.Print "Headers created"
        ' Note: Start with no rows - user clicks "+" to add rows
        ' This prevents crashes from dynamic control creation
        Debug.Print "Table initialized - use + button to add rows"
        On Error GoTo 0
    Else
        Debug.Print "WARNING: frameSignTable does not exist!"
    End If

    Debug.Print "UserForm_Initialize completed successfully"
    Call BuildWZTCOrderTable
    Exit Sub
InitError:
    MsgBox "Error initializing form at line: " & Erl & vbCrLf & "Error: " & Err.Description & vbCrLf & "Number: " & Err.Number, vbCritical, "Initialization Error"
    Debug.Print "CRASH at line: " & Erl & " - " & Err.Description
End Sub

' ============================================================
' CHECK IF CONTROL EXISTS
' ============================================================
Private Function ControlExists(controlName As String) As Boolean
    On Error GoTo NotFound
    Dim ctrl As Object
    Set ctrl = Me.Controls(controlName)
    ControlExists = True
    Exit Function
NotFound:
    ControlExists = False
End Function
Private Sub PopulateCategories()
    cboCategory.Clear
    cboCategory.AddItem "001-020: General Information (8 sheets)"
    cboCategory.AddItem "021-099: Special Operations (13 sheets)"
    cboCategory.AddItem "101-109: Stop & Go Operations (4 sheets)"
    cboCategory.AddItem "110-200: Mobile Operations (5 sheets)"
    cboCategory.AddItem "201-300: Short Duration Operations (12 sheets)"
    cboCategory.AddItem "301-400: Short Term Operations (25 sheets)"
    cboCategory.AddItem "401-500: Intermediate Operations (17 sheets)"
    cboCategory.AddItem "501-600: Long Term Operations (10 sheets)"
    cboCategory.AddItem "ALL: Show All Sheets (91 total)"
    cboCategory.ListIndex = -1
End Sub

Private Sub cboCategory_Change()
    If cboCategory.ListIndex >= 0 Then
        Call PopulateSheets(cboCategory.List(cboCategory.ListIndex))
        lblStatus.Caption = "Category selected - Please choose a standard sheet number."
    End If
End Sub


' ============================================================
' POPULATE SHEET NUMBERS BASED ON CATEGORY
' ============================================================
Private Sub PopulateSheets(categoryName As String)
    cboSheet.Clear
    
    Select Case categoryName
        Case "001-020: General Information (8 sheets)"
            cboSheet.AddItem "619-001: Temporary Positive Barrier (6 Sheets)"
            cboSheet.AddItem "619-002: Type III Construction Barricades (2 Sheets)"
            cboSheet.AddItem "619-004: Portable Temporary Wooden Sign Support"
            cboSheet.AddItem "619-005: Details on Placement of Portable Temporary Rumble Strips"
            cboSheet.AddItem "619-006: Speed Feedback in Work Zones"
            cboSheet.AddItem "619-010: Work Zone Traffic Control General Notes"
            cboSheet.AddItem "619-011: Work Zone Traffic Control General Tables and Legend"
            cboSheet.AddItem "619-012: Sign Table (2 Sheets)"
            
        Case "021-099: Special Operations (13 sheets)"
            cboSheet.AddItem "619-021: Work Beyond Shoulder - Non-Freeway Mowing"
            cboSheet.AddItem "619-022: Shoulder Encroachment - Non-Freeway Mowing"
            cboSheet.AddItem "619-023: Lane Closure/Encroachment - Two-Lane Mowing (2 Sheets)"
            cboSheet.AddItem "619-031: Work Beyond Shoulder - Freeway Mowing"
            cboSheet.AddItem "619-032: Shoulder Encroachment - Freeway Mowing"
            cboSheet.AddItem "619-033: Lane Encroachment - Freeway Mowing"
            cboSheet.AddItem "619-041: Lane Closure/Encroachment - Parkway Mowing"
            cboSheet.AddItem "619-050: Lane Closure/Encroachment - Two-Lane Mulching/Herbicide (2 Sheets)"
            cboSheet.AddItem "619-051: Lane Encroachment or Shoulder Closure - Freeway Mulching/Herbicide"
            cboSheet.AddItem "619-060: Lane Closure - Two-Lane Pavement Marking (2 Sheets)"
            cboSheet.AddItem "619-080: Work Beyond Shoulder - All Roadways All Durations"
            cboSheet.AddItem "619-090: Temporary Road Closure - Two-Lane Two-Way"
            cboSheet.AddItem "619-091: Temporary Intersection Closure - Two-Lane Two-Way"
            
        Case "101-109: Stop & Go Operations (4 sheets)"
            cboSheet.AddItem "619-101: Right Shoulder Closure - Non-Freeway Stop and Go"
            cboSheet.AddItem "619-102: Lane Closure - Non-Freeway Stop and Go"
            cboSheet.AddItem "619-103: Left Lane and Shoulder Closure - Freeway Stop and Go"
            cboSheet.AddItem "619-104: Left Two Lane and Shoulder Closure - Freeway Stop and Go"
            
        Case "110-200: Mobile Operations (5 sheets)"
            cboSheet.AddItem "619-110: Lane Encroachment/Shoulder Closure - Freeway Mobile (2 Sheets)"
            cboSheet.AddItem "619-111: Right Lane Closure - Freeway Mobile (2 Sheets)"
            cboSheet.AddItem "619-112: Right Two Lane Closure - Freeway Mobile (2 Sheets)"
            cboSheet.AddItem "619-113: Left Shoulder Closure on Ramp - Freeway Mobile"
            cboSheet.AddItem "619-114: Lane Closure - Parkway Mobile"
            
        Case "201-300: Short Duration Operations (12 sheets)"
            cboSheet.AddItem "619-201: Right Shoulder Closure - Non-Freeway Short Duration"
            cboSheet.AddItem "619-202: Left Lane Closure - Multilane Undivided Short Duration"
            cboSheet.AddItem "619-203: Right Lane Closure - Multilane Undivided Short Duration"
            cboSheet.AddItem "619-204: Two Way Left Turn Lane Closure - Multilane Undivided Short Duration"
            cboSheet.AddItem "619-205: Right Shoulder Closure - Freeway Short Duration (2 Sheets)"
            cboSheet.AddItem "619-206: Right Lane Closure - Freeway Short Duration (2 Sheets)"
            cboSheet.AddItem "619-207: Right Two Lane Closure - Freeway Short Duration (2 Sheets)"
            cboSheet.AddItem "619-208: Left Lane Closure - Freeway Short Duration"
            cboSheet.AddItem "619-209: Left Two Lane Closure - Freeway Short Duration"
            cboSheet.AddItem "619-211: Left Shoulder Closure on Exit Ramp - Freeway Short Duration"
            cboSheet.AddItem "619-212: Right/Left Lane Closure - Parkway Short Duration"
            
        Case "301-400: Short Term Operations (25 sheets)"
            cboSheet.AddItem "619-301: Right Shoulder Closure - Freeway Short Term"
            cboSheet.AddItem "619-302: Right Lane Closure - All Roadways Short Term"
            cboSheet.AddItem "619-303: Right (or Left) Two Lane Closure - All Roadways Short Term"
            cboSheet.AddItem "619-304: Left Lane Closure - Freeway Short Term"
            cboSheet.AddItem "619-305: Left Two Lane Closure - Freeway Short Term"
            cboSheet.AddItem "619-306: Right Lane Closure - Parkway Short Term"
            cboSheet.AddItem "619-307: Lane Closure with Flaggers - Two-Lane Short Term (2 Sheets)"
            cboSheet.AddItem "619-308: Lane Closure with Flagger Prior to Intersection - Two-Lane Short Term"
            cboSheet.AddItem "619-309: Lane Closure with AFAD's - Two-Lane Short Term (2 Sheets)"
            cboSheet.AddItem "619-310: Shoulder Closure - Non-Freeway Short Term"
            cboSheet.AddItem "619-311: Right Lane Closure - Multilane Undivided Short Term"
            cboSheet.AddItem "619-312: Two Way Left Turn Lane Closure - Multilane Undivided Short Term (2 Sheets)"
            cboSheet.AddItem "619-313: Right Lane Closure Thru Intersection - Multilane One Way Short Term"
            cboSheet.AddItem "619-314: Lane Closure with Moving Flaggers - Two-Lane Short Term"
            cboSheet.AddItem "619-315: Shoulder Closure at Ramp Approach - Freeway Short Term (2 Sheets)"
            cboSheet.AddItem "619-316: Partial Exit Ramp Closure - Freeway Short Term (2 Sheets)"
            cboSheet.AddItem "619-317: Single Lane Closure - Multilane Undivided Short Term (2 Sheets)"
            cboSheet.AddItem "619-318: Single Lane Closure Near Entrance Ramp - Freeway Short Term (2 Sheets)"
            cboSheet.AddItem "619-319: Single Lane Closure Near Exit Ramp - Freeway Short Term (2 Sheets)"
            cboSheet.AddItem "619-321: Sidewalk Detour or Diversion - Two-Lane Short Term (2 Sheets)"
            cboSheet.AddItem "619-322: Crosswalk Closure and Pedestrian Detour - Two-Lane Short Term"
            cboSheet.AddItem "619-323: Flagging Operation at Intersection - Two-Lane Short Term"
            cboSheet.AddItem "619-324: Single Lane Shift with Two Way Left Turn Lane - Two-Lane Short Term"
            cboSheet.AddItem "619-325: Double Interior Lane Closure - Multilane Two-Way Short Term"
            
        Case "401-500: Intermediate Operations (17 sheets)"
            cboSheet.AddItem "619-401: Right Shoulder Closure - Freeway Intermediate (2 Sheets)"
            cboSheet.AddItem "619-402: Right Lane Closure - All Roadways Intermediate (2 Sheets)"
            cboSheet.AddItem "619-403: Right (or Left) Two Lane Closure - All Roadways Intermediate (2 Sheets)"
            cboSheet.AddItem "619-407: Lane Closure with Flaggers - Two-Lane Intermediate (2 Sheets)"
            cboSheet.AddItem "619-410: Shoulder Closure - Two-Lane Intermediate (2 Sheets)"
            cboSheet.AddItem "619-412: Two Way Left Turn Lane Closure - Multilane Undivided Intermediate (2 Sheets)"
            cboSheet.AddItem "619-414: Single Lane Closure - Multilane Undivided Intermediate (2 Sheets)"
            cboSheet.AddItem "619-415: Shoulder Closure at Ramp Approach - Freeway Intermediate (2 Sheets)"
            cboSheet.AddItem "619-416: Partial Exit Ramp Closure - Freeway Intermediate (2 Sheets)"
            cboSheet.AddItem "619-417: Single Lane Closure Near Entrance Ramp - Freeway Intermediate (2 Sheets)"
            cboSheet.AddItem "619-418: Single Lane Closure Near Exit Ramp - Freeway Intermediate (2 Sheets)"
            cboSheet.AddItem "619-419: Sidewalk Detour or Diversion - Two-Lane Intermediate (2 Sheets)"
            cboSheet.AddItem "619-420: Crosswalk Closure and Pedestrian Detour - Two-Lane Intermediate"
            cboSheet.AddItem "619-421: Flagging Operation at Intersection - Two-Lane Intermediate (2 Sheets)"
            cboSheet.AddItem "619-422: Single Lane Shift with Two Way Left Turn Lane - Two-Lane Intermediate"
            cboSheet.AddItem "619-423: Double Interior Lane Closure - Multilane Two-Way Intermediate"
            
        Case "501-600: Long Term Operations (10 sheets)"
            cboSheet.AddItem "619-501: Right Shoulder Closure - Freeway Long Term (2 Sheets)"
            cboSheet.AddItem "619-502: Multi Lane Shift - Freeway Long Term (2 Sheets)"
            cboSheet.AddItem "619-503: Offsite Detour Road/Bridge - Non-Freeway Long Term (3 Sheets)"
            cboSheet.AddItem "619-504: Single Lane Closure - Multilane Divided/Freeway Long Term (2 Sheets)"
            cboSheet.AddItem "619-517: Single Lane Closure Near Entrance Ramp - Freeway Long Term (2 Sheets)"
            cboSheet.AddItem "619-518: Single Lane Closure Near Exit Ramp - Freeway Long Term (2 Sheets)"
            cboSheet.AddItem "619-519: Sidewalk Detour or Diversion - Two-Lane Long Term (2 Sheets)"
            cboSheet.AddItem "619-520: Crosswalk Closure and Pedestrian Detour - Two-Lane Long Term"
            cboSheet.AddItem "619-523: Double Interior Lane Closure - Multilane Two-Way Long Term"
            cboSheet.AddItem "619-524: Temporary Traffic Signal - Two-Lane Long Term"
    End Select
    
    If cboSheet.ListCount > 0 Then
        cboSheet.ListIndex = 0
    End If
End Sub

Private Sub cboSheet_Change()
    If cboSheet.ListIndex >= 0 Then
        lblStatus.Caption = "Standard Sheet Number selected - Please select Road Speed."
        Call CheckAllSelectionsComplete
    End If
End Sub

' ============================================================
' POPULATE ROAD TYPE
' ============================================================
Private Sub PopulateRoadType()
    On Error GoTo PopError
    If ControlExists("cboRoadType") Then
        cboRoadType.Clear
        cboRoadType.AddItem "Select road type..."
        cboRoadType.AddItem "Freeway"
        cboRoadType.AddItem "Non-Freeway"
        cboRoadType.ListIndex = 0
    End If
    Exit Sub
PopError:
    MsgBox "Error populating road types: " & Err.Description, vbExclamation
End Sub

Private Sub cboRoadType_Change()
    If cboRoadType.ListIndex > 0 Then
        selectedRoadType = cboRoadType.Value
        lblStatus.Caption = "Road Type selected - Please select Lane Width."
        Call CheckAllSelectionsComplete
    End If
End Sub

' ============================================================
' POPULATE LANE WIDTH
' ============================================================
Private Sub PopulateLaneWidth()
    On Error GoTo PopError
    If ControlExists("cboLaneWidth") Then
        cboLaneWidth.Clear
        cboLaneWidth.AddItem "Select lane width..."
        cboLaneWidth.AddItem "10 ft"
        cboLaneWidth.AddItem "11 ft"
        cboLaneWidth.AddItem "12 ft"
        cboLaneWidth.ListIndex = 0
    End If
    Exit Sub
PopError:
    MsgBox "Error populating lane widths: " & Err.Description, vbExclamation
End Sub

Private Sub cboLaneWidth_Change()
    If cboLaneWidth.ListIndex > 0 Then
        If cboRoadSpeed.ListIndex > 0 And cboCategory.ListIndex > 0 Then
            selectedSpeed = cboRoadSpeed.Value
            Call GenerateSpacingTable
        End If
        lblStatus.Caption = "Lane Width selected - Please select Shoulder Width."
        Call CheckAllSelectionsComplete
    End If
End Sub

' ============================================================
' POPULATE SHOULDER WIDTHS
' ============================================================
Private Sub PopulateShoulderWidth()
    On Error GoTo PopError
    If ControlExists("cboShoulderWidth") Then
        cboShoulderWidth.Clear
        cboShoulderWidth.AddItem "Select shoulder width..."
        cboShoulderWidth.AddItem "<= 4 ft"
        cboShoulderWidth.AddItem "5-7 ft"
        cboShoulderWidth.AddItem "8 ft"
        cboShoulderWidth.AddItem "9 ft"
        cboShoulderWidth.AddItem "10 ft"
        cboShoulderWidth.AddItem "11 ft"
        cboShoulderWidth.AddItem "12 ft"
        cboShoulderWidth.ListIndex = 0
    End If
    Exit Sub
PopError:
    MsgBox "Error populating shoulder widths: " & Err.Description, vbExclamation
End Sub

Private Sub cboShoulderWidth_Change()
    If cboShoulderWidth.ListIndex > 0 Then
        If cboRoadSpeed.ListIndex > 0 And cboCategory.ListIndex > 0 Then
            selectedSpeed = cboRoadSpeed.Value
            Call GenerateSpacingTable
        End If
        lblStatus.Caption = "Shoulder Width selected - Checking all selections..."
        Call CheckAllSelectionsComplete
    End If
End Sub

' ============================================================
' POPULATE ROAD SPEEDS
' ============================================================
Private Sub PopulateRoadSpeeds()
    On Error GoTo PopError
    If ControlExists("cboRoadSpeed") Then
        cboRoadSpeed.Clear
        cboRoadSpeed.AddItem "Select a speed..."
        cboRoadSpeed.AddItem "25 mph"
        cboRoadSpeed.AddItem "30 mph"
        cboRoadSpeed.AddItem "35 mph"
        cboRoadSpeed.AddItem "40 mph"
        cboRoadSpeed.AddItem "45 mph"
        cboRoadSpeed.AddItem "50 mph"
        cboRoadSpeed.AddItem "55 mph"
        cboRoadSpeed.AddItem "65 mph"
        cboRoadSpeed.ListIndex = 0
    End If
    Exit Sub
PopError:
    MsgBox "Error populating road speeds: " & Err.Description, vbExclamation
End Sub


' ============================================================
' ROAD SPEED CHANGE EVENT - TRIGGERS TABLE GENERATION
' ============================================================
Private Sub cboRoadSpeed_Change()
    If cboRoadSpeed.ListIndex > 0 And cboCategory.ListIndex > 0 Then
        selectedSpeed = cboRoadSpeed.Value
        Call GenerateSpacingTable
        Call PopulateSignTable
        lblStatus.Caption = "Road Speed selected - Please select Road Type."
        Call CheckAllSelectionsComplete
    End If
End Sub

' ============================================================
' CREATE SPACING LABELS IN FRAME
' ============================================================
Private Sub CreateSpacingLabels()
    On Error GoTo SpacingError

    ' Create labels and textboxes for the spacing values
    Dim lblDownstream As MSForms.Label
    Dim txtDownstream As MSForms.TextBox
    Dim lblRollAhead As MSForms.Label
    Dim txtRollAhead As MSForms.TextBox
    Dim lblVehicle As MSForms.Label
    Dim txtVehicle As MSForms.TextBox
    Dim lblBuffer As MSForms.Label
    Dim txtBuffer As MSForms.TextBox
    Dim lblMerging As MSForms.Label
    Dim txtMerging As MSForms.TextBox
    Dim lblShoulder As MSForms.Label
    Dim txtShoulder As MSForms.TextBox
    Dim lblAdvanced As MSForms.Label
    Dim txtAdvanced As MSForms.TextBox
    Dim lblSkipLines As MSForms.Label
    Dim txtSkipLines As MSForms.TextBox
    Dim lblChannelizing As MSForms.Label
    Dim txtChannelizing As MSForms.TextBox
    Dim lblFlareBarrier As MSForms.Label
    Dim txtFlareBarrier As MSForms.TextBox
    Dim lblFlareBeam As MSForms.Label
    Dim txtFlareBeam As MSForms.TextBox

    ' Label: downstreamTaper
    Set lblDownstream = frameSpacingValues.Controls.Add("Forms.Label.1", "lblDownstreamTaper")
    With lblDownstream
        .Caption = "Downstream Taper (ft):"
        .Top = 20
        .Left = 10
        .Width = 110
        .Height = 18
    End With

    Set txtDownstream = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtDownstreamTaper")
    With txtDownstream
        .Top = 20
        .Left = 125
        .Width = 50
        .Height = 18
        .Enabled = False
    End With

    ' Label: Roll Ahead Distance (inserted below downstream taper)
    Set lblRollAhead = frameSpacingValues.Controls.Add("Forms.Label.1", "lblRollAhead")
    With lblRollAhead
        .Caption = "Roll Ahead Distance (ft):"
        .Top = 40
        .Left = 10
        .Width = 110
        .Height = 18
    End With

    Set txtRollAhead = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtRollAhead")
    With txtRollAhead
        .Top = 40
        .Left = 125
        .Width = 50
        .Height = 18
        .Enabled = False
    End With

    ' Label: vehicleSpace
    Set lblVehicle = frameSpacingValues.Controls.Add("Forms.Label.1", "lblVehicleSpace")
    With lblVehicle
        .Caption = "Vehicle Space (ft):"
        .Top = 60
        .Left = 10
        .Width = 110
        .Height = 18
    End With

    Set txtVehicle = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtVehicleSpace")
    With txtVehicle
        .Top = 60
        .Left = 125
        .Width = 50
        .Height = 18
        .Enabled = False
    End With

    ' Label: bufferSpace
    Set lblBuffer = frameSpacingValues.Controls.Add("Forms.Label.1", "lblBufferSpace")
    With lblBuffer
        .Caption = "Buffer Space (ft):"
        .Top = 80
        .Left = 10
        .Width = 110
        .Height = 18
    End With

    Set txtBuffer = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtBufferSpace")
    With txtBuffer
        .Top = 80
        .Left = 125
        .Width = 50
        .Height = 18
        .Enabled = False
    End With

    ' Label: mergingTaper
    Set lblMerging = frameSpacingValues.Controls.Add("Forms.Label.1", "lblMergingTaper")
    With lblMerging
        .Caption = "Merging/Shifting Taper (ft):"
        .Top = 100
        .Left = 10
        .Width = 110
        .Height = 18
    End With

    Set txtMerging = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtMergingTaper")
    With txtMerging
        .Top = 100
        .Left = 125
        .Width = 50
        .Height = 18
        .Enabled = False
    End With

    ' Label: shoulderTapers
    Set lblShoulder = frameSpacingValues.Controls.Add("Forms.Label.1", "lblShoulderTapers")
    With lblShoulder
        .Caption = "Shoulder Tapers (ft):"
        .Top = 120
        .Left = 10
        .Width = 110
        .Height = 18
    End With

    Set txtShoulder = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtShoulderTapers")
    With txtShoulder
        .Top = 120
        .Left = 125
        .Width = 50
        .Height = 18
        .Enabled = False
    End With

    ' Label: advancedWarningSpacing
    Set lblAdvanced = frameSpacingValues.Controls.Add("Forms.Label.1", "lblAdvancedWarningSpacing")
    With lblAdvanced
        .Caption = "Advanced Warning Spacing (ft):"
        .Top = 140
        .Left = 10
        .Width = 110
        .Height = 18
    End With

    Set txtAdvanced = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtAdvancedWarningSpacing")
    With txtAdvanced
        .Top = 140
        .Left = 125
        .Width = 50
        .Height = 18
        .Enabled = False
    End With

    ' Label: # of Skip Lines
    Set lblSkipLines = frameSpacingValues.Controls.Add("Forms.Label.1", "lblSkipLines")
    With lblSkipLines
        .Caption = "# of Skip Lines:"
        .Top = 160
        .Left = 10
        .Width = 110
        .Height = 18
    End With

    Set txtSkipLines = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtSkipLines")
    With txtSkipLines
        .Top = 160
        .Left = 125
        .Width = 50
        .Height = 18
    End With

    ' Label: # of Channelizing Devices
    Set lblChannelizing = frameSpacingValues.Controls.Add("Forms.Label.1", "lblChannelizing")
    With lblChannelizing
        .Caption = "# of Channelizing Devices:"
        .Top = 180
        .Left = 10
        .Width = 110
        .Height = 18
    End With

    Set txtChannelizing = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtChannelizing")
    With txtChannelizing
        .Top = 180
        .Left = 125
        .Width = 50
        .Height = 18
    End With

    ' Label: Flare Rate Temporary Positive Barrier
    Set lblFlareBarrier = frameSpacingValues.Controls.Add("Forms.Label.1", "lblFlareBarrier")
    With lblFlareBarrier
        .Caption = "Flare Rate Temp Barrier:"
        .Top = 200
        .Left = 10
        .Width = 110
        .Height = 18
    End With

    Set txtFlareBarrier = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtFlareBarrier")
    With txtFlareBarrier
        .Top = 200
        .Left = 125
        .Width = 50
        .Height = 18
    End With

    ' Label: Flare Rate Box Beam/Corrugated Beam
    Set lblFlareBeam = frameSpacingValues.Controls.Add("Forms.Label.1", "lblFlareBeam")
    With lblFlareBeam
        .Caption = "Flare Rate Box/Corr Beam:"
        .Top = 220
        .Left = 10
        .Width = 110
        .Height = 18
    End With

    Set txtFlareBeam = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtFlareBeam")
    With txtFlareBeam
        .Top = 220
        .Left = 125
        .Width = 50
        .Height = 18
    End With

    Exit Sub
SpacingError:
    MsgBox "Error in CreateSpacingLabels: " & Err.Description, vbExclamation
End Sub

' ============================================================
' CHECK IF ALL REQUIRED SELECTIONS ARE COMPLETE
' ============================================================
Private Sub CheckAllSelectionsComplete()
    Dim allComplete As Boolean
    
    ' Check if all 6 required fields have valid selections:
    ' 1-Category, 2-Sheet, 3-Speed, 4-Road Type, 5-Lane Width, 6-Shoulder Width
    allComplete = (cboCategory.ListIndex >= 0) And _
                  (cboSheet.ListIndex >= 0) And _
                  (cboRoadSpeed.ListIndex > 0) And _
                  (cboRoadType.ListIndex > 0) And _
                  (cboLaneWidth.ListIndex > 0) And _
                  (cboShoulderWidth.ListIndex > 0)
    
    If allComplete Then
        lblStatus.Caption = "Ready to configure signs for: " & cboSheet.List(cboSheet.ListIndex)
    End If
End Sub

' ============================================================
' GENERATE SPACING TABLE BASED ON MUTCD NY STANDARDS
' ============================================================
Private Sub GenerateSpacingTable()
    Dim speed As Integer
    Dim laneWidth As Integer
    Dim downstreamTaper As Double
    Dim vehicleSpace As Double
    Dim bufferSpace As Double
    Dim mergingTaper As Double
    Dim shoulderTapers As Double
    Dim advancedWarningSpacing As Double
    Dim skipMerge As Integer, chanMerge As Integer
    Dim skipShoulder As Integer, chanShoulder As Integer
    Dim skipBuffer As Integer
    Dim skipRollAhead As Integer
    Dim flareBarrierStr As String, flareBeamStr As String
    Dim skipTotal As Integer, chanTotal As Integer
    
    ' Extract speed value from dropdown
    speed = Val(Left(selectedSpeed, 2))
    
    ' Extract lane width value from dropdown
    laneWidth = Val(Left(cboLaneWidth.Value, 2))
    
    ' MUTCD NY formulas for spacing based on speed
    ' Downstream taper depends on road type (Freeway vs Non-Freeway)
    If LCase(Trim(selectedRoadType)) = "non-freeway" Then
        downstreamTaper = 50
    Else
        downstreamTaper = 100
    End If
    vehicleSpace = 50
    
    ' Buffer Space based on preconstruction posted speed limit table
    Select Case speed
        Case 25
            bufferSpace = 155
        Case 30
            bufferSpace = 200
        Case 35
            bufferSpace = 250
        Case 40
            bufferSpace = 305
        Case 45
            bufferSpace = 360
        Case 50
            bufferSpace = 425
        Case 55
            bufferSpace = 495
        Case 65
            bufferSpace = 645
        Case Else
            bufferSpace = speed * 70  ' Default formula if speed not in table
    End Select
    
    ' Buffer Space Skip Lines based on speed only
    Select Case speed
        Case 25
            skipBuffer = 4
        Case 30
            skipBuffer = 5
        Case 35
            skipBuffer = 6
        Case 40
            skipBuffer = 8
        Case 45
            skipBuffer = 9
        Case 50
            skipBuffer = 11
        Case 55
            skipBuffer = 13
        Case 65
            skipBuffer = 16
        Case Else
            skipBuffer = 0
    End Select
    
    ' Merging/Shifting Taper based on speed and lane width table
    ' Also captures skip lines and channelizing device counts
    Select Case speed
        Case 25
            Select Case laneWidth
                Case 10: mergingTaper = 120: skipMerge = 3: chanMerge = 4
                Case 11: mergingTaper = 120: skipMerge = 3: chanMerge = 4
                Case 12: mergingTaper = 120: skipMerge = 3: chanMerge = 4
                Case Else: mergingTaper = 120: skipMerge = 3: chanMerge = 4
            End Select
        Case 30
            Select Case laneWidth
                Case 10: mergingTaper = 160: skipMerge = 4: chanMerge = 5
                Case 11: mergingTaper = 160: skipMerge = 4: chanMerge = 5
                Case 12: mergingTaper = 200: skipMerge = 5: chanMerge = 6
                Case Else: mergingTaper = 160: skipMerge = 4: chanMerge = 5
            End Select
        Case 35
            Select Case laneWidth
                Case 10: mergingTaper = 200: skipMerge = 5: chanMerge = 6
                Case 11: mergingTaper = 240: skipMerge = 6: chanMerge = 7
                Case 12: mergingTaper = 240: skipMerge = 6: chanMerge = 7
                Case Else: mergingTaper = 200: skipMerge = 5: chanMerge = 6
            End Select
        Case 40
            Select Case laneWidth
                Case 10: mergingTaper = 280: skipMerge = 7: chanMerge = 8
                Case 11: mergingTaper = 320: skipMerge = 8: chanMerge = 9
                Case 12: mergingTaper = 320: skipMerge = 8: chanMerge = 9
                Case Else: mergingTaper = 280: skipMerge = 7: chanMerge = 8
            End Select
        Case 45
            Select Case laneWidth
                Case 10: mergingTaper = 440: skipMerge = 11: chanMerge = 12
                Case 11: mergingTaper = 520: skipMerge = 13: chanMerge = 14
                Case 12: mergingTaper = 560: skipMerge = 14: chanMerge = 15
                Case Else: mergingTaper = 440: skipMerge = 11: chanMerge = 12
            End Select
        Case 50
            Select Case laneWidth
                Case 10: mergingTaper = 520: skipMerge = 13: chanMerge = 14
                Case 11: mergingTaper = 560: skipMerge = 14: chanMerge = 15
                Case 12: mergingTaper = 600: skipMerge = 15: chanMerge = 16
                Case Else: mergingTaper = 520: skipMerge = 13: chanMerge = 14
            End Select
        Case 55
            Select Case laneWidth
                Case 10: mergingTaper = 560: skipMerge = 14: chanMerge = 15
                Case 11: mergingTaper = 600: skipMerge = 15: chanMerge = 16
                Case 12: mergingTaper = 680: skipMerge = 17: chanMerge = 18
                Case Else: mergingTaper = 560: skipMerge = 14: chanMerge = 15
            End Select
        Case 65
            Select Case laneWidth
                Case 10: mergingTaper = 640: skipMerge = 16: chanMerge = 17
                Case 11: mergingTaper = 720: skipMerge = 18: chanMerge = 19
                Case 12: mergingTaper = 800: skipMerge = 19: chanMerge = 20
                Case Else: mergingTaper = 640: skipMerge = 16: chanMerge = 17
            End Select
        Case Else
            mergingTaper = (speed * (laneWidth)^2)/60: skipMerge = 0: chanMerge = 0
    End Select
    
    ' Shoulder Taper based on speed and shoulder width table
    ' Also captures skip lines and channelizing device counts
    Select Case speed
        Case 25
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "5-7 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "8 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "9 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "10 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "11 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "12 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case Else: shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
            End Select
        Case 30
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "5-7 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "8 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "9 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "10 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "11 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "12 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case Else: shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
            End Select
        Case 35
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "5-7 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "8 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "9 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "10 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "11 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "12 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case Else: shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
            End Select
        Case 40
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "5-7 ft": shoulderTapers = 80: skipShoulder = 1: chanShoulder = 2
                Case "8 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "9 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "10 ft": shoulderTapers = 120: skipShoulder = 2: chanShoulder = 3
                Case "11 ft": shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "12 ft": shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case Else: shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
            End Select
        Case 45
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "5-7 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "8 ft": shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "9 ft": shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "10 ft": shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "11 ft": shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "12 ft": shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case Else: shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
            End Select
        Case 50
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "5-7 ft": shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "8 ft": shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "9 ft": shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "10 ft": shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "11 ft": shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "12 ft": shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case Else: shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
            End Select
        Case 55
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "5-7 ft": shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "8 ft": shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "9 ft": shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "10 ft": shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "11 ft": shoulderTapers = 200: skipShoulder = 5: chanShoulder = 6
                Case "12 ft": shoulderTapers = 200: skipShoulder = 5: chanShoulder = 6
                Case Else: shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
            End Select
        Case 65
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "5-7 ft": shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "8 ft": shoulderTapers = 200: skipShoulder = 5: chanShoulder = 6
                Case "9 ft": shoulderTapers = 240: skipShoulder = 6: chanShoulder = 7
                Case "10 ft": shoulderTapers = 240: skipShoulder = 6: chanShoulder = 7
                Case "11 ft": shoulderTapers = 280: skipShoulder = 7: chanShoulder = 8
                Case "12 ft": shoulderTapers = 280: skipShoulder = 7: chanShoulder = 8
                Case Else: shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
            End Select
        Case Else
            shoulderTapers = speed * 0.8: skipShoulder = 0: chanShoulder = 0
    End Select
    
    ' Advanced placement of warning sign based on speed (lookup)
    Select Case speed
        Case 25
            advancedWarningSpacing = 515
        Case 30
            advancedWarningSpacing = 620
        Case 35
            advancedWarningSpacing = 720
        Case 40
            advancedWarningSpacing = 825
        Case 45
            advancedWarningSpacing = 930
        Case 50
            advancedWarningSpacing = 1030
        Case 55
            advancedWarningSpacing = 1135
        Case 65
            advancedWarningSpacing = 1365
        Case Else
            advancedWarningSpacing = speed * 10
    End Select

    ' Roll ahead distance based on speed (lookup)
    Dim rollAhead As Double
    Select Case speed
        Case 25, 30, 35, 40
            rollAhead = 120
        Case 45, 50, 55
            rollAhead = 160
        Case 65
            rollAhead = 200
        Case Else
            rollAhead = 120
    End Select
    
    ' Roll ahead distance skip lines based on speed
    Select Case speed
        Case 25, 30, 35, 40
            skipRollAhead = 3
        Case 45, 50, 55
            skipRollAhead = 4
        Case 65
            skipRollAhead = 5
        Case Else
            skipRollAhead = 0
    End Select
    
    ' Determine flare rates based on speed
    Select Case speed
        Case 25, 30, 35
            flareBarrierStr = "8:1"
        Case 40, 45
            flareBarrierStr = "11:1"
        Case 50
            flareBarrierStr = "14:1"
        Case 55
            flareBarrierStr = "16:1"
        Case 65
            flareBarrierStr = "20:1"
        Case Else
            flareBarrierStr = ""
    End Select

    Select Case speed
        Case 25, 30, 35
            flareBeamStr = "7:1"
        Case 40, 45
            flareBeamStr = "9:1"
        Case 50
            flareBeamStr = "11:1"
        Case 55
            flareBeamStr = "12:1"
        Case 65
            flareBeamStr = "15:1"
        Case Else
            flareBeamStr = ""
    End Select

    ' Sum skip lines and channelizing devices from merging, shoulder, buffer and roll-ahead contributions
    skipTotal = skipMerge + skipShoulder + skipBuffer + skipRollAhead
    chanTotal = chanMerge + chanShoulder
    
    ' Populate the text boxes
    frameSpacingValues.Controls("txtDownstreamTaper").Value = Format(downstreamTaper, "0.0")
    frameSpacingValues.Controls("txtRollAhead").Value = Format(rollAhead, "0.0")
    frameSpacingValues.Controls("txtVehicleSpace").Value = Format(vehicleSpace, "0.0")
    frameSpacingValues.Controls("txtBufferSpace").Value = Format(bufferSpace, "0.0")
    frameSpacingValues.Controls("txtMergingTaper").Value = Format(mergingTaper, "0.0")
    frameSpacingValues.Controls("txtShoulderTapers").Value = Format(shoulderTapers, "0.0")
    frameSpacingValues.Controls("txtAdvancedWarningSpacing").Value = Format(advancedWarningSpacing, "0.0")
    frameSpacingValues.Controls("txtSkipLines").Value = Format(skipTotal, "0")
    frameSpacingValues.Controls("txtChannelizing").Value = Format(chanTotal, "0")
    frameSpacingValues.Controls("txtFlareBarrier").Value = flareBarrierStr
    frameSpacingValues.Controls("txtFlareBeam").Value = flareBeamStr

    Call BuildWZTCOrderTable

End Sub

' ============================================================
' CREATE TABLE HEADERS IN FRAME
' ============================================================
Private Sub CreateTableHeaders()
    Dim lblSignNum As MSForms.Label
    Dim lblSpacing As MSForms.Label
    Dim lblSize As MSForms.Label
    Dim lblSide As MSForms.Label
    
    ' Create header labels
    Set lblSignNum = frameSignTable.Controls.Add("Forms.Label.1", "lblHeaderSignNum")
    With lblSignNum
        .Caption = "Sign Number"
        .Top = 15
        .Left = TABLE_LEFT
        .Width = COL1_WIDTH
        .Height = 20
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackColor = &H8000000F  ' Light gray
    End With
    
    Set lblSpacing = frameSignTable.Controls.Add("Forms.Label.1", "lblHeaderSpacing")
    With lblSpacing
        .Caption = "Spacing (ft)"
        .Top = 15
        .Left = TABLE_LEFT + COL1_WIDTH + 5
        .Width = COL2_WIDTH
        .Height = 20
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackColor = &H8000000F
    End With
    
    Set lblSize = frameSignTable.Controls.Add("Forms.Label.1", "lblHeaderSize")
    With lblSize
        .Caption = "Size"
        .Top = 15
        .Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + 10
        .Width = COL3_WIDTH
        .Height = 20
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackColor = &H8000000F
    End With
    
    Set lblSide = frameSignTable.Controls.Add("Forms.Label.1", "lblHeaderSide")
    With lblSide
        .Caption = "Road Side"
        .Top = 15
        .Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + COL3_WIDTH + 15
        .Width = COL4_WIDTH
        .Height = 20
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackColor = &H8000000F
    End With
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
    On Error GoTo RowError

    Dim txtSignNum As MSForms.TextBox
    Dim txtSpacing As MSForms.TextBox
    Dim txtSize As MSForms.TextBox
    Dim cboSide As MSForms.ComboBox
    Dim currentTop As Integer

    Debug.Print "Adding row " & (rowCount + 1)

    rowCount = rowCount + 1
    currentTop = TABLE_START_TOP + (rowCount - 1) * ROW_HEIGHT

    ' Resize arrays
    ReDim Preserve signNumberBoxes(1 To rowCount)
    ReDim Preserve signSpacingBoxes(1 To rowCount)
    ReDim Preserve signSizeBoxes(1 To rowCount)
    ReDim Preserve signSideComboBoxes(1 To rowCount)

    ' Sign Number textbox
    Debug.Print "Creating sign number textbox..."
    Set txtSignNum = frameSignTable.Controls.Add("Forms.TextBox.1", "txtSignNum" & rowCount)
    With txtSignNum
        .Top = currentTop
        .Left = TABLE_LEFT
        .Width = COL1_WIDTH
        .Height = 22
        .Text = ""
    End With
    Set signNumberBoxes(rowCount) = txtSignNum
    ' Wire Exit event so typing a sign number auto-fills spacing and size from library
    Dim handler As SignNumBox
    Set handler = New SignNumBox
    handler.RowIndex = rowCount
    Set handler.Txt = txtSignNum
    Set handler.ParentForm = Me
    signNumberHandlers.Add handler, CStr(rowCount)
    Debug.Print "Sign number textbox created"

    ' Spacing textbox
    Debug.Print "Creating spacing textbox..."
    Set txtSpacing = frameSignTable.Controls.Add("Forms.TextBox.1", "txtSpacing" & rowCount)
    With txtSpacing
        .Top = currentTop
        .Left = TABLE_LEFT + COL1_WIDTH + 5
        .Width = COL2_WIDTH
        .Height = 22
        .Text = ""
    End With
    Set signSpacingBoxes(rowCount) = txtSpacing
    Debug.Print "Spacing textbox created"

    ' Size textbox
    Debug.Print "Creating size textbox..."
    Set txtSize = frameSignTable.Controls.Add("Forms.TextBox.1", "txtSize" & rowCount)
    With txtSize
        .Top = currentTop
        .Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + 10
        .Width = COL3_WIDTH
        .Height = 22
        .Text = ""
    End With
    Set signSizeBoxes(rowCount) = txtSize
    Debug.Print "Size textbox created"

    ' Side selection combobox (replaces option buttons for reliability)
    Debug.Print "Creating side selection combobox..."
    Set cboSide = frameSignTable.Controls.Add("Forms.ComboBox.1", "cboSide" & rowCount)
    With cboSide
        .Top = currentTop
        .Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + COL3_WIDTH + 15
        .Width = COL4_WIDTH
        .Height = 22
        .AddItem "One Side"
        .AddItem "Both Sides"
        .ListIndex = 0  ' Default to "One Side"
        .Style = fmStyleDropDownList  ' Dropdown list style
    End With
    Set signSideComboBoxes(rowCount) = cboSide
    Debug.Print "Side combobox created"

    ' Adjust frame height if needed
    If currentTop + 40 > frameSignTable.Height Then
        frameSignTable.Height = currentTop + 40
    End If

    Debug.Print "Row " & rowCount & " added successfully"
    Exit Sub

RowError:
    MsgBox "Error adding row at line " & Erl & ":" & vbCrLf & Err.Description & vbCrLf & "Row count: " & rowCount, vbCritical, "Add Row Error"
    Debug.Print "ERROR adding row: " & Err.Description & " at line " & Erl
    ' Rollback the row count if we failed
    If rowCount > 0 Then rowCount = rowCount - 1
End Sub

' ============================================================
' POPULATE SIGN TABLE (CLEAR ONLY - USER TYPES SIGN NUMBERS)
' User types sign number in each row; spacing and size auto-fill from library (SignLibrary) on Exit.
' ============================================================
Private Sub PopulateSignTable()
    Dim i As Integer
    ' Clear existing data only; user adds sign numbers and library fills spacing/size
    For i = 1 To rowCount
        signNumberBoxes(i).Value = ""
        signSpacingBoxes(i).Value = ""
        signSizeBoxes(i).Value = ""
        signSideComboBoxes(i).ListIndex = 0  ' Default to "One Side"
    Next i
End Sub

' ============================================================
' APPLY SIGN LIBRARY TO ROW (called when user leaves sign number field or presses Enter)
' Looks up sign in SignLibrary and auto-fills spacing and size for that row.
' ============================================================
Public Sub ApplySignLibraryToRow(rowIndex As Integer)
    On Error GoTo ApplyLibError
    If rowIndex < 1 Or rowIndex > rowCount Then Exit Sub
    Dim s As String
    Dim sd As signData
    Dim allSigns() As String
    Dim i As Long
    Dim matchKey As String
    s = Trim(signNumberBoxes(rowIndex).Value)
    If s = "" Then Exit Sub
    sd = GetSignData(s)
    If sd.SignNumber = "" Then
        ' Try case-insensitive match (library keys are case-sensitive)
        allSigns = GetAllSignNumbers
        matchKey = ""
        For i = LBound(allSigns) To UBound(allSigns)
            If allSigns(i) <> "" And StrComp(s, allSigns(i), vbTextCompare) = 0 Then
                matchKey = allSigns(i)
                Exit For
            End If
        Next i
        If matchKey = "" Then Exit Sub
        sd = GetSignData(matchKey)
        If sd.SignNumber = "" Then Exit Sub
    End If
    signSpacingBoxes(rowIndex).Value = CStr(sd.DefaultSpacing)
    signSizeBoxes(rowIndex).Value = sd.TextLine2
    Exit Sub
ApplyLibError:
    Debug.Print "ApplySignLibraryToRow error: " & Err.Description
End Sub

' Move focus to the spacing textbox for the given row (after Enter in sign number).
Public Sub MoveFocusToSpacing(rowIndex As Integer)
    On Error Resume Next
    If rowIndex >= 1 And rowIndex <= rowCount Then
        signSpacingBoxes(rowIndex).SetFocus
    End If
End Sub

' ============================================================
' ADD ROW BUTTON CLICK EVENT
' ============================================================
Private Sub btnAddRow_Click()
    Call AddTableRow
    lblStatus.Caption = "Row " & rowCount & " added"
End Sub

' ============================================================
' REMOVE ROW BUTTON CLICK EVENT
' ============================================================
Private Sub btnRemoveRow_Click()
    Call RemoveTableRow
End Sub

' ============================================================
' REMOVE TABLE ROW
' ============================================================
Private Sub RemoveTableRow()
    On Error GoTo RemoveError
    
    If rowCount <= 0 Then
        MsgBox "No rows to remove", vbExclamation
        Exit Sub
    End If
    
    Dim lastRowNum As Integer
    lastRowNum = rowCount
    
    ' Remove the controls for the last row
    Debug.Print "Removing row " & lastRowNum
    
    ' Remove event handler for this row's sign number box
    On Error Resume Next
    signNumberHandlers.Remove CStr(lastRowNum)
    On Error GoTo 0

    ' Remove controls in reverse order
    On Error Resume Next  ' Continue if a control doesn't exist
    frameSignTable.Controls.Remove "cboSide" & lastRowNum
    frameSignTable.Controls.Remove "txtSize" & lastRowNum
    frameSignTable.Controls.Remove "txtSpacing" & lastRowNum
    frameSignTable.Controls.Remove "txtSignNum" & lastRowNum
    On Error GoTo RemoveError  ' Resume error handling
    
    ' Clear array references before resizing
    Set signSideComboBoxes(lastRowNum) = Nothing
    Set signSizeBoxes(lastRowNum) = Nothing
    Set signSpacingBoxes(lastRowNum) = Nothing
    Set signNumberBoxes(lastRowNum) = Nothing
    
    ' Decrement row count
    rowCount = rowCount - 1
    
    ' Resize arrays to exclude the deleted row
    ReDim Preserve signNumberBoxes(1 To rowCount)
    ReDim Preserve signSpacingBoxes(1 To rowCount)
    ReDim Preserve signSizeBoxes(1 To rowCount)
    ReDim Preserve signSideComboBoxes(1 To rowCount)
    
    Debug.Print "Row removed successfully. Total rows: " & rowCount
    lblStatus.Caption = "Row " & lastRowNum & " removed. Total rows: " & rowCount
    Exit Sub

RemoveError:
    MsgBox "Error removing row at line " & Erl & ":" & vbCrLf & Err.Description, vbCritical, "Remove Row Error"
    Debug.Print "ERROR removing row: " & Err.Description & " at line " & Erl
End Sub

' ============================================================
' SIDE SELECTION - Uses ComboBox for better reliability
' ComboBoxes are more stable than option buttons in MicroStation VBA
' ============================================================

' ============================================================
' REFERENCE BUTTON CLICK EVENT - OPENS REFERENCE MUTCD FORM
' ============================================================
Private Sub btnReference_Click()
    SheetViewer.Show vbModal
End Sub

' ============================================================
' SUBMIT BUTTON CLICK EVENT
' Validates input and processes selection
' ============================================================
Private Sub btnSubmit_Click()
    ' Validate that selections are made
    If cboCategory.ListIndex <= 0 Then
        MsgBox "Please select a workzone category", vbExclamation
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

    If cboRoadType.ListIndex <= 0 Then
        MsgBox "Please select a road type (Freeway or Non-Freeway)", vbExclamation
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
    
    ' Save spacing & clearances to public storage
    wztcDownstreamTaper = frameSpacingValues.Controls("txtDownstreamTaper").Value
    wztcRollAhead = frameSpacingValues.Controls("txtRollAhead").Value
    wztcVehicleSpace = frameSpacingValues.Controls("txtVehicleSpace").Value
    wztcBufferSpace = frameSpacingValues.Controls("txtBufferSpace").Value
    wztcMergingTaper = frameSpacingValues.Controls("txtMergingTaper").Value
    wztcShoulderTapers = frameSpacingValues.Controls("txtShoulderTapers").Value
    wztcAdvancedWarningSpacing = frameSpacingValues.Controls("txtAdvancedWarningSpacing").Value
    wztcSkipLines = frameSpacingValues.Controls("txtSkipLines").Value
    wztcChannelizing = frameSpacingValues.Controls("txtChannelizing").Value
    wztcFlareBarrier = frameSpacingValues.Controls("txtFlareBarrier").Value
    wztcFlareBeam = frameSpacingValues.Controls("txtFlareBeam").Value

    ' Save user selections to public storage
    wztcCategory = cboCategory.Value
    wztcSheet = cboSheet.Value
    wztcSpeed = cboRoadSpeed.Value
    wztcRoadType = cboRoadType.Value
    wztcLaneWidth = cboLaneWidth.Value
    wztcShoulderWidth = cboShoulderWidth.Value

    ' Save sign table to public storage (sign number, spacing, size string, One Side/Both Sides)
    wztcSignCount = rowCount
    ReDim wztcSignNumbers(1 To rowCount)
    ReDim wztcSignSpacings(1 To rowCount)
    ReDim wztcSignSizes(1 To rowCount)
    ReDim wztcSignSides(1 To rowCount)
    For i = 1 To rowCount
        wztcSignNumbers(i) = signNumberBoxes(i).Value
        wztcSignSpacings(i) = signSpacingBoxes(i).Value
        wztcSignSizes(i) = signSizeBoxes(i).Value      ' string e.g. "48"" x 48"""
        wztcSignSides(i) = signSideComboBoxes(i).Value ' "One Side" or "Both Sides"
    Next i

    ' Store the WZTC Order table (all parameter labels and sign labels in current order)
    wztcOrderLabelCount = wztcOrderCount
    If wztcOrderCount > 0 Then
        ReDim wztcOrderLabels(0 To wztcOrderCount - 1)
        For i = 0 To wztcOrderCount - 1
            wztcOrderLabels(i) = wztcOrderTexts(i)
        Next i
    Else
        ReDim wztcOrderLabels(0 To -1)
    End If

    ' Close form and launch alignment drawing tool
    Unload Me
    StartWZTCDrawing

End Sub

' ============================================================
' DRAW WORKZONE TRAFFIC CONTROL
' Main routine to coordinate sign placement
' ============================================================
Private Sub DrawWorkzoneTrafficControl()
    Dim i As Integer
    Dim signNum As String
    Dim spacing As Double
    Dim bothSides As Boolean

    ' Process each sign in the table
    For i = 1 To rowCount
        signNum = signNumberBoxes(i).Value

        ' Skip empty rows
        If signNum = "" Then
            GoTo NextSign
        End If

        ' Get spacing and side information
        spacing = Val(signSpacingBoxes(i).Value)
        bothSides = (signSideComboBoxes(i).Value = "Both Sides")

        ' Log the sign information (placeholder for actual drawing logic)
        ' In a real implementation, this would place sign symbols in MicroStation
        Debug.Print "Sign: " & signNum & ", Spacing: " & spacing & ", Both Sides: " & bothSides

NextSign:
    Next i

End Sub

' ============================================================
' BUILD WZTC ORDER TABLE
' Collects spacing labels + non-empty sign numbers + Work Area
' ============================================================
Private Sub BuildWZTCOrderTable()
    If Not ControlExists("lstWZTCOrder") Then Exit Sub

    Dim i As Integer
    ReDim wztcOrderTexts(0 To 7 + rowCount)
    wztcOrderCount = 0

    ' Fixed spacing section labels (no values, just names)
    wztcOrderTexts(wztcOrderCount) = "Downstream Taper"
    wztcOrderCount = wztcOrderCount + 1
    wztcOrderTexts(wztcOrderCount) = "Roll Ahead Distance"
    wztcOrderCount = wztcOrderCount + 1
    wztcOrderTexts(wztcOrderCount) = "Vehicle Space"
    wztcOrderCount = wztcOrderCount + 1
    wztcOrderTexts(wztcOrderCount) = "Buffer Space"
    wztcOrderCount = wztcOrderCount + 1
    wztcOrderTexts(wztcOrderCount) = "Merging/Shifting Taper"
    wztcOrderCount = wztcOrderCount + 1
    wztcOrderTexts(wztcOrderCount) = "Shoulder Taper"
    wztcOrderCount = wztcOrderCount + 1

    ' Non-empty sign rows from sign selection table
    For i = 1 To rowCount
        If signNumberBoxes(i).Value <> "" Then
            wztcOrderTexts(wztcOrderCount) = signNumberBoxes(i).Value
            wztcOrderCount = wztcOrderCount + 1
        End If
    Next i

    ' Work Area (always present)
    wztcOrderTexts(wztcOrderCount) = "Work Area"
    wztcOrderCount = wztcOrderCount + 1

    If wztcOrderCount > 0 Then
        ReDim Preserve wztcOrderTexts(0 To wztcOrderCount - 1)
    End If

    Call RenderWZTCOrder
End Sub

' ============================================================
' RENDER WZTC ORDER INTO LISTBOX
' ============================================================
Private Sub RenderWZTCOrder()
    If Not ControlExists("lstWZTCOrder") Then Exit Sub
    Dim savedIdx As Integer
    savedIdx = lstWZTCOrder.ListIndex
    lstWZTCOrder.Clear
    Dim i As Integer
    For i = 0 To wztcOrderCount - 1
        lstWZTCOrder.AddItem wztcOrderTexts(i)
    Next i
    If savedIdx >= 0 And savedIdx < wztcOrderCount Then
        lstWZTCOrder.ListIndex = savedIdx
    End If
End Sub

' ============================================================
' WZTC ORDER - MOVE SELECTED ITEM UP
' ============================================================
Private Sub btnOrderUp_Click()
    If Not ControlExists("lstWZTCOrder") Then Exit Sub
    Dim idx As Integer
    idx = lstWZTCOrder.ListIndex
    If idx <= 0 Then Exit Sub
    Dim temp As String
    temp = wztcOrderTexts(idx)
    wztcOrderTexts(idx) = wztcOrderTexts(idx - 1)
    wztcOrderTexts(idx - 1) = temp
    Call RenderWZTCOrder
    lstWZTCOrder.ListIndex = idx - 1
End Sub

' ============================================================
' WZTC ORDER - MOVE SELECTED ITEM DOWN
' ============================================================
Private Sub btnOrderDown_Click()
    If Not ControlExists("lstWZTCOrder") Then Exit Sub
    Dim idx As Integer
    idx = lstWZTCOrder.ListIndex
    If idx < 0 Or idx >= wztcOrderCount - 1 Then Exit Sub
    Dim temp As String
    temp = wztcOrderTexts(idx)
    wztcOrderTexts(idx) = wztcOrderTexts(idx + 1)
    wztcOrderTexts(idx + 1) = temp
    Call RenderWZTCOrder
    lstWZTCOrder.ListIndex = idx + 1
End Sub

' ============================================================
' WZTC ORDER - DELETE SELECTED ITEM
' ============================================================
Private Sub btnOrderDelete_Click()
    If Not ControlExists("lstWZTCOrder") Then Exit Sub
    Dim idx As Integer
    idx = lstWZTCOrder.ListIndex
    If idx < 0 Or wztcOrderCount = 0 Then Exit Sub
    Dim i As Integer
    For i = idx To wztcOrderCount - 2
        wztcOrderTexts(i) = wztcOrderTexts(i + 1)
    Next i
    wztcOrderCount = wztcOrderCount - 1
    If wztcOrderCount > 0 Then
        ReDim Preserve wztcOrderTexts(0 To wztcOrderCount - 1)
    End If
    Call RenderWZTCOrder
    If wztcOrderCount > 0 Then
        If idx < wztcOrderCount Then
            lstWZTCOrder.ListIndex = idx
        Else
            lstWZTCOrder.ListIndex = wztcOrderCount - 1
        End If
    End If
End Sub

' ============================================================
' WZTC ORDER - REFRESH FROM CURRENT SIGN TABLE
' ============================================================
Private Sub btnRefreshOrder_Click()
    Call BuildWZTCOrderTable
End Sub







