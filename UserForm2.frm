VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   14010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21405
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' MICROSTATION VBA - WORKZONE DESIGN TOOL
' Version 1.2 - Scrollable table, option buttons, resizable
' ============================================================
'
' UserForm Controls needed:
' - UserForm1 (Form) - SET ShowModal = False for resizing
' - cboCategory (ComboBox) - workzone category selection
' - cboSheet (ComboBox) - specific sheet selection
' - frameSignTable (Frame) - container for sign table
' - lblCategory (Label) - "Workzone Category:"
' - lblSheet (Label) - "Standard Sheet:"
' - btnAddRow (CommandButton) - "+" button to add rows
' - btnReference (CommandButton) - "Reference" button
' - lblStatus (Label) - status display
'
' IMPORTANT: Set frameSignTable ScrollBars property to:
'   fmScrollBarsVertical (2) or fmScrollBarsBoth (3)
' ============================================================

Option Explicit

' Table control arrays
Private signNumberBoxes() As MSForms.TextBox
Private signSpacingBoxes() As MSForms.TextBox
Private signWidthBoxes() As MSForms.TextBox
Private signHeightBoxes() As MSForms.TextBox
Private signSideFrames() As MSForms.Frame
Private signOptionOne() As MSForms.OptionButton
Private signOptionBoth() As MSForms.OptionButton
Private rowCount As Integer

' Constants for table layout
Private Const TABLE_START_TOP As Integer = 45
Private Const TABLE_LEFT As Integer = 10
Private Const COL1_WIDTH As Integer = 90   ' Sign Number
Private Const COL2_WIDTH As Integer = 80   ' Spacing
Private Const COL3_WIDTH As Integer = 70   ' Width
Private Const COL4_WIDTH As Integer = 70   ' Height
Private Const COL5_WIDTH As Integer = 130  ' Side selection (option buttons)
Private Const ROW_HEIGHT As Integer = 30
Private Const INITIAL_ROWS As Integer = 10

Private Sub UserForm_Initialize()
    Me.Caption = "Workzone Design Tool"
    Me.Width = 620
    Me.Height = 650
    
    ' Make form resizable (if ShowModal = False)
    ' Note: In VBA, forms shown with .Show are resizable by default when ShowModal = False
    
    ' Initialize dropdown position and size
    lblCategory.Caption = "Workzone Category:"
    lblCategory.Top = 10
    lblCategory.Left = 20
    lblCategory.Width = 120
    
    cboCategory.Top = 10
    cboCategory.Left = 145
    cboCategory.Width = 450
    
    lblSheet.Caption = "Standard Sheet:"
    lblSheet.Top = 40
    lblSheet.Left = 20
    lblSheet.Width = 120
    
    cboSheet.Top = 40
    cboSheet.Left = 145
    cboSheet.Width = 450
    
    ' Initialize frame for table - FIXED SIZE with scrollbars
    frameSignTable.Top = 75
    frameSignTable.Left = 10
    frameSignTable.Width = 590
    frameSignTable.Height = 420  ' Fixed height
    frameSignTable.Caption = "Sign Selection Table"
    frameSignTable.ScrollBars = fmScrollBarsVertical  ' Enable vertical scrolling
    frameSignTable.ScrollHeight = 1000  ' Will be adjusted as rows are added
    frameSignTable.KeepScrollBarsVisible = fmScrollBarsVertical
    
    ' Position Add Row button - FIXED POSITION (won't move)
    btnAddRow.Top = 505
    btnAddRow.Left = 20
    btnAddRow.Width = 40
    btnAddRow.Height = 25
    btnAddRow.Caption = "+"
    btnAddRow.Font.Size = 12
    btnAddRow.Font.Bold = True
    
    ' Position Reference button - FIXED POSITION (won't move)
    btnReference.Top = 505
    btnReference.Left = 70
    btnReference.Width = 100
    btnReference.Height = 25
    btnReference.Caption = "Reference"
    
    ' Position status label - FIXED POSITION (won't move)
    lblStatus.Top = 540
    lblStatus.Left = 20
    lblStatus.Width = 570
    lblStatus.Caption = "Ready - Select workzone category and sheet"
    
    ' Populate categories
    Call PopulateCategories
    
    ' Initialize table
    rowCount = 0
    Call CreateTableHeaders
    Call AddInitialRows
End Sub

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
        lblStatus.Caption = "Category selected - Choose a standard sheet"
    End If
End Sub

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
        lblStatus.Caption = "Ready to configure signs for: " & cboSheet.List(cboSheet.ListIndex)
    End If
End Sub

Private Sub CreateTableHeaders()
    Dim lblSignNum As MSForms.Label
    Dim lblSpacing As MSForms.Label
    Dim lblWidth As MSForms.Label
    Dim lblHeight As MSForms.Label
    Dim lblSide As MSForms.Label
    
    ' Create header labels
    Set lblSignNum = frameSignTable.Controls.Add("Forms.Label.1", "lblHeaderSignNum")
    With lblSignNum
        .Caption = "Sign Number"
        .Top = 10
        .Left = TABLE_LEFT
        .Width = COL1_WIDTH
        .Height = 25
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackColor = &H8000000F  ' Light gray
        .BorderStyle = fmBorderStyleSingle
    End With
    
    Set lblSpacing = frameSignTable.Controls.Add("Forms.Label.1", "lblHeaderSpacing")
    With lblSpacing
        .Caption = "Spacing (ft)"
        .Top = 10
        .Left = TABLE_LEFT + COL1_WIDTH + 5
        .Width = COL2_WIDTH
        .Height = 25
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackColor = &H8000000F
        .BorderStyle = fmBorderStyleSingle
    End With
    
    Set lblWidth = frameSignTable.Controls.Add("Forms.Label.1", "lblHeaderWidth")
    With lblWidth
        .Caption = "Width (in)"
        .Top = 10
        .Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + 10
        .Width = COL3_WIDTH
        .Height = 25
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackColor = &H8000000F
        .BorderStyle = fmBorderStyleSingle
    End With
    
    Set lblHeight = frameSignTable.Controls.Add("Forms.Label.1", "lblHeaderHeight")
    With lblHeight
        .Caption = "Height (in)"
        .Top = 10
        .Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + COL3_WIDTH + 15
        .Width = COL4_WIDTH
        .Height = 25
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackColor = &H8000000F
        .BorderStyle = fmBorderStyleSingle
    End With
    
    Set lblSide = frameSignTable.Controls.Add("Forms.Label.1", "lblHeaderSide")
    With lblSide
        .Caption = "Road Side"
        .Top = 10
        .Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + COL3_WIDTH + COL4_WIDTH + 20
        .Width = COL5_WIDTH
        .Height = 25
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackColor = &H8000000F
        .BorderStyle = fmBorderStyleSingle
    End With
End Sub

Private Sub AddInitialRows()
    Dim i As Integer
    For i = 1 To INITIAL_ROWS
        Call AddTableRow
    Next i
End Sub

Private Sub AddTableRow()
    Dim txtSignNum As MSForms.TextBox
    Dim txtSpacing As MSForms.TextBox
    Dim txtWidth As MSForms.TextBox
    Dim txtHeight As MSForms.TextBox
    Dim frameSide As MSForms.Frame
    Dim optOne As MSForms.OptionButton
    Dim optBoth As MSForms.OptionButton
    Dim currentTop As Integer
    
    rowCount = rowCount + 1
    currentTop = TABLE_START_TOP + (rowCount - 1) * ROW_HEIGHT
    
    ' Resize arrays
    ReDim Preserve signNumberBoxes(1 To rowCount)
    ReDim Preserve signSpacingBoxes(1 To rowCount)
    ReDim Preserve signWidthBoxes(1 To rowCount)
    ReDim Preserve signHeightBoxes(1 To rowCount)
    ReDim Preserve signSideFrames(1 To rowCount)
    ReDim Preserve signOptionOne(1 To rowCount)
    ReDim Preserve signOptionBoth(1 To rowCount)
    
    ' Sign Number textbox
    Set txtSignNum = frameSignTable.Controls.Add("Forms.TextBox.1", "txtSignNum" & rowCount)
    With txtSignNum
        .Top = currentTop
        .Left = TABLE_LEFT
        .Width = COL1_WIDTH
        .Height = 22
        .Text = ""
    End With
    Set signNumberBoxes(rowCount) = txtSignNum
    
    ' Spacing textbox
    Set txtSpacing = frameSignTable.Controls.Add("Forms.TextBox.1", "txtSpacing" & rowCount)
    With txtSpacing
        .Top = currentTop
        .Left = TABLE_LEFT + COL1_WIDTH + 5
        .Width = COL2_WIDTH
        .Height = 22
        .Text = ""
    End With
    Set signSpacingBoxes(rowCount) = txtSpacing
    
    ' Width textbox
    Set txtWidth = frameSignTable.Controls.Add("Forms.TextBox.1", "txtWidth" & rowCount)
    With txtWidth
        .Top = currentTop
        .Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + 10
        .Width = COL3_WIDTH
        .Height = 22
        .Text = ""
    End With
    Set signWidthBoxes(rowCount) = txtWidth
    
    ' Height textbox
    Set txtHeight = frameSignTable.Controls.Add("Forms.TextBox.1", "txtHeight" & rowCount)
    With txtHeight
        .Top = currentTop
        .Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + COL3_WIDTH + 15
        .Width = COL4_WIDTH
        .Height = 22
        .Text = ""
    End With
    Set signHeightBoxes(rowCount) = txtHeight
    
    ' Frame for option buttons (contains the two options)
    Set frameSide = frameSignTable.Controls.Add("Forms.Frame.1", "frameSide" & rowCount)
    With frameSide
        .Top = currentTop
        .Left = TABLE_LEFT + COL1_WIDTH + COL2_WIDTH + COL3_WIDTH + COL4_WIDTH + 20
        .Width = COL5_WIDTH
        .Height = 24
        .Caption = ""
        .BorderStyle = fmBorderStyleNone
    End With
    Set signSideFrames(rowCount) = frameSide
    
    ' Option button for "One Side"
    Set optOne = frameSide.Controls.Add("Forms.OptionButton.1", "optOne" & rowCount)
    With optOne
        .Caption = "One Side"
        .Top = 2
        .Left = 5
        .Width = 60
        .Height = 18
        .GroupName = "SideGroup" & rowCount  ' Each row has its own group
        .Value = True  ' Default selection
    End With
    Set signOptionOne(rowCount) = optOne
    
    ' Option button for "Both Sides"
    Set optBoth = frameSide.Controls.Add("Forms.OptionButton.1", "optBoth" & rowCount)
    With optBoth
        .Caption = "Both Sides"
        .Top = 2
        .Left = 68
        .Width = 60
        .Height = 18
        .GroupName = "SideGroup" & rowCount  ' Same group as optOne
        .Value = False
    End With
    Set signOptionBoth(rowCount) = optBoth
    
    ' Update the scroll height of the frame to accommodate all rows
    frameSignTable.ScrollHeight = currentTop + 50
End Sub

Private Sub btnAddRow_Click()
    Call AddTableRow
    lblStatus.Caption = "Row " & rowCount & " added (scroll down to see it)"
    
    ' Scroll to the bottom to show the new row
    frameSignTable.Scroll 0, frameSignTable.ScrollHeight
End Sub

Private Sub btnReference_Click()
    Dim refData As String
    Dim i As Integer
    Dim signCount As Integer
    Dim sideText As String
    
    ' Validate that category and sheet are selected
    If cboCategory.ListIndex < 0 Or cboSheet.ListIndex < 0 Then
        MsgBox "Please select both a Workzone Category and Standard Sheet before creating reference.", vbExclamation, "Selection Required"
        Exit Sub
    End If
    
    ' Count filled rows
    signCount = 0
    For i = 1 To rowCount
        If signNumberBoxes(i).Text <> "" Then
            signCount = signCount + 1
        End If
    Next i
    
    If signCount = 0 Then
        MsgBox "Please enter at least one sign in the table before creating reference.", vbExclamation, "No Signs Entered"
        Exit Sub
    End If
    
    ' Build reference data string
    refData = "WORKZONE DESIGN REFERENCE" & vbCrLf & vbCrLf
    refData = refData & "Category: " & cboCategory.List(cboCategory.ListIndex) & vbCrLf
    refData = refData & "Sheet: " & cboSheet.List(cboSheet.ListIndex) & vbCrLf & vbCrLf
    refData = refData & "SIGN CONFIGURATION:" & vbCrLf
    refData = refData & String(80, "-") & vbCrLf
    
    For i = 1 To rowCount
        If signNumberBoxes(i).Text <> "" Then
            ' Determine which option is selected
            If signOptionBoth(i).Value = True Then
                sideText = "Both Sides"
            Else
                sideText = "One Side"
            End If
            
            refData = refData & "Sign #" & i & ": "
            refData = refData & signNumberBoxes(i).Text & " | "
            refData = refData & "Spacing: " & signSpacingBoxes(i).Text & " ft | "
            refData = refData & "Size: " & signWidthBoxes(i).Text & "W x " & signHeightBoxes(i).Text & "H | "
            refData = refData & "Side: " & sideText
            refData = refData & vbCrLf
        End If
    Next i
    
    ' Show the reference
    MsgBox refData, vbInformation, "Workzone Reference Created"
    
    lblStatus.Caption = signCount & " sign(s) referenced successfully"
    
    ' Optional: Export to file or place in MicroStation drawing
    Call ExportToMicroStation(refData)
End Sub

Private Sub ExportToMicroStation(refData As String)
    ' This is where you would integrate with MicroStation API
    ' to place text elements, create reference files, etc.
    
    ' Example placeholder for MicroStation integration:
    ' Dim app As MicroStationDGN.Application
    ' Dim activeModel As Model
    ' Set app = Application
    ' Set activeModel = app.ActiveModelReference
    
    ' Add your MicroStation-specific code here to:
    ' - Create text elements
    ' - Place sign symbols at specified locations
    ' - Create reference attachments
    ' - Generate plan sheets
    
    Debug.Print "Reference data ready for MicroStation:"
    Debug.Print refData
End Sub

