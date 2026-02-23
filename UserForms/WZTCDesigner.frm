Option Explicit

' ============================================================
' ALIGNMENT TABLE DATA
' Each alignment (1 to MAX_ALIGNMENTS) has its own scrollable
' frame containing rows with Type / Label / Spacing / Size / Side.
' ============================================================
Private Const MAX_ALIGNMENTS As Integer = 10
Private Const MAX_ALIGN_ROWS  As Integer = 50

Private alignTypeBoxes(1 To 10, 1 To 50)    As MSForms.ComboBox
Private alignLabelBoxes(1 To 10, 1 To 50)   As MSForms.TextBox
Private alignSpacingBoxes(1 To 10, 1 To 50) As MSForms.TextBox
Private alignSizeBoxes(1 To 10, 1 To 50)    As MSForms.TextBox
Private alignSideBoxes(1 To 10, 1 To 50)    As MSForms.ComboBox

Private alignRowCounts(1 To 10)     As Integer
Private alignCount                  As Integer
Private alignFrames(1 To 10)        As Object    ' Frame objects (created in code)
Private alignSectionLabels(1 To 10) As Object    ' Section title Label objects

Private activeAlignIdx As Integer   ' 0 = no row selected
Private activeRowIdx   As Integer   ' 0 = no row selected
Private inSelectAlignRow As Boolean  ' recursion guard for SelectAlignRow

Private alignRowHandlers(1 To 10)    As Collection  ' AlignRowBox instances per alignment
Private alignTypeHandlers(1 To 10)  As Collection  ' AlignTypeBox instances per alignment
Private alignAddBtnHandlers(1 To 10) As Object     ' AlignSectionBtns instances per alignment
Private alignRowSelBoxes(1 To 10)    As Object     ' Row selector ComboBoxes per alignment
Private spacingBoxHandlers           As Collection  ' SpacingBox instances

' Cached spacing values from GenerateSpacingTable — used by AutoPopulateWZTCItems
Private cachedDownstreamTaper  As Double
Private cachedRollAhead        As Double
Private cachedVehicleSpace     As Double
Private cachedBufferSpace      As Double
Private cachedMergingTaper     As Double
Private cachedShoulderTapers   As Double
Private cachedUpTaperBarrier   As Double
Private cachedUpTaperBeam      As Double
Private hasSpacingData         As Boolean

' Status tracking
Private selectedSpeed    As String
Private selectedRoadType As String

' ---- Column layout constants (inside alignment Frame) ----
Private Const ALN_FRAME_LEFT    As Integer = 440   ' right column start on form
Private Const ALN_FRAME_WIDTH   As Integer = 540   ' frame visible width
Private Const ALN_FRAME_VIS_H   As Integer = 210   ' frame visible height (scrolls internally)
Private Const ALN_SECTION_TOP0  As Integer = 175   ' top of first alignment section on form
Private Const ALN_SECTION_H     As Integer = 275   ' vertical space per alignment section (includes footer buttons)
Private Const ALN_COL_TYPE_L    As Integer = 10
Private Const ALN_COL_TYPE_W    As Integer = 75
Private Const ALN_COL_LABEL_L   As Integer = 90
Private Const ALN_COL_LABEL_W   As Integer = 110
Private Const ALN_COL_SPACE_L   As Integer = 205
Private Const ALN_COL_SPACE_W   As Integer = 80
Private Const ALN_COL_SIZE_L    As Integer = 290
Private Const ALN_COL_SIZE_W    As Integer = 70
Private Const ALN_COL_SIDE_L    As Integer = 365
Private Const ALN_COL_SIDE_W    As Integer = 110
Private Const ALN_HDR_TOP       As Integer = 15
Private Const ALN_ROW_START_TOP As Integer = 40
Private Const ALN_ROW_H         As Integer = 25
Private Const HIGHLIGHT_COLOR   As Long = 13166847  ' &HC8E8FF light blue

' ============================================================
' INITIALIZE FORM
' ============================================================
Private Sub UserForm_Initialize()
    On Error GoTo InitError

    Debug.Print "WZTCDesigner: Starting Initialize..."

    Me.Caption = "Workzone Traffic Control Designer - MUTCD NY"
    Me.Width = 1030
    Me.Height = 740
    Set spacingBoxHandlers = New Collection

    ' ========== LEFT COLUMN: INPUT DROPDOWNS ==========
    If ControlExists("lblCategory") Then
        lblCategory.Caption = "Workzone Category:"
        lblCategory.Top = 10: lblCategory.Left = 20
        lblCategory.Width = 120: lblCategory.Font.Bold = True
    End If
    If ControlExists("cboCategory") Then
        cboCategory.Top = 10: cboCategory.Left = 150: cboCategory.Width = 250
        Call PopulateCategories
    End If

    If ControlExists("lblSheet") Then
        lblSheet.Caption = "Standard Sheet Number:"
        lblSheet.Top = 40: lblSheet.Left = 20
        lblSheet.Width = 120: lblSheet.Font.Bold = True
    End If
    If ControlExists("cboSheet") Then
        cboSheet.Top = 40: cboSheet.Left = 150: cboSheet.Width = 250
    End If

    If ControlExists("lblRoadSpeed") Then
        lblRoadSpeed.Caption = "Road Speed (mph):"
        lblRoadSpeed.Top = 70: lblRoadSpeed.Left = 20
        lblRoadSpeed.Width = 120: lblRoadSpeed.Font.Bold = True
    End If
    If ControlExists("cboRoadSpeed") Then
        cboRoadSpeed.Top = 70: cboRoadSpeed.Left = 150: cboRoadSpeed.Width = 250
        Call PopulateRoadSpeeds
    End If

    If ControlExists("lblRoadType") Then
        lblRoadType.Caption = "Road Type:"
        lblRoadType.Top = 100: lblRoadType.Left = 20
        lblRoadType.Width = 120: lblRoadType.Font.Bold = True
    End If
    If ControlExists("cboRoadType") Then
        cboRoadType.Top = 100: cboRoadType.Left = 150: cboRoadType.Width = 250
        Call PopulateRoadType
    End If

    If ControlExists("lblLaneWidth") Then
        lblLaneWidth.Caption = "Lane Width (ft):"
        lblLaneWidth.Top = 130: lblLaneWidth.Left = 20
        lblLaneWidth.Width = 120: lblLaneWidth.Font.Bold = True
    End If
    If ControlExists("cboLaneWidth") Then
        cboLaneWidth.Top = 130: cboLaneWidth.Left = 150: cboLaneWidth.Width = 250
        Call PopulateLaneWidth
    End If

    If ControlExists("lblShoulderWidth") Then
        lblShoulderWidth.Caption = "Shoulder Width (ft):"
        lblShoulderWidth.Top = 160: lblShoulderWidth.Left = 20
        lblShoulderWidth.Width = 120: lblShoulderWidth.Font.Bold = True
    End If
    If ControlExists("cboShoulderWidth") Then
        cboShoulderWidth.Top = 160: cboShoulderWidth.Left = 150: cboShoulderWidth.Width = 250
        Call PopulateShoulderWidth
    End If

    ' ========== SPACING & CLEARANCES SECTION ==========
    If ControlExists("frameSpacingValues") Then
        frameSpacingValues.Caption = "Calculated Recommended Spacing & Clearances (per MUTCD NY)"
        frameSpacingValues.Top = 195: frameSpacingValues.Left = 10
        frameSpacingValues.Width = 280: frameSpacingValues.Height = 360
        Call CreateSpacingLabels
    End If

    ' ========== RIGHT COLUMN: GLOBAL ROW BUTTONS (hidden — now per-section below each table) ==========
    If ControlExists("lblRowActions") Then  lblRowActions.Visible = False  End If
    If ControlExists("btnAddRow") Then      btnAddRow.Visible = False      End If
    If ControlExists("btnDeleteRow") Then   btnDeleteRow.Visible = False   End If
    If ControlExists("btnMoveUp") Then      btnMoveUp.Visible = False      End If
    If ControlExists("btnMoveDown") Then    btnMoveDown.Visible = False    End If

    ' ========== ALIGNMENT MANAGEMENT BUTTONS ==========
    If ControlExists("lblAlignActions") Then
        lblAlignActions.Caption = "Alignments:"
        lblAlignActions.Top = 10: lblAlignActions.Left = 440
        lblAlignActions.Width = 110: lblAlignActions.Font.Bold = True
    End If
    If ControlExists("btnAddAlignment") Then
        btnAddAlignment.Caption = "Add Alignment +"
        btnAddAlignment.Top = 30: btnAddAlignment.Left = 440
        btnAddAlignment.Width = 115: btnAddAlignment.Height = 22
        btnAddAlignment.Font.Size = 8
    End If
    If ControlExists("btnRemoveAlignment") Then
        btnRemoveAlignment.Caption = "Remove Alignment"
        btnRemoveAlignment.Top = 30: btnRemoveAlignment.Left = 560
        btnRemoveAlignment.Width = 125: btnRemoveAlignment.Height = 22
        btnRemoveAlignment.Font.Size = 8
    End If

    ' ========== ACTION BUTTONS (top-right) ==========
    If ControlExists("btnReference") Then
        btnReference.Caption = "Reference (MUTCD)"
        btnReference.Top = 10: btnReference.Left = 875
        btnReference.Width = 130: btnReference.Height = 22
    End If
    If ControlExists("btnSubmit") Then
        btnSubmit.Caption = "Submit & Draw"
        btnSubmit.Top = 38: btnSubmit.Left = 875
        btnSubmit.Width = 130: btnSubmit.Height = 22
        btnSubmit.Font.Bold = True
    End If
    If ControlExists("btnClear") Then
        btnClear.Caption = "Clear All"
        btnClear.Top = 62: btnClear.Left = 875
        btnClear.Width = 130: btnClear.Height = 22
    End If

    ' ========== STATUS LABEL ==========
    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Select road parameters, then use the buttons below each alignment table to add and arrange rows."
        lblStatus.Top = 60: lblStatus.Left = 440
        lblStatus.Width = 425: lblStatus.Height = 100
        lblStatus.Font.Size = 9: lblStatus.WordWrap = True
    End If

    ' ========== INITIALIZE ALIGNMENT TABLES ==========
    alignCount = 0
    activeAlignIdx = 0: activeRowIdx = 0
    hasSpacingData = False
    Call InitAlignments

    Debug.Print "WZTCDesigner: Initialize complete"
    Call RestoreState
    Exit Sub

InitError:
    MsgBox "Error initializing form at line: " & Erl & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & "Number: " & Err.Number, _
           vbCritical, "Initialization Error"
    Debug.Print "CRASH at line: " & Erl & " - " & Err.Description
End Sub

' ============================================================
' CONTROL EXISTS HELPER
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

' ============================================================
' INITIALIZE ALIGNMENTS — creates Upstream and Downstream sections
' ============================================================
Private Sub InitAlignments()
    Call CreateAlignSection(1, "Upstream Alignment")
    Call CreateAlignSection(2, "Downstream Alignment")
    alignCount = 2
    Call UpdateFormScrollHeight
End Sub

' ============================================================
' ALIGNMENT SECTION TOP POSITION
' ============================================================
Private Function AlignSectionTop(aIdx As Integer) As Integer
    AlignSectionTop = ALN_SECTION_TOP0 + (aIdx - 1) * ALN_SECTION_H
End Function

' ============================================================
' CREATE ALIGNMENT SECTION (title label + scrollable frame + headers)
' ============================================================
Private Sub CreateAlignSection(aIdx As Integer, sectionName As String)
    On Error GoTo CreateError
    Dim topPos As Integer: topPos = AlignSectionTop(aIdx)

    Set alignRowHandlers(aIdx) = New Collection
    Set alignTypeHandlers(aIdx) = New Collection
    alignRowCounts(aIdx) = 0

    ' Section title label on the form
    Dim lblTitle As MSForms.Label
    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblAlign_" & aIdx)
    With lblTitle
        .Caption = sectionName
        .Top = topPos: .Left = ALN_FRAME_LEFT
        .Width = 220: .Height = 18
        .Font.Bold = True: .Font.Size = 9
    End With
    Set alignSectionLabels(aIdx) = lblTitle

    ' Scrollable frame
    Dim frm As Object
    Set frm = Me.Controls.Add("Forms.Frame.1", "frAlign_" & aIdx)
    With frm
        .Caption = ""
        .Top = topPos + 20: .Left = ALN_FRAME_LEFT
        .Width = ALN_FRAME_WIDTH: .Height = ALN_FRAME_VIS_H
        .ScrollBars = fmScrollBarsVertical
        .KeepScrollBarsVisible = fmScrollBarsVertical
        .ScrollHeight = 500
    End With
    Set alignFrames(aIdx) = frm

    Call CreateAlignHeaders(aIdx)

    ' Per-section row action buttons in footer below frame
    Dim footerTop As Integer: footerTop = topPos + 20 + ALN_FRAME_VIS_H + 4
    Dim secBtns As AlignSectionBtns
    Set secBtns = New AlignSectionBtns
    secBtns.AlignIdx = aIdx
    Set secBtns.ParentForm = Me

    Dim bAdd As MSForms.CommandButton
    Set bAdd = Me.Controls.Add("Forms.CommandButton.1", "btnSecAdd_" & aIdx)
    With bAdd: .Caption = "+ Add Row": .Top = footerTop: .Left = ALN_FRAME_LEFT
               .Width = 85: .Height = 22: .Font.Size = 7: End With
    Set secBtns.BtnAdd = bAdd

    Dim bDel As MSForms.CommandButton
    Set bDel = Me.Controls.Add("Forms.CommandButton.1", "btnSecDel_" & aIdx)
    With bDel: .Caption = "Del Row": .Top = footerTop: .Left = ALN_FRAME_LEFT + 88
               .Width = 65: .Height = 22: .Font.Size = 7: End With
    Set secBtns.BtnDel = bDel

    ' Row # selector — user picks which row to move up/down
    Dim lblRowSel As MSForms.Label
    Set lblRowSel = Me.Controls.Add("Forms.Label.1", "lblRowSel_" & aIdx)
    With lblRowSel: .Caption = "Row #:": .Top = footerTop + 4: .Left = ALN_FRAME_LEFT + 157
                    .Width = 38: .Height = 16: .Font.Size = 7: End With
    Dim cboRowSel As MSForms.ComboBox
    Set cboRowSel = Me.Controls.Add("Forms.ComboBox.1", "cboRowSel_" & aIdx)
    With cboRowSel: .Top = footerTop: .Left = ALN_FRAME_LEFT + 198: .Width = 42: .Height = 22
                    .Style = fmStyleDropDownList: .Font.Size = 8: End With
    Set alignRowSelBoxes(aIdx) = cboRowSel
    Set secBtns.CboRowSel = cboRowSel

    Dim bUp As MSForms.CommandButton
    Set bUp = Me.Controls.Add("Forms.CommandButton.1", "btnSecUp_" & aIdx)
    With bUp: .Caption = "Move Up": .Top = footerTop: .Left = ALN_FRAME_LEFT + 243
              .Width = 68: .Height = 22: .Font.Size = 7: End With
    Set secBtns.BtnUp = bUp

    Dim bDown As MSForms.CommandButton
    Set bDown = Me.Controls.Add("Forms.CommandButton.1", "btnSecDown_" & aIdx)
    With bDown: .Caption = "Move Down": .Top = footerTop: .Left = ALN_FRAME_LEFT + 314
                .Width = 75: .Height = 22: .Font.Size = 7: End With
    Set secBtns.BtnDown = bDown

    Set alignAddBtnHandlers(aIdx) = secBtns
    Exit Sub

CreateError:
    Debug.Print "CreateAlignSection error aIdx=" & aIdx & ": " & Err.Description
End Sub

' ============================================================
' CREATE COLUMN HEADERS INSIDE ALIGNMENT FRAME
' ============================================================
Private Sub CreateAlignHeaders(aIdx As Integer)
    Dim frm As Object: Set frm = alignFrames(aIdx)
    Dim lbl As MSForms.Label

    Set lbl = frm.Controls.Add("Forms.Label.1", "lblHdrType_" & aIdx)
    With lbl: .Caption = "Type": .Top = ALN_HDR_TOP: .Left = ALN_COL_TYPE_L
        .Width = ALN_COL_TYPE_W: .Height = 18: .Font.Bold = True: .BackColor = &H8000000F: End With

    Set lbl = frm.Controls.Add("Forms.Label.1", "lblHdrLabel_" & aIdx)
    With lbl: .Caption = "Sign # / Description": .Top = ALN_HDR_TOP: .Left = ALN_COL_LABEL_L
        .Width = ALN_COL_LABEL_W: .Height = 18: .Font.Bold = True: .BackColor = &H8000000F: End With

    Set lbl = frm.Controls.Add("Forms.Label.1", "lblHdrSpacing_" & aIdx)
    With lbl: .Caption = "Spacing (ft)": .Top = ALN_HDR_TOP: .Left = ALN_COL_SPACE_L
        .Width = ALN_COL_SPACE_W: .Height = 18: .Font.Bold = True: .BackColor = &H8000000F: End With

    Set lbl = frm.Controls.Add("Forms.Label.1", "lblHdrSize_" & aIdx)
    With lbl: .Caption = "Size (Signs)": .Top = ALN_HDR_TOP: .Left = ALN_COL_SIZE_L
        .Width = ALN_COL_SIZE_W: .Height = 18: .Font.Bold = True: .BackColor = &H8000000F: End With

    Set lbl = frm.Controls.Add("Forms.Label.1", "lblHdrSide_" & aIdx)
    With lbl: .Caption = "Road Side": .Top = ALN_HDR_TOP: .Left = ALN_COL_SIDE_L
        .Width = ALN_COL_SIDE_W: .Height = 18: .Font.Bold = True: .BackColor = &H8000000F: End With
End Sub

' ============================================================
' ADD A BLANK SIGN ROW TO AN ALIGNMENT
' ============================================================
Private Sub AddAlignRow(aIdx As Integer)
    Call AddAlignRowWithData(aIdx, "Sign", "", "", "", "One Side")
End Sub

' ============================================================
' ADD A ROW WITH SPECIFIED DATA TO AN ALIGNMENT
' ============================================================
Private Sub AddAlignRowWithData(aIdx As Integer, rowType As String, rowLabel As String, _
                                 rowSpacing As String, rowSize As String, rowSide As String)
    On Error GoTo RowError
    If aIdx < 1 Or aIdx > MAX_ALIGNMENTS Then Exit Sub
    If alignRowCounts(aIdx) >= MAX_ALIGN_ROWS Then
        MsgBox "Maximum " & MAX_ALIGN_ROWS & " rows per alignment.", vbExclamation: Exit Sub
    End If

    alignRowCounts(aIdx) = alignRowCounts(aIdx) + 1
    Dim rIdx As Integer: rIdx = alignRowCounts(aIdx)
    Dim frm As Object: Set frm = alignFrames(aIdx)
    Dim rowTop As Integer: rowTop = ALN_ROW_START_TOP + (rIdx - 1) * ALN_ROW_H

    ' --- Type ComboBox ---
    Dim cboType As MSForms.ComboBox
    Set cboType = frm.Controls.Add("Forms.ComboBox.1", "cboType_" & aIdx & "_" & rIdx)
    With cboType
        .Top = rowTop: .Left = ALN_COL_TYPE_L: .Width = ALN_COL_TYPE_W: .Height = 20
        .AddItem "Sign": .AddItem "Non-Sign"
        .Style = fmStyleDropDownList
        If rowType = "Non-Sign" Then .ListIndex = 1 Else .ListIndex = 0
        .BackColor = &HFFFFFF
    End With
    Set alignTypeBoxes(aIdx, rIdx) = cboType

    Dim typeHandler As AlignTypeBox
    Set typeHandler = New AlignTypeBox
    typeHandler.AlignIdx = aIdx: typeHandler.RowIdx = rIdx
    Set typeHandler.Cbo = cboType: Set typeHandler.ParentForm = Me
    alignTypeHandlers(aIdx).Add typeHandler, CStr(rIdx)

    ' --- Label/Sign# TextBox ---
    Dim txtLabel As MSForms.TextBox
    Set txtLabel = frm.Controls.Add("Forms.TextBox.1", "txtLabel_" & aIdx & "_" & rIdx)
    With txtLabel
        .Top = rowTop: .Left = ALN_COL_LABEL_L: .Width = ALN_COL_LABEL_W: .Height = 20
        .Text = rowLabel: .BackColor = &HFFFFFF
    End With
    Set alignLabelBoxes(aIdx, rIdx) = txtLabel

    Dim rowHandler As AlignRowBox
    Set rowHandler = New AlignRowBox
    rowHandler.AlignIdx = aIdx: rowHandler.RowIdx = rIdx
    Set rowHandler.Txt = txtLabel: Set rowHandler.ParentForm = Me
    alignRowHandlers(aIdx).Add rowHandler, CStr(rIdx)

    ' --- Spacing TextBox ---
    Dim txtSpacing As MSForms.TextBox
    Set txtSpacing = frm.Controls.Add("Forms.TextBox.1", "txtSpacing_" & aIdx & "_" & rIdx)
    With txtSpacing
        .Top = rowTop: .Left = ALN_COL_SPACE_L: .Width = ALN_COL_SPACE_W: .Height = 20
        .Text = rowSpacing: .BackColor = &HFFFFFF
    End With
    Set alignSpacingBoxes(aIdx, rIdx) = txtSpacing

    ' --- Size TextBox ---
    Dim txtSize As MSForms.TextBox
    Set txtSize = frm.Controls.Add("Forms.TextBox.1", "txtSize_" & aIdx & "_" & rIdx)
    With txtSize
        .Top = rowTop: .Left = ALN_COL_SIZE_L: .Width = ALN_COL_SIZE_W: .Height = 20
        .Text = rowSize: .BackColor = &HFFFFFF
    End With
    Set alignSizeBoxes(aIdx, rIdx) = txtSize

    ' --- Side ComboBox ---
    Dim cboSide As MSForms.ComboBox
    Set cboSide = frm.Controls.Add("Forms.ComboBox.1", "cboSide_" & aIdx & "_" & rIdx)
    With cboSide
        .Top = rowTop: .Left = ALN_COL_SIDE_L: .Width = ALN_COL_SIDE_W: .Height = 20
        .AddItem "One Side": .AddItem "Both Sides"
        .Style = fmStyleDropDownList
        If rowSide = "Both Sides" Then .ListIndex = 1 Else .ListIndex = 0
        .BackColor = &HFFFFFF
    End With
    Set alignSideBoxes(aIdx, rIdx) = cboSide

    ' Apply Non-Sign disabled state
    If rowType = "Non-Sign" Then
        txtSize.Enabled = False:  txtSize.BackColor = &HE0E0E0
        cboSide.Enabled = False: cboSide.BackColor = &HE0E0E0
    End If

    ' Update frame scroll height so new row is reachable
    Dim needed As Integer: needed = ALN_ROW_START_TOP + rIdx * ALN_ROW_H + 15
    If needed > frm.ScrollHeight Then frm.ScrollHeight = needed
    Call UpdateRowSelector(aIdx)
    Exit Sub
RowError:
    Debug.Print "AddAlignRowWithData error aIdx=" & aIdx & " rIdx=" & rIdx & ": " & Err.Description
    alignRowCounts(aIdx) = alignRowCounts(aIdx) - 1
End Sub

' ============================================================
' REBUILD ALIGNMENT TABLE FROM DATA ARRAYS
' Delete, Move Up, Move Down all use read-modify-rebuild pattern.
' ============================================================
Private Sub RebuildAlignTable(aIdx As Integer, _
                               types() As String, labels() As String, _
                               spacings() As String, sizes() As String, _
                               sides() As String, rCount As Integer)
    On Error GoTo RebuildError
    Dim frm As Object: Set frm = alignFrames(aIdx)

    ' Clear all existing row controls
    Dim r As Integer
    For r = 1 To alignRowCounts(aIdx)
        On Error Resume Next
        frm.Controls.Remove "cboType_" & aIdx & "_" & r
        frm.Controls.Remove "txtLabel_" & aIdx & "_" & r
        frm.Controls.Remove "txtSpacing_" & aIdx & "_" & r
        frm.Controls.Remove "txtSize_" & aIdx & "_" & r
        frm.Controls.Remove "cboSide_" & aIdx & "_" & r
        Set alignTypeBoxes(aIdx, r) = Nothing
        Set alignLabelBoxes(aIdx, r) = Nothing
        Set alignSpacingBoxes(aIdx, r) = Nothing
        Set alignSizeBoxes(aIdx, r) = Nothing
        Set alignSideBoxes(aIdx, r) = Nothing
        On Error GoTo RebuildError
    Next r
    Set alignRowHandlers(aIdx) = New Collection
    Set alignTypeHandlers(aIdx) = New Collection
    alignRowCounts(aIdx) = 0
    frm.ScrollHeight = 500   ' reset scroll

    ' Recreate from supplied arrays
    Dim i As Integer
    For i = 1 To rCount
        Call AddAlignRowWithData(aIdx, types(i), labels(i), spacings(i), sizes(i), sides(i))
    Next i
    Call UpdateRowSelector(aIdx)   ' ensure combobox is correct even when rCount=0
    Exit Sub

RebuildError:
    Debug.Print "RebuildAlignTable error aIdx=" & aIdx & ": " & Err.Description
End Sub

' ============================================================
' READ ALL ROW DATA FROM AN ALIGNMENT INTO TEMP ARRAYS
' ============================================================
Private Sub ReadAlignData(aIdx As Integer, _
                           types() As String, labels() As String, _
                           spacings() As String, sizes() As String, _
                           sides() As String, rCount As Integer)
    rCount = alignRowCounts(aIdx)
    If rCount = 0 Then Exit Sub
    ReDim types(1 To rCount)
    ReDim labels(1 To rCount)
    ReDim spacings(1 To rCount)
    ReDim sizes(1 To rCount)
    ReDim sides(1 To rCount)
    Dim r As Integer
    For r = 1 To rCount
        types(r) = alignTypeBoxes(aIdx, r).Value
        labels(r) = alignLabelBoxes(aIdx, r).Text
        spacings(r) = alignSpacingBoxes(aIdx, r).Text
        sizes(r) = alignSizeBoxes(aIdx, r).Text
        sides(r) = alignSideBoxes(aIdx, r).Value
    Next r
End Sub

' ============================================================
' ROW SELECTION — Public so AlignRowBox class can call it
' ============================================================
Public Sub SelectAlignRow(aIdx As Integer, rIdx As Integer)
    If inSelectAlignRow Then Exit Sub
    inSelectAlignRow = True
    On Error Resume Next
    If activeAlignIdx > 0 And activeRowIdx > 0 Then
        Call SetRowHighlight(activeAlignIdx, activeRowIdx, False)
    End If
    activeAlignIdx = aIdx
    activeRowIdx = rIdx
    Call SetRowHighlight(aIdx, rIdx, True)
    ' Sync row selector combobox without retriggering this sub
    If Not alignRowSelBoxes(aIdx) Is Nothing Then
        If alignRowSelBoxes(aIdx).ListIndex <> rIdx - 1 Then
            alignRowSelBoxes(aIdx).ListIndex = rIdx - 1
        End If
    End If
    inSelectAlignRow = False
End Sub

' ============================================================
' CALLED BY CboRowSel_Change IN AlignSectionBtns
' ============================================================
Public Sub SelectAlignRowFromSel(aIdx As Integer, rIdx As Integer)
    If aIdx < 1 Or rIdx < 1 Then Exit Sub
    If aIdx > alignCount Or rIdx > alignRowCounts(aIdx) Then Exit Sub
    Call SelectAlignRow(aIdx, rIdx)
End Sub

' ============================================================
' REPOPULATE ROW SELECTOR COMBOBOX FOR AN ALIGNMENT
' ============================================================
Private Sub UpdateRowSelector(aIdx As Integer)
    If aIdx < 1 Or aIdx > MAX_ALIGNMENTS Then Exit Sub
    On Error Resume Next
    If alignRowSelBoxes(aIdx) Is Nothing Then Exit Sub
    Dim cbo As Object: Set cbo = alignRowSelBoxes(aIdx)
    cbo.Clear
    Dim r As Integer
    For r = 1 To alignRowCounts(aIdx)
        cbo.AddItem CStr(r)
    Next r
    ' Leave at -1 (no selection); caller sets selection explicitly if needed
End Sub

' ============================================================
' SYNC SPACING TABLE VALUE TO MATCHING ALIGNMENT ROW
' Called by SpacingBox.cls on Enter or focus-loss
' ============================================================
Public Sub SyncSpacingToAlignment(aIdx As Integer, rowLabel As String, newValue As String)
    On Error Resume Next
    If aIdx < 1 Or aIdx > alignCount Then Exit Sub
    Dim r As Integer
    For r = 1 To alignRowCounts(aIdx)
        If Not alignLabelBoxes(aIdx, r) Is Nothing Then
            If Trim(alignLabelBoxes(aIdx, r).Text) = rowLabel Then
                alignSpacingBoxes(aIdx, r).Text = newValue
                Exit For
            End If
        End If
    Next r
End Sub

Private Sub SetRowHighlight(aIdx As Integer, rIdx As Integer, highlighted As Boolean)
    If aIdx < 1 Or rIdx < 1 Then Exit Sub
    If aIdx > alignCount Or rIdx > alignRowCounts(aIdx) Then Exit Sub
    On Error Resume Next
    Dim c As Long
    If highlighted Then
        c = HIGHLIGHT_COLOR
    Else
        c = CLng(&HFFFFFF)
    End If
    alignTypeBoxes(aIdx, rIdx).BackColor = c
    alignLabelBoxes(aIdx, rIdx).BackColor = c
    alignSpacingBoxes(aIdx, rIdx).BackColor = c
    If highlighted Then
        alignSizeBoxes(aIdx, rIdx).BackColor = c
        alignSideBoxes(aIdx, rIdx).BackColor = c
    Else
        If alignSizeBoxes(aIdx, rIdx).Enabled Then
            alignSizeBoxes(aIdx, rIdx).BackColor = CLng(&HFFFFFF)
        Else
            alignSizeBoxes(aIdx, rIdx).BackColor = CLng(&HE0E0E0)
        End If
        If alignSideBoxes(aIdx, rIdx).Enabled Then
            alignSideBoxes(aIdx, rIdx).BackColor = CLng(&HFFFFFF)
        Else
            alignSideBoxes(aIdx, rIdx).BackColor = CLng(&HE0E0E0)
        End If
    End If
    On Error GoTo 0
End Sub

' ============================================================
' TYPE CHANGE — Public so AlignTypeBox class can call it
' ============================================================
Public Sub OnTypeChange(aIdx As Integer, rIdx As Integer)
    If aIdx < 1 Or rIdx < 1 Then Exit Sub
    If aIdx > alignCount Or rIdx > alignRowCounts(aIdx) Then Exit Sub
    On Error Resume Next
    Dim isSign As Boolean: isSign = (alignTypeBoxes(aIdx, rIdx).Value = "Sign")
    alignSizeBoxes(aIdx, rIdx).Enabled = isSign
    alignSideBoxes(aIdx, rIdx).Enabled = isSign
    If isSign Then
        alignSizeBoxes(aIdx, rIdx).BackColor = CLng(&HFFFFFF)
        alignSideBoxes(aIdx, rIdx).BackColor = CLng(&HFFFFFF)
    Else
        alignSizeBoxes(aIdx, rIdx).BackColor = CLng(&HE0E0E0)
        alignSideBoxes(aIdx, rIdx).BackColor = CLng(&HE0E0E0)
        alignSizeBoxes(aIdx, rIdx).Text = ""
        alignSideBoxes(aIdx, rIdx).ListIndex = 0
    End If
    On Error GoTo 0
End Sub

' ============================================================
' SIGN LIBRARY AUTO-FILL — Public so AlignRowBox can call it
' ============================================================
Public Sub ApplySignLibraryToAlignRow(aIdx As Integer, rIdx As Integer)
    On Error GoTo LibError
    If aIdx < 1 Or rIdx < 1 Then Exit Sub
    If aIdx > alignCount Or rIdx > alignRowCounts(aIdx) Then Exit Sub
    If alignTypeBoxes(aIdx, rIdx).Value <> "Sign" Then Exit Sub

    Dim s As String: s = Trim(alignLabelBoxes(aIdx, rIdx).Text)
    If s = "" Then Exit Sub

    Dim roadType As String: roadType = "Non-Freeway"
    If ControlExists("cboRoadType") And cboRoadType.ListIndex > 0 Then
        roadType = cboRoadType.Value
    End If

    Dim sd As signData
    sd = GetSignData(s, roadType)
    If sd.SignNumber = "" Then
        Dim allSigns() As String: allSigns = GetAllSignNumbers
        Dim i As Long, matchKey As String: matchKey = ""
        For i = LBound(allSigns) To UBound(allSigns)
            If allSigns(i) <> "" And StrComp(s, allSigns(i), vbTextCompare) = 0 Then
                matchKey = allSigns(i): Exit For
            End If
        Next i
        If matchKey = "" Then Exit Sub
        sd = GetSignData(matchKey, roadType)
        If sd.SignNumber = "" Then Exit Sub
    End If
    alignSpacingBoxes(aIdx, rIdx).Text = CStr(sd.DefaultSpacing)
    alignSizeBoxes(aIdx, rIdx).Text = sd.TextLine2
    Exit Sub
LibError:
    Debug.Print "ApplySignLibraryToAlignRow error: " & Err.Description
End Sub

' ============================================================
' AUTO-POPULATE WZTC SPACING ITEMS INTO UPSTREAM ALIGNMENT
' Called from GenerateSpacingTable after spacing values are computed.
' ============================================================
Private Sub AutoPopulateWZTCItems(aIdx As Integer)
    If Not hasSpacingData Then Exit Sub
    If aIdx < 1 Or aIdx > alignCount Then Exit Sub

    If aIdx = 1 Then
        ' Upstream: Roll Ahead, Vehicle Space, Buffer Space, Merging Taper, Shoulder Taper,
        '           Upstream Taper Temp Barrier, Upstream Taper Box/Corr Beam
        Dim ut(1 To 7)  As String, ul(1 To 7)  As String
        Dim us(1 To 7)  As String, usz(1 To 7) As String, usd(1 To 7) As String
        Dim i As Integer
        For i = 1 To 7: ut(i) = "Non-Sign": usz(i) = "": usd(i) = "One Side": Next i
        ul(1) = "Roll Ahead Distance":          us(1) = Format(cachedRollAhead, "0.0")
        ul(2) = "Vehicle Space":                us(2) = Format(cachedVehicleSpace, "0.0")
        ul(3) = "Buffer Space":                 us(3) = Format(cachedBufferSpace, "0.0")
        ul(4) = "Merging/Shifting Taper":       us(4) = Format(cachedMergingTaper, "0.0")
        ul(5) = "Shoulder Taper":               us(5) = Format(cachedShoulderTapers, "0.0")
        ul(6) = "Upstream Taper Temp Barrier":  us(6) = Format(cachedUpTaperBarrier, "0.0")
        ul(7) = "Upstream Taper Box/Corr Beam": us(7) = Format(cachedUpTaperBeam, "0.0")
        Call RebuildAlignTable(1, ut, ul, us, usz, usd, 7)
        If ControlExists("lblStatus") Then
            lblStatus.Caption = "Upstream populated with 7 spacing items. " & _
                                "Use '+ Add Row' next to each alignment title to add Sign rows."
        End If

    ElseIf aIdx = 2 Then
        ' Downstream: Downstream Taper only
        Dim dt(1 To 1)  As String, dl(1 To 1)  As String
        Dim ds(1 To 1)  As String, dsz(1 To 1) As String, dsd(1 To 1) As String
        dt(1) = "Non-Sign": dl(1) = "Downstream Taper"
        ds(1) = Format(cachedDownstreamTaper, "0.0"): dsz(1) = "": dsd(1) = "One Side"
        Call RebuildAlignTable(2, dt, dl, ds, dsz, dsd, 1)
    End If
End Sub

' ============================================================
' PER-SECTION ADD ROW — called by AlignAddBtn click handler
' ============================================================
Public Sub AddRowForAlign(aIdx As Integer)
    If aIdx < 1 Or aIdx > alignCount Then Exit Sub
    Call AddAlignRow(aIdx)
    Call SelectAlignRow(aIdx, alignRowCounts(aIdx))   ' auto-select the new row
    If ControlExists("lblStatus") Then
        Dim nm As String
        Select Case aIdx
            Case 1:    nm = "Upstream"
            Case 2:    nm = "Downstream"
            Case Else: nm = "Alignment " & aIdx
        End Select
        lblStatus.Caption = "Row added to " & nm & " alignment."
    End If
End Sub

' ============================================================
' PER-SECTION DEL ROW
' ============================================================
Public Sub DelRowForAlign(aIdx As Integer)
    If alignRowCounts(aIdx) = 0 Then
        If ControlExists("lblStatus") Then lblStatus.Caption = "No rows to delete." End If
        Exit Sub
    End If
    ' Delete the row selected in this alignment's combobox, or the last row if none selected
    Dim rIdx As Integer
    If activeAlignIdx = aIdx And activeRowIdx >= 1 And activeRowIdx <= alignRowCounts(aIdx) Then
        rIdx = activeRowIdx
    Else
        rIdx = alignRowCounts(aIdx)
    End If
    Dim types() As String, labels() As String
    Dim spacings() As String, sizes() As String
    Dim sides() As String, rCount As Integer
    Call ReadAlignData(aIdx, types, labels, spacings, sizes, sides, rCount)
    If rCount = 0 Then Exit Sub
    Dim i As Integer
    For i = rIdx To rCount - 1
        types(i) = types(i + 1): labels(i) = labels(i + 1)
        spacings(i) = spacings(i + 1): sizes(i) = sizes(i + 1): sides(i) = sides(i + 1)
    Next i
    rCount = rCount - 1
    activeAlignIdx = 0: activeRowIdx = 0
    If rCount > 0 Then
        Call RebuildAlignTable(aIdx, types, labels, spacings, sizes, sides, rCount)
    Else
        Dim emptyT(1 To 1) As String, emptyL(1 To 1) As String
        Dim emptySp(1 To 1) As String, emptySz(1 To 1) As String
        Dim emptySd(1 To 1) As String
        Call RebuildAlignTable(aIdx, emptyT, emptyL, emptySp, emptySz, emptySd, 0)
    End If
    If ControlExists("lblStatus") Then lblStatus.Caption = "Row deleted." End If
End Sub

' ============================================================
' PER-SECTION MOVE UP
' ============================================================
Public Sub MoveRowUpForAlign(aIdx As Integer)
    If activeAlignIdx <> aIdx Or activeRowIdx <= 1 Then
        If ControlExists("lblStatus") Then
            lblStatus.Caption = "Select a row below the first in this alignment, then click Move Up."
        End If
        Exit Sub
    End If
    Dim rIdx As Integer: rIdx = activeRowIdx
    Dim types() As String, labels() As String
    Dim spacings() As String, sizes() As String
    Dim sides() As String, rCount As Integer
    Call ReadAlignData(aIdx, types, labels, spacings, sizes, sides, rCount)
    Dim tmp As String
    tmp = types(rIdx):    types(rIdx) = types(rIdx - 1):       types(rIdx - 1) = tmp
    tmp = labels(rIdx):   labels(rIdx) = labels(rIdx - 1):     labels(rIdx - 1) = tmp
    tmp = spacings(rIdx): spacings(rIdx) = spacings(rIdx - 1): spacings(rIdx - 1) = tmp
    tmp = sizes(rIdx):    sizes(rIdx) = sizes(rIdx - 1):       sizes(rIdx - 1) = tmp
    tmp = sides(rIdx):    sides(rIdx) = sides(rIdx - 1):       sides(rIdx - 1) = tmp
    activeAlignIdx = 0: activeRowIdx = 0
    Call RebuildAlignTable(aIdx, types, labels, spacings, sizes, sides, rCount)
    Call SelectAlignRow(aIdx, rIdx - 1)
    If ControlExists("lblStatus") Then lblStatus.Caption = "Row moved up." End If
End Sub

' ============================================================
' PER-SECTION MOVE DOWN
' ============================================================
Public Sub MoveRowDownForAlign(aIdx As Integer)
    If activeAlignIdx <> aIdx Or activeRowIdx = 0 Then
        If ControlExists("lblStatus") Then
            lblStatus.Caption = "Select a row in this alignment first, then click Move Down."
        End If
        Exit Sub
    End If
    Dim rIdx As Integer: rIdx = activeRowIdx
    If rIdx >= alignRowCounts(aIdx) Then
        If ControlExists("lblStatus") Then lblStatus.Caption = "Row is already at the bottom." End If
        Exit Sub
    End If
    Dim types() As String, labels() As String
    Dim spacings() As String, sizes() As String
    Dim sides() As String, rCount As Integer
    Call ReadAlignData(aIdx, types, labels, spacings, sizes, sides, rCount)
    Dim tmp As String
    tmp = types(rIdx):    types(rIdx) = types(rIdx + 1):       types(rIdx + 1) = tmp
    tmp = labels(rIdx):   labels(rIdx) = labels(rIdx + 1):     labels(rIdx + 1) = tmp
    tmp = spacings(rIdx): spacings(rIdx) = spacings(rIdx + 1): spacings(rIdx + 1) = tmp
    tmp = sizes(rIdx):    sizes(rIdx) = sizes(rIdx + 1):       sizes(rIdx + 1) = tmp
    tmp = sides(rIdx):    sides(rIdx) = sides(rIdx + 1):       sides(rIdx + 1) = tmp
    activeAlignIdx = 0: activeRowIdx = 0
    Call RebuildAlignTable(aIdx, types, labels, spacings, sizes, sides, rCount)
    Call SelectAlignRow(aIdx, rIdx + 1)
    If ControlExists("lblStatus") Then lblStatus.Caption = "Row moved down." End If
End Sub

' ============================================================
' UPDATE FORM SCROLL HEIGHT
' ============================================================
Private Sub UpdateFormScrollHeight()
    Dim neededHeight As Integer
    neededHeight = AlignSectionTop(alignCount) + ALN_SECTION_H + 40
    Me.ScrollBars = fmScrollBarsVertical
    If neededHeight > Me.Height Then
        Me.ScrollHeight = neededHeight
    End If
End Sub

' ============================================================
' UPDATE SIGN SIZES WHEN ROAD TYPE CHANGES
' ============================================================
Private Sub UpdateSignSizesForAllAlignments()
    If ControlExists("cboRoadType") Then
        If cboRoadType.ListIndex <= 0 Then Exit Sub
    End If
    Dim roadType As String: roadType = cboRoadType.Value
    Dim a As Integer, r As Integer
    Dim sd As signData
    For a = 1 To alignCount
        For r = 1 To alignRowCounts(a)
            If Not alignTypeBoxes(a, r) Is Nothing Then
                If alignTypeBoxes(a, r).Value = "Sign" Then
                    Dim s As String: s = Trim(alignLabelBoxes(a, r).Text)
                    If s <> "" Then
                        sd = GetSignData(s, roadType)
                        If sd.SignNumber <> "" Then
                            alignSizeBoxes(a, r).Text = sd.TextLine2
                        End If
                    End If
                End If
            End If
        Next r
    Next a
End Sub

' ============================================================
' ROW ACTION BUTTON HANDLERS
' ============================================================
Private Sub btnAddRow_Click()
    If activeAlignIdx > 0 Then
        Call AddAlignRow(activeAlignIdx)
        Dim nm As String
        Select Case activeAlignIdx
            Case 1:    nm = "Upstream"
            Case 2:    nm = "Downstream"
            Case Else: nm = "Alignment " & activeAlignIdx
        End Select
        lblStatus.Caption = "Row added to " & nm & " alignment."
    Else
        lblStatus.Caption = "Use the '+ Add Row' button beside each alignment's title to add rows, " & _
                            "or click any existing row's Sign # field first to target that alignment."
    End If
End Sub

Private Sub btnDeleteRow_Click()
    If activeAlignIdx = 0 Or activeRowIdx = 0 Then
        lblStatus.Caption = "Click the Sign # / Description field in a row to select it, then click Del Row."
        Exit Sub
    End If
    Dim aIdx As Integer: aIdx = activeAlignIdx
    Dim rIdx As Integer: rIdx = activeRowIdx
    If alignRowCounts(aIdx) = 0 Then Exit Sub

    Dim types()    As String, labels()   As String
    Dim spacings() As String, sizes()    As String
    Dim sides()    As String, rCount As Integer
    Call ReadAlignData(aIdx, types, labels, spacings, sizes, sides, rCount)

    If rCount = 0 Then Exit Sub

    ' Remove rIdx from arrays by shifting down
    Dim i As Integer
    For i = rIdx To rCount - 1
        types(i) = types(i + 1): labels(i) = labels(i + 1)
        spacings(i) = spacings(i + 1): sizes(i) = sizes(i + 1): sides(i) = sides(i + 1)
    Next i
    rCount = rCount - 1

    activeAlignIdx = 0: activeRowIdx = 0
    If rCount > 0 Then
        Call RebuildAlignTable(aIdx, types, labels, spacings, sizes, sides, rCount)
    Else
        Dim emptyT(1 To 1) As String, emptyL(1 To 1) As String
        Dim emptySp(1 To 1) As String, emptySz(1 To 1) As String
        Dim emptySd(1 To 1) As String
        Call RebuildAlignTable(aIdx, emptyT, emptyL, emptySp, emptySz, emptySd, 0)
    End If
    lblStatus.Caption = "Row deleted."
End Sub

Private Sub btnMoveUp_Click()
    If activeAlignIdx = 0 Or activeRowIdx <= 1 Then
        lblStatus.Caption = "Select a row below the first, then click Move Up."
        Exit Sub
    End If
    Dim aIdx As Integer: aIdx = activeAlignIdx
    Dim rIdx As Integer: rIdx = activeRowIdx

    Dim types()    As String, labels()   As String
    Dim spacings() As String, sizes()    As String
    Dim sides()    As String, rCount As Integer
    Call ReadAlignData(aIdx, types, labels, spacings, sizes, sides, rCount)

    Dim tmp As String
    tmp = types(rIdx): types(rIdx) = types(rIdx - 1): types(rIdx - 1) = tmp
    tmp = labels(rIdx): labels(rIdx) = labels(rIdx - 1): labels(rIdx - 1) = tmp
    tmp = spacings(rIdx): spacings(rIdx) = spacings(rIdx - 1): spacings(rIdx - 1) = tmp
    tmp = sizes(rIdx): sizes(rIdx) = sizes(rIdx - 1): sizes(rIdx - 1) = tmp
    tmp = sides(rIdx): sides(rIdx) = sides(rIdx - 1): sides(rIdx - 1) = tmp

    activeAlignIdx = 0: activeRowIdx = 0
    Call RebuildAlignTable(aIdx, types, labels, spacings, sizes, sides, rCount)
    Call SelectAlignRow(aIdx, rIdx - 1)
    lblStatus.Caption = "Row moved up."
End Sub

Private Sub btnMoveDown_Click()
    If activeAlignIdx = 0 Or activeRowIdx = 0 Then
        lblStatus.Caption = "Select a row first, then click Move Down."
        Exit Sub
    End If
    Dim aIdx As Integer: aIdx = activeAlignIdx
    Dim rIdx As Integer: rIdx = activeRowIdx
    If rIdx >= alignRowCounts(aIdx) Then
        lblStatus.Caption = "Row is already at the bottom."
        Exit Sub
    End If

    Dim types()    As String, labels()   As String
    Dim spacings() As String, sizes()    As String
    Dim sides()    As String, rCount As Integer
    Call ReadAlignData(aIdx, types, labels, spacings, sizes, sides, rCount)

    Dim tmp As String
    tmp = types(rIdx): types(rIdx) = types(rIdx + 1): types(rIdx + 1) = tmp
    tmp = labels(rIdx): labels(rIdx) = labels(rIdx + 1): labels(rIdx + 1) = tmp
    tmp = spacings(rIdx): spacings(rIdx) = spacings(rIdx + 1): spacings(rIdx + 1) = tmp
    tmp = sizes(rIdx): sizes(rIdx) = sizes(rIdx + 1): sizes(rIdx + 1) = tmp
    tmp = sides(rIdx): sides(rIdx) = sides(rIdx + 1): sides(rIdx + 1) = tmp

    activeAlignIdx = 0: activeRowIdx = 0
    Call RebuildAlignTable(aIdx, types, labels, spacings, sizes, sides, rCount)
    Call SelectAlignRow(aIdx, rIdx + 1)
    lblStatus.Caption = "Row moved down."
End Sub

' ============================================================
' ALIGNMENT MANAGEMENT BUTTON HANDLERS
' ============================================================
Private Sub btnAddAlignment_Click()
    If alignCount >= MAX_ALIGNMENTS Then
        MsgBox "Maximum " & MAX_ALIGNMENTS & " alignments reached.", vbExclamation
        Exit Sub
    End If
    alignCount = alignCount + 1
    Call CreateAlignSection(alignCount, "Alignment " & alignCount)
    Call AddAlignRow(alignCount)   ' seed with one empty row so user can target this alignment
    Call UpdateFormScrollHeight
    lblStatus.Caption = "Alignment " & alignCount & " added. Use '+ Add Row' beside its title to add more rows."
End Sub

Private Sub btnRemoveAlignment_Click()
    If alignCount <= 2 Then
        MsgBox "Upstream and Downstream alignments cannot be removed.", vbExclamation
        Exit Sub
    End If
    Dim aIdx As Integer: aIdx = alignCount

    If activeAlignIdx = aIdx Then
        activeAlignIdx = 0: activeRowIdx = 0
    End If

    ' Remove all row controls from this alignment's frame
    Dim r As Integer
    Dim frm As Object: Set frm = alignFrames(aIdx)
    For r = 1 To alignRowCounts(aIdx)
        On Error Resume Next
        frm.Controls.Remove "cboType_" & aIdx & "_" & r
        frm.Controls.Remove "txtLabel_" & aIdx & "_" & r
        frm.Controls.Remove "txtSpacing_" & aIdx & "_" & r
        frm.Controls.Remove "txtSize_" & aIdx & "_" & r
        frm.Controls.Remove "cboSide_" & aIdx & "_" & r
        On Error GoTo 0
    Next r
    alignRowCounts(aIdx) = 0

    On Error Resume Next
    Me.Controls.Remove "frAlign_" & aIdx
    Me.Controls.Remove "lblAlign_" & aIdx
    On Error GoTo 0
    Set alignFrames(aIdx) = Nothing
    Set alignSectionLabels(aIdx) = Nothing

    alignCount = alignCount - 1
    lblStatus.Caption = "Alignment removed."
End Sub

' ============================================================
' POPULATE DROPDOWN LISTS
' ============================================================
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
    If cboSheet.ListCount > 0 Then cboSheet.ListIndex = 0
End Sub

Private Sub cboSheet_Change()
    If cboSheet.ListIndex >= 0 Then
        lblStatus.Caption = "Sheet selected - Please select Road Speed."
        Call CheckAllSelectionsComplete
    End If
End Sub

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
        Call UpdateSignSizesForAllAlignments
        lblStatus.Caption = "Road Type selected - Please select Lane Width."
        Call CheckAllSelectionsComplete
    End If
End Sub

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

Private Sub cboRoadSpeed_Change()
    If cboRoadSpeed.ListIndex > 0 And cboCategory.ListIndex > 0 Then
        selectedSpeed = cboRoadSpeed.Value
        Call GenerateSpacingTable
        lblStatus.Caption = "Road Speed selected - Please select Road Type."
        Call CheckAllSelectionsComplete
    End If
End Sub

' ============================================================
' REFERENCE BUTTON — opens MUTCD sheet viewer
' ============================================================
Private Sub btnReference_Click()
    On Error Resume Next
    Dim catText As String, sheetText As String
    catText = cboCategory.Value
    sheetText = cboSheet.Value
    SheetViewer.Left = Me.Left + Me.Width + 8
    SheetViewer.Top = Me.Top
    If catText <> "" Then
        SheetViewer.SelectAndShow catText, sheetText
    Else
        Me.Hide
        SheetViewer.Show vbModeless
    End If
End Sub

' ============================================================
' CHECK IF ALL REQUIRED SELECTIONS ARE COMPLETE
' ============================================================
Private Sub CheckAllSelectionsComplete()
    Dim allComplete As Boolean
    allComplete = (cboCategory.ListIndex >= 0) And _
                  (cboSheet.ListIndex >= 0) And _
                  (cboRoadSpeed.ListIndex > 0) And _
                  (cboRoadType.ListIndex > 0) And _
                  (cboLaneWidth.ListIndex > 0) And _
                  (cboShoulderWidth.ListIndex > 0)
    If allComplete Then
        lblStatus.Caption = "All parameters selected. Add signs and spacing items to alignments, then click Submit & Draw."
    End If
End Sub

' ============================================================
' Wire a SpacingBox handler for a spacing textbox → alignment row sync
' ============================================================
Private Sub AddSpacingHandler(txt As MSForms.TextBox, aIdx As Integer, rowLabel As String)
    Dim h As SpacingBox
    Set h = New SpacingBox
    h.AlignIdx = aIdx
    h.RowLabel = rowLabel
    Set h.Txt = txt
    Set h.ParentForm = Me
    spacingBoxHandlers.Add h
End Sub

' ============================================================
' CREATE SPACING LABELS IN FRAME (unchanged from original)
' ============================================================
Private Sub CreateSpacingLabels()
    On Error GoTo SpacingError

    Dim lblDownstream As MSForms.Label, txtDownstream As MSForms.TextBox
    Dim lblRollAhead As MSForms.Label, txtRollAhead As MSForms.TextBox
    Dim lblVehicle As MSForms.Label, txtVehicle As MSForms.TextBox
    Dim lblBuffer As MSForms.Label, txtBuffer As MSForms.TextBox
    Dim lblMerging As MSForms.Label, txtMerging As MSForms.TextBox
    Dim lblShoulder As MSForms.Label, txtShoulder As MSForms.TextBox
    Dim lblAdvanced As MSForms.Label, txtAdvanced As MSForms.TextBox
    Dim lblSkipLines As MSForms.Label, txtSkipLines As MSForms.TextBox
    Dim lblChannelizing As MSForms.Label, txtChannelizing As MSForms.TextBox
    Dim lblFlareBarrier As MSForms.Label, txtFlareBarrier As MSForms.TextBox
    Dim lblFlareBeam As MSForms.Label, txtFlareBeam As MSForms.TextBox

    Set lblDownstream = frameSpacingValues.Controls.Add("Forms.Label.1", "lblDownstreamTaper")
    With lblDownstream: .Caption = "Downstream Taper (ft):": .Top = 20: .Left = 10: .Width = 110: .Height = 18: End With
    Set txtDownstream = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtDownstreamTaper")
    With txtDownstream: .Top = 20: .Left = 125: .Width = 50: .Height = 18: End With
    Call AddSpacingHandler(txtDownstream, 2, "Downstream Taper")

    Set lblRollAhead = frameSpacingValues.Controls.Add("Forms.Label.1", "lblRollAhead")
    With lblRollAhead: .Caption = "Roll Ahead Distance (ft):": .Top = 40: .Left = 10: .Width = 110: .Height = 18: End With
    Set txtRollAhead = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtRollAhead")
    With txtRollAhead: .Top = 40: .Left = 125: .Width = 50: .Height = 18: End With
    Call AddSpacingHandler(txtRollAhead, 1, "Roll Ahead Distance")

    Set lblVehicle = frameSpacingValues.Controls.Add("Forms.Label.1", "lblVehicleSpace")
    With lblVehicle: .Caption = "Vehicle Space (ft):": .Top = 60: .Left = 10: .Width = 110: .Height = 18: End With
    Set txtVehicle = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtVehicleSpace")
    With txtVehicle: .Top = 60: .Left = 125: .Width = 50: .Height = 18: End With
    Call AddSpacingHandler(txtVehicle, 1, "Vehicle Space")

    Set lblBuffer = frameSpacingValues.Controls.Add("Forms.Label.1", "lblBufferSpace")
    With lblBuffer: .Caption = "Buffer Space (ft):": .Top = 80: .Left = 10: .Width = 110: .Height = 18: End With
    Set txtBuffer = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtBufferSpace")
    With txtBuffer: .Top = 80: .Left = 125: .Width = 50: .Height = 18: End With
    Call AddSpacingHandler(txtBuffer, 1, "Buffer Space")

    Set lblMerging = frameSpacingValues.Controls.Add("Forms.Label.1", "lblMergingTaper")
    With lblMerging: .Caption = "Merging/Shifting Taper (ft):": .Top = 100: .Left = 10: .Width = 110: .Height = 18: End With
    Set txtMerging = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtMergingTaper")
    With txtMerging: .Top = 100: .Left = 125: .Width = 50: .Height = 18: End With
    Call AddSpacingHandler(txtMerging, 1, "Merging/Shifting Taper")

    Set lblShoulder = frameSpacingValues.Controls.Add("Forms.Label.1", "lblShoulderTapers")
    With lblShoulder: .Caption = "Shoulder Tapers (ft):": .Top = 120: .Left = 10: .Width = 110: .Height = 18: End With
    Set txtShoulder = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtShoulderTapers")
    With txtShoulder: .Top = 120: .Left = 125: .Width = 50: .Height = 18: End With
    Call AddSpacingHandler(txtShoulder, 1, "Shoulder Taper")

    Set lblAdvanced = frameSpacingValues.Controls.Add("Forms.Label.1", "lblAdvancedWarningSpacing")
    With lblAdvanced: .Caption = "Adv. Warning Spacing (ft):": .Top = 140: .Left = 10: .Width = 110: .Height = 18: End With
    Set txtAdvanced = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtAdvancedWarningSpacing")
    With txtAdvanced: .Top = 140: .Left = 125: .Width = 50: .Height = 18: End With

    Set lblSkipLines = frameSpacingValues.Controls.Add("Forms.Label.1", "lblSkipLines")
    With lblSkipLines: .Caption = "# of Skip Lines:": .Top = 160: .Left = 10: .Width = 110: .Height = 18: End With
    Set txtSkipLines = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtSkipLines")
    With txtSkipLines: .Top = 160: .Left = 125: .Width = 50: .Height = 18: End With

    Set lblChannelizing = frameSpacingValues.Controls.Add("Forms.Label.1", "lblChannelizing")
    With lblChannelizing: .Caption = "# of Channelizing Devices:": .Top = 180: .Left = 10: .Width = 110: .Height = 18: End With
    Set txtChannelizing = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtChannelizing")
    With txtChannelizing: .Top = 180: .Left = 125: .Width = 50: .Height = 18: End With

    Set lblFlareBarrier = frameSpacingValues.Controls.Add("Forms.Label.1", "lblFlareBarrier")
    With lblFlareBarrier: .Caption = "Flare Rate Temp Barrier:": .Top = 200: .Left = 10: .Width = 110: .Height = 18: End With
    Set txtFlareBarrier = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtFlareBarrier")
    With txtFlareBarrier: .Top = 200: .Left = 125: .Width = 50: .Height = 18: End With

    Set lblFlareBeam = frameSpacingValues.Controls.Add("Forms.Label.1", "lblFlareBeam")
    With lblFlareBeam: .Caption = "Flare Rate Box/Corr Beam:": .Top = 220: .Left = 10: .Width = 110: .Height = 18: End With
    Set txtFlareBeam = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtFlareBeam")
    With txtFlareBeam: .Top = 220: .Left = 125: .Width = 50: .Height = 18: End With

    Dim lblUpTaperBarrier As MSForms.Label, txtUpTaperBarrier As MSForms.TextBox
    Dim lblUpTaperBeam    As MSForms.Label, txtUpTaperBeam    As MSForms.TextBox

    Set lblUpTaperBarrier = frameSpacingValues.Controls.Add("Forms.Label.1", "lblUpTaperBarrier")
    With lblUpTaperBarrier: .Caption = "Upstream Taper Temp Barrier (ft):": .Top = 242: .Left = 10: .Width = 110: .Height = 18: End With
    Set txtUpTaperBarrier = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtUpTaperBarrier")
    With txtUpTaperBarrier: .Top = 242: .Left = 125: .Width = 50: .Height = 18: End With
    Call AddSpacingHandler(txtUpTaperBarrier, 1, "Upstream Taper Temp Barrier")

    Set lblUpTaperBeam = frameSpacingValues.Controls.Add("Forms.Label.1", "lblUpTaperBeam")
    With lblUpTaperBeam: .Caption = "Upstream Taper Box/Corr Beam (ft):": .Top = 264: .Left = 10: .Width = 110: .Height = 18: End With
    Set txtUpTaperBeam = frameSpacingValues.Controls.Add("Forms.TextBox.1", "txtUpTaperBeam")
    With txtUpTaperBeam: .Top = 264: .Left = 125: .Width = 50: .Height = 18: End With
    Call AddSpacingHandler(txtUpTaperBeam, 1, "Upstream Taper Box/Corr Beam")

    Exit Sub
SpacingError:
    MsgBox "Error in CreateSpacingLabels: " & Err.Description, vbExclamation
End Sub

' ============================================================
' PARSE UPSTREAM TAPER — converts "X:Y" flare rate + lane width to taper length
' Formula: upstreamTaper = laneWidth × (X / Y)
' Example: "8:1" + 12ft lane → 12 × 8 = 96 ft
' ============================================================
Private Function ParseUpstreamTaper(flareStr As String, laneWid As Integer) As Double
    On Error Resume Next
    ParseUpstreamTaper = 0
    If flareStr = "" Then Exit Function
    Dim parts() As String: parts = Split(flareStr, ":")
    If UBound(parts) < 1 Then Exit Function
    Dim num As Double: num = CDbl(Trim(parts(0)))
    Dim den As Double: den = CDbl(Trim(parts(1)))
    If den = 0 Then Exit Function
    ParseUpstreamTaper = laneWid * (num / den)
End Function

' ============================================================
' GENERATE SPACING TABLE BASED ON MUTCD NY STANDARDS
' Caches computed values, then auto-populates Upstream alignment.
' ============================================================
Private Sub GenerateSpacingTable()
    Dim speed As Integer, laneWidth As Integer
    Dim downstreamTaper As Double, vehicleSpace As Double
    Dim bufferSpace As Double, mergingTaper As Double
    Dim shoulderTapers As Double, advancedWarningSpacing As Double
    Dim skipMerge As Integer, chanMerge As Integer
    Dim skipShoulder As Integer, chanShoulder As Integer
    Dim skipBuffer As Integer, skipRollAhead As Integer
    Dim flareBarrierStr As String, flareBeamStr As String
    Dim skipTotal As Integer, chanTotal As Integer
    Dim upTaperBarrierVal As Double, upTaperBeamVal As Double

    speed = Val(Left(selectedSpeed, 2))
    laneWidth = Val(Left(cboLaneWidth.Value, 2))

    If LCase(Trim(selectedRoadType)) = "non-freeway" Then
        downstreamTaper = 50
    Else
        downstreamTaper = 100
    End If
    vehicleSpace = 50

    Select Case speed
        Case 25: bufferSpace = 155
        Case 30: bufferSpace = 200
        Case 35: bufferSpace = 250
        Case 40: bufferSpace = 305
        Case 45: bufferSpace = 360
        Case 50: bufferSpace = 425
        Case 55: bufferSpace = 495
        Case 65: bufferSpace = 645
        Case Else: bufferSpace = speed * 70
    End Select

    Select Case speed
        Case 25: skipBuffer = 4
        Case 30: skipBuffer = 5
        Case 35: skipBuffer = 6
        Case 40: skipBuffer = 8
        Case 45: skipBuffer = 9
        Case 50: skipBuffer = 11
        Case 55: skipBuffer = 13
        Case 65: skipBuffer = 16
        Case Else: skipBuffer = 0
    End Select

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
            mergingTaper = (speed * (laneWidth) ^ 2) / 60: skipMerge = 0: chanMerge = 0
    End Select

    Select Case speed
        Case 25
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "5-7 ft":  shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "8 ft":    shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "9 ft":    shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "10 ft":   shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "11 ft":   shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "12 ft":   shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case Else:      shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
            End Select
        Case 30
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "5-7 ft":  shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "8 ft":    shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "9 ft":    shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "10 ft":   shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "11 ft":   shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "12 ft":   shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case Else:      shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
            End Select
        Case 35
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "5-7 ft":  shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
                Case "8 ft":    shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "9 ft":    shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "10 ft":   shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "11 ft":   shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case "12 ft":   shoulderTapers = 80: skipShoulder = 2: chanShoulder = 3
                Case Else:      shoulderTapers = 40: skipShoulder = 1: chanShoulder = 2
            End Select
        Case 40
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 40:  skipShoulder = 1: chanShoulder = 2
                Case "5-7 ft":  shoulderTapers = 80:  skipShoulder = 1: chanShoulder = 2
                Case "8 ft":    shoulderTapers = 80:  skipShoulder = 2: chanShoulder = 3
                Case "9 ft":    shoulderTapers = 80:  skipShoulder = 2: chanShoulder = 3
                Case "10 ft":   shoulderTapers = 120: skipShoulder = 2: chanShoulder = 3
                Case "11 ft":   shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "12 ft":   shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case Else:      shoulderTapers = 40:  skipShoulder = 1: chanShoulder = 2
            End Select
        Case 45
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 80:  skipShoulder = 2: chanShoulder = 3
                Case "5-7 ft":  shoulderTapers = 80:  skipShoulder = 2: chanShoulder = 3
                Case "8 ft":    shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "9 ft":    shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "10 ft":   shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "11 ft":   shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "12 ft":   shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case Else:      shoulderTapers = 80:  skipShoulder = 2: chanShoulder = 3
            End Select
        Case 50
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 80:  skipShoulder = 2: chanShoulder = 3
                Case "5-7 ft":  shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "8 ft":    shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "9 ft":    shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "10 ft":   shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "11 ft":   shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "12 ft":   shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case Else:      shoulderTapers = 80:  skipShoulder = 2: chanShoulder = 3
            End Select
        Case 55
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 80:  skipShoulder = 2: chanShoulder = 3
                Case "5-7 ft":  shoulderTapers = 120: skipShoulder = 3: chanShoulder = 4
                Case "8 ft":    shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "9 ft":    shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "10 ft":   shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "11 ft":   shoulderTapers = 200: skipShoulder = 5: chanShoulder = 6
                Case "12 ft":   shoulderTapers = 200: skipShoulder = 5: chanShoulder = 6
                Case Else:      shoulderTapers = 80:  skipShoulder = 2: chanShoulder = 3
            End Select
        Case 65
            Select Case cboShoulderWidth.Value
                Case "<= 4 ft": shoulderTapers = 80:  skipShoulder = 2: chanShoulder = 3
                Case "5-7 ft":  shoulderTapers = 160: skipShoulder = 4: chanShoulder = 5
                Case "8 ft":    shoulderTapers = 200: skipShoulder = 5: chanShoulder = 6
                Case "9 ft":    shoulderTapers = 240: skipShoulder = 6: chanShoulder = 7
                Case "10 ft":   shoulderTapers = 240: skipShoulder = 6: chanShoulder = 7
                Case "11 ft":   shoulderTapers = 280: skipShoulder = 7: chanShoulder = 8
                Case "12 ft":   shoulderTapers = 280: skipShoulder = 7: chanShoulder = 8
                Case Else:      shoulderTapers = 80:  skipShoulder = 2: chanShoulder = 3
            End Select
        Case Else
            shoulderTapers = speed * 0.8: skipShoulder = 0: chanShoulder = 0
    End Select

    Select Case speed
        Case 25: advancedWarningSpacing = 515
        Case 30: advancedWarningSpacing = 620
        Case 35: advancedWarningSpacing = 720
        Case 40: advancedWarningSpacing = 825
        Case 45: advancedWarningSpacing = 930
        Case 50: advancedWarningSpacing = 1030
        Case 55: advancedWarningSpacing = 1135
        Case 65: advancedWarningSpacing = 1365
        Case Else: advancedWarningSpacing = speed * 10
    End Select

    Dim rollAhead As Double
    Select Case speed
        Case 25, 30, 35, 40: rollAhead = 120
        Case 45, 50, 55:     rollAhead = 160
        Case 65:             rollAhead = 200
        Case Else:           rollAhead = 120
    End Select

    Dim skipRollAheadVal As Integer
    Select Case speed
        Case 25, 30, 35, 40: skipRollAheadVal = 3
        Case 45, 50, 55:     skipRollAheadVal = 4
        Case 65:             skipRollAheadVal = 5
        Case Else:           skipRollAheadVal = 0
    End Select
    skipRollAhead = skipRollAheadVal

    Select Case speed
        Case 25, 30, 35: flareBarrierStr = "8:1"
        Case 40, 45:     flareBarrierStr = "11:1"
        Case 50:         flareBarrierStr = "14:1"
        Case 55:         flareBarrierStr = "16:1"
        Case 65:         flareBarrierStr = "20:1"
        Case Else:       flareBarrierStr = ""
    End Select

    Select Case speed
        Case 25, 30, 35: flareBeamStr = "7:1"
        Case 40, 45:     flareBeamStr = "9:1"
        Case 50:         flareBeamStr = "11:1"
        Case 55:         flareBeamStr = "12:1"
        Case 65:         flareBeamStr = "15:1"
        Case Else:       flareBeamStr = ""
    End Select

    skipTotal = skipMerge + skipShoulder + skipBuffer + skipRollAhead
    chanTotal = chanMerge + chanShoulder

    ' Upstream taper = laneWidth × (flare numerator / flare denominator)
    upTaperBarrierVal = ParseUpstreamTaper(flareBarrierStr, laneWidth)
    upTaperBeamVal    = ParseUpstreamTaper(flareBeamStr, laneWidth)

    ' Populate spacing display textboxes
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
    frameSpacingValues.Controls("txtUpTaperBarrier").Value = Format(upTaperBarrierVal, "0.0")
    frameSpacingValues.Controls("txtUpTaperBeam").Value = Format(upTaperBeamVal, "0.0")

    ' Cache values and auto-populate Upstream and Downstream alignments
    cachedDownstreamTaper = downstreamTaper
    cachedRollAhead = rollAhead
    cachedVehicleSpace = vehicleSpace
    cachedBufferSpace = bufferSpace
    cachedMergingTaper = mergingTaper
    cachedShoulderTapers = shoulderTapers
    cachedUpTaperBarrier = upTaperBarrierVal
    cachedUpTaperBeam = upTaperBeamVal
    hasSpacingData = True
    Call AutoPopulateWZTCItems(1)
    Call AutoPopulateWZTCItems(2)
End Sub

' ============================================================
' SUBMIT BUTTON
' ============================================================
Private Sub btnSubmit_Click()
    ' Basic validation
    If cboCategory.ListIndex <= 0 Then
        MsgBox "Please select a workzone category.", vbExclamation: Exit Sub
    End If
    If cboSheet.ListIndex < 0 Then
        MsgBox "Please select a sheet number.", vbExclamation: Exit Sub
    End If
    If cboRoadSpeed.ListIndex <= 0 Then
        MsgBox "Please select a road speed.", vbExclamation: Exit Sub
    End If
    If cboRoadType.ListIndex <= 0 Then
        MsgBox "Please select a road type (Freeway or Non-Freeway).", vbExclamation: Exit Sub
    End If

    ' Require at least one Sign row in alignment 1
    Dim hasSign As Boolean: hasSign = False
    Dim r As Integer
    For r = 1 To alignRowCounts(1)
        If Not alignTypeBoxes(1, r) Is Nothing Then
            If alignTypeBoxes(1, r).Value = "Sign" And Trim(alignLabelBoxes(1, r).Text) <> "" Then
                hasSign = True: Exit For
            End If
        End If
    Next r
    If Not hasSign Then
        MsgBox "Please add at least one Sign row to the Upstream alignment.", vbExclamation
        Exit Sub
    End If

    ' Save spacing values to SharedState
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
    wztcUpTaperBarrier = frameSpacingValues.Controls("txtUpTaperBarrier").Value
    wztcUpTaperBeam = frameSpacingValues.Controls("txtUpTaperBeam").Value

    ' Save user selections
    wztcCategory = cboCategory.Value
    wztcSheet = cboSheet.Value
    wztcSpeed = cboRoadSpeed.Value
    wztcRoadType = cboRoadType.Value
    wztcLaneWidth = cboLaneWidth.Value
    wztcShoulderWidth = cboShoulderWidth.Value

    ' Save multi-alignment data to SharedState
    wztcAlignCount = alignCount
    Dim a As Integer
    For a = 1 To alignCount
        If a = 1 Then
            wztcAlignNames(a) = "Upstream"
        ElseIf a = 2 Then
            wztcAlignNames(a) = "Downstream"
        Else
            wztcAlignNames(a) = "Alignment " & a
        End If
        wztcAlignRowCounts(a) = alignRowCounts(a)
        For r = 1 To alignRowCounts(a)
            wztcAlignRowTypes(a, r) = alignTypeBoxes(a, r).Value
            wztcAlignRowLabels(a, r) = Trim(alignLabelBoxes(a, r).Text)
            wztcAlignRowSpacings(a, r) = Trim(alignSpacingBoxes(a, r).Text)
            wztcAlignRowSizes(a, r) = Trim(alignSizeBoxes(a, r).Text)
            wztcAlignRowSides(a, r) = alignSideBoxes(a, r).Value
        Next r
    Next a

    ' Backward compatibility: populate wztcOrderLabels from alignment 1
    wztcOrderLabelCount = alignRowCounts(1)
    If alignRowCounts(1) > 0 Then
        ReDim wztcOrderLabels(0 To alignRowCounts(1) - 1)
        For r = 1 To alignRowCounts(1)
            wztcOrderLabels(r - 1) = Trim(alignLabelBoxes(1, r).Text)
        Next r
    Else
        ReDim wztcOrderLabels(0 To -1)
    End If

    ' Backward compatibility: populate wztcSignNumbers etc. from Sign rows in alignment 1
    Dim signIdx As Integer: signIdx = 0
    For r = 1 To alignRowCounts(1)
        If Not alignTypeBoxes(1, r) Is Nothing Then
            If alignTypeBoxes(1, r).Value = "Sign" And Trim(alignLabelBoxes(1, r).Text) <> "" Then
                signIdx = signIdx + 1
            End If
        End If
    Next r
    wztcSignCount = signIdx
    If signIdx > 0 Then
        ReDim wztcSignNumbers(1 To signIdx)
        ReDim wztcSignSpacings(1 To signIdx)
        ReDim wztcSignSizes(1 To signIdx)
        ReDim wztcSignSides(1 To signIdx)
        signIdx = 0
        For r = 1 To alignRowCounts(1)
            If Not alignTypeBoxes(1, r) Is Nothing Then
                If alignTypeBoxes(1, r).Value = "Sign" And Trim(alignLabelBoxes(1, r).Text) <> "" Then
                    signIdx = signIdx + 1
                    wztcSignNumbers(signIdx) = Trim(alignLabelBoxes(1, r).Text)
                    wztcSignSpacings(signIdx) = Trim(alignSpacingBoxes(1, r).Text)
                    wztcSignSizes(signIdx) = Replace(Trim(alignSizeBoxes(1, r).Text), "'", Chr(34))
                    wztcSignSides(signIdx) = alignSideBoxes(1, r).Value
                End If
            End If
        Next r
    End If

    ' Confirm
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Submit & Draw will close this form and start alignment drawing mode." & vbCrLf & vbCrLf & _
                     "Are you ready to submit?", vbYesNo + vbQuestion, "Confirm Submit & Draw")
    If confirm = vbNo Then Exit Sub

    Unload Me
    DrawWorkSpace.Show vbModeless
End Sub

' ============================================================
' RESTORE PREVIOUS SUBMISSION STATE
' ============================================================
Private Sub RestoreState()
    If wztcCategory = "" Then Exit Sub
    On Error Resume Next

    ' ---- Dropdowns ----
    Dim i As Integer
    For i = 0 To cboCategory.ListCount - 1
        If cboCategory.List(i) = wztcCategory Then cboCategory.ListIndex = i: Exit For
    Next i
    For i = 0 To cboSheet.ListCount - 1
        If cboSheet.List(i) = wztcSheet Then cboSheet.ListIndex = i: Exit For
    Next i
    For i = 0 To cboRoadSpeed.ListCount - 1
        If cboRoadSpeed.List(i) = wztcSpeed Then cboRoadSpeed.ListIndex = i: Exit For
    Next i
    For i = 0 To cboRoadType.ListCount - 1
        If cboRoadType.List(i) = wztcRoadType Then cboRoadType.ListIndex = i: Exit For
    Next i
    For i = 0 To cboLaneWidth.ListCount - 1
        If cboLaneWidth.List(i) = wztcLaneWidth Then cboLaneWidth.ListIndex = i: Exit For
    Next i
    For i = 0 To cboShoulderWidth.ListCount - 1
        If cboShoulderWidth.List(i) = wztcShoulderWidth Then cboShoulderWidth.ListIndex = i: Exit For
    Next i

    ' ---- Spacing values ----
    frameSpacingValues.Controls("txtDownstreamTaper").Value = wztcDownstreamTaper
    frameSpacingValues.Controls("txtRollAhead").Value = wztcRollAhead
    frameSpacingValues.Controls("txtVehicleSpace").Value = wztcVehicleSpace
    frameSpacingValues.Controls("txtBufferSpace").Value = wztcBufferSpace
    frameSpacingValues.Controls("txtMergingTaper").Value = wztcMergingTaper
    frameSpacingValues.Controls("txtShoulderTapers").Value = wztcShoulderTapers
    frameSpacingValues.Controls("txtAdvancedWarningSpacing").Value = wztcAdvancedWarningSpacing
    frameSpacingValues.Controls("txtSkipLines").Value = wztcSkipLines
    frameSpacingValues.Controls("txtChannelizing").Value = wztcChannelizing
    frameSpacingValues.Controls("txtFlareBarrier").Value = wztcFlareBarrier
    frameSpacingValues.Controls("txtFlareBeam").Value = wztcFlareBeam
    frameSpacingValues.Controls("txtUpTaperBarrier").Value = wztcUpTaperBarrier
    frameSpacingValues.Controls("txtUpTaperBeam").Value = wztcUpTaperBeam

    ' ---- Alignment table rows from saved SharedState ----
    If wztcAlignCount > 0 Then
        ' Create any extra alignments beyond the initial 2
        Do While alignCount < wztcAlignCount
            alignCount = alignCount + 1
            Call CreateAlignSection(alignCount, "Alignment " & alignCount)
        Loop
        Call UpdateFormScrollHeight

        Dim a As Integer, r As Integer
        For a = 1 To wztcAlignCount
            Dim rc As Integer: rc = wztcAlignRowCounts(a)
            If rc > 0 Then
                Dim types()    As String: ReDim types(1 To rc)
                Dim labels()   As String: ReDim labels(1 To rc)
                Dim spacings() As String: ReDim spacings(1 To rc)
                Dim sizes()    As String: ReDim sizes(1 To rc)
                Dim sides()    As String: ReDim sides(1 To rc)
                For r = 1 To rc
                    types(r) = wztcAlignRowTypes(a, r)
                    labels(r) = wztcAlignRowLabels(a, r)
                    spacings(r) = wztcAlignRowSpacings(a, r)
                    sizes(r) = wztcAlignRowSizes(a, r)
                    sides(r) = wztcAlignRowSides(a, r)
                    If sides(r) = "" Then sides(r) = "One Side"
                Next r
                Call RebuildAlignTable(a, types, labels, spacings, sizes, sides, rc)
            End If
        Next a
    End If

    On Error GoTo 0
    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Previous session restored. Review your alignments, then click Submit & Draw."
    End If
End Sub

' ============================================================
' CLEAR ALL BUTTON
' ============================================================
Private Sub btnClear_Click()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Clear all selections and start fresh?" & vbCrLf & _
                 "All alignment rows, spacing values, and sign selections will be removed.", _
                 vbYesNo + vbQuestion, "Clear All")
    If ans = vbNo Then Exit Sub

    wztcCategory = "": wztcSheet = "": wztcSpeed = "": wztcRoadType = ""
    wztcLaneWidth = "": wztcShoulderWidth = ""
    wztcSignCount = 0: wztcOrderLabelCount = 0: wztcAlignCount = 0
    wztcDownstreamTaper = "": wztcRollAhead = "": wztcVehicleSpace = ""
    wztcBufferSpace = "": wztcMergingTaper = "": wztcShoulderTapers = ""
    wztcAdvancedWarningSpacing = "": wztcSkipLines = ""
    wztcChannelizing = "": wztcFlareBarrier = "": wztcFlareBeam = ""

    Unload Me
    WZTCDesigner.Show vbModeless
End Sub
