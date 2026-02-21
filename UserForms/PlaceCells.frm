Option Explicit

' ============================================================
' WZTC CELL LIBRARY FORM
' ------------------------------------------------------------
' Controls to add manually in the VBA IDE form designer:
'   lblCellTitle     - Label          "Select WZTC Symbol:"
'   cmbCellSelect    - ComboBox       (dropdown of cell names)
'   lblCellInstr     - Label          (placement instructions)
'   btnPlaceCell     - CommandButton  "Place Cell"
'   btnFinish        - CommandButton  "Finish"
'   lblStatus        - Label          (status / error messages)
'   lblCalloutTitle  - Label          "Place Callout:"
'   cmbCallout       - ComboBox       (wide dropdown of callout types)
'   lblCalloutInstr  - Label          (callout placement instructions)
'   btnPlaceCallout  - CommandButton  "Place Callout"
'   btnBack          - CommandButton  "< Back"
'   btnReturnToDesigner - CommandButton "Return to Designer"
' ============================================================

Private Function ControlExists(ctrlName As String) As Boolean
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = Me.Controls(ctrlName)
    ControlExists = Not (ctrl Is Nothing)
    On Error GoTo 0
End Function

' ============================================================
' INITIALIZE - populate ComboBoxes and lay out controls
' ============================================================
Private Sub UserForm_Initialize()
    Me.Caption = "WZTC Cell Library"
    Me.Width  = 570
    Me.Height = 350

    ' ========== TITLE LABEL ==========
    If ControlExists("lblCellTitle") Then
        lblCellTitle.Caption   = "Select WZTC Symbol:"
        lblCellTitle.Top       = 8
        lblCellTitle.Left      = 10
        lblCellTitle.Width     = 540
        lblCellTitle.Height    = 16
        lblCellTitle.Font.Size = 9
        lblCellTitle.Font.Bold = True
    End If

    ' ========== CELL COMBO BOX ==========
    If ControlExists("cmbCellSelect") Then
        cmbCellSelect.Top    = 27
        cmbCellSelect.Left   = 10
        cmbCellSelect.Width  = 540
        cmbCellSelect.Height = 20
        cmbCellSelect.Style  = 2   ' fmStyleDropDownList (read-only)

        ' Populate from catalogue
        Dim cats() As String
        cats = GetCellCatalogue()
        Dim i As Integer
        For i = 1 To UBound(cats)
            cmbCellSelect.AddItem cats(i)
        Next i
        cmbCellSelect.ListIndex = 0
    End If

    ' ========== INSTRUCTION LABEL ==========
    If ControlExists("lblCellInstr") Then
        lblCellInstr.Caption   = "Click 'Place Cell', then click in MicroStation to place the symbol. Right-click to stop placing."
        lblCellInstr.Top       = 55
        lblCellInstr.Left      = 10
        lblCellInstr.Width     = 540
        lblCellInstr.Height    = 34
        lblCellInstr.Font.Size = 8
        lblCellInstr.WordWrap  = True
        lblCellInstr.ForeColor = RGB(80, 80, 80)
    End If

    ' ========== ACTION BUTTONS ==========
    If ControlExists("btnPlaceCell") Then
        btnPlaceCell.Caption   = "Place Cell"
        btnPlaceCell.Top       = 98
        btnPlaceCell.Left      = 10
        btnPlaceCell.Width     = 100
        btnPlaceCell.Height    = 23
        btnPlaceCell.Font.Bold = True
    End If

    If ControlExists("btnFinish") Then
        btnFinish.Caption = "Finish"
        btnFinish.Top     = 98
        btnFinish.Left    = 118
        btnFinish.Width   = 80
        btnFinish.Height  = 23
    End If

    ' ========== STATUS LABEL ==========
    If ControlExists("lblStatus") Then
        lblStatus.Caption   = "Select a symbol from the dropdown, then click 'Place Cell'."
        lblStatus.Top       = 132
        lblStatus.Left      = 10
        lblStatus.Width     = 540
        lblStatus.Height    = 50
        lblStatus.Font.Size = 8
        lblStatus.WordWrap  = True
    End If

    ' ========== CALLOUT SECTION TITLE ==========
    If ControlExists("lblCalloutTitle") Then
        lblCalloutTitle.Caption   = "Place Callout:"
        lblCalloutTitle.Top       = 195
        lblCalloutTitle.Left      = 10
        lblCalloutTitle.Width     = 540
        lblCalloutTitle.Height    = 16
        lblCalloutTitle.Font.Size = 9
        lblCalloutTitle.Font.Bold = True
    End If

    ' ========== CALLOUT COMBO BOX (wide — text can be long) ==========
    If ControlExists("cmbCallout") Then
        cmbCallout.Top    = 215
        cmbCallout.Left   = 10
        cmbCallout.Width  = 540
        cmbCallout.Height = 20
        cmbCallout.Style  = 2   ' fmStyleDropDownList (read-only)

        cmbCallout.AddItem "DOWNSTREAM TAPER"
        cmbCallout.AddItem "WORK AREA"
        cmbCallout.AddItem "VEHICLE SPACE"
        cmbCallout.AddItem "BUFFER SPACE"
        cmbCallout.AddItem "MERGING TAPER"
        cmbCallout.AddItem "SHOULDER TAPER"
        cmbCallout.AddItem "SHIFTING TAPER"
        cmbCallout.AddItem "LANE SHIFT"
        cmbCallout.AddItem "MULTILANE SHIFT"
        cmbCallout.AddItem "TEMPORARY CONCRETE BARRIER (UNPINNED) TO REMAIN IN PLACE ITEM 619.090100"
        cmbCallout.AddItem "INTERIM PAVEMENT MARKING STRIPE (TRAFFIC PAINT) 6" & Chr(34) & " YELLOW LANE LINE ITEM 619.100101"
        cmbCallout.AddItem "INTERIM PAVEMENT MARKING STRIPE (TRAFFIC PAINT) 4" & Chr(34) & " SOLID WHITE LINE"
        cmbCallout.AddItem "PAVEMENT MARKING LIMIT MEET EXISTING MARKINGS"
        cmbCallout.AddItem "TEMPORARY IMPACT ATTENUOATOR ITEM 619.1803"
        cmbCallout.AddItem "TYPE III BARRICADE"
        cmbCallout.AddItem "TEMPORARY POSITIVE BARRIER (PINNING PROHIBITED) (TYP.) ITEM 619.1711"
        cmbCallout.AddItem "REMOVE EXISTING PAVEMENT MARKINGS CONFLICTING WITH INTERIM SHIFT MARKINGS"
        cmbCallout.AddItem "CHANNELIZING DEVICES SPACED @ 10' O.C."
        cmbCallout.AddItem "CHANNELIZING DEVICES SPACED @ 20' O.C."
        cmbCallout.AddItem "CHANNELIZING DEVICES SPACED @ 40' O.C."
        cmbCallout.AddItem "MAINTAIN ONE XX' LANE (TYP.)"
        cmbCallout.ListIndex = 0
    End If

    ' ========== CALLOUT INSTRUCTION LABEL ==========
    If ControlExists("lblCalloutInstr") Then
        lblCalloutInstr.Caption   = "Select a callout type, click 'Place Callout', then click three points in MicroStation: (1) note tip, (2) leader elbow, (3) text position."
        lblCalloutInstr.Top       = 243
        lblCalloutInstr.Left      = 10
        lblCalloutInstr.Width     = 540
        lblCalloutInstr.Height    = 30
        lblCalloutInstr.Font.Size = 8
        lblCalloutInstr.WordWrap  = True
        lblCalloutInstr.ForeColor = RGB(80, 80, 80)
    End If

    ' ========== PLACE CALLOUT BUTTON ==========
    If ControlExists("btnPlaceCallout") Then
        btnPlaceCallout.Caption   = "Place Callout"
        btnPlaceCallout.Top       = 280
        btnPlaceCallout.Left      = 10
        btnPlaceCallout.Width     = 130
        btnPlaceCallout.Height    = 23
        btnPlaceCallout.Font.Bold = True
    End If

    ' ========== NAVIGATION BUTTONS ==========
    If ControlExists("btnBack") Then
        btnBack.Caption   = "< Back"
        btnBack.Top       = 312
        btnBack.Left      = 10
        btnBack.Width     = 90
        btnBack.Height    = 23
    End If

    If ControlExists("btnReturnToDesigner") Then
        btnReturnToDesigner.Caption = "Return to Designer"
        btnReturnToDesigner.Top     = 312
        btnReturnToDesigner.Left    = 108
        btnReturnToDesigner.Width   = 145
        btnReturnToDesigner.Height  = 23
    End If

    Me.Height = 348
End Sub

' ============================================================
' PLACE CELL - hide form, place selected cell, re-show
' ============================================================
Private Sub btnPlaceCell_Click()
    On Error GoTo PlaceErr

    If Not ControlExists("cmbCellSelect") Then Exit Sub
    If cmbCellSelect.ListIndex < 0 Then
        If ControlExists("lblStatus") Then lblStatus.Caption = "Please select a symbol from the dropdown first."
        Exit Sub
    End If

    Dim entry As String
    entry = cmbCellSelect.Value
    Dim cellName As String
    cellName = ExtractCellName(entry)

    If ControlExists("lblStatus") Then lblStatus.Caption = "Placing " & cellName & " — click in MicroStation. Right-click to stop."

    ' Hide form so MicroStation gets full mouse focus for cell placement
    Me.Hide
    PlaceSelectedCell cellName
    Me.Show

    ' Increment count for this cell type
    Dim cellIdx As Integer
    cellIdx = cmbCellSelect.ListIndex + 1   ' 1-based to match catalogue array
    If cellIdx >= 1 And cellIdx <= 16 Then
        wztcCellPlacementCounts(cellIdx) = wztcCellPlacementCounts(cellIdx) + 1
    End If

    Dim placedCount As Integer
    If cellIdx >= 1 And cellIdx <= 16 Then
        placedCount = wztcCellPlacementCounts(cellIdx)
    End If

    If ControlExists("lblStatus") Then
        lblStatus.Caption = cellName & " placed (" & placedCount & " total). Select another symbol or click 'Finish'."
    End If
    If ControlExists("btnPlaceCell") Then btnPlaceCell.Enabled = True
    If ControlExists("btnFinish") Then btnFinish.Enabled = True
    If ControlExists("btnBack") Then btnBack.Enabled = True
    If ControlExists("btnReturnToDesigner") Then btnReturnToDesigner.Enabled = True
    Exit Sub

PlaceErr:
    If ControlExists("lblStatus") Then lblStatus.Caption = "Error placing cell: " & Err.Description
    If ControlExists("btnPlaceCell") Then btnPlaceCell.Enabled = True
    If ControlExists("btnFinish") Then btnFinish.Enabled = True
    If ControlExists("btnBack") Then btnBack.Enabled = True
    If ControlExists("btnReturnToDesigner") Then btnReturnToDesigner.Enabled = True
End Sub

' ============================================================
' PLACE CALLOUT - hide form, collect 3 user clicks, place note
' ============================================================
Private Sub btnPlaceCallout_Click()
    On Error GoTo CalloutErr

    If Not ControlExists("cmbCallout") Then Exit Sub
    If cmbCallout.ListIndex < 0 Then
        If ControlExists("lblStatus") Then lblStatus.Caption = "Please select a callout type first."
        Exit Sub
    End If

    Dim selectedCallout As String
    selectedCallout = cmbCallout.Value

    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Click in MicroStation: (1) note tip, (2) leader elbow, (3) text position."
    End If

    Me.Hide

    ' Start the Place Note command
    CadInputQueue.SendCommand "TEXTEDITOR PLACENOTE"

    ' Collect first click: note tip / arrowhead origin
    Dim oMsg As CadInputMessage
    Set oMsg = CadInputQueue.GetInput
    Do While oMsg.InputType <> msdCadInputTypeDataPoint
        If oMsg.InputType = msdCadInputTypeReset Then
            CadInputQueue.SendReset
            CommandState.StartDefaultCommand
            GoTo CalloutDone
        End If
        Set oMsg = CadInputQueue.GetInput
    Loop
    CadInputQueue.SendDataPoint oMsg.Point, 1
    CadInputQueue.SendReset

    ' Insert the selected callout text.
    ' Escape any embedded double-quotes by doubling them (MicroStation keyin convention).
    Dim escapedText As String
    escapedText = Replace(selectedCallout, Chr(34), Chr(34) & Chr(34))
    CadInputQueue.SendKeyin "TEXTEDITOR PLAYCOMMAND INSERT_TEXT """ & escapedText & """"

    ' Collect second click: leader elbow point
    Set oMsg = CadInputQueue.GetInput
    Do While oMsg.InputType <> msdCadInputTypeDataPoint
        If oMsg.InputType = msdCadInputTypeReset Then
            CadInputQueue.SendReset
            CommandState.StartDefaultCommand
            GoTo CalloutDone
        End If
        Set oMsg = CadInputQueue.GetInput
    Loop
    CadInputQueue.SendDataPoint oMsg.Point, 1

    ' Collect third click: text box position
    Set oMsg = CadInputQueue.GetInput
    Do While oMsg.InputType <> msdCadInputTypeDataPoint
        If oMsg.InputType = msdCadInputTypeReset Then
            CadInputQueue.SendReset
            CommandState.StartDefaultCommand
            GoTo CalloutDone
        End If
        Set oMsg = CadInputQueue.GetInput
    Loop
    CadInputQueue.SendDataPoint oMsg.Point, 1

    CadInputQueue.SendReset
    CommandState.StartDefaultCommand

CalloutDone:
    Me.Show
    If ControlExists("lblStatus") Then lblStatus.Caption = "Callout placed. Select another or click Finish."
    Exit Sub

CalloutErr:
    CadInputQueue.SendReset
    CommandState.StartDefaultCommand
    Me.Show
    If ControlExists("lblStatus") Then lblStatus.Caption = "Error placing callout: " & Err.Description
End Sub

' ============================================================
' FINISH - delete perp reference lines, then close
' ============================================================
Private Sub btnFinish_Click()
    ' Remove the perpendicular reference lines that were placed during
    ' the alignment placement step — they are no longer needed in the final drawing.
    DeletePerpLines

    MsgBox "WZTC design complete!" & vbCrLf & vbCrLf & _
           "All elements have been placed. You may save the design file.", _
           vbInformation, "Done"
    Unload Me
End Sub

' ============================================================
' NAVIGATION - BACK AND RETURN TO DESIGNER
' ============================================================
Private Sub btnBack_Click()
    Unload Me
    StartWZTCElementsPlacement
End Sub

Private Sub btnReturnToDesigner_Click()
    Unload Me
    WZTCDesigner.Show vbModeless
End Sub

' ============================================================
' CONFIRM CLOSE VIA X BUTTON
' ============================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Close WZTC Cell Library tool?", vbYesNo + vbQuestion, "Close")
        If ans = vbNo Then Cancel = True
    End If
End Sub
