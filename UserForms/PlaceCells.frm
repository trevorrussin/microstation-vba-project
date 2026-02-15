Option Explicit

' ============================================================
' WZTC CELL LIBRARY FORM
' ------------------------------------------------------------
' Controls to add manually in the VBA IDE form designer:
'   lblCellTitle  - Label          "Select WZTC Symbol:"
'   cmbCellSelect - ComboBox       (dropdown of cell names)
'   lblCellInstr  - Label          (placement instructions)
'   btnPlaceCell  - CommandButton  "Place Cell"
'   btnFinish     - CommandButton  "Finish"
'   lblStatus     - Label          (status / error messages)
' ============================================================

Private Function ControlExists(ctrlName As String) As Boolean
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = Me.Controls(ctrlName)
    ControlExists = Not (ctrl Is Nothing)
    On Error GoTo 0
End Function

' ============================================================
' INITIALIZE - populate ComboBox and lay out controls
' ============================================================
Private Sub UserForm_Initialize()
    Me.Caption = "WZTC Cell Library"
    Me.Width  = 320
    Me.Height = 235

    ' ========== TITLE LABEL ==========
    If ControlExists("lblCellTitle") Then
        lblCellTitle.Caption   = "Select WZTC Symbol:"
        lblCellTitle.Top       = 8
        lblCellTitle.Left      = 10
        lblCellTitle.Width     = 290
        lblCellTitle.Height    = 16
        lblCellTitle.Font.Size = 9
        lblCellTitle.Font.Bold = True
    End If

    ' ========== COMBO BOX ==========
    If ControlExists("cmbCellSelect") Then
        cmbCellSelect.Top    = 27
        cmbCellSelect.Left   = 10
        cmbCellSelect.Width  = 290
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
        lblCellInstr.Width     = 290
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
        lblStatus.Width     = 290
        lblStatus.Height    = 50
        lblStatus.Font.Size = 8
        lblStatus.WordWrap  = True
    End If

    ' ========== NAVIGATION BUTTONS ==========
    If ControlExists("btnBack") Then
        btnBack.Caption   = "< Back"
        btnBack.Top       = 190
        btnBack.Left      = 10
        btnBack.Width     = 90
        btnBack.Height    = 23
    End If

    If ControlExists("btnReturnToDesigner") Then
        btnReturnToDesigner.Caption = "Return to Designer"
        btnReturnToDesigner.Top     = 190
        btnReturnToDesigner.Left    = 108
        btnReturnToDesigner.Width   = 145
        btnReturnToDesigner.Height  = 23
    End If

    Me.Height = 240
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
' FINISH - close the form; workflow is complete
' ============================================================
Private Sub btnFinish_Click()
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
    frmWorkzoneDesigner.Show vbModeless
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
