Option Explicit

' ============================================================
' WZTC ALIGNMENT PLACEMENT FORM
' ------------------------------------------------------------
' Controls to add manually in the VBA IDE form designer:
'   lblAlignProgress - Label       (shows current alignment name)
'   lblItemOf        - Label       (counter, e.g. "Item 1 of 5:")
'   lblItemName      - Label       (item name, large/bold)
'   lblSpacingHint   - Label       "Spacing to this item (ft):"
'   txtSpacing       - TextBox     (editable spacing value)
'   lblCumulative    - Label       (next position along alignment)
'   lblTotalLen      - Label       (total alignment length)
'   btnPlace         - CommandButton  "Place Line"
'   btnSkip          - CommandButton  "Skip"
'   btnCancel        - CommandButton  "Cancel All"
'   btnNext          - CommandButton  "Next: Draw Signs"
'   btnNextAlign     - CommandButton  "Next Alignment >"
'   lblStatus        - Label       (status / error messages)
' ============================================================

Private Function ControlExists(ctrlName As String) As Boolean
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = Me.Controls(ctrlName)
    ControlExists = Not (ctrl Is Nothing)
    On Error GoTo 0
End Function

' ============================================================
' INITIALIZE
' ============================================================
Private Sub UserForm_Initialize()
    Me.Caption = "WZTC Alignment Placement"
    Me.Width  = 360
    Me.Height = 250

    ' ========== ALIGNMENT PROGRESS LABEL (shows current alignment name) ==========
    If ControlExists("lblAlignProgress") Then
        lblAlignProgress.Caption   = ""
        lblAlignProgress.Top       = 8
        lblAlignProgress.Left      = 10
        lblAlignProgress.Width     = 330
        lblAlignProgress.Height    = 14
        lblAlignProgress.Font.Size = 8
        lblAlignProgress.ForeColor = RGB(0, 80, 160)
    End If

    ' ========== ITEM COUNTER LABEL ==========
    If ControlExists("lblItemOf") Then
        lblItemOf.Caption   = "Initialising..."
        lblItemOf.Top       = 24
        lblItemOf.Left      = 10
        lblItemOf.Width     = 330
        lblItemOf.Height    = 16
        lblItemOf.Font.Size = 9
    End If

    ' ========== ITEM NAME LABEL (large, bold) ==========
    If ControlExists("lblItemName") Then
        lblItemName.Caption   = ""
        lblItemName.Top       = 42
        lblItemName.Left      = 10
        lblItemName.Width     = 330
        lblItemName.Height    = 22
        lblItemName.Font.Size = 11
        lblItemName.Font.Bold = True
        lblItemName.ForeColor = RGB(0, 0, 160)
    End If

    ' ========== SPACING ROW ==========
    If ControlExists("lblSpacingHint") Then
        lblSpacingHint.Caption   = "Spacing to this item (ft):"
        lblSpacingHint.Top       = 76
        lblSpacingHint.Left      = 10
        lblSpacingHint.Width     = 160
        lblSpacingHint.Height    = 16
        lblSpacingHint.Font.Size = 9
    End If

    If ControlExists("txtSpacing") Then
        txtSpacing.Text      = "0.0"
        txtSpacing.Top       = 73
        txtSpacing.Left      = 175
        txtSpacing.Width     = 70
        txtSpacing.Height    = 20
        txtSpacing.Font.Size = 9
    End If

    ' ========== PROGRESS LABELS ==========
    If ControlExists("lblCumulative") Then
        lblCumulative.Caption   = "Next position along alignment:  0.0 ft"
        lblCumulative.Top       = 101
        lblCumulative.Left      = 10
        lblCumulative.Width     = 330
        lblCumulative.Height    = 14
        lblCumulative.Font.Size = 8
        lblCumulative.ForeColor = RGB(0, 120, 0)
    End If

    If ControlExists("lblTotalLen") Then
        lblTotalLen.Caption   = "Total alignment length:  " & _
                                Format(GetTotalPathLength(), "0.0") & " ft"
        lblTotalLen.Top       = 117
        lblTotalLen.Left      = 10
        lblTotalLen.Width     = 330
        lblTotalLen.Height    = 14
        lblTotalLen.Font.Size = 8
        lblTotalLen.ForeColor = RGB(100, 100, 100)
    End If

    ' ========== ACTION BUTTONS ==========
    If ControlExists("btnPlace") Then
        btnPlace.Caption   = "Place Line"
        btnPlace.Top       = 141
        btnPlace.Left      = 10
        btnPlace.Width     = 90
        btnPlace.Height    = 23
        btnPlace.Font.Bold = True
    End If

    If ControlExists("btnSkip") Then
        btnSkip.Caption = "Skip"
        btnSkip.Top     = 141
        btnSkip.Left    = 108
        btnSkip.Width   = 70
        btnSkip.Height  = 23
    End If

    If ControlExists("btnCancel") Then
        btnCancel.Caption = "Cancel All"
        btnCancel.Top     = 141
        btnCancel.Left    = 186
        btnCancel.Width   = 80
        btnCancel.Height  = 23
    End If

    If ControlExists("btnNext") Then
        btnNext.Caption    = "Next: Draw Signs"
        btnNext.Top        = 141
        btnNext.Left       = 274
        btnNext.Width      = 76
        btnNext.Height     = 23
        btnNext.Font.Bold  = True
        btnNext.Enabled    = False
    End If

    ' ========== NEXT ALIGNMENT BUTTON (hidden; shown when current alignment is done but more remain) ==========
    If ControlExists("btnNextAlign") Then
        btnNextAlign.Caption   = "Next Alignment >"
        btnNextAlign.Top       = 141
        btnNextAlign.Left      = 274
        btnNextAlign.Width     = 76
        btnNextAlign.Height    = 23
        btnNextAlign.Font.Bold = True
        btnNextAlign.Visible   = False
    End If

    ' ========== STATUS LABEL ==========
    If ControlExists("lblStatus") Then
        lblStatus.Caption   = "Ready"
        lblStatus.Top       = 174
        lblStatus.Left      = 10
        lblStatus.Width     = 330
        lblStatus.Height    = 52
        lblStatus.Font.Size = 8
        lblStatus.WordWrap  = True
    End If

    ' ========== NAVIGATION BUTTONS ==========
    If ControlExists("btnBack") Then
        btnBack.Caption   = "< Back"
        btnBack.Top       = 234
        btnBack.Left      = 10
        btnBack.Width     = 90
        btnBack.Height    = 23
    End If

    If ControlExists("btnReturnToDesigner") Then
        btnReturnToDesigner.Caption = "Return to Designer"
        btnReturnToDesigner.Top     = 234
        btnReturnToDesigner.Left    = 108
        btnReturnToDesigner.Width   = 145
        btnReturnToDesigner.Height  = 23
    End If

    Me.Height = 283
    Call RefreshDisplay
End Sub

' ============================================================
' BUTTON EVENTS
' ============================================================
Private Sub btnPlace_Click()
    On Error GoTo PlaceErr

    Dim spacing As Double
    spacing = ParseSpacing(txtSpacing.Text)
    If spacing < 0 Then
        If ControlExists("lblStatus") Then lblStatus.Caption = "Please enter a spacing of 0 or greater."
        Exit Sub
    End If

    PlaceLineForCurrentItem spacing

    If IsAllDone() Then
        Call ShowAllDone
    Else
        Call RefreshDisplay
    End If
    Exit Sub

PlaceErr:
    If ControlExists("lblStatus") Then lblStatus.Caption = "Error placing line: " & Err.Description
End Sub

Private Sub btnSkip_Click()
    If ControlExists("lblStatus") Then lblStatus.Caption = "Skipped: " & GetCurrentItemLabel()

    SkipCurrentItem

    If IsAllDone() Then
        Call ShowAllDone
    Else
        Call RefreshDisplay
    End If
End Sub

Private Sub btnCancel_Click()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Cancel placement?" & vbCrLf & _
                 "Lines placed so far will remain in the drawing.", _
                 vbYesNo + vbQuestion, "Cancel Placement")
    If ans = vbYes Then
        Unload Me
    End If
End Sub

Private Sub btnNext_Click()
    Unload Me
    StartSignPlacement
End Sub

' ============================================================
' REFRESH DISPLAY FOR CURRENT ITEM
' ============================================================
Private Sub RefreshDisplay()
    If IsAllDone() Then
        Call ShowAllDone
        Exit Sub
    End If

    Dim idx   As Integer
    Dim lbl   As String
    Dim sp    As Double
    Dim nextP As Double

    idx   = GetCurrentItemNumber()
    lbl   = GetCurrentItemLabel()
    sp    = GetCurrentItemSuggestedSpacing()
    nextP = GetCurrentPosition() + sp

    If ControlExists("lblAlignProgress") Then
        lblAlignProgress.Caption = GetCurrentAlignmentName() & " — Item " & idx & " of " & GetTotalItemCount()
    End If
    If ControlExists("btnNextAlign") Then btnNextAlign.Visible = False
    If ControlExists("lblItemOf") Then lblItemOf.Caption = "Item " & idx & " of " & GetTotalItemCount() & ":"
    If ControlExists("lblItemName") Then lblItemName.Caption = lbl
    If ControlExists("txtSpacing") Then txtSpacing.Text = Format(sp, "0.0")

    If ControlExists("lblCumulative") Then
        lblCumulative.Caption = _
            "Next position along alignment:  " & Format(nextP, "0.0") & " ft" & _
            "   (current:  " & Format(GetCurrentPosition(), "0.0") & " ft)"
    End If

    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Confirm the spacing (or type a new value), " & _
                            "then click 'Place Line'."
    End If

    If ControlExists("btnPlace") Then btnPlace.Enabled = True
    If ControlExists("btnSkip") Then btnSkip.Enabled = True
    If ControlExists("btnNext") Then btnNext.Enabled = False
End Sub

' ============================================================
' SHOW COMPLETION STATE
' ============================================================
Private Sub ShowAllDone()
    If ControlExists("lblItemOf") Then lblItemOf.Caption = "Complete!"
    If ControlExists("lblItemName") Then
        lblItemName.Caption = "All " & GetTotalItemCount() & " reference lines placed."
        lblItemName.ForeColor = RGB(0, 140, 0)
    End If
    If ControlExists("txtSpacing") Then txtSpacing.Text = ""
    If ControlExists("lblCumulative") Then
        lblCumulative.Caption = "Total distance used:  " & _
                                Format(GetCurrentPosition(), "0.0") & " ft"
    End If

    If Not IsLastAlignment() Then
        ' More alignments remain — show Next Alignment button, hide Next: Draw Signs
        If ControlExists("lblAlignProgress") Then
            lblAlignProgress.Caption = GetCurrentAlignmentName() & " complete."
        End If
        If ControlExists("lblStatus") Then
            lblStatus.Caption = GetCurrentAlignmentName() & " done. Click 'Next Alignment >' to continue."
        End If
        If ControlExists("btnNextAlign") Then btnNextAlign.Visible = True
        If ControlExists("btnNext") Then btnNext.Enabled = False
    ElseIf wztcPlacedSignCount > 0 Then
        ' All alignments done and signs were placed
        If ControlExists("lblAlignProgress") Then
            lblAlignProgress.Caption = "All alignments complete."
        End If
        If ControlExists("lblStatus") Then
            lblStatus.Caption = wztcPlacedSignCount & " sign line(s) placed." & vbCrLf & _
                                "Click 'Next: Draw Signs' to place sign graphics."
        End If
        If ControlExists("btnNextAlign") Then btnNextAlign.Visible = False
        If ControlExists("btnNext") Then btnNext.Enabled = True
    Else
        ' All alignments done, no signs — still allow the user to proceed or go back
        If ControlExists("lblAlignProgress") Then
            lblAlignProgress.Caption = "All alignments complete."
        End If
        If ControlExists("lblStatus") Then
            lblStatus.Caption = "All reference lines placed (no signs in this WZTC order). " & _
                                "Click 'Next: Draw Signs' to continue, or use the Back / Return buttons below."
        End If
        If ControlExists("btnNextAlign") Then btnNextAlign.Visible = False
        If ControlExists("btnNext") Then btnNext.Enabled = True
    End If

    If ControlExists("btnPlace") Then btnPlace.Enabled = False
    If ControlExists("btnSkip") Then btnSkip.Enabled = False
End Sub

' ============================================================
' NEXT ALIGNMENT - advance to the next committed alignment
' ============================================================
Private Sub btnNextAlign_Click()
    Call AdvanceToNextAlignment
    If ControlExists("btnNextAlign") Then btnNextAlign.Visible = False
    If ControlExists("lblTotalLen") Then
        lblTotalLen.Caption = "Total alignment length:  " & _
                              Format(GetTotalPathLength(), "0.0") & " ft"
    End If
    Call RefreshDisplay
End Sub

' ============================================================
' PARSE SPACING TEXT BOX - returns -1 on invalid input
' ============================================================
Private Function ParseSpacing(txt As String) As Double
    Dim v As Double
    On Error GoTo BadInput
    v = CDbl(txt)
    If v < 0 Then GoTo BadInput
    ParseSpacing = v
    Exit Function
BadInput:
    ParseSpacing = -1
End Function

' ============================================================
' NAVIGATION - BACK AND RETURN TO DESIGNER
' ============================================================
Private Sub btnBack_Click()
    Unload Me
    AlignDraw.Show vbModeless
End Sub

Private Sub btnReturnToDesigner_Click()
    Unload Me
    WZTCDesigner.Show vbModeless
End Sub

' ============================================================
' CONFIRM CLOSE VIA X BUTTON
' ============================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then    ' user clicked the X
        If Not IsAllDone() Then
            Dim ans As VbMsgBoxResult
            ans = MsgBox("Close the placement tool?" & vbCrLf & _
                         "Lines placed so far will remain.", _
                         vbYesNo + vbQuestion, "Close")
            If ans = vbNo Then Cancel = True
        End If
    End If
End Sub
