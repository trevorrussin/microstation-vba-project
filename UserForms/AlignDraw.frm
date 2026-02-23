Option Explicit

' ============================================================
' ALIGNDRAW FORM
' ------------------------------------------------------------
' Allows the user to draw line/arc alignment segments for each
' alignment (Upstream, Downstream, plus any extras added in
' WZTCDesigner). Controls added manually in the VBA IDE:
'
'   cboAlignSelect  - ComboBox   (alignment name dropdown)
'   optLine         - OptionButton "Line"
'   optArc          - OptionButton "Arc"
'   cmdStartSegment - CommandButton "Start Segment"
'   cmdCommit       - CommandButton "Commit This Alignment"
'   cmdNextStep     - CommandButton "Next: Place Perp Lines >"
'   lblStatus       - Label (status/instruction text)
'   lblRightClick   - Label (right-click tip)
'   btnBack         - CommandButton "< Back"
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
' INITIALIZE
' ============================================================
Private Sub UserForm_Initialize()
    Me.Caption = "Draw Alignments"
    Me.Width  = 230
    Me.Height = 420

    ' ---- Alignment selector dropdown ----
    If ControlExists("cboAlignSelect") Then
        cboAlignSelect.Top   = 12
        cboAlignSelect.Left  = 10
        cboAlignSelect.Width = 200
        cboAlignSelect.Height = 20
        ' Populate from SharedState alignment names
        cboAlignSelect.Clear
        Dim i As Integer
        Dim cnt As Integer
        cnt = wztcAlignCount
        If cnt < 1 Then cnt = 2  ' default: Upstream + Downstream
        For i = 1 To cnt
            If Len(Trim(wztcAlignNames(i))) > 0 Then
                cboAlignSelect.AddItem wztcAlignNames(i)
            Else
                If i = 1 Then
                    cboAlignSelect.AddItem "Upstream Alignment"
                ElseIf i = 2 Then
                    cboAlignSelect.AddItem "Downstream Alignment"
                Else
                    cboAlignSelect.AddItem "Alignment " & i
                End If
            End If
        Next i
        cboAlignSelect.ListIndex = 0
    End If

    ' ---- Line / Arc radio buttons ----
    If ControlExists("optLine") Then
        optLine.Caption = "Line"
        optLine.Top = 42: optLine.Left = 10: optLine.Width = 80
        optLine.Value = True
    End If
    If ControlExists("optArc") Then
        optArc.Caption = "Arc"
        optArc.Top = 64: optArc.Left = 10: optArc.Width = 80
    End If

    ' ---- Start Segment ----
    If ControlExists("cmdStartSegment") Then
        cmdStartSegment.Caption = "Start Segment"
        cmdStartSegment.Top = 92: cmdStartSegment.Left = 10
        cmdStartSegment.Width = 200: cmdStartSegment.Height = 28
        cmdStartSegment.Font.Bold = True
    End If

    ' ---- Commit This Alignment ----
    If ControlExists("cmdCommit") Then
        cmdCommit.Caption = "Commit This Alignment"
        cmdCommit.Top = 128: cmdCommit.Left = 10
        cmdCommit.Width = 200: cmdCommit.Height = 28
    End If

    ' ---- Next Step ----
    If ControlExists("cmdNextStep") Then
        cmdNextStep.Caption = "Next: Place Perp Lines >"
        cmdNextStep.Top = 163: cmdNextStep.Left = 10
        cmdNextStep.Width = 200: cmdNextStep.Height = 28
        cmdNextStep.Font.Bold = True
        cmdNextStep.Enabled = False   ' enabled after first commit
    End If

    ' ---- Status label ----
    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Select an alignment, draw segments, then click 'Commit'."
        lblStatus.Top = 200: lblStatus.Left = 10
        lblStatus.Width = 205: lblStatus.Height = 48
        lblStatus.WordWrap = True
        lblStatus.ForeColor = RGB(0, 0, 160)
    End If

    ' ---- Right-click tip ----
    If ControlExists("lblRightClick") Then
        lblRightClick.Caption = "Tip: Right-click in MicroStation to finish each segment."
        lblRightClick.Top = 252: lblRightClick.Left = 10
        lblRightClick.Width = 205: lblRightClick.Height = 30
        lblRightClick.Font.Size = 8
        lblRightClick.WordWrap = True
        lblRightClick.ForeColor = RGB(80, 80, 80)
    End If

    ' ---- Navigation buttons ----
    If ControlExists("btnBack") Then
        btnBack.Caption = "< Back"
        btnBack.Top = 292: btnBack.Left = 10
        btnBack.Width = 95: btnBack.Height = 22
    End If
    If ControlExists("btnReturnToDesigner") Then
        btnReturnToDesigner.Caption = "Return to Designer"
        btnReturnToDesigner.Top = 292: btnReturnToDesigner.Left = 115
        btnReturnToDesigner.Width = 100: btnReturnToDesigner.Height = 22
    End If

    ' Initialize AlignmentTool to alignment 1
    Call SetCurrentAlignment(1)
End Sub

' ============================================================
' ALIGNMENT DROPDOWN CHANGE — switch active alignment
' ============================================================
Private Sub cboAlignSelect_Change()
    Dim aIdx As Integer
    aIdx = cboAlignSelect.ListIndex + 1
    If aIdx >= 1 Then
        Call SetCurrentAlignment(aIdx)
        If ControlExists("lblStatus") Then
            Dim nm As String
            If aIdx <= cboAlignSelect.ListCount Then nm = cboAlignSelect.List(aIdx - 1)
            If wztcAlignDrawn(aIdx) Then
                lblStatus.Caption = nm & " already committed. Draw more segments and commit again to add to it."
            Else
                lblStatus.Caption = "Drawing: " & nm & ". Click 'Start Segment' to draw."
            End If
        End If
    End If
End Sub

' ============================================================
' START SEGMENT — hide form, call StartLineSegment or StartArcSegment
' ============================================================
Private Sub cmdStartSegment_Click()
    Me.Hide

    If ControlExists("optArc") Then
        If optArc.Value Then
            StartArcSegment
        Else
            StartLineSegment
        End If
    Else
        StartLineSegment
    End If

    Me.Show
End Sub

' ============================================================
' COMMIT THIS ALIGNMENT
' ============================================================
Private Sub cmdCommit_Click()
    Call CommitCurrentAlignment

    ' Enable Next Step button if at least one alignment is drawn
    Dim i As Integer
    For i = 1 To 10
        If wztcAlignDrawn(i) Then
            If ControlExists("cmdNextStep") Then cmdNextStep.Enabled = True
            Exit For
        End If
    Next i

    If ControlExists("lblStatus") Then
        Dim aIdx As Integer
        aIdx = cboAlignSelect.ListIndex + 1
        Dim nm As String
        If ControlExists("cboAlignSelect") And aIdx >= 1 And aIdx <= cboAlignSelect.ListCount Then
            nm = cboAlignSelect.List(aIdx - 1)
        Else
            nm = "Alignment " & aIdx
        End If
        If wztcAlignDrawn(aIdx) Then
            lblStatus.Caption = nm & " committed! Select another alignment or click 'Next: Place Perp Lines'."
            lblStatus.ForeColor = RGB(0, 120, 0)
        End If
    End If
End Sub

' ============================================================
' NEXT STEP — launch perp placement for all committed alignments
' ============================================================
Private Sub cmdNextStep_Click()
    On Error Resume Next
    CadInputQueue.SendReset
    CommandState.StartDefaultCommand
    Unload Me
    StartAlignmentPlacement   ' PerpPlacement.bas
End Sub

' ============================================================
' NAVIGATION
' ============================================================
Private Sub btnBack_Click()
    Unload Me
    WZTCDesigner.Show vbModeless
End Sub

Private Sub btnReturnToDesigner_Click()
    Unload Me
    WZTCDesigner.Show vbModeless
End Sub
