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
'   cmdCommitAll    - CommandButton "Commit All Alignments"
'   cmdNextStep     - CommandButton "Next: Place Perp Lines >"
'   lblStatus       - Label (status/instruction text)
'   lblRightClick   - Label (right-click tip)
'   btnBack         - CommandButton "< Back"
'   btnReturnToDesigner - CommandButton "Return to Designer"
'
' NOTE: cmdCommit ("Commit This Alignment") has been removed.
'       Use cmdCommitAll to commit all drawn alignments at once.
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

    ' ---- Commit All Alignments (single commit button) ----
    If ControlExists("cmdCommitAll") Then
        cmdCommitAll.Caption = "Commit All Alignments"
        cmdCommitAll.Top = 128: cmdCommitAll.Left = 10
        cmdCommitAll.Width = 200: cmdCommitAll.Height = 28
    End If

    ' ---- Next Step ----
    If ControlExists("cmdNextStep") Then
        cmdNextStep.Caption = "Next: Place Perp Lines >"
        cmdNextStep.Top = 163: cmdNextStep.Left = 10
        cmdNextStep.Width = 200: cmdNextStep.Height = 28
        cmdNextStep.Font.Bold = True
        cmdNextStep.Enabled = False   ' enabled after commit
    End If

    ' ---- Status label ----
    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Select an alignment, draw segments, then click 'Commit All Alignments'."
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

    Me.Height = 345

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
' START SEGMENT — hide form, draw, record session element IDs
' Snapshot the max element ID before and after drawing so
' CommitCurrentAlignment knows exactly which elements belong
' to this alignment (regardless of drawing order).
' ============================================================
Private Sub cmdStartSegment_Click()
    Me.Hide

    Dim aIdx As Integer
    aIdx = cboAlignSelect.ListIndex + 1
    If aIdx < 1 Then aIdx = 1

    Dim startID As Double
    startID = GetCurrentMaxID()

    If ControlExists("optArc") Then
        If optArc.Value Then
            StartArcSegment
        Else
            StartLineSegment
        End If
    Else
        StartLineSegment
    End If

    Dim endID As Double
    endID = GetCurrentMaxID()
    Call RecordAlignmentSession(aIdx, startID, endID)

    Me.Show
End Sub

' ============================================================
' COMMIT ALL ALIGNMENTS
' Commits every alignment that has recorded drawing sessions.
' Works correctly regardless of drawing order — each alignment
' only receives elements drawn during its own sessions.
' ============================================================
Private Sub cmdCommitAll_Click()
    Dim i As Integer
    Dim nCommitted As Integer
    nCommitted = 0

    Dim total As Integer
    total = wztcAlignCount
    If total < 1 Then total = 2

    For i = 1 To total
        If wztcAlignSessionCount(i) > 0 Then
            Call SetCurrentAlignment(i)   ' sets currentAlignIdx; snapshot harmless with session approach
            Call CommitCurrentAlignment
            If wztcAlignDrawn(i) Then nCommitted = nCommitted + 1
        End If
    Next i

    ' Restore dropdown selection
    If ControlExists("cboAlignSelect") Then
        Dim sel As Integer
        sel = cboAlignSelect.ListIndex + 1
        If sel >= 1 Then Call SetCurrentAlignment(sel)
    End If

    If nCommitted > 0 Then
        If ControlExists("cmdNextStep") Then cmdNextStep.Enabled = True
        If ControlExists("lblStatus") Then
            lblStatus.Caption = CStr(nCommitted) & " alignment(s) committed. Click 'Next: Place Perp Lines' to continue."
            lblStatus.ForeColor = RGB(0, 120, 0)
        End If
    Else
        If ControlExists("lblStatus") Then
            lblStatus.Caption = "No alignment segments found to commit. Draw at least one segment first."
            lblStatus.ForeColor = RGB(160, 0, 0)
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
