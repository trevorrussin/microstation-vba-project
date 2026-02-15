Option Explicit

' ============================================================
' ALIGNMENTFORM BEHAVIOR
' ------------------------------------------------------------
' Controls to add manually in the VBA IDE form designer:
'   optLine         - OptionButton  "Line"
'   optArc          - OptionButton  "Arc"
'   cmdStartSegment - CommandButton "Start Segment"
'   cmdDone         - CommandButton "Done"
' ============================================================

Private Function ControlExists(ctrlName As String) As Boolean
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = Me.Controls(ctrlName)
    ControlExists = Not (ctrl Is Nothing)
    On Error GoTo 0
End Function

Private Sub UserForm_Initialize()
    Me.Caption = "Alignment Tool"
    Me.Width  = 210
    Me.Height = 290

    If ControlExists("optLine") Then
        optLine.Caption = "Line"
        optLine.Top = 20
        optLine.Left = 20
        optLine.Width = 80
        optLine.Value = True
    End If

    If ControlExists("optArc") Then
        optArc.Caption = "Arc"
        optArc.Top = 50
        optArc.Left = 20
        optArc.Width = 80
        optArc.Value = False
    End If

    If ControlExists("cmdStartSegment") Then
        cmdStartSegment.Caption = "Start Segment"
        cmdStartSegment.Top = 90
        cmdStartSegment.Left = 20
        cmdStartSegment.Width = 160
        cmdStartSegment.Height = 30
        cmdStartSegment.Font.Bold = True
    End If

    If ControlExists("cmdDone") Then
        cmdDone.Caption = "Done"
        cmdDone.Top = 130
        cmdDone.Left = 20
        cmdDone.Width = 160
        cmdDone.Height = 30
    End If

    ' ========== RIGHT-CLICK INSTRUCTION ==========
    If ControlExists("lblRightClick") Then
        lblRightClick.Caption   = "Tip: Right-click in MicroStation to finish a segment and return to this form."
        lblRightClick.Top       = 172
        lblRightClick.Left      = 10
        lblRightClick.Width     = 185
        lblRightClick.Height    = 36
        lblRightClick.Font.Size = 8
        lblRightClick.WordWrap  = True
        lblRightClick.ForeColor = RGB(80, 80, 80)
    End If

    ' ========== NAVIGATION BUTTONS ==========
    If ControlExists("btnBack") Then
        btnBack.Caption   = "< Back to Designer"
        btnBack.Top       = 218
        btnBack.Left      = 10
        btnBack.Width     = 175
        btnBack.Height    = 23
    End If

    If ControlExists("btnReturnToDesigner") Then
        btnReturnToDesigner.Caption = "Return to Designer"
        btnReturnToDesigner.Top     = 250
        btnReturnToDesigner.Left    = 10
        btnReturnToDesigner.Width   = 175
        btnReturnToDesigner.Height  = 23
    End If
End Sub

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

Private Sub cmdDone_Click()
    On Error Resume Next
    CadInputQueue.SendReset
    CommandState.StartDefaultCommand
    Unload Me
    On Error GoTo 0
    ' Group alignment elements and launch perpendicular placement tool
    GroupAndLaunchPlacement
End Sub

' ============================================================
' NAVIGATION - BACK AND RETURN TO DESIGNER
' Both go to WorkzoneDesigner since this is the first step
' ============================================================
Private Sub btnBack_Click()
    Unload Me
    WZTCDesigner.Show vbModeless
End Sub

Private Sub btnReturnToDesigner_Click()
    Unload Me
    WZTCDesigner.Show vbModeless
End Sub
