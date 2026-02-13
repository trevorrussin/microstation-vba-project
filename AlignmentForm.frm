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
        cmdStartSegment.Width = 120
        cmdStartSegment.Height = 30
        cmdStartSegment.Font.Bold = True
    End If

    If ControlExists("cmdDone") Then
        cmdDone.Caption = "Done"
        cmdDone.Top = 130
        cmdDone.Left = 20
        cmdDone.Width = 120
        cmdDone.Height = 30
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
End Sub
