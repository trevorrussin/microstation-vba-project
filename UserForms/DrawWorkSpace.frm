Option Explicit

' ============================================================
' DrawWorkSpace — Standalone Work Space (hatch) drawing form.
' First step after WZTCDesigner Submit & Draw.
' User traces the closed boundary of their work zone, then
' clicks inside it to place a hatch fill.
' Pattern mirrors DrawElements.bas Work Space logic exactly.
' Controls must be added manually in the VBA IDE:
'   lblInstructions, btnDrawWorkSpace, btnNextStep,
'   btnBack, btnReturnToDesigner, lblStatus
' ============================================================

Private Function ControlExists(ctrlName As String) As Boolean
    On Error Resume Next
    Dim c As Object
    Set c = Me.Controls(ctrlName)
    ControlExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' ============================================================
' INITIALIZE
' ============================================================
Private Sub UserForm_Initialize()
    On Error Resume Next

    Me.Caption = "Draw Work Space"
    Me.Width  = 320
    Me.Height = 280

    If ControlExists("lblInstructions") Then
        With lblInstructions
            .Caption = "Draw the boundary of your Work Zone:" & vbCrLf & _
                       "1. Click ""Draw Work Space"" below." & vbCrLf & _
                       "2. Click corners of the work zone area." & vbCrLf & _
                       "3. Right-click to close the shape." & vbCrLf & _
                       "4. Click the BORDER of the shape to apply hatching." & vbCrLf & _
                       "You may draw multiple work spaces. Click ""Next"" when done."
            .Top = 10: .Left = 10: .Width = 290: .Height = 90
            .WordWrap = True
        End With
    End If

    If ControlExists("btnDrawWorkSpace") Then
        With btnDrawWorkSpace
            .Caption = "Draw Work Space"
            .Top = 110: .Left = 10: .Width = 150: .Height = 25
            .Font.Bold = True
        End With
    End If

    If ControlExists("btnNextStep") Then
        With btnNextStep
            .Caption = "Next: Draw Alignments >"
            .Top = 145: .Left = 10: .Width = 180: .Height = 25
            .Font.Bold = True
        End With
    End If

    If ControlExists("btnBack") Then
        With btnBack
            .Caption = "< Back to Designer"
            .Top = 185: .Left = 10: .Width = 130: .Height = 22
        End With
    End If

    If ControlExists("btnReturnToDesigner") Then
        With btnReturnToDesigner
            .Caption = "Return to Designer"
            .Top = 185: .Left = 150: .Width = 130: .Height = 22
        End With
    End If

    If ControlExists("lblStatus") Then
        With lblStatus
            .Caption = "Ready. Click 'Draw Work Space' to begin."
            .Top = 215: .Left = 10: .Width = 290: .Height = 36
            .WordWrap = True
            .ForeColor = RGB(0, 100, 0)
        End With
    End If
End Sub

' ============================================================
' DRAW WORK SPACE BUTTON
' Activates PLACE SHAPE CONSTRAINED; user clicks corners.
' Right-click finishes shape; then HATCH ICON is activated
' and the user clicks inside the shape to fill it.
' ============================================================
Private Sub btnDrawWorkSpace_Click()
    On Error GoTo DrawErr

    ' Disable buttons so user must interact with MicroStation
    Call SetButtons(False)
    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Click corners of work zone. Right-click to finish shape."
        lblStatus.ForeColor = RGB(0, 0, 180)
    End If

    ' Set element properties: TWZWS2_P level, color 6, weight 2
    CadInputQueue.SendKeyin "ACTIVE LEVEL ""TWZWS2_P"""
    CadInputQueue.SendKeyin "ACTIVE COLOR 6"
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 2"

    ' Start closed shape command
    CadInputQueue.SendCommand "PLACE SHAPE CONSTRAINED"

    ' Collect user clicks until right-click (Reset)
    Dim oMsg As CadInputMessage
    Dim nPts As Integer
    nPts = 0
    Set oMsg = CadInputQueue.GetInput
    Do While oMsg.InputType <> msdCadInputTypeReset
        If oMsg.InputType = msdCadInputTypeDataPoint Then
            CadInputQueue.SendDataPoint oMsg.Point, 1
            nPts = nPts + 1
        End If
        Set oMsg = CadInputQueue.GetInput
    Loop

    ' If enough points for a shape, offer hatching
    If nPts >= 3 Then
        If ControlExists("lblStatus") Then
            lblStatus.Caption = "Shape closed. Click the BORDER of the work zone shape to apply hatching."
            lblStatus.ForeColor = RGB(0, 0, 180)
        End If

        ' Activate hatch tool
        CadInputQueue.SendCommand "HATCH ICON"

        ' Wait for user to click inside the shape
        Set oMsg = CadInputQueue.GetInput
        Do While oMsg.InputType <> msdCadInputTypeDataPoint And _
                 oMsg.InputType <> msdCadInputTypeReset
            Set oMsg = CadInputQueue.GetInput
        Loop

        If oMsg.InputType = msdCadInputTypeDataPoint Then
            ' Send the interior point twice (LegacyElements.bas pattern)
            CadInputQueue.SendDataPoint oMsg.Point, 1
            CadInputQueue.SendDataPoint oMsg.Point, 1
        End If

        CadInputQueue.SendReset
    End If

    CommandState.StartDefaultCommand

    If ControlExists("lblStatus") Then
        If nPts >= 3 Then
            lblStatus.Caption = "Work space drawn. Draw another or click 'Next: Draw Alignments'."
        Else
            lblStatus.Caption = "Not enough points for a shape (need 3+). Try again."
        End If
        lblStatus.ForeColor = RGB(0, 100, 0)
    End If

    Call SetButtons(True)
    Exit Sub

DrawErr:
    CadInputQueue.SendReset
    CommandState.StartDefaultCommand
    Call SetButtons(True)
    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Error: " & Err.Description
        lblStatus.ForeColor = RGB(180, 0, 0)
    End If
End Sub

' ============================================================
' NEXT STEP — proceed to draw alignments
' ============================================================
Private Sub btnNextStep_Click()
    Unload Me
    Call StartWZTCDrawing   ' AlignmentTool.bas
End Sub

' ============================================================
' BACK — return to WZTCDesigner
' ============================================================
Private Sub btnBack_Click()
    Unload Me
    WZTCDesigner.Show vbModeless
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
        Dim ans As Integer
        ans = MsgBox("Exit the WZTC Designer workflow?", vbYesNo + vbQuestion, "Confirm Exit")
        If ans = vbNo Then Cancel = 1
    End If
End Sub

' ============================================================
' HELPER: Enable/disable all action buttons together
' ============================================================
Private Sub SetButtons(enabled As Boolean)
    If ControlExists("btnDrawWorkSpace")   Then btnDrawWorkSpace.Enabled   = enabled
    If ControlExists("btnNextStep")        Then btnNextStep.Enabled        = enabled
    If ControlExists("btnBack")            Then btnBack.Enabled            = enabled
    If ControlExists("btnReturnToDesigner") Then btnReturnToDesigner.Enabled = enabled
End Sub
