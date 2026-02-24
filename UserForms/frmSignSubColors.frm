Option Explicit

' ============================================================
' frmSignSubColors — Optional: change sign face background color.
' Step 6 of 8 in the WZTC workflow (after PlaceSign, before PlaceElements).
' Controls to add manually in the VBA IDE form designer:
'   lblInstructions   - Label
'   cmdChangeAttrib   - CommandButton  "Apply Attributes to Sign"
'   lblStatus         - Label
'   btnBack           - CommandButton  "< Back"
'   btnReturnToDesigner - CommandButton "Return to Designer"
'   btnNext           - CommandButton  "Next: WZTC Elements >"
' ============================================================

Private Function ControlExists(ctrlName As String) As Boolean
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = Me.Controls(ctrlName)
    ControlExists = Not (ctrl Is Nothing)
    On Error GoTo 0
End Function

Private Sub UserForm_Initialize()
    Me.Caption = "Sign Attribute Editor  (Optional)"
    Me.Width = 360: Me.Height = 220

    If ControlExists("lblInstructions") Then
        With lblInstructions
            .Caption = "OPTIONAL STEP — Click 'Next: WZTC Elements' at any time to skip." & vbCrLf & vbCrLf & _
                       "If you want to change the background color of any sign face in your drawing, " & _
                       "click 'Apply Attributes to Sign' and then click the sign cell in MicroStation. " & _
                       "You can click multiple signs. Right-click when done."
            .Top = 8: .Left = 10: .Width = 330: .Height = 80
            .WordWrap = True
            .Font.Size = 9
        End With
    End If

    If ControlExists("cmdChangeAttrib") Then
        With cmdChangeAttrib
            .Caption = "Apply Attributes to Sign"
            .Top = 96: .Left = 10: .Width = 330: .Height = 26
            .Font.Bold = True
        End With
    End If

    If ControlExists("lblStatus") Then
        With lblStatus
            .Caption = "Click 'Apply Attributes to Sign' to change a sign's background color, or click Next to continue."
            .Top = 130: .Left = 10: .Width = 330: .Height = 32
            .WordWrap = True
            .ForeColor = RGB(0, 0, 160)
            .Font.Size = 8
        End With
    End If

    If ControlExists("btnBack") Then
        With btnBack
            .Caption = "< Back"
            .Top = 170: .Left = 10: .Width = 70: .Height = 23
        End With
    End If

    If ControlExists("btnReturnToDesigner") Then
        With btnReturnToDesigner
            .Caption = "Return to Designer"
            .Top = 170: .Left = 88: .Width = 145: .Height = 23
        End With
    End If

    If ControlExists("btnNext") Then
        With btnNext
            .Caption = "Next: WZTC Elements >"
            .Top = 170: .Left = 241: .Width = 109: .Height = 23
            .Font.Bold = True
            .Enabled = True
        End With
    End If
End Sub

' --- Setup CHANGE ATTRIBUTES with presets, then let user click sign(s) directly ---
Private Sub cmdChangeAttrib_Click()
    On Error GoTo AttribErr
    StatusBlue "Click sign(s) in MicroStation to apply attributes. Right-click when done."
    Me.Hide

    CadInputQueue.SendCommand "CHANGE ATTRIBUTES"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES USEACTIVE OFF"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE LEVEL"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET LEVEL ""SF_P"""
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE COLOR"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET COLOR ""240"""
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE LINESTYLE"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET LINESTYLE ""ByLevel"""
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE WEIGHT"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET WEIGHT 3"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE TRANSPARENCY"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET TRANSPARENCY 0"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE PRIORITY"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET PRIORITY 0"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE ELEMENTCLASS"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET ELEMENTCLASS CONSTRUCTION"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE TEMPLATE"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET TEMPLATE """""
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE FILLCOLOR"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET FILLCOLOR ""6"""
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES MAKECOPY OFF"
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENTIREELEMENT OFF"

    Dim oMsg As CadInputMessage
    Dim nApplied As Integer
    nApplied = 0
    Set oMsg = CadInputQueue.GetInput
    Do
        If oMsg.InputType = msdCadInputTypeDataPoint Then
            CadInputQueue.SendDataPoint oMsg.Point, 1   ' identify element
            CadInputQueue.SendDataPoint oMsg.Point, 1   ' accept / apply
            nApplied = nApplied + 1
        ElseIf oMsg.InputType = msdCadInputTypeReset Then
            Exit Do
        End If
        Set oMsg = CadInputQueue.GetInput
    Loop

    CommandState.StartDefaultCommand
    Me.Show
    If nApplied > 0 Then
        StatusGreen "Attributes applied to " & nApplied & " element(s). Apply more or click Next."
    Else
        StatusBlue "No elements clicked. Click 'Apply Attributes to Sign' to try again, or click Next to continue."
    End If
    Exit Sub

AttribErr:
    On Error Resume Next
    CommandState.StartDefaultCommand
    Me.Show
    StatusBlue "Error applying attributes: " & Err.Description
End Sub

' ============================================================
' NAVIGATION
' ============================================================
Private Sub btnNext_Click()
    Unload Me
    StartWZTCElementsPlacementFromElement2
End Sub

Private Sub btnBack_Click()
    Unload Me
    StartSignPlacement
End Sub

Private Sub btnReturnToDesigner_Click()
    Unload Me
    WZTCDesigner.Show vbModeless
End Sub

' ============================================================
' HELPERS
' ============================================================
Private Sub StatusBlue(ByVal msg As String)
    If ControlExists("lblStatus") Then
        lblStatus.Caption = msg
        lblStatus.ForeColor = RGB(0, 0, 160)
    End If
End Sub

Private Sub StatusGreen(ByVal msg As String)
    If ControlExists("lblStatus") Then
        lblStatus.Caption = msg
        lblStatus.ForeColor = RGB(0, 120, 0)
    End If
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    CommandState.StartDefaultCommand
    On Error GoTo 0
End Sub
