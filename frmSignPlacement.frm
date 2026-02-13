Option Explicit

' ============================================================
' WZTC SIGN PLACEMENT FORM
' ------------------------------------------------------------
' Controls to add manually in the VBA IDE form designer:
'   lblSignOf      - Label          (e.g. "Sign 1 of 3:")
'   lblSignName    - Label          (sign number, large/bold)
'   lblSignSide    - Label          ("One Side" or "Both Sides" description)
'   lblInstruction - Label          (click instructions)
'   btnDraw        - CommandButton  "Draw Sign"
'   btnNextSign    - CommandButton  "Next Sign"
'   btnCancel      - CommandButton  "Cancel"
'   lblStatus      - Label          (status / error messages)
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
    Me.Caption = "WZTC Sign Placement"

    If ControlExists("lblSignOf") Then
        lblSignOf.Caption = "Initialising..."
        lblSignOf.Font.Size = 9
    End If

    If ControlExists("lblSignName") Then
        lblSignName.Caption = ""
        lblSignName.Font.Size = 11
        lblSignName.Font.Bold = True
        lblSignName.ForeColor = RGB(0, 0, 160)
    End If

    If ControlExists("lblSignSide") Then
        lblSignSide.Caption = ""
        lblSignSide.Font.Size = 9
        lblSignSide.ForeColor = RGB(120, 60, 0)
    End If

    If ControlExists("lblInstruction") Then
        lblInstruction.Caption = "Click 'Draw Sign' then click post location(s) on the perpendicular line in MicroStation."
        lblInstruction.Font.Size = 8
        lblInstruction.WordWrap = True
    End If

    If ControlExists("btnDraw") Then
        btnDraw.Caption = "Draw Sign"
        btnDraw.Font.Bold = True
    End If

    If ControlExists("btnNextSign") Then
        btnNextSign.Caption = "Next Sign"
    End If

    If ControlExists("btnCancel") Then
        btnCancel.Caption = "Cancel"
    End If

    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Ready"
        lblStatus.Font.Size = 8
        lblStatus.WordWrap = True
    End If

    Call RefreshDisplay
End Sub

' ============================================================
' DRAW SIGN - hide form, run drawing sub, re-show
' ============================================================
Private Sub btnDraw_Click()
    On Error GoTo DrawErr

    If ControlExists("btnDraw") Then btnDraw.Enabled = False
    If ControlExists("btnNextSign") Then btnNextSign.Enabled = False
    If ControlExists("lblStatus") Then lblStatus.Caption = "Drawing — click post location(s) in MicroStation..."

    Me.Hide
    DrawCurrentSign
    Me.Show

    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Sign drawn. Verify in MicroStation, then click 'Next Sign' (or 'Draw Sign' again to redraw)."
    End If
    If ControlExists("btnDraw") Then btnDraw.Enabled = True
    If ControlExists("btnNextSign") Then btnNextSign.Enabled = True
    Exit Sub

DrawErr:
    Me.Show
    If ControlExists("lblStatus") Then lblStatus.Caption = "Error drawing sign: " & Err.Description
    If ControlExists("btnDraw") Then btnDraw.Enabled = True
    If ControlExists("btnNextSign") Then btnNextSign.Enabled = True
End Sub

' ============================================================
' NEXT SIGN - advance index and refresh
' ============================================================
Private Sub btnNextSign_Click()
    AdvanceSign
    If IsAllSignsDone() Then
        Call ShowAllSignsDone
    Else
        Call RefreshDisplay
    End If
End Sub

' ============================================================
' CANCEL
' ============================================================
Private Sub btnCancel_Click()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Cancel sign placement?" & vbCrLf & _
                 "Signs already drawn will remain in the drawing.", _
                 vbYesNo + vbQuestion, "Cancel")
    If ans = vbYes Then
        Unload Me
    End If
End Sub

' ============================================================
' REFRESH DISPLAY FOR CURRENT SIGN
' ============================================================
Private Sub RefreshDisplay()
    If IsAllSignsDone() Then
        Call ShowAllSignsDone
        Exit Sub
    End If

    If ControlExists("lblSignOf") Then
        lblSignOf.Caption = "Sign " & GetCurrentSignNumber() & " of " & GetTotalSignCount() & ":"
    End If

    If ControlExists("lblSignName") Then
        lblSignName.Caption = GetCurrentSignNum()
        lblSignName.ForeColor = RGB(0, 0, 160)
    End If

    If ControlExists("lblSignSide") Then
        Dim s As String
        s = GetCurrentSignSide()
        If s = "Both Sides" Then
            lblSignSide.Caption = "Both Sides — click 2 points on the perpendicular line"
        Else
            lblSignSide.Caption = "One Side — click 1 point on the perpendicular line"
        End If
    End If

    If ControlExists("lblInstruction") Then
        lblInstruction.Caption = "Click 'Draw Sign' then click the post location(s) on the perpendicular line in MicroStation."
    End If

    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Ready. The size for this sign is: " & GetCurrentSignSize()
    End If

    If ControlExists("btnDraw") Then btnDraw.Enabled = True
    If ControlExists("btnNextSign") Then btnNextSign.Enabled = True
End Sub

' ============================================================
' ALL SIGNS COMPLETE
' ============================================================
Private Sub ShowAllSignsDone()
    If ControlExists("lblSignOf") Then lblSignOf.Caption = "Complete!"
    If ControlExists("lblSignName") Then
        lblSignName.Caption = "All " & GetTotalSignCount() & " signs placed."
        lblSignName.ForeColor = RGB(0, 140, 0)
    End If
    If ControlExists("lblSignSide") Then lblSignSide.Caption = ""
    If ControlExists("lblInstruction") Then lblInstruction.Caption = ""
    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Done! All signs have been drawn. You may close this window."
    End If
    If ControlExists("btnDraw") Then btnDraw.Enabled = False
    If ControlExists("btnNextSign") Then btnNextSign.Enabled = False
End Sub

' ============================================================
' CONFIRM CLOSE VIA X BUTTON
' ============================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        If Not IsAllSignsDone() Then
            Dim ans As VbMsgBoxResult
            ans = MsgBox("Close sign placement tool?" & vbCrLf & _
                         "Signs already drawn will remain.", _
                         vbYesNo + vbQuestion, "Close")
            If ans = vbNo Then Cancel = True
        End If
    End If
End Sub
