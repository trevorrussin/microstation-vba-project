Option Explicit

' ============================================================
' WZTC SIGN PLACEMENT FORM
' ------------------------------------------------------------
' Sign placement state and drawing: DrawSign.bas (StartSignPlacement,
' DrawCurrentSign, AdvanceSign, GetCurrentSignNum, GetCurrentSignSide,
' GetCurrentSignSize, GetCurrentSignNumber, GetTotalSignCount, IsAllSignsDone).
'
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
    Me.Width  = 330
    Me.Height = 230

    ' ========== SIGN COUNTER LABEL ==========
    If ControlExists("lblSignOf") Then
        lblSignOf.Caption   = "Initialising..."
        lblSignOf.Top       = 8
        lblSignOf.Left      = 10
        lblSignOf.Width     = 300
        lblSignOf.Height    = 16
        lblSignOf.Font.Size = 9
    End If

    ' ========== SIGN NUMBER LABEL (large, bold) ==========
    If ControlExists("lblSignName") Then
        lblSignName.Caption   = ""
        lblSignName.Top       = 26
        lblSignName.Left      = 10
        lblSignName.Width     = 300
        lblSignName.Height    = 22
        lblSignName.Font.Size = 11
        lblSignName.Font.Bold = True
        lblSignName.ForeColor = RGB(0, 0, 160)
    End If

    ' ========== SIDE DESCRIPTION LABEL ==========
    If ControlExists("lblSignSide") Then
        lblSignSide.Caption   = ""
        lblSignSide.Top       = 52
        lblSignSide.Left      = 10
        lblSignSide.Width     = 300
        lblSignSide.Height    = 16
        lblSignSide.Font.Size = 9
        lblSignSide.ForeColor = RGB(120, 60, 0)
    End If

    ' ========== INSTRUCTION LABEL ==========
    If ControlExists("lblInstruction") Then
        lblInstruction.Caption   = "Click 'Draw Sign' then click the post location(s) on the perpendicular line in MicroStation."
        lblInstruction.Top       = 76
        lblInstruction.Left      = 10
        lblInstruction.Width     = 300
        lblInstruction.Height    = 34
        lblInstruction.Font.Size = 8
        lblInstruction.WordWrap  = True
    End If

    ' ========== ACTION BUTTONS ==========
    If ControlExists("btnDraw") Then
        btnDraw.Caption   = "Draw Sign"
        btnDraw.Top       = 120
        btnDraw.Left      = 10
        btnDraw.Width     = 90
        btnDraw.Height    = 23
        btnDraw.Font.Bold = True
    End If

    If ControlExists("btnNextSign") Then
        btnNextSign.Caption = "Next Sign"
        btnNextSign.Top     = 120
        btnNextSign.Left    = 108
        btnNextSign.Width   = 90
        btnNextSign.Height  = 23
    End If

    If ControlExists("btnCancel") Then
        btnCancel.Caption = "Cancel"
        btnCancel.Top     = 120
        btnCancel.Left    = 206
        btnCancel.Width   = 75
        btnCancel.Height  = 23
    End If

    ' ========== NEXT STEP BUTTON ==========
    If ControlExists("btnWZTCElements") Then
        btnWZTCElements.Caption   = "Next: WZTC Elements"
        btnWZTCElements.Top       = 151
        btnWZTCElements.Left      = 10
        btnWZTCElements.Width     = 145
        btnWZTCElements.Height    = 23
        btnWZTCElements.Font.Bold = True
    End If

    ' ========== STATUS LABEL ==========
    If ControlExists("lblStatus") Then
        lblStatus.Caption   = "Ready"
        lblStatus.Top       = 182
        lblStatus.Left      = 10
        lblStatus.Width     = 300
        lblStatus.Height    = 42
        lblStatus.Font.Size = 8
        lblStatus.WordWrap  = True
    End If

    ' ========== NAVIGATION BUTTONS ==========
    If ControlExists("btnBack") Then
        btnBack.Caption   = "< Back"
        btnBack.Top       = 232
        btnBack.Left      = 10
        btnBack.Width     = 90
        btnBack.Height    = 23
    End If

    If ControlExists("btnReturnToDesigner") Then
        btnReturnToDesigner.Caption = "Return to Designer"
        btnReturnToDesigner.Top     = 232
        btnReturnToDesigner.Left    = 108
        btnReturnToDesigner.Width   = 145
        btnReturnToDesigner.Height  = 23
    End If

    Me.Height = 275
    Call RefreshDisplay
End Sub

' ============================================================
' DRAW SIGN - form stays visible (modeless); buttons disabled during input
' ============================================================
Private Sub btnDraw_Click()
    On Error GoTo DrawErr

    If ControlExists("btnDraw") Then btnDraw.Enabled = False
    If ControlExists("btnNextSign") Then btnNextSign.Enabled = False
    If ControlExists("btnWZTCElements") Then btnWZTCElements.Enabled = False
    If ControlExists("lblStatus") Then lblStatus.Caption = "Click post location(s) in MicroStation..."

    DrawCurrentSign

    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Sign drawn. Click 'Next Sign' to continue to the next sign."
    End If
    ' Draw Sign stays grayed — user must click Next Sign to proceed
    If ControlExists("btnNextSign") Then btnNextSign.Enabled = True
    If ControlExists("btnWZTCElements") Then btnWZTCElements.Enabled = True
    Exit Sub

DrawErr:
    If ControlExists("lblStatus") Then lblStatus.Caption = "Error drawing sign: " & Err.Description
    ' Re-enable Draw Sign on error so user can retry
    If ControlExists("btnDraw") Then btnDraw.Enabled = True
    If ControlExists("btnNextSign") Then btnNextSign.Enabled = True
    If ControlExists("btnWZTCElements") Then btnWZTCElements.Enabled = True
End Sub

' ============================================================
' NEXT STEP - unload this form and launch WZTC elements drawing
' ============================================================
Private Sub btnWZTCElements_Click()
    Unload Me
    frmSignSubColors.Show vbModeless
End Sub

' ============================================================
' NEXT SIGN - advance index and refresh
' ============================================================
Private Sub btnNextSign_Click()
    If ControlExists("btnDraw") Then btnDraw.Enabled = True
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
        lblStatus.Caption = "All signs placed. Click 'Next: WZTC Elements' to continue."
    End If
    If ControlExists("btnDraw") Then btnDraw.Enabled = False
    If ControlExists("btnNextSign") Then btnNextSign.Enabled = False
    If ControlExists("btnWZTCElements") Then btnWZTCElements.Enabled = True
End Sub

' ============================================================
' NAVIGATION - BACK AND RETURN TO DESIGNER
' ============================================================
Private Sub btnBack_Click()
    Unload Me
    StartAlignmentPlacement
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
        If Not IsAllSignsDone() Then
            Dim ans As VbMsgBoxResult
            ans = MsgBox("Close sign placement tool?" & vbCrLf & _
                         "Signs already drawn will remain.", _
                         vbYesNo + vbQuestion, "Close")
            If ans = vbNo Then Cancel = True
        End If
    End If
End Sub
