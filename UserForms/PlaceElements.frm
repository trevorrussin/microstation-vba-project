Option Explicit

' ============================================================
' WZTC DRAWING ELEMENTS FORM
' ------------------------------------------------------------
' Controls to add manually in the VBA IDE form designer:
'   lblElementOf    - Label          (e.g. "Element 1 of 5:")
'   lblElementName  - Label          (element name, large/bold)
'   lblElementInstr - Label          (drawing instructions)
'   btnStartDrawing - CommandButton  "Start Drawing"
'   btnNextElement  - CommandButton  "Next WZTC Element"
'   btnGoCellLib    - CommandButton  "Next: Cell Library"
'   btnCancelElem   - CommandButton  "Cancel"
'   lblStatus       - Label          (status / error messages)
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
    Me.Caption = "WZTC Drawing Elements"
    Me.Width  = 340
    Me.Height = 255

    ' ========== ELEMENT COUNTER LABEL ==========
    If ControlExists("lblElementOf") Then
        lblElementOf.Caption   = "Initialising..."
        lblElementOf.Top       = 8
        lblElementOf.Left      = 10
        lblElementOf.Width     = 310
        lblElementOf.Height    = 16
        lblElementOf.Font.Size = 9
    End If

    ' ========== ELEMENT NAME LABEL (large, bold) ==========
    If ControlExists("lblElementName") Then
        lblElementName.Caption   = ""
        lblElementName.Top       = 26
        lblElementName.Left      = 10
        lblElementName.Width     = 310
        lblElementName.Height    = 22
        lblElementName.Font.Size = 11
        lblElementName.Font.Bold = True
        lblElementName.ForeColor = RGB(0, 0, 160)
    End If

    ' ========== INSTRUCTION LABEL ==========
    If ControlExists("lblElementInstr") Then
        lblElementInstr.Caption   = ""
        lblElementInstr.Top       = 52
        lblElementInstr.Left      = 10
        lblElementInstr.Width     = 310
        lblElementInstr.Height    = 34
        lblElementInstr.Font.Size = 8
        lblElementInstr.WordWrap  = True
        lblElementInstr.ForeColor = RGB(80, 80, 80)
    End If

    ' ========== PRIMARY ACTION BUTTONS ==========
    If ControlExists("btnStartDrawing") Then
        btnStartDrawing.Caption   = "Start Drawing"
        btnStartDrawing.Top       = 96
        btnStartDrawing.Left      = 10
        btnStartDrawing.Width     = 100
        btnStartDrawing.Height    = 23
        btnStartDrawing.Font.Bold = True
    End If

    If ControlExists("btnNextElement") Then
        btnNextElement.Caption = "Next WZTC Element"
        btnNextElement.Top     = 96
        btnNextElement.Left    = 118
        btnNextElement.Width   = 125
        btnNextElement.Height  = 23
    End If

    If ControlExists("btnCancelElem") Then
        btnCancelElem.Caption = "Cancel"
        btnCancelElem.Top     = 96
        btnCancelElem.Left    = 251
        btnCancelElem.Width   = 60
        btnCancelElem.Height  = 23
    End If

    ' ========== NEXT STEP BUTTON (enabled when all elements done) ==========
    If ControlExists("btnGoCellLib") Then
        btnGoCellLib.Caption   = "Next: Cell Library"
        btnGoCellLib.Top       = 128
        btnGoCellLib.Left      = 10
        btnGoCellLib.Width     = 140
        btnGoCellLib.Height    = 23
        btnGoCellLib.Font.Bold = True
        btnGoCellLib.Enabled   = True   ' always available; user may skip elements
    End If

    ' ========== STATUS LABEL ==========
    If ControlExists("lblStatus") Then
        lblStatus.Caption   = "Ready"
        lblStatus.Top       = 160
        lblStatus.Left      = 10
        lblStatus.Width     = 310
        lblStatus.Height    = 50
        lblStatus.Font.Size = 8
        lblStatus.WordWrap  = True
    End If

    ' ========== NAVIGATION BUTTONS ==========
    If ControlExists("btnBack") Then
        btnBack.Caption   = "< Back"
        btnBack.Top       = 218
        btnBack.Left      = 10
        btnBack.Width     = 90
        btnBack.Height    = 23
    End If

    If ControlExists("btnReturnToDesigner") Then
        btnReturnToDesigner.Caption = "Return to Designer"
        btnReturnToDesigner.Top     = 218
        btnReturnToDesigner.Left    = 108
        btnReturnToDesigner.Width   = 145
        btnReturnToDesigner.Height  = 23
    End If

    Me.Height = 268
    Call RefreshDisplay
End Sub

' ============================================================
' START DRAWING - hide form, draw segment, re-show
' ============================================================
Private Sub btnStartDrawing_Click()
    On Error GoTo DrawErr

    If ControlExists("btnStartDrawing") Then btnStartDrawing.Enabled = False
    If ControlExists("btnNextElement") Then btnNextElement.Enabled = False
    If ControlExists("btnGoCellLib") Then btnGoCellLib.Enabled = False
    If ControlExists("btnBack") Then btnBack.Enabled = False
    If ControlExists("btnReturnToDesigner") Then btnReturnToDesigner.Enabled = False
    If ControlExists("lblStatus") Then lblStatus.Caption = "Drawing in MicroStation — right-click when done..."

    DrawCurrentElementSegment

    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Segment drawn. Draw another segment or click 'Next WZTC Element'."
    End If
    If ControlExists("btnStartDrawing") Then btnStartDrawing.Enabled = True
    If ControlExists("btnNextElement") Then btnNextElement.Enabled = True
    If ControlExists("btnGoCellLib") Then btnGoCellLib.Enabled = True
    If ControlExists("btnBack") Then btnBack.Enabled = True
    If ControlExists("btnReturnToDesigner") Then btnReturnToDesigner.Enabled = True
    Exit Sub

DrawErr:
    If ControlExists("lblStatus") Then lblStatus.Caption = "Error drawing: " & Err.Description
    If ControlExists("btnStartDrawing") Then btnStartDrawing.Enabled = True
    If ControlExists("btnNextElement") Then btnNextElement.Enabled = True
    If ControlExists("btnGoCellLib") Then btnGoCellLib.Enabled = True
    If ControlExists("btnBack") Then btnBack.Enabled = True
    If ControlExists("btnReturnToDesigner") Then btnReturnToDesigner.Enabled = True
End Sub

' ============================================================
' NEXT WZTC ELEMENT - advance to next element type
' ============================================================
Private Sub btnNextElement_Click()
    AdvanceElement
    If IsAllElementsDone() Then
        Call ShowAllElementsDone
    Else
        Call RefreshDisplay
    End If
End Sub

' ============================================================
' NEXT: CELL LIBRARY - launch cell placement form
' ============================================================
Private Sub btnGoCellLib_Click()
    Unload Me
    StartWZTCCellPlacement
End Sub

' ============================================================
' CANCEL
' ============================================================
Private Sub btnCancelElem_Click()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Cancel WZTC element drawing?" & vbCrLf & _
                 "Elements already drawn will remain in the drawing.", _
                 vbYesNo + vbQuestion, "Cancel")
    If ans = vbYes Then
        Unload Me
    End If
End Sub

' ============================================================
' REFRESH DISPLAY FOR CURRENT ELEMENT
' ============================================================
Private Sub RefreshDisplay()
    If IsAllElementsDone() Then
        Call ShowAllElementsDone
        Exit Sub
    End If

    If ControlExists("lblElementOf") Then
        lblElementOf.Caption = "Element " & GetCurrentElementNumber() & " of " & GetTotalElementCount() & ":"
    End If

    If ControlExists("lblElementName") Then
        lblElementName.Caption   = GetCurrentElementDisplayName()
        lblElementName.ForeColor = RGB(0, 0, 160)
    End If

    If ControlExists("lblElementInstr") Then
        lblElementInstr.Caption = GetCurrentElementInstructions()
    End If

    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Click 'Start Drawing', then click points in MicroStation."
    End If

    If ControlExists("btnStartDrawing") Then btnStartDrawing.Enabled = True
    If ControlExists("btnNextElement") Then btnNextElement.Enabled = True
End Sub

' ============================================================
' ALL ELEMENTS COMPLETE
' ============================================================
Private Sub ShowAllElementsDone()
    If ControlExists("lblElementOf") Then lblElementOf.Caption = "Complete!"
    If ControlExists("lblElementName") Then
        lblElementName.Caption   = "All " & GetTotalElementCount() & " elements drawn."
        lblElementName.ForeColor = RGB(0, 140, 0)
    End If
    If ControlExists("lblElementInstr") Then lblElementInstr.Caption = ""
    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Done! Click 'Next: Cell Library' to place additional WZTC symbols."
    End If
    If ControlExists("btnStartDrawing") Then btnStartDrawing.Enabled = False
    If ControlExists("btnNextElement") Then btnNextElement.Enabled = False
    If ControlExists("btnGoCellLib") Then btnGoCellLib.Enabled = True
End Sub

' ============================================================
' NAVIGATION - BACK AND RETURN TO DESIGNER
' ============================================================
Private Sub btnBack_Click()
    Unload Me
    StartSignPlacement
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
        If Not IsAllElementsDone() Then
            Dim ans As VbMsgBoxResult
            ans = MsgBox("Close WZTC element drawing tool?" & vbCrLf & _
                         "Elements already drawn will remain.", _
                         vbYesNo + vbQuestion, "Close")
            If ans = vbNo Then Cancel = True
        End If
    End If
End Sub
