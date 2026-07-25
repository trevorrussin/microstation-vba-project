Option Explicit

' ============================================================
' frmBBCalibrate.frm — Coordinate Reference Point Calibration
' ------------------------------------------------------------
' Controls to add manually in the VBA IDE form designer:
'   lblTitle          - Label          "Set PDF → MicroStation Reference Points"
'   lblInstr          - Label          (instructions)
'   lblPt1Header      - Label          "Reference Point 1:"
'   lblPdf1X          - Label          "PDF X:"
'   txtPdf1X          - TextBox        (PDF X coord of point 1)
'   lblPdf1Y          - Label          "PDF Y:"
'   txtPdf1Y          - TextBox        (PDF Y coord of point 1)
'   cmdClickPt1       - CommandButton  "Click Point 1 in Drawing"
'   lblMstn1          - Label          (shows captured MicroStation coords)
'   lblPt2Header      - Label          "Reference Point 2:"
'   lblPdf2X          - Label          "PDF X:"
'   txtPdf2X          - TextBox        (PDF X coord of point 2)
'   lblPdf2Y          - Label          "PDF Y:"
'   txtPdf2Y          - TextBox        (PDF Y coord of point 2)
'   cmdClickPt2       - CommandButton  "Click Point 2 in Drawing"
'   lblMstn2          - Label          (shows captured MicroStation coords)
'   cmdCompute        - CommandButton  "Calibrate"
'   lblCalibResult    - Label          (scale / offset result)
'   cmdClose          - CommandButton  "Close"
' ============================================================

Private mstn1Captured As Boolean
Private mstn2Captured As Boolean

Private Function ControlExists(ctrlName As String) As Boolean
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = Me.Controls(ctrlName)
    ControlExists = Not (ctrl Is Nothing)
    On Error GoTo 0
End Function

' ============================================================
' INITIALIZE — layout controls
' ============================================================
Private Sub UserForm_Initialize()
    Me.Caption = "Set Reference Points"
    Me.Width  = 400
    Me.Height = 400

    mstn1Captured = False
    mstn2Captured = False

    ' ---- TITLE ----
    If ControlExists("lblTitle") Then
        lblTitle.Caption   = "PDF to MicroStation Coordinate Calibration"
        lblTitle.Top       = 8
        lblTitle.Left      = 10
        lblTitle.Width     = 375
        lblTitle.Height    = 16
        lblTitle.Font.Bold = True
        lblTitle.Font.Size = 9
    End If

    If ControlExists("lblInstr") Then
        lblInstr.Caption   = "Choose two points visible on both the exported PDF and in MicroStation " & _
                             "(e.g., survey corners, benchmark stations). Enter their PDF page coordinates " & _
                             "below, then click each point in the MicroStation drawing."
        lblInstr.Top       = 28
        lblInstr.Left      = 10
        lblInstr.Width     = 375
        lblInstr.Height    = 45
        lblInstr.Font.Size = 8
        lblInstr.WordWrap  = True
        lblInstr.ForeColor = RGB(60, 60, 60)
    End If

    ' ---- POINT 1 ----
    If ControlExists("lblPt1Header") Then
        lblPt1Header.Caption   = "Reference Point 1:"
        lblPt1Header.Top       = 80
        lblPt1Header.Left      = 10
        lblPt1Header.Width     = 375
        lblPt1Header.Height    = 15
        lblPt1Header.Font.Bold = True
        lblPt1Header.Font.Size = 8
    End If

    If ControlExists("lblPdf1X") Then
        lblPdf1X.Caption = "PDF X:"
        lblPdf1X.Top     = 100
        lblPdf1X.Left    = 10
        lblPdf1X.Width   = 45
        lblPdf1X.Height  = 16
        lblPdf1X.Font.Size = 8
    End If

    If ControlExists("txtPdf1X") Then
        txtPdf1X.Top    = 98
        txtPdf1X.Left   = 58
        txtPdf1X.Width  = 80
        txtPdf1X.Height = 18
        txtPdf1X.Font.Size = 8
        txtPdf1X.Text   = ""
    End If

    If ControlExists("lblPdf1Y") Then
        lblPdf1Y.Caption = "PDF Y:"
        lblPdf1Y.Top     = 100
        lblPdf1Y.Left    = 150
        lblPdf1Y.Width   = 45
        lblPdf1Y.Height  = 16
        lblPdf1Y.Font.Size = 8
    End If

    If ControlExists("txtPdf1Y") Then
        txtPdf1Y.Top    = 98
        txtPdf1Y.Left   = 198
        txtPdf1Y.Width  = 80
        txtPdf1Y.Height = 18
        txtPdf1Y.Font.Size = 8
        txtPdf1Y.Text   = ""
    End If

    If ControlExists("cmdClickPt1") Then
        cmdClickPt1.Caption = "Click Point 1 in Drawing"
        cmdClickPt1.Top     = 124
        cmdClickPt1.Left    = 10
        cmdClickPt1.Width   = 165
        cmdClickPt1.Height  = 22
        cmdClickPt1.Font.Size = 8
    End If

    If ControlExists("lblMstn1") Then
        lblMstn1.Caption   = "MicroStation: not captured"
        lblMstn1.Top       = 152
        lblMstn1.Left      = 10
        lblMstn1.Width     = 375
        lblMstn1.Height    = 14
        lblMstn1.Font.Size = 8
        lblMstn1.ForeColor = RGB(100, 100, 100)
    End If

    ' ---- POINT 2 ----
    If ControlExists("lblPt2Header") Then
        lblPt2Header.Caption   = "Reference Point 2:"
        lblPt2Header.Top       = 174
        lblPt2Header.Left      = 10
        lblPt2Header.Width     = 375
        lblPt2Header.Height    = 15
        lblPt2Header.Font.Bold = True
        lblPt2Header.Font.Size = 8
    End If

    If ControlExists("lblPdf2X") Then
        lblPdf2X.Caption = "PDF X:"
        lblPdf2X.Top     = 194
        lblPdf2X.Left    = 10
        lblPdf2X.Width   = 45
        lblPdf2X.Height  = 16
        lblPdf2X.Font.Size = 8
    End If

    If ControlExists("txtPdf2X") Then
        txtPdf2X.Top    = 192
        txtPdf2X.Left   = 58
        txtPdf2X.Width  = 80
        txtPdf2X.Height = 18
        txtPdf2X.Font.Size = 8
        txtPdf2X.Text   = ""
    End If

    If ControlExists("lblPdf2Y") Then
        lblPdf2Y.Caption = "PDF Y:"
        lblPdf2Y.Top     = 194
        lblPdf2Y.Left    = 150
        lblPdf2Y.Width   = 45
        lblPdf2Y.Height  = 16
        lblPdf2Y.Font.Size = 8
    End If

    If ControlExists("txtPdf2Y") Then
        txtPdf2Y.Top    = 192
        txtPdf2Y.Left   = 198
        txtPdf2Y.Width  = 80
        txtPdf2Y.Height = 18
        txtPdf2Y.Font.Size = 8
        txtPdf2Y.Text   = ""
    End If

    If ControlExists("cmdClickPt2") Then
        cmdClickPt2.Caption = "Click Point 2 in Drawing"
        cmdClickPt2.Top     = 216
        cmdClickPt2.Left    = 10
        cmdClickPt2.Width   = 165
        cmdClickPt2.Height  = 22
        cmdClickPt2.Font.Size = 8
    End If

    If ControlExists("lblMstn2") Then
        lblMstn2.Caption   = "MicroStation: not captured"
        lblMstn2.Top       = 244
        lblMstn2.Left      = 10
        lblMstn2.Width     = 375
        lblMstn2.Height    = 14
        lblMstn2.Font.Size = 8
        lblMstn2.ForeColor = RGB(100, 100, 100)
    End If

    ' ---- CALIBRATE BUTTON ----
    If ControlExists("cmdCompute") Then
        cmdCompute.Caption   = "Calibrate"
        cmdCompute.Top       = 266
        cmdCompute.Left      = 10
        cmdCompute.Width     = 80
        cmdCompute.Height    = 23
        cmdCompute.Font.Bold = True
    End If

    If ControlExists("lblCalibResult") Then
        lblCalibResult.Caption   = ""
        lblCalibResult.Top       = 296
        lblCalibResult.Left      = 10
        lblCalibResult.Width     = 375
        lblCalibResult.Height    = 40
        lblCalibResult.Font.Size = 8
        lblCalibResult.WordWrap  = True
        lblCalibResult.ForeColor = RGB(60, 60, 60)
    End If

    ' ---- CLOSE BUTTON ----
    If ControlExists("cmdClose") Then
        cmdClose.Caption = "Close"
        cmdClose.Top     = 340
        cmdClose.Left    = 10
        cmdClose.Width   = 70
        cmdClose.Height  = 22
    End If

    Me.Height = 380
End Sub

' ============================================================
' cmdClickPt1_Click — collect MicroStation point 1 via CadInputQueue
' ============================================================
Private Sub cmdClickPt1_Click()
    Me.Hide
    CadInputQueue.SendKeyin "ECHO Calibration: Click reference point 1 in the MicroStation drawing (right-click to cancel)"

    Dim oMsg As CadInputMessage
    Set oMsg = CadInputQueue.GetInput

    Do While oMsg.InputType <> msdCadInputTypeDataPoint
        If oMsg.InputType = msdCadInputTypeReset Then
            CommandState.StartDefaultCommand
            Me.Show vbModeless
            Exit Sub
        End If
        Set oMsg = CadInputQueue.GetInput
    Loop

    CommandState.StartDefaultCommand

    bbMstnRef1X = oMsg.Point.X
    bbMstnRef1Y = oMsg.Point.Y
    mstn1Captured = True

    Me.Show vbModeless

    If ControlExists("lblMstn1") Then
        lblMstn1.Caption   = "MicroStation: X = " & Format(bbMstnRef1X, "0.00") & _
                             ",  Y = " & Format(bbMstnRef1Y, "0.00") & "  (captured)"
        lblMstn1.ForeColor = RGB(0, 100, 0)
    End If
End Sub

' ============================================================
' cmdClickPt2_Click — collect MicroStation point 2 via CadInputQueue
' ============================================================
Private Sub cmdClickPt2_Click()
    Me.Hide
    CadInputQueue.SendKeyin "ECHO Calibration: Click reference point 2 in the MicroStation drawing (right-click to cancel)"

    Dim oMsg As CadInputMessage
    Set oMsg = CadInputQueue.GetInput

    Do While oMsg.InputType <> msdCadInputTypeDataPoint
        If oMsg.InputType = msdCadInputTypeReset Then
            CommandState.StartDefaultCommand
            Me.Show vbModeless
            Exit Sub
        End If
        Set oMsg = CadInputQueue.GetInput
    Loop

    CommandState.StartDefaultCommand

    bbMstnRef2X = oMsg.Point.X
    bbMstnRef2Y = oMsg.Point.Y
    mstn2Captured = True

    Me.Show vbModeless

    If ControlExists("lblMstn2") Then
        lblMstn2.Caption   = "MicroStation: X = " & Format(bbMstnRef2X, "0.00") & _
                             ",  Y = " & Format(bbMstnRef2Y, "0.00") & "  (captured)"
        lblMstn2.ForeColor = RGB(0, 100, 0)
    End If
End Sub

' ============================================================
' cmdCompute_Click — validate inputs and compute calibration
' ============================================================
Private Sub cmdCompute_Click()
    On Error GoTo ComputeErr

    ' Validate PDF coordinate inputs
    If Not ControlExists("txtPdf1X") Or Not ControlExists("txtPdf1Y") Or _
       Not ControlExists("txtPdf2X") Or Not ControlExists("txtPdf2Y") Then
        MsgBox "Required text boxes are missing. Check form setup.", vbExclamation, "Bluebeam Importer"
        Exit Sub
    End If

    If Len(Trim(txtPdf1X.Text)) = 0 Or Len(Trim(txtPdf1Y.Text)) = 0 Or _
       Len(Trim(txtPdf2X.Text)) = 0 Or Len(Trim(txtPdf2Y.Text)) = 0 Then
        MsgBox "Enter PDF X and Y coordinates for both reference points.", vbExclamation, "Bluebeam Importer"
        Exit Sub
    End If

    If Not mstn1Captured Then
        MsgBox "Click 'Click Point 1 in Drawing' to capture the MicroStation location for reference point 1.", _
               vbExclamation, "Bluebeam Importer"
        Exit Sub
    End If

    If Not mstn2Captured Then
        MsgBox "Click 'Click Point 2 in Drawing' to capture the MicroStation location for reference point 2.", _
               vbExclamation, "Bluebeam Importer"
        Exit Sub
    End If

    ' Parse PDF coordinates
    bbPdfRef1X = CDbl(Trim(txtPdf1X.Text))
    bbPdfRef1Y = CDbl(Trim(txtPdf1Y.Text))
    bbPdfRef2X = CDbl(Trim(txtPdf2X.Text))
    bbPdfRef2Y = CDbl(Trim(txtPdf2Y.Text))

    ' Check points are not identical
    If Abs(bbPdfRef2X - bbPdfRef1X) < 0.001 And Abs(bbPdfRef2Y - bbPdfRef1Y) < 0.001 Then
        MsgBox "PDF reference points appear to be the same location. " & _
               "Choose two distinct points far apart on the sheet.", vbExclamation, "Bluebeam Importer"
        Exit Sub
    End If

    ' Compute calibration (sets bbScaleX, bbScaleY, bbOffsetX, bbOffsetY, bbCalibrated)
    ComputeCalibration

    If bbCalibrated Then
        If ControlExists("lblCalibResult") Then
            lblCalibResult.Caption   = "Calibration complete." & vbCrLf & _
                                       "Scale: " & Format(bbScaleX, "0.000") & " ft/PDF-unit" & vbCrLf & _
                                       "Offset: X = " & Format(bbOffsetX, "0.0") & _
                                       ", Y = " & Format(bbOffsetY, "0.0") & " ft"
            lblCalibResult.ForeColor = RGB(0, 100, 0)
        End If

        ' Notify main form to refresh its calibration status
        frmBBImport.UpdateCalibStatus
    End If

    Exit Sub

ComputeErr:
    MsgBox "Error during calibration: " & Err.Description, vbCritical, "Bluebeam Importer"
End Sub

' ============================================================
' cmdClose_Click
' ============================================================
Private Sub cmdClose_Click()
    Unload Me
End Sub
