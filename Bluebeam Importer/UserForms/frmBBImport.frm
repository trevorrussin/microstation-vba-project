Option Explicit

' ============================================================
' frmBBImport.frm — Bluebeam Markup Importer (main form)
' ------------------------------------------------------------
' Controls to add manually in the VBA IDE form designer:
'   lblFileTitle      - Label          "Markup File:"
'   txtFilePath       - TextBox        (read-only; shows loaded file path)
'   cmdBrowse         - CommandButton  "Browse..."
'   cmdCalibrate      - CommandButton  "Set Reference Points"
'   lblCalibStatus    - Label          (calibration status)
'   lblListTitle      - Label          "Markups:"
'   lstMarkups        - ListBox        (4 columns: #, Operation, Status, Comment)
'   cmdProcessAll     - CommandButton  "Process All"
'   cmdProcessSel     - CommandButton  "Process Selected"
'   cmdSkip           - CommandButton  "Skip"
'   cmdReprocess      - CommandButton  "Reprocess"
'   lblStatus         - Label          (running status / errors)
' ============================================================

Private Function ControlExists(ctrlName As String) As Boolean
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = Me.Controls(ctrlName)
    ControlExists = Not (ctrl Is Nothing)
    On Error GoTo 0
End Function

' ============================================================
' INITIALIZE — layout and default state
' ============================================================
Private Sub UserForm_Initialize()
    Me.Caption = "Bluebeam Markup Importer"
    Me.Width  = 560
    Me.Height = 430

    ' ---- FILE PATH SECTION ----
    If ControlExists("lblFileTitle") Then
        lblFileTitle.Caption   = "Markup File:"
        lblFileTitle.Top       = 8
        lblFileTitle.Left      = 10
        lblFileTitle.Width     = 80
        lblFileTitle.Height    = 16
        lblFileTitle.Font.Bold = True
    End If

    If ControlExists("txtFilePath") Then
        txtFilePath.Top     = 7
        txtFilePath.Left    = 90
        txtFilePath.Width   = 350
        txtFilePath.Height  = 18
        txtFilePath.Locked  = True
        txtFilePath.Text    = "(no file loaded)"
        txtFilePath.BackColor = RGB(240, 240, 240)
    End If

    If ControlExists("cmdBrowse") Then
        cmdBrowse.Caption   = "Browse..."
        cmdBrowse.Top       = 5
        cmdBrowse.Left      = 447
        cmdBrowse.Width     = 90
        cmdBrowse.Height    = 22
    End If

    ' ---- CALIBRATION ROW ----
    If ControlExists("cmdCalibrate") Then
        cmdCalibrate.Caption   = "Set Reference Points"
        cmdCalibrate.Top       = 34
        cmdCalibrate.Left      = 10
        cmdCalibrate.Width     = 150
        cmdCalibrate.Height    = 22
    End If

    If ControlExists("lblCalibStatus") Then
        lblCalibStatus.Caption   = Chr(9679) & " Not calibrated — coordinates will require manual click"
        lblCalibStatus.Top       = 38
        lblCalibStatus.Left      = 168
        lblCalibStatus.Width     = 368
        lblCalibStatus.Height    = 16
        lblCalibStatus.Font.Size = 8
        lblCalibStatus.ForeColor = RGB(160, 0, 0)
    End If

    ' ---- MARKUP LIST ----
    If ControlExists("lblListTitle") Then
        lblListTitle.Caption   = "Markups:"
        lblListTitle.Top       = 63
        lblListTitle.Left      = 10
        lblListTitle.Width     = 535
        lblListTitle.Height    = 14
        lblListTitle.Font.Bold = True
        lblListTitle.Font.Size = 8
    End If

    If ControlExists("lstMarkups") Then
        lstMarkups.Top          = 80
        lstMarkups.Left         = 10
        lstMarkups.Width        = 535
        lstMarkups.Height       = 220
        lstMarkups.ColumnCount  = 4
        lstMarkups.ColumnWidths = "38 pt;80 pt;55 pt;0 pt"
        lstMarkups.Font.Size    = 8
        lstMarkups.MultiSelect  = 0   ' fmMultiSelectSingle
    End If

    ' ---- ACTION BUTTONS ----
    If ControlExists("cmdProcessAll") Then
        cmdProcessAll.Caption   = "Process All"
        cmdProcessAll.Top       = 308
        cmdProcessAll.Left      = 10
        cmdProcessAll.Width     = 95
        cmdProcessAll.Height    = 23
        cmdProcessAll.Font.Bold = True
        cmdProcessAll.Enabled   = False
    End If

    If ControlExists("cmdProcessSel") Then
        cmdProcessSel.Caption   = "Process Selected"
        cmdProcessSel.Top       = 308
        cmdProcessSel.Left      = 113
        cmdProcessSel.Width     = 115
        cmdProcessSel.Height    = 23
        cmdProcessSel.Enabled   = False
    End If

    If ControlExists("cmdSkip") Then
        cmdSkip.Caption   = "Skip"
        cmdSkip.Top       = 308
        cmdSkip.Left      = 236
        cmdSkip.Width     = 60
        cmdSkip.Height    = 23
        cmdSkip.Enabled   = False
    End If

    If ControlExists("cmdReprocess") Then
        cmdReprocess.Caption   = "Reprocess"
        cmdReprocess.Top       = 308
        cmdReprocess.Left      = 304
        cmdReprocess.Width     = 80
        cmdReprocess.Height    = 23
        cmdReprocess.Enabled   = False
    End If

    ' ---- STATUS LABEL ----
    If ControlExists("lblStatus") Then
        lblStatus.Caption   = "Ready. Browse for a Bluebeam XML or CSV export file, then click Process All."
        lblStatus.Top       = 338
        lblStatus.Left      = 10
        lblStatus.Width     = 535
        lblStatus.Height    = 50
        lblStatus.Font.Size = 8
        lblStatus.WordWrap  = True
        lblStatus.ForeColor = RGB(60, 60, 60)
    End If

    Me.Height = 410
End Sub

' ============================================================
' cmdBrowse_Click — open file path input dialog
' Uses InputBox as a universal MicroStation-compatible fallback.
' User can copy/paste path from Windows Explorer (Shift+Right-click > "Copy as path").
' ============================================================
Private Sub cmdBrowse_Click()
    Dim filePath As String
    filePath = InputBox("Paste the full path to your Bluebeam markup export file." & vbCrLf & _
                        "Supported formats: .xml (Bluebeam XML Summary) or .csv (Custom CSV Export)" & vbCrLf & vbCrLf & _
                        "Tip: In Windows Explorer, Shift+Right-click the file > 'Copy as path'." & vbCrLf & _
                        "Then paste here with Ctrl+V.", _
                        "Select Markup File", bbLoadedFilePath)

    filePath = Trim(filePath)
    If Len(filePath) = 0 Then Exit Sub

    ' Strip surrounding quotes (from "Copy as path")
    If Left(filePath, 1) = Chr(34) And Right(filePath, 1) = Chr(34) Then
        filePath = Mid(filePath, 2, Len(filePath) - 2)
    End If

    ' Verify file exists
    If Len(Dir(filePath)) = 0 Then
        SetStatus "File not found: " & filePath, True
        Exit Sub
    End If

    ' Parse the file
    SetStatus "Loading " & filePath & "..."
    bbMarkupCount = 0

    If Not ParseMarkupFile(filePath) Then
        SetStatus "Failed to parse file. Verify it is a valid Bluebeam XML or CSV export.", True
        Exit Sub
    End If

    ' Classify all markups
    ParseAllMarkups

    ' Convert coordinates if already calibrated
    If bbCalibrated Then
        ConvertAllMarkups
    End If

    ' Update UI
    If ControlExists("txtFilePath") Then
        txtFilePath.Text = filePath
    End If

    RefreshMarkupList
    UpdateButtons

    Dim coordCount As Integer, unknownCount As Integer
    Dim i As Integer
    For i = 1 To bbMarkupCount
        If bbMarkups(i).HasCoord Then coordCount = coordCount + 1
        If bbMarkups(i).OpType = "UNKNOWN" Then unknownCount = unknownCount + 1
    Next i

    Dim msg As String
    msg = bbMarkupCount & " markup(s) loaded."
    If coordCount < bbMarkupCount Then
        msg = msg & " " & (bbMarkupCount - coordCount) & " without coordinates (manual click required)."
    End If
    If unknownCount > 0 Then
        msg = msg & " " & unknownCount & " unrecognized (will be skipped)."
    End If
    SetStatus msg, False
End Sub

' ============================================================
' cmdCalibrate_Click — open coordinate calibration form
' ============================================================
Private Sub cmdCalibrate_Click()
    frmBBCalibrate.Show vbModeless
End Sub

' ============================================================
' cmdProcessAll_Click — process all Pending markups in sequence
' ============================================================
Private Sub cmdProcessAll_Click()
    If bbMarkupCount = 0 Then
        SetStatus "No markups loaded. Browse for a file first.", True
        Exit Sub
    End If

    If Not bbCalibrated Then
        Dim ans As Integer
        ans = MsgBox("Coordinate calibration has not been set." & vbCrLf & _
                     "Markups without inline coordinates will pause and ask you to click" & vbCrLf & _
                     "the target element manually in MicroStation." & vbCrLf & vbCrLf & _
                     "Continue anyway?", vbYesNo Or vbQuestion, "Bluebeam Importer")
        If ans = vbNo Then Exit Sub
    End If

    DisableButtons

    Dim i As Integer
    Dim doneCount As Integer, skipCount As Integer, errCount As Integer
    doneCount = 0: skipCount = 0: errCount = 0

    For i = 1 To bbMarkupCount
        If bbMarkups(i).Status = "Pending" Then
            SetStatus "Processing " & bbMarkups(i).MarkupID & " of " & bbMarkupCount & _
                      " [" & bbMarkups(i).OpType & "]: " & Left(bbMarkups(i).RawText, 50) & "..."
            DoEvents
            ProcessMarkup i
            RefreshMarkupRow i
        End If

        Select Case bbMarkups(i).Status
            Case "Done":    doneCount = doneCount + 1
            Case "Skipped": skipCount = skipCount + 1
            Case "Error":   errCount = errCount + 1
        End Select
    Next i

    UpdateButtons

    SetStatus "Complete. " & doneCount & " done, " & skipCount & " skipped, " & errCount & " errors." & _
              IIf(errCount > 0, " Click an error row to see details.", ""), (errCount > 0)
End Sub

' ============================================================
' cmdProcessSel_Click — process the currently selected markup
' ============================================================
Private Sub cmdProcessSel_Click()
    If Not ControlExists("lstMarkups") Then Exit Sub
    If lstMarkups.ListIndex < 0 Then
        SetStatus "Select a markup row first.", True
        Exit Sub
    End If

    Dim idx As Integer
    idx = lstMarkups.ListIndex + 1   ' ListBox is 0-based; array is 1-based

    If idx < 1 Or idx > bbMarkupCount Then Exit Sub

    If bbMarkups(idx).Status = "Done" Then
        SetStatus bbMarkups(idx).MarkupID & " is already done. Use Reprocess to run again.", False
        Exit Sub
    End If

    bbMarkups(idx).Status = "Pending"   ' reset before processing
    ProcessMarkup idx
    RefreshMarkupRow idx
    SetStatus bbMarkups(idx).MarkupID & ": " & bbMarkups(idx).Status & _
              IIf(Len(bbMarkups(idx).StatusNote) > 0, " — " & bbMarkups(idx).StatusNote, ""), _
              (bbMarkups(idx).Status = "Error")
End Sub

' ============================================================
' cmdSkip_Click — mark selected markup as Skipped
' ============================================================
Private Sub cmdSkip_Click()
    If Not ControlExists("lstMarkups") Then Exit Sub
    If lstMarkups.ListIndex < 0 Then Exit Sub
    Dim idx As Integer
    idx = lstMarkups.ListIndex + 1
    If idx < 1 Or idx > bbMarkupCount Then Exit Sub
    bbMarkups(idx).Status = "Skipped"
    bbMarkups(idx).StatusNote = "Manually skipped"
    RefreshMarkupRow idx
    SetStatus bbMarkups(idx).MarkupID & " skipped.", False
End Sub

' ============================================================
' cmdReprocess_Click — reset selected markup to Pending
' ============================================================
Private Sub cmdReprocess_Click()
    If Not ControlExists("lstMarkups") Then Exit Sub
    If lstMarkups.ListIndex < 0 Then Exit Sub
    Dim idx As Integer
    idx = lstMarkups.ListIndex + 1
    If idx < 1 Or idx > bbMarkupCount Then Exit Sub
    bbMarkups(idx).Status = "Pending"
    bbMarkups(idx).StatusNote = ""
    RefreshMarkupRow idx
    SetStatus bbMarkups(idx).MarkupID & " reset to Pending. Click Process Selected to run.", False
End Sub

' ============================================================
' lstMarkups_Click — show full detail of selected markup in status
' ============================================================
Private Sub lstMarkups_Click()
    If Not ControlExists("lstMarkups") Then Exit Sub
    If lstMarkups.ListIndex < 0 Then Exit Sub
    Dim idx As Integer
    idx = lstMarkups.ListIndex + 1
    If idx < 1 Or idx > bbMarkupCount Then Exit Sub

    Dim m As BBMarkup
    m = bbMarkups(idx)

    Dim detail As String
    detail = m.MarkupID & " [" & m.OpType & "]"
    If Len(m.Param1) > 0 Then detail = detail & "  Param: " & m.Param1
    If Len(m.Param2) > 0 Then detail = detail & ", " & m.Param2 & " ft"
    detail = detail & vbCrLf & "Text: " & m.RawText
    If m.HasCoord Then
        detail = detail & vbCrLf & "PDF: (" & Format(m.PdfX, "0.0") & ", " & Format(m.PdfY, "0.0") & ")"
        If bbCalibrated Then
            detail = detail & "  Mstn: (" & Format(m.MstnX, "0.0") & ", " & Format(m.MstnY, "0.0") & ")"
        End If
    Else
        detail = detail & vbCrLf & "No coordinates — will prompt for click"
    End If
    If Len(m.StatusNote) > 0 Then
        detail = detail & vbCrLf & "Status: " & m.StatusNote
    End If

    SetStatus detail, (m.Status = "Error")
End Sub

' ============================================================
' RefreshMarkupList — rebuild lstMarkups from bbMarkups array
' ============================================================
Public Sub RefreshMarkupList()
    If Not ControlExists("lstMarkups") Then Exit Sub

    lstMarkups.Clear

    Dim i As Integer
    For i = 1 To bbMarkupCount
        Dim m As BBMarkup
        m = bbMarkups(i)
        lstMarkups.AddItem m.MarkupID
        lstMarkups.List(i - 1, 1) = m.OpType
        lstMarkups.List(i - 1, 2) = m.Status
        lstMarkups.List(i - 1, 3) = Left(m.RawText, 55)
    Next i

    If ControlExists("lblListTitle") Then
        lblListTitle.Caption = "Markups (" & bbMarkupCount & " loaded):"
    End If
End Sub

' ============================================================
' RefreshMarkupRow — update a single row in lstMarkups
' ============================================================
Private Sub RefreshMarkupRow(idx As Integer)
    If Not ControlExists("lstMarkups") Then Exit Sub
    If idx < 1 Or idx > bbMarkupCount Then Exit Sub

    Dim row As Integer
    row = idx - 1   ' 0-based ListBox row

    Dim m As BBMarkup
    m = bbMarkups(idx)

    lstMarkups.List(row, 0) = m.MarkupID
    lstMarkups.List(row, 1) = m.OpType
    lstMarkups.List(row, 2) = m.Status
    lstMarkups.List(row, 3) = Left(m.RawText, 55)
End Sub

' ============================================================
' Public UpdateCalibStatus — called by frmBBCalibrate after calibration
' ============================================================
Public Sub UpdateCalibStatus()
    If ControlExists("lblCalibStatus") Then
        If bbCalibrated Then
            lblCalibStatus.Caption   = Chr(9679) & " " & CalibrationSummary()
            lblCalibStatus.ForeColor = RGB(0, 120, 0)
        Else
            lblCalibStatus.Caption   = Chr(9679) & " Not calibrated — coordinates will require manual click"
            lblCalibStatus.ForeColor = RGB(160, 0, 0)
        End If
    End If

    ' Re-convert all markup coordinates with new calibration
    If bbCalibrated And bbMarkupCount > 0 Then
        ConvertAllMarkups
    End If
End Sub

' ============================================================
' SetStatus — update lblStatus; red if isError
' ============================================================
Private Sub SetStatus(msg As String, Optional isError As Boolean = False)
    If ControlExists("lblStatus") Then
        lblStatus.Caption   = msg
        lblStatus.ForeColor = IIf(isError, RGB(180, 0, 0), RGB(60, 60, 60))
    End If
    DoEvents
End Sub

' ============================================================
' UpdateButtons — enable/disable based on load state
' ============================================================
Private Sub UpdateButtons()
    Dim hasData As Boolean
    hasData = (bbMarkupCount > 0)

    If ControlExists("cmdProcessAll") Then cmdProcessAll.Enabled = hasData
    If ControlExists("cmdProcessSel") Then cmdProcessSel.Enabled = hasData
    If ControlExists("cmdSkip") Then cmdSkip.Enabled = hasData
    If ControlExists("cmdReprocess") Then cmdReprocess.Enabled = hasData
End Sub

Private Sub DisableButtons()
    If ControlExists("cmdProcessAll") Then cmdProcessAll.Enabled = False
    If ControlExists("cmdProcessSel") Then cmdProcessSel.Enabled = False
    If ControlExists("cmdSkip") Then cmdSkip.Enabled = False
    If ControlExists("cmdReprocess") Then cmdReprocess.Enabled = False
End Sub
