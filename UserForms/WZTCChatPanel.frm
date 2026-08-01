Option Explicit

' ============================================================
' WZTCChatPanel (M7) -- in-MicroStation agent chat panel.
' Modeless. Launch via Modules/Launcher.bas:LaunchChatPanel --
' NOT via F5 on this form's designer/code window in the VBA IDE.
' F5 on a UserForm runs an implicit Show with no argument, which
' defaults to vbModal and blocks clicking anywhere else until
' the form closes (this was the "can't click out" symptom).
'
' Thin by design: this form has zero knowledge of the Anthropic
' API or the agent loop. It only does three things --
'   1. On Send, append the typed line to Bridge/chat-input.tsv.
'   2. Poll Bridge/chat-log.tsv (via WZTCChatTimer.bas) for new
'      lines as chat_driver.py (M7 Stage 4) appends them.
'   3. Parse each line's TYPE (see FormatLogLine) and route it
'      to one of two panes (see AppendChatLine), instead of
'      showing the raw timestamp\tTYPE\tkey=val... line.
' The actual thinking/tool-calling happens in that separate
' Python process, driving the same WZTCBridge.ExecuteOp dispatch
' every other op already uses -- so watching a tool call happen
' here is the same visible MicroStation activity as today,
' nothing new to prove there.
'
' Layout -- two read-only panes plus the input row, closest VBA
' equivalent of the Claude Code in VS Code chat panel:
'   txtConversation -- the actual back-and-forth: your messages,
'                      the agent's final answers, ASK_USER
'                      questions, and errors.
'   txtActivity     -- the agent's live work: THINKING text and
'                      TOOL_CALL/TOOL_RESULT lines.
' Both are MultiLine + WordWrap TextBoxes, not a ListBox -- a
' ListBox item doesn't word-wrap in MSForms, so a long agent
' answer became one unreadable, horizontally-scrolling line
' (the root cause of "can't see the full output").
'
' Controls must be added manually in the VBA IDE (File Sync
' Protocol, CLAUDE.md). Delete the old lstTranscript ListBox if
' it's still on the form -- replaced by the two textboxes below.
'   txtConversation  TextBox: MultiLine=True, WordWrap=True,
'                    ScrollBars=2-fmScrollBarsVertical,
'                    Locked=True, TabStop=False
'   txtActivity      TextBox: same settings as txtConversation
'   txtInput         TextBox: plain single-line, as before
'   btnSend          CommandButton, as before
'   lblStatus        Label, as before
' ============================================================

Private Const CHAT_LOG_FILE As String = "c:\repos\microstation-vba-project\Bridge\chat-log.tsv"
Private Const CHAT_INPUT_FILE As String = "c:\repos\microstation-vba-project\Bridge\chat-input.tsv"

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

    Me.Caption = "WZTC Agent Chat"
    Me.Width = 620
    Me.Height = 640

    If ControlExists("txtConversation") Then
        With txtConversation
            .Top = 10: .Left = 10: .Width = 590: .Height = 300
            .MultiLine = True
            .WordWrap = True
            .ScrollBars = 2   ' fmScrollBarsVertical
            .Locked = True
            .TabStop = False
            .Text = ""
        End With
    End If

    If ControlExists("txtActivity") Then
        With txtActivity
            .Top = 320: .Left = 10: .Width = 590: .Height = 170
            .MultiLine = True
            .WordWrap = True
            .ScrollBars = 2   ' fmScrollBarsVertical
            .Locked = True
            .TabStop = False
            .Text = ""
        End With
    End If

    If ControlExists("txtInput") Then
        With txtInput
            .Top = 500: .Left = 10: .Width = 505: .Height = 24
            .Text = ""
        End With
    End If

    If ControlExists("btnSend") Then
        With btnSend
            .Caption = "Send"
            .Top = 500: .Left = 520: .Width = 80: .Height = 24
            .Font.Bold = True
        End With
    End If

    If ControlExists("lblStatus") Then
        With lblStatus
            .Caption = "Ready. (If nothing responds, make sure chat_driver.py is running.)"
            .Top = 530: .Left = 10: .Width = 590: .Height = 34
            .WordWrap = True
            .ForeColor = RGB(0, 100, 0)
        End With
    End If

    Call WZTCChatTimer.StartChatTimer(Me, CHAT_LOG_FILE)
End Sub

' ============================================================
' APPEND A LINE -- called by WZTCChatTimer.ChatTimerProc for
' every new line it finds in CHAT_LOG_FILE. Public so the timer
' callback (in a standard module) can reach it. Parses the
' line's TYPE, routes it to txtConversation (the actual
' back-and-forth) or txtActivity (the agent's live work), and
' updates lblStatus so "waiting for your reply" vs "ready" is
' visible without reading either pane.
' ============================================================
Public Sub AppendChatLine(rawLine As String)
    Dim lineType As String
    Dim display As String
    display = FormatLogLine(rawLine, lineType)

    Select Case lineType
        Case "THINKING", "TOOL_CALL", "TOOL_RESULT"
            If ControlExists("txtActivity") Then AppendTo txtActivity, display
        Case Else
            If ControlExists("txtConversation") Then AppendTo txtConversation, display
    End Select

    If ControlExists("lblStatus") Then
        Select Case lineType
            Case "ASK_USER"
                lblStatus.Caption = "Agent is waiting for your reply..."
                lblStatus.ForeColor = RGB(180, 120, 0)
            Case "FINAL"
                lblStatus.Caption = "Ready."
                lblStatus.ForeColor = RGB(0, 100, 0)
            Case "ERROR"
                lblStatus.Caption = "Agent hit an error -- see transcript. Still ready for another message."
                lblStatus.ForeColor = RGB(180, 0, 0)
        End Select
    End If
End Sub

' Appends display text to a read-only textbox pane, separated
' from the previous entry by a blank line, and scrolls it into
' view. A TextBox has no TopIndex like a ListBox -- setting
' SelStart past the end of .Text (with the box still Locked)
' moves the caret there and the visible scroll position follows
' it, which is the standard MSForms way to auto-scroll a
' read-only multiline textbox.
Private Sub AppendTo(box As Object, display As String)
    If box Is Nothing Then Exit Sub
    If box.Text = "" Then
        box.Text = display
    Else
        box.Text = box.Text & vbCrLf & vbCrLf & display
    End If
    box.SelStart = Len(box.Text)
End Sub

' ============================================================
' PARSE ONE chat-log.tsv LINE INTO DISPLAY TEXT
' Schema (chat_driver.py's ChatLog class):
'   timestamp<TAB>TYPE<TAB>key=val<TAB>key=val...
' Returns the formatted display string; outLineType receives the
' raw TYPE so AppendChatLine can react to it (e.g. ASK_USER).
' An unrecognized TYPE renders the raw line rather than being
' dropped silently -- same honesty-over-silence principle used
' throughout this bridge (an unexpected line is still shown, not
' hidden from the engineer).
' ============================================================
Private Function FormatLogLine(rawLine As String, ByRef outLineType As String) As String
    outLineType = ""
    Dim parts() As String
    parts = Split(rawLine, vbTab)
    If UBound(parts) < 1 Then
        FormatLogLine = rawLine
        Exit Function
    End If

    Dim lineType As String: lineType = parts(1)
    outLineType = lineType

    Dim fields As Object
    Set fields = CreateObject("Scripting.Dictionary")
    Dim i As Integer
    For i = 2 To UBound(parts)
        Dim eqPos As Integer: eqPos = InStr(parts(i), "=")
        If eqPos > 0 Then
            fields(Left(parts(i), eqPos - 1)) = Mid(parts(i), eqPos + 1)
        End If
    Next i

    Select Case lineType
        Case "USER_ECHO"
            FormatLogLine = "[you] " & FieldOrBlank(fields, "text")
        Case "THINKING"
            FormatLogLine = "  " & FieldOrBlank(fields, "text")
        Case "TOOL_CALL"
            FormatLogLine = "[tool] " & FieldOrBlank(fields, "name") & "(" & FieldOrBlank(fields, "input") & ")"
        Case "TOOL_RESULT"
            FormatLogLine = "   -> " & FieldOrBlank(fields, "status") & " " & FieldOrBlank(fields, "summary")
        Case "ASK_USER"
            FormatLogLine = "[agent asks] " & FieldOrBlank(fields, "question")
        Case "FINAL"
            FormatLogLine = "[agent] " & FieldOrBlank(fields, "text")
        Case "ERROR"
            FormatLogLine = "[ERROR] " & FieldOrBlank(fields, "note")
        Case Else
            FormatLogLine = rawLine
    End Select
End Function

Private Function FieldOrBlank(fields As Object, key As String) As String
    If fields.Exists(key) Then
        FieldOrBlank = fields(key)
    Else
        FieldOrBlank = ""
    End If
End Function

' ============================================================
' SEND -- append the typed line to Bridge/chat-input.tsv, which
' chat_driver.py polls (InputWatcher). VBA's Print# already
' writes CRLF-terminated lines natively, matching what
' chat_driver.py expects (Line Input#-safe on the VBA-read side
' too, for symmetry with everything else in this bridge).
' ============================================================
Private Sub btnSend_Click()
    On Error GoTo SendErr
    If Not ControlExists("txtInput") Then Exit Sub

    Dim msg As String
    msg = Trim(txtInput.Text)
    If msg = "" Then Exit Sub

    Dim fnum As Integer
    fnum = FreeFile
    Open CHAT_INPUT_FILE For Append As #fnum
    Print #fnum, Now & vbTab & msg
    Close #fnum

    AppendChatLine Now & vbTab & "USER_ECHO" & vbTab & "text=" & msg
    txtInput.Text = ""
    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Waiting for the agent..."
        lblStatus.ForeColor = RGB(0, 0, 180)
    End If
    Exit Sub

SendErr:
    If fnum <> 0 Then Close #fnum
    If ControlExists("lblStatus") Then lblStatus.Caption = "Send failed: " & Err.Description
End Sub

Private Sub txtInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnSend_Click
        KeyCode = 0
    End If
End Sub

' ============================================================
' TEARDOWN -- KillTimer on every path out of the form, guarded
' against double-kill inside StopChatTimer itself.
' ============================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call WZTCChatTimer.StopChatTimer
End Sub

Private Sub UserForm_Terminate()
    Call WZTCChatTimer.StopChatTimer
End Sub
