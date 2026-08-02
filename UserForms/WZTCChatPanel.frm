Option Explicit

' ============================================================
' WZTCChatPanel (M7) -- in-MicroStation agent chat panel.
' Modeless. Prefer launching via Modules/Launcher.bas:LaunchChatPanel
' (explicit Show vbModeless) over F5 on this form's designer/code
' window in the VBA IDE -- F5 on a UserForm runs an implicit Show
' with no argument, which defaults to vbModal and disables
' MicroStation's main window (and everything else) until the form
' closes (the "can't click anything in MicroStation" symptom,
' confirmed live 2026-08-02).
'
' UserForm_Activate below makes this self-correcting either way: the
' first time the form activates, it unconditionally hides and
' re-Shows itself vbModeless, then never repeats. If it was already
' modeless (the Launcher path), this is a harmless no-op flicker; if
' it came up modal (F5), Me.Hide unblocks the implicit modal Show
' call (Hide, unlike Unload, lets a blocked Show return while
' keeping the form loaded) and the immediate Show vbModeless
' re-displays it non-blocking. There's no way to detect *how* Show
' was called from inside the form itself, so this always runs rather
' than branching on it -- simpler than it sounds, and confirmed not
' to recurse (see the mForcedModeless guard).
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
'   imgScreenshot    Image control (added 2026-08-02): shows the most
'                    recent auto-captured screenshot of what the agent
'                    just did (see chat_driver.py's _auto_focus_and_
'                    capture -- pans/zooms the view to everything the
'                    turn touched, then screenshots it). Set
'                    PictureSizeMode = 3 - fmPictureSizeModeZoom in the
'                    IDE so it scales to fit the control without
'                    distortion; BorderStyle = 1 - fmBorderStyleSingle
'                    looks best against the modernized panel background.
'   txtActivity      TextBox: same settings as txtConversation
'   txtInput         TextBox: plain single-line, as before
'   btnSend          CommandButton, as before
'   lblStatus        Label, as before
' ============================================================

Private Const CHAT_LOG_FILE As String = "c:\repos\microstation-vba-project\Bridge\chat-log.tsv"
Private Const CHAT_INPUT_FILE As String = "c:\repos\microstation-vba-project\Bridge\chat-input.tsv"

' ============================================================
' DOCK TO RIGHT EDGE (2026-08-02 feedback) -- MSForms has no real docking
' (that's MicroStation's own tool-palette framework, not available to a
' VBA UserForm), so this is the closest equivalent: position against the
' right edge of the screen and use its full height every time the panel
' loads, rather than a small fixed-size floating window. GetSystemMetrics
' returns PIXELS; UserForm Left/Top/Width/Height are all in POINTS
' (1/72in) -- PIXELS_TO_POINTS is the standard 96-DPI (72/96) conversion.
' ============================================================
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const PIXELS_TO_POINTS As Double = 0.75
Private Const TASKBAR_MARGIN_PTS As Double = 40   ' leaves room for the taskbar

' Guards UserForm_Activate's self-correcting Hide+Show vbModeless below
' against recursion -- Show always fires Activate again, so without this
' the second (already-modeless) Activate would re-trigger the same
' Hide+Show forever.
Private mForcedModeless As Boolean

' ============================================================
' ACTIVATE -- see the header comment above for why this exists.
' ============================================================
Private Sub UserForm_Activate()
    If mForcedModeless Then Exit Sub
    mForcedModeless = True
    Me.Hide
    Me.Show vbModeless
End Sub

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
' Modernized 2026-08-02: Segoe UI throughout (Windows' modern system UI
' font, universally available, a plain visual upgrade over MSForms'
' default MS Sans Serif) and a light neutral palette that visually
' separates the three panes -- white for the actual conversation
' (primary content), a soft blue-gray for the screenshot frame, a soft
' gray for the activity/reasoning trace (secondary/detail content).
' MSForms TextBox has no rich-text/per-line color support, so within a
' single box every line shares one color -- that's a real, unavoidable
' constraint of this UI framework, not an oversight.
Private Sub UserForm_Initialize()
    On Error Resume Next

    Me.Caption = "WZTC Agent Chat"
    Me.Width = 620
    Me.Height = GetSystemMetrics(SM_CYSCREEN) * PIXELS_TO_POINTS - TASKBAR_MARGIN_PTS
    Me.Left = GetSystemMetrics(SM_CXSCREEN) * PIXELS_TO_POINTS - Me.Width   ' dock to right edge
    Me.Top = 0
    Me.BackColor = RGB(244, 244, 247)

    ' Bottom-up layout: txtActivity/txtInput/btnSend/lblStatus are a fixed-
    ' height block anchored to the bottom of whatever Me.Height turned out
    ' to be; imgScreenshot fills everything between txtConversation and
    ' that block. This is what makes the image box grow to use the full
    ' docked height instead of leaving dead space below lblStatus (exactly
    ' the space the 2026-08-02 feedback flagged as wasted under the old
    ' fixed-position layout).
    Const ACTIVITY_HEIGHT As Double = 140
    Const INPUT_HEIGHT As Double = 24
    Const STATUS_HEIGHT As Double = 34
    Const GAP As Double = 8
    Const GAP_SMALL As Double = 6
    Const BOTTOM_MARGIN As Double = 8
    Const IMG_TOP As Double = 188

    Dim bottomBlockTop As Double
    bottomBlockTop = Me.Height - (ACTIVITY_HEIGHT + GAP + INPUT_HEIGHT + GAP_SMALL + STATUS_HEIGHT + BOTTOM_MARGIN)

    Dim imgHeight As Double
    imgHeight = bottomBlockTop - GAP - IMG_TOP
    If imgHeight < 100 Then imgHeight = 100   ' safety floor on an unexpectedly small screen

    If ControlExists("txtConversation") Then
        With txtConversation
            .Top = 10: .Left = 10: .Width = 590: .Height = 170
            .MultiLine = True
            .WordWrap = True
            .ScrollBars = 2   ' fmScrollBarsVertical
            .Locked = True
            .TabStop = False
            .Text = ""
            .BackColor = RGB(255, 255, 255)
            .ForeColor = RGB(30, 30, 30)
            .Font.Name = "Segoe UI"
            .Font.Size = 9.5
        End With
    End If

    If ControlExists("imgScreenshot") Then
        With Me.Controls("imgScreenshot")
            .Top = IMG_TOP: .Left = 10: .Width = 590: .Height = imgHeight
            .BackColor = RGB(230, 234, 240)
            .BorderStyle = 1   ' fmBorderStyleSingle
            .PictureSizeMode = 3   ' fmPictureSizeModeZoom -- fit without distortion
        End With
    End If

    If ControlExists("txtActivity") Then
        With txtActivity
            .Top = bottomBlockTop: .Left = 10: .Width = 590: .Height = ACTIVITY_HEIGHT
            .MultiLine = True
            .WordWrap = True
            .ScrollBars = 2   ' fmScrollBarsVertical
            .Locked = True
            .TabStop = False
            .Text = ""
            .BackColor = RGB(248, 248, 250)
            .ForeColor = RGB(70, 70, 75)
            .Font.Name = "Segoe UI"
            .Font.Size = 9
        End With
    End If

    Dim inputTop As Double
    inputTop = bottomBlockTop + ACTIVITY_HEIGHT + GAP

    If ControlExists("txtInput") Then
        With txtInput
            .Top = inputTop: .Left = 10: .Width = 505: .Height = INPUT_HEIGHT
            .Text = ""
            .Font.Name = "Segoe UI"
            .Font.Size = 9.5
        End With
    End If

    If ControlExists("btnSend") Then
        With btnSend
            .Caption = "Send"
            .Top = inputTop: .Left = 520: .Width = 80: .Height = INPUT_HEIGHT
            .Font.Name = "Segoe UI"
            .Font.Bold = True
        End With
    End If

    If ControlExists("lblStatus") Then
        With lblStatus
            .Caption = "Ready. (If nothing responds, make sure chat_driver.py is running.)"
            .Top = inputTop + INPUT_HEIGHT + GAP_SMALL: .Left = 10: .Width = 590: .Height = STATUS_HEIGHT
            .WordWrap = True
            .ForeColor = RGB(0, 100, 0)
            .Font.Name = "Segoe UI"
            .Font.Size = 8.5
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
        Case "SCREENSHOT"
            Call ShowScreenshot(display)   ' display holds the raw file path for this type
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
' view. Setting SelStart alone does NOT reliably scroll an
' unfocused textbox in this MSForms host -- confirmed live
' 2026-08-02 (txtActivity kept showing the oldest content; the
' engineer had to scroll down manually to see new THINKING/
' TOOL_CALL lines as the agent worked). The box must actually
' have focus for the scroll-to-caret to visually take effect;
' focus is returned to txtInput immediately after so typing the
' next message isn't interrupted.
Private Sub AppendTo(box As Object, display As String)
    If box Is Nothing Then Exit Sub
    If box.Text = "" Then
        box.Text = display
    Else
        box.Text = box.Text & vbCrLf & vbCrLf & display
    End If

    On Error Resume Next
    box.SetFocus
    box.SelStart = Len(box.Text)
    If ControlExists("txtInput") Then txtInput.SetFocus
    On Error GoTo 0
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
        Case "SCREENSHOT"
            ' Raw path, not a "[icon] text" display string -- AppendChatLine
            ' routes this straight to ShowScreenshot instead of a textbox.
            FormatLogLine = FieldOrBlank(fields, "path")
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

' ============================================================
' DISPLAY A SCREENSHOT THE AGENT JUST TOOK (chat_driver.py's
' _auto_focus_and_capture, once per completed turn). LoadPicture
' reads the PNG from Bridge/captures/; wrapped in On Error Resume
' Next since a screenshot file that's still mid-write when this
' fires shouldn't crash the panel -- worst case this one doesn't
' render and the next turn's does.
' ============================================================
Private Sub ShowScreenshot(imgPath As String)
    If Trim(imgPath) = "" Then Exit Sub
    If Not ControlExists("imgScreenshot") Then Exit Sub
    On Error Resume Next
    Me.Controls("imgScreenshot").Picture = LoadPicture(imgPath)
    On Error GoTo 0
End Sub

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
