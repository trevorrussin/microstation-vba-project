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
'   btnChoice1..4    CommandButton (added 2026-08-02): ask_user_choice's
'                    option buttons. No Designer properties needed beyond
'                    existing -- Top/Left/Width/Height/Font/Visible are all
'                    set in code (UserForm_Initialize + ShowChoiceButtons),
'                    same as imgScreenshot above. Just add with these exact
'                    names, nothing else to configure.
'   btnPickPoint     CommandButton (added 2026-08-02): the "click a point
'                    in the drawing" option for ask_user_choice. Same deal
'                    -- add with this name, code sets everything else.
'   lblConversationHeader, lblImageHeader, lblActivityHeader,
'   lblInputHeader   Label controls (added 2026-08-02, feedback that the
'                    four panes needed labels): small bold captions above
'                    txtConversation / imgScreenshot / txtActivity /
'                    txtInput respectively. Add with these exact names,
'                    nothing else to configure -- Caption/Top/Left/Width/
'                    Height/Font are all set in code (UserForm_Initialize).
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

    ' Must be set before Left/Top below -- StartUpPosition otherwise
    ' defaults to 2 (CenterScreen), which silently re-centers the form on
    ' every .Show call (including UserForm_Activate's own Hide+Show
    ' vbModeless below), undoing the dock-right positioning entirely. This
    ' was the actual bug behind "the panel appears centered instead of
    ' docked" -- the docking code below was already correct and already
    ' running, just always overridden immediately after.
    Me.StartUpPosition = 0   ' 0 = Manual -- respect the Left/Top set below

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
    ' fixed-position layout). Each of the 4 sections now has a small header
    ' Label above it (2026-08-02 feedback) -- LABEL_HEIGHT+LABEL_GAP is
    ' folded into every section's own reserved footprint below, rather than
    ' bolted on separately, so this stays one consistent formula.
    Const ACTIVITY_HEIGHT As Double = 140
    Const INPUT_HEIGHT As Double = 24
    Const STATUS_HEIGHT As Double = 34
    Const GAP As Double = 8
    Const GAP_SMALL As Double = 6
    Const BOTTOM_MARGIN As Double = 8
    Const LABEL_HEIGHT As Double = 14
    Const LABEL_GAP As Double = 2

    Dim bottomBlockTop As Double
    bottomBlockTop = Me.Height - ((LABEL_HEIGHT + LABEL_GAP + ACTIVITY_HEIGHT) + GAP + _
        (LABEL_HEIGHT + LABEL_GAP + INPUT_HEIGHT) + GAP_SMALL + STATUS_HEIGHT + BOTTOM_MARGIN)

    Dim imgHeaderTop As Double
    imgHeaderTop = 10 + LABEL_HEIGHT + LABEL_GAP + 170 + GAP   ' below txtConversation's header+box
    Dim imgTop As Double
    imgTop = imgHeaderTop + LABEL_HEIGHT + LABEL_GAP

    Dim imgHeight As Double
    imgHeight = bottomBlockTop - GAP - imgTop
    If imgHeight < 100 Then imgHeight = 100   ' safety floor on an unexpectedly small screen

    If ControlExists("lblConversationHeader") Then
        With Me.Controls("lblConversationHeader")
            .Caption = "Conversation -- your messages and the agent's answers"
            .Top = 10: .Left = 10: .Width = 590: .Height = LABEL_HEIGHT
            .Font.Name = "Segoe UI": .Font.Size = 8: .Font.Bold = True
            .ForeColor = RGB(110, 110, 115)
        End With
    End If

    If ControlExists("txtConversation") Then
        With txtConversation
            .Top = 10 + LABEL_HEIGHT + LABEL_GAP: .Left = 10: .Width = 590: .Height = 170
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

    If ControlExists("lblImageHeader") Then
        With Me.Controls("lblImageHeader")
            .Caption = "Screenshot / Reference -- what the agent just drew or looked up"
            .Top = imgHeaderTop: .Left = 10: .Width = 590: .Height = LABEL_HEIGHT
            .Font.Name = "Segoe UI": .Font.Size = 8: .Font.Bold = True
            .ForeColor = RGB(110, 110, 115)
        End With
    End If

    If ControlExists("imgScreenshot") Then
        With Me.Controls("imgScreenshot")
            .Top = imgTop: .Left = 10: .Width = 590: .Height = imgHeight
            .BackColor = RGB(230, 234, 240)
            .BorderStyle = 1   ' fmBorderStyleSingle
            .PictureSizeMode = 3   ' fmPictureSizeModeZoom -- fit without distortion
        End With
    End If

    If ControlExists("lblActivityHeader") Then
        With Me.Controls("lblActivityHeader")
            .Caption = "Agent Activity -- reasoning and tool calls as it works"
            .Top = bottomBlockTop: .Left = 10: .Width = 590: .Height = LABEL_HEIGHT
            .Font.Name = "Segoe UI": .Font.Size = 8: .Font.Bold = True
            .ForeColor = RGB(110, 110, 115)
        End With
    End If

    Dim activityTop As Double
    activityTop = bottomBlockTop + LABEL_HEIGHT + LABEL_GAP

    If ControlExists("txtActivity") Then
        With txtActivity
            .Top = activityTop: .Left = 10: .Width = 590: .Height = ACTIVITY_HEIGHT
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

    ' ask_user_choice option buttons (added 2026-08-02, repositioned same
    ' day per feedback) -- overlay imgScreenshot's rectangle, vertically
    ' CENTERED within it, rather than covering txtActivity: a pending
    ' question shouldn't hide the agent's live reasoning trace, and there's
    ' nothing lost by temporarily replacing whatever screenshot/reference
    ' image is showing, which ShowChoiceButtons also hides outright (not
    ' just visually covered) so it truly reads as "replaced," not layered
    ' on top. Visible=False until ShowChoiceButtons turns specific ones on.
    ' Late-bound Me.Controls(...) throughout -- see that sub's header
    ' comment for why.
    Const CHOICE_ROW_HEIGHT As Double = 24
    Const CHOICE_ROW_GAP As Double = 2
    Const CHOICE_BLOCK_ROWS As Double = 5   ' btnChoice1-4 + btnPickPoint
    Dim choiceBlockHeight As Double
    choiceBlockHeight = CHOICE_BLOCK_ROWS * CHOICE_ROW_HEIGHT + (CHOICE_BLOCK_ROWS - 1) * CHOICE_ROW_GAP
    Dim choiceBlockTop As Double
    choiceBlockTop = imgTop + (imgHeight - choiceBlockHeight) / 2
    If choiceBlockTop < imgTop Then choiceBlockTop = imgTop   ' guard a very short image box

    Dim choiceNum As Integer
    For choiceNum = 1 To 4
        Dim choiceCtrl As String: choiceCtrl = "btnChoice" & choiceNum
        If ControlExists(choiceCtrl) Then
            With Me.Controls(choiceCtrl)
                .Top = choiceBlockTop + (choiceNum - 1) * (CHOICE_ROW_HEIGHT + CHOICE_ROW_GAP)
                .Left = 10: .Width = 590: .Height = CHOICE_ROW_HEIGHT
                .Visible = False
                .Font.Name = "Segoe UI"
                .Font.Size = 9
            End With
        End If
    Next choiceNum
    If ControlExists("btnPickPoint") Then
        With Me.Controls("btnPickPoint")
            .Caption = "Click a point in the drawing"
            .Top = choiceBlockTop + 4 * (CHOICE_ROW_HEIGHT + CHOICE_ROW_GAP)
            .Left = 10: .Width = 590: .Height = CHOICE_ROW_HEIGHT
            .Visible = False
            .Font.Name = "Segoe UI"
            .Font.Bold = True
        End With
    End If

    Dim inputHeaderTop As Double
    inputHeaderTop = activityTop + ACTIVITY_HEIGHT + GAP

    If ControlExists("lblInputHeader") Then
        With Me.Controls("lblInputHeader")
            .Caption = "Your Message"
            .Top = inputHeaderTop: .Left = 10: .Width = 590: .Height = LABEL_HEIGHT
            .Font.Name = "Segoe UI": .Font.Size = 8: .Font.Bold = True
            .ForeColor = RGB(110, 110, 115)
        End With
    End If

    Dim inputTop As Double
    inputTop = inputHeaderTop + LABEL_HEIGHT + LABEL_GAP

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
    Dim imgPath As String
    Dim fields As Object
    display = FormatLogLine(rawLine, lineType, imgPath, fields)

    Select Case lineType
        Case "THINKING", "TOOL_CALL", "TOOL_RESULT"
            If ControlExists("txtActivity") Then AppendTo txtActivity, display
        Case "SCREENSHOT"
            Call ShowScreenshot(display)   ' display holds the raw file path for this type
        Case "REFERENCE_IMAGE"
            ' Unlike SCREENSHOT, this one shows in BOTH panes: a citation
            ' line in the activity trace (so there's a readable record even
            ' without looking at the image) and the actual manual/sheet
            ' page in imgScreenshot -- display is the citation text here,
            ' imgPath carries the raw file path separately.
            If ControlExists("txtActivity") Then AppendTo txtActivity, display
            Call ShowScreenshot(imgPath)
        Case "ASK_USER_CHOICE"
            If ControlExists("txtConversation") Then AppendTo txtConversation, display
            Call ShowChoiceButtons(fields)
        Case Else
            If ControlExists("txtConversation") Then AppendTo txtConversation, display
    End Select

    If ControlExists("lblStatus") Then
        Select Case lineType
            Case "ASK_USER", "ASK_USER_CHOICE"
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
Private Function FormatLogLine(rawLine As String, ByRef outLineType As String, ByRef outImagePath As String, ByRef outFields As Object) As String
    outLineType = ""
    outImagePath = ""
    Set outFields = Nothing
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
    Set outFields = fields

    Select Case lineType
        Case "SCREENSHOT"
            ' Raw path, not a "[icon] text" display string -- AppendChatLine
            ' routes this straight to ShowScreenshot instead of a textbox.
            FormatLogLine = FieldOrBlank(fields, "path")
        Case "REFERENCE_IMAGE"
            outImagePath = FieldOrBlank(fields, "path")
            FormatLogLine = "[reference] " & FieldOrBlank(fields, "source") & " -- " & _
                             FieldOrBlank(fields, "heading") & " (page " & FieldOrBlank(fields, "page") & ")"
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
        Case "ASK_USER_CHOICE"
            Dim choiceText As String
            choiceText = "[agent asks] " & FieldOrBlank(fields, "question")
            Dim optNum As Integer
            For optNum = 1 To 4
                Dim optLabel As String: optLabel = FieldOrBlank(fields, "option" & optNum & "Label")
                If optLabel <> "" Then
                    Dim optDetail As String: optDetail = FieldOrBlank(fields, "option" & optNum & "Detail")
                    choiceText = choiceText & vbCrLf & "  " & optNum & ". " & optLabel
                    If optDetail <> "" Then choiceText = choiceText & " -- " & optDetail
                End If
            Next optNum
            If FieldOrBlank(fields, "allowPointPick") = "Y" Then
                choiceText = choiceText & vbCrLf & "  (or click 'Pick Point' to click a location in the drawing)"
            End If
            choiceText = choiceText & vbCrLf & "  (or just type your own reply)"
            FormatLogLine = choiceText
        Case "FINAL"
            FormatLogLine = "[agent] " & FieldOrBlank(fields, "text")
        Case "MODE_CHANGED"
            FormatLogLine = "-- Switched to " & FieldOrBlank(fields, "mode") & " mode -- " & FieldOrBlank(fields, "description")
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

' ============================================================
' SHOW/HIDE THE ask_user_choice OPTION BUTTONS (added 2026-08-02,
' repositioned same day per feedback). btnChoice1..btnChoice4 +
' btnPickPoint overlay imgScreenshot's rectangle, centered within
' it (Top math lives in UserForm_Initialize) -- a pending question
' shouldn't cover the activity/reasoning trace, and there's
' nothing lost by replacing whatever screenshot/reference image is
' currently showing. ShowChoiceButtons hides imgScreenshot itself
' (not just visually covering it) so it reads as truly replaced;
' HideChoiceButtons restores it -- the underlying .Picture is
' untouched throughout, so whatever was showing before reappears
' as-is once the choice is resolved.
'
' Late-bound Me.Controls("...") throughout, not bare identifiers
' -- these controls must be added manually in the VBA IDE Designer
' (File Sync Protocol, CLAUDE.md; same one-time step already done
' for imgScreenshot) and referencing a not-yet-existing control by
' bare name fails to COMPILE, not just fails at runtime, even
' inside a runtime ControlExists guard (hit and fixed once already
' this session for imgScreenshot -- see Claude Code memory
' project_chat_panel_ux_improvements.md).
' ============================================================
Private Sub ShowChoiceButtons(fields As Object)
    If fields Is Nothing Then Exit Sub
    If ControlExists("imgScreenshot") Then Me.Controls("imgScreenshot").Visible = False
    Dim i As Integer
    For i = 1 To 4
        Dim ctrlName As String: ctrlName = "btnChoice" & i
        If ControlExists(ctrlName) Then
            Dim lbl As String: lbl = FieldOrBlank(fields, "option" & i & "Label")
            If lbl <> "" Then
                Me.Controls(ctrlName).Caption = lbl
                Me.Controls(ctrlName).Visible = True
            Else
                Me.Controls(ctrlName).Visible = False
            End If
        End If
    Next i
    If ControlExists("btnPickPoint") Then
        Me.Controls("btnPickPoint").Visible = (FieldOrBlank(fields, "allowPointPick") = "Y")
    End If
End Sub

Private Sub HideChoiceButtons()
    Dim i As Integer
    For i = 1 To 4
        If ControlExists("btnChoice" & i) Then Me.Controls("btnChoice" & i).Visible = False
    Next i
    If ControlExists("btnPickPoint") Then Me.Controls("btnPickPoint").Visible = False
    If ControlExists("imgScreenshot") Then Me.Controls("imgScreenshot").Visible = True
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
    If Not ControlExists("txtInput") Then Exit Sub

    Dim msg As String
    msg = Trim(txtInput.Text)
    If msg = "" Then Exit Sub

    Call SendTextAsReply(msg)
    txtInput.Text = ""
End Sub

' ============================================================
' SEND A REPLY -- shared by btnSend, the ask_user_choice option
' buttons, and the pick-point button (added 2026-08-02). All
' three ways of answering a pending question converge here:
' append to Bridge/chat-input.tsv (same file/format chat_driver.
' py's InputWatcher already polls for both the main turn loop
' and ask_user/ask_user_choice), echo it into the transcript, and
' clear any pending choice buttons, since any one of these three
' paths resolves the same pending question.
' ============================================================
Private Sub SendTextAsReply(msg As String)
    On Error GoTo SendErr
    If Trim(msg) = "" Then Exit Sub

    Dim fnum As Integer
    fnum = FreeFile
    Open CHAT_INPUT_FILE For Append As #fnum
    Print #fnum, Now & vbTab & msg
    Close #fnum

    AppendChatLine Now & vbTab & "USER_ECHO" & vbTab & "text=" & msg
    Call HideChoiceButtons

    If ControlExists("lblStatus") Then
        lblStatus.Caption = "Waiting for the agent..."
        lblStatus.ForeColor = RGB(0, 0, 180)
    End If
    Exit Sub

SendErr:
    If fnum <> 0 Then Close #fnum
    If ControlExists("lblStatus") Then lblStatus.Caption = "Send failed: " & Err.Description
End Sub

' ============================================================
' ASK_USER_CHOICE OPTION BUTTONS -- each just sends its own
' caption (the option's label) as the reply, exactly like typing
' that label and hitting Send. Late-bound Me.Controls(...), see
' the ShowChoiceButtons header comment for why.
' ============================================================
Private Sub btnChoice1_Click()
    If ControlExists("btnChoice1") Then SendTextAsReply Me.Controls("btnChoice1").Caption
End Sub

Private Sub btnChoice2_Click()
    If ControlExists("btnChoice2") Then SendTextAsReply Me.Controls("btnChoice2").Caption
End Sub

Private Sub btnChoice3_Click()
    If ControlExists("btnChoice3") Then SendTextAsReply Me.Controls("btnChoice3").Caption
End Sub

Private Sub btnChoice4_Click()
    If ControlExists("btnChoice4") Then SendTextAsReply Me.Controls("btnChoice4").Caption
End Sub

' ============================================================
' PICK A POINT IN THE DRAWING -- same GetInput loop pattern
' Modules/DrawSign.bas already uses for sign placement (~line
' 152-166), run from a native button-Click event so it only
' blocks this form's own VBA thread (confirmed safe in
' production for the whole 8-step workflow), never the chat
' bridge/Python side -- ask_user_choice's Python-side wait is a
' plain file poll (INPUT.wait_for_next()), completely separate
' from this click-capture. Deliberately NOT routed through
' WZTCBridge.ExecuteOp: that path is triggered by chat_driver.py
' via a blocking, timeout-less SendKeyin COM call, and a GetInput
' wait inside it would hang the entire chat_driver.py process
' with no way to cancel if the engineer doesn't click right away
' -- confirmed via research as the same failure mode that already
' hung a different feature once (Modules/WZTCViewCapture.bas).
' ============================================================
Private Sub btnPickPoint_Click()
    On Error GoTo PickErr
    Dim oMsg As CadInputMessage
    CadInputQueue.SendKeyin "ECHO Click a point in the drawing"
    CadInputQueue.SendCommand "NULL"
    Set oMsg = CadInputQueue.GetInput
    Do While oMsg.InputType <> msdCadInputTypeDataPoint
        If oMsg.InputType = msdCadInputTypeReset Then
            If ControlExists("lblStatus") Then
                lblStatus.Caption = "Point pick cancelled -- pick an option or type your reply."
            End If
            Exit Sub
        End If
        Set oMsg = CadInputQueue.GetInput
    Loop

    Dim coordText As String
    coordText = "(" & Format(oMsg.Point.X, "0.00") & ", " & Format(oMsg.Point.Y, "0.00") & ", " & Format(oMsg.Point.Z, "0.00") & ")"
    SendTextAsReply coordText
    Exit Sub

PickErr:
    If ControlExists("lblStatus") Then lblStatus.Caption = "Point pick failed: " & Err.Description
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
