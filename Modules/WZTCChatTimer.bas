Option Explicit

' ============================================================
' WZTC CHAT TIMER (M7 Stage 1)
' ------------------------------------------------------------
' Polls Bridge/chat-log.tsv and delivers new lines to the open
' WZTCChatPanel. Originally used Win32 SetTimer + AddressOf.
' That callback form triggers VBA "Unexpected error (35010)" on
' 64-bit VBE compile/run in this MicroStation host (confirmed
' 2026-08-21) -- a known VBE bug around AddressOf, not a logic
' bug in the panel. Replaced with Sleep + DoEvents pump: no
' function pointer, so Compile/Run no longer hit 35010.
'
' StartChatTimer only stores state. RunPollPump (called once
' from WZTCChatPanel's post-modeless Activate) blocks that
' Activate with DoEvents so the modeless form stays clickable
' while lines are delivered every intervalMs. StopChatTimer
' clears mRunning so the pump exits on QueryClose/Terminate.
'
' Watched files MUST be CRLF, not bare LF -- same requirement as
' Data/sheet-registry.tsv (see Data/README.md).
'
' Re-importing this module: Remove the old module first, then
' Import, then rename the component to WZTCChatTimer in the
' VBA project tree (Import without Attribute VB_Name lands as
' ModuleN -- confirmed 2026-08-21).
' ============================================================

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private mPanel As Object           ' the open WZTCChatPanel instance
Private mWatchFile As String       ' file being polled for new lines
Private mLastLineCount As Long     ' lines already delivered to the panel
Private mIntervalMs As Long        ' Sleep between ticks
Private mRunning As Boolean        ' pump guard -- StopChatTimer clears this
Private mPumpActive As Boolean     ' true while RunPollPump's Do-loop is on stack

' ============================================================
' START -- store panel/file/interval. Does NOT start polling;
' call RunPollPump once the form is modeless and Activate can
' afford to block with DoEvents.
' ============================================================
Public Sub StartChatTimer(panel As Object, watchFile As String, _
                          Optional intervalMs As Long = 300)
    Call StopChatTimer   ' clear any prior session state

    Set mPanel = panel
    mWatchFile = watchFile
    mLastLineCount = CountLines(mWatchFile)   ' don't replay pre-existing lines
    If intervalMs < 50 Then intervalMs = 50
    mIntervalMs = intervalMs
    mRunning = True
End Sub

' ============================================================
' PUMP -- blocks the caller (panel Activate) until StopChatTimer.
' DoEvents keeps the modeless form responsive between ticks.
' ============================================================
Public Sub RunPollPump()
    If mPumpActive Then Exit Sub   ' never nest
    If Not mRunning Then Exit Sub
    If mPanel Is Nothing Then Exit Sub

    mPumpActive = True
    On Error GoTo PumpErr

    Do While mRunning
        Call PollOnce
        Sleep mIntervalMs
        DoEvents
    Loop

PumpDone:
    mPumpActive = False
    Exit Sub

PumpErr:
    On Error Resume Next
    If Not mPanel Is Nothing Then mPanel.AppendChatLine "[WZTCChatTimer] " & Err.Description
    Resume PumpDone
End Sub

' ============================================================
' STOP -- call from the panel's QueryClose/Terminate. Clears
' mRunning so RunPollPump exits on the next DoEvents pass.
' ============================================================
Public Sub StopChatTimer()
    mRunning = False
    Set mPanel = Nothing
    mWatchFile = ""
    mLastLineCount = 0
End Sub

' ============================================================
' ONE TICK -- same deliver-new-lines logic as the old SetTimer
' callback. Safe to call from the pump or manually.
' ============================================================
Public Sub PollOnce()
    On Error GoTo ProcErr

    If mPanel Is Nothing Then Exit Sub
    If mWatchFile = "" Then Exit Sub

    Dim lines() As String
    Dim n As Long
    n = ReadAllLines(mWatchFile, lines)

    ' Only treat a SHRINK as rotation when we still got a positive line
    ' count. ReadAllLines returns 0 on any I/O error (file locked mid-
    ' append by chat_driver.py is the common case) -- treating that as
    ' "rotated" used to reset mLastLineCount to 0 and replay the entire
    ' chat-log into the panel (confirmed live 2026-08-02).
    If n > 0 And n < mLastLineCount Then
        On Error Resume Next
        mPanel.ResetTranscriptPanes
        On Error GoTo ProcErr
        mLastLineCount = 0
    End If

    If n > mLastLineCount Then
        Dim i As Long
        For i = mLastLineCount + 1 To n
            mPanel.AppendChatLine lines(i)
        Next i
        mLastLineCount = n
    End If
    Exit Sub

ProcErr:
    On Error Resume Next
    If Not mPanel Is Nothing Then mPanel.AppendChatLine "[WZTCChatTimer] " & Err.Description
End Sub

' ============================================================
' FILE I/O -- ADODB.Stream UTF-8 (chat-log is written UTF-8 by
' chat_driver.py; Open/Line Input# would mojibake em-dashes).
' ============================================================
Private Function ReadAllLines(path As String, ByRef outLines() As String) As Long
    On Error GoTo ReadErr
    If Dir(path) = "" Then
        ReadAllLines = 0
        Exit Function
    End If

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    Dim content As String
    content = stream.ReadText(-1) ' adReadAll
    stream.Close

    Dim rawLines() As String
    rawLines = Split(content, vbCrLf)

    Dim n As Long: n = 0
    Dim i As Long
    Dim ln As String
    For i = LBound(rawLines) To UBound(rawLines)
        ln = rawLines(i)
        If Len(Trim(ln)) > 0 Then
            n = n + 1
            ReDim Preserve outLines(1 To n)
            outLines(n) = ln
        End If
    Next i
    ReadAllLines = n
    Exit Function

ReadErr:
    ReadAllLines = 0
End Function

Private Function CountLines(path As String) As Long
    Dim lines() As String
    CountLines = ReadAllLines(path, lines)
End Function
