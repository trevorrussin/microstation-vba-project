Option Explicit

' ============================================================
' WZTC CHAT TIMER (M7 Stage 1)
' ------------------------------------------------------------
' Win32 SetTimer/KillTimer polling loop for WZTCChatPanel.frm.
' No Timer control exists in this MicroStation VBA host (unlike
' Access/Excel), and nothing resembling a polling loop existed
' anywhere in this repo before M7 -- this is the first one.
' Confirmed working live (2026-08-01): the SetTimer/AddressOf
' mechanism fires reliably every intervalMs with no drift or
' missed ticks across a 700+ tick session.
'
' hWnd=0 creates a timer with no associated window; Win32's
' message pump calls ChatTimerProc directly out of the existing
' VBA idle loop -- no blocking Do/Loop of our own, so MicroStation
' and every other open modeless form (WZTCDesigner, PlacePerp,
' etc.) stay fully responsive between ticks.
'
' AddressOf only works on a Public Sub in a standard module (not
' a class or form), so the callback lives here and reaches the
' open panel via a module-level object reference -- same
' reach-back-into-host-form indirection this repo already uses
' for WithEvents handlers (Class Modules/SignNumBox.cls is the
' live example; Class Modules/PlaceButtons.cls is confirmed dead
' code despite being documented -- do not copy that one).
'
' Watched files MUST be CRLF, not bare LF -- same requirement as
' Data/sheet-registry.tsv and Data/command-registry.tsv (see
' Data/README.md). A bare-LF file reads as a single giant line
' under Line Input#, which silently breaks the line-count-based
' new-content detection below. Confirmed by testing, not
' theoretical -- this exact failure mode cost significant bring-up
' time on 2026-08-01.
'
' Re-importing this module over an existing copy of the same name
' does not reliably replace it in this VBA IDE -- Remove the old
' module first, then File -> Import File. Confirmed by testing.
'
' Stage 1: watches a throwaway test file, no LLM/bridge involved.
' Stage 5: WZTCChatPanel points this at the real Bridge/chat-log.tsv
' -- this module does not change.
' ============================================================

#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, _
        ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal uIDEvent As LongPtr) As Long
#Else
    Private Declare Function SetTimer Lib "user32" ( _
        ByVal hwnd As Long, ByVal nIDEvent As Long, _
        ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" ( _
        ByVal hwnd As Long, ByVal uIDEvent As Long) As Long
#End If

#If VBA7 Then
    Private mTimerID As LongPtr
#Else
    Private mTimerID As Long
#End If

Private mPanel As Object           ' the open WZTCChatPanel instance
Private mWatchFile As String       ' file being polled for new lines
Private mLastLineCount As Long     ' lines already delivered to the panel

' ============================================================
' START -- begin polling watchFile every intervalMs, delivering
' each new line (beyond whatever's in the file at start time) to
' panel.AppendChatLine. Safe to call again after StopChatTimer.
' ============================================================
Public Sub StartChatTimer(panel As Object, watchFile As String, _
                          Optional intervalMs As Long = 300)
    Call StopChatTimer   ' guard against double-registration on reopen

    Set mPanel = panel
    mWatchFile = watchFile
    mLastLineCount = CountLines(mWatchFile)   ' don't replay pre-existing lines

    mTimerID = SetTimer(0, 0, intervalMs, AddressOf ChatTimerProc)
    If mTimerID = 0 Then
        If Not mPanel Is Nothing Then mPanel.AppendChatLine "[WZTCChatTimer] SetTimer failed to register"
        Set mPanel = Nothing
    End If
End Sub

' ============================================================
' STOP -- call from the panel's QueryClose/Terminate. Guarded
' against being called when no timer is running (mTimerID = 0),
' and against a stale mTimerID firing into a panel reference
' that's already gone.
' ============================================================
Public Sub StopChatTimer()
    If mTimerID <> 0 Then
        KillTimer 0, mTimerID
        mTimerID = 0
    End If
    Set mPanel = Nothing
    mWatchFile = ""
    mLastLineCount = 0
End Sub

' ============================================================
' TIMER CALLBACK -- must be Public Sub in a standard module.
' Re-reads mWatchFile each tick (same whole-file-reread pattern
' WZTCCommandRegistry.ReadAllLines / the journal readers already
' use elsewhere in this codebase -- file sizes here are small
' enough that this isn't a performance concern) and delivers only
' lines beyond mLastLineCount, so a human editing the file
' externally (or Python appending to it) shows up incrementally.
' ============================================================
#If VBA7 Then
Public Sub ChatTimerProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, _
                         ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Public Sub ChatTimerProc(ByVal hwnd As Long, ByVal uMsg As Long, _
                         ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    On Error GoTo ProcErr

    If mPanel Is Nothing Then Exit Sub
    If mWatchFile = "" Then Exit Sub

    Dim lines() As String
    Dim n As Long
    n = ReadAllLines(mWatchFile, lines)

    If n > mLastLineCount Then
        Dim i As Long
        For i = mLastLineCount + 1 To n
            mPanel.AppendChatLine lines(i)
        Next i
        mLastLineCount = n
    End If
    Exit Sub

ProcErr:
    ' Never let an error inside the timer callback propagate --
    ' that would be an error inside MicroStation's own message
    ' pump. Surface it in the transcript instead and keep polling.
    On Error Resume Next
    If Not mPanel Is Nothing Then mPanel.AppendChatLine "[WZTCChatTimer] " & Err.Description
End Sub

' ============================================================
' FILE I/O -- same pattern as WZTCSheetRegistry.ReadAllLines /
' WZTCCommandRegistry.ReadAllLines. Blank lines are skipped, same
' as those, so an empty line in the chat log can't desync the
' line-count bookkeeping against what's actually rendered.
' ============================================================
Private Function ReadAllLines(path As String, ByRef outLines() As String) As Long
    On Error GoTo ReadErr
    Dim fnum As Integer: fnum = 0
    If Dir(path) = "" Then
        ReadAllLines = 0
        Exit Function
    End If

    fnum = FreeFile
    Open path For Input As #fnum

    Dim n As Long: n = 0
    Dim ln As String
    Do While Not EOF(fnum)
        Line Input #fnum, ln
        If Len(Trim(ln)) > 0 Then
            n = n + 1
            ReDim Preserve outLines(1 To n)
            outLines(n) = ln
        End If
    Loop
    Close #fnum
    ReadAllLines = n
    Exit Function

ReadErr:
    If fnum <> 0 Then Close #fnum
    ReadAllLines = 0
End Function

Private Function CountLines(path As String) As Long
    Dim lines() As String
    CountLines = ReadAllLines(path, lines)
End Function
