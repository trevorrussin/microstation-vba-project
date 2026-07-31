Option Explicit

' ============================================================
' WZTCBridge M1 SMOKE TEST
' ------------------------------------------------------------
' Writes a PLACE_CELL request, calls WZTCBridge.RunRequest
' directly (in-process — no external trigger involved yet),
' and prints the response. This isolates the VBA-side half of
' the bridge (file I/O, param parsing, programmatic cell
' placement) from the external-trigger half (the "VBA RUN"
' keyin, or later, COM automation from Python).
'
' Run from the VBA IDE: place cursor in TestBridgePlaceCell and
' press F5. Watch the Immediate window (Ctrl+G). Then check
' MicroStation — a TWZAP_P (Arrow Panel) cell should appear at
' design coordinates (1000, 1000).
'
' Once this passes, the manual keyin test in Bridge/README.md
' proves the second half (external trigger).
' ============================================================

Public Sub TestBridgePlaceCell()
    Debug.Print "=== WZTCBridge M1 smoke test ==="

    Dim bridgeDir As String
    bridgeDir = "c:\repos\microstation-vba-project\Bridge\"
    Dim reqPath As String, respPath As String
    reqPath = bridgeDir & "request.tsv"
    respPath = bridgeDir & "response.tsv"

    ' Snapshot element count before, so we can confirm exactly one
    ' new graphical element appeared (not more, not fewer).
    Dim countBefore As Long
    countBefore = CountGraphicalElements()

    ' Write a single PLACE_CELL request line
    Dim fnum As Integer: fnum = FreeFile
    Open reqPath For Output As #fnum
    Print #fnum, "0001" & vbTab & "PLACE_CELL" & vbTab & "cellName=TWZAP_P" & vbTab & _
                 "ptX=1000" & vbTab & "ptY=1000" & vbTab & "ptZ=0" & vbTab & "angleDeg=0"
    Close #fnum
    Debug.Print "Wrote request: " & reqPath

    ' Execute in-process (same code path RunRequest uses when triggered
    ' externally — this just skips the keyin/COM trigger step)
    Call WZTCBridge.RunRequest
    Debug.Print "RunRequest executed."

    ' Read back the response
    If Dir(respPath) = "" Then
        Debug.Print "FAIL: response.tsv was not written"
        Exit Sub
    End If

    Dim respLine As String
    fnum = FreeFile
    Open respPath For Input As #fnum
    Line Input #fnum, respLine
    Close #fnum
    Debug.Print "Response: " & respLine

    Dim countAfter As Long
    countAfter = CountGraphicalElements()
    Debug.Print "Elements before: " & countBefore & "  after: " & countAfter

    ' ---- Assertions ----
    Dim failCount As Integer: failCount = 0

    If InStr(respLine, vbTab & "OK" & vbTab) = 0 Then
        Debug.Print "FAIL: response status is not OK"
        failCount = failCount + 1
    End If
    If InStr(respLine, "elementId=") = 0 Then
        Debug.Print "FAIL: response missing elementId"
        failCount = failCount + 1
    End If
    If countAfter <> countBefore + 1 Then
        Debug.Print "FAIL: expected exactly 1 new element, got " & (countAfter - countBefore)
        failCount = failCount + 1
    End If

    Dim journalPath As String: journalPath = bridgeDir & "wztc-journal.tsv"
    If Dir(journalPath) = "" Then
        Debug.Print "FAIL: wztc-journal.tsv was not created"
        failCount = failCount + 1
    End If

    If failCount = 0 Then
        Debug.Print "=== PASS - VBA-side bridge works. Check MicroStation for the placed cell. ==="
        Debug.Print "=== Next: try the manual keyin trigger from Bridge/README.md to test the external half. ==="
    Else
        Debug.Print "=== FAIL - " & failCount & " assertion(s) failed ==="
    End If
End Sub

Private Function CountGraphicalElements() As Long
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical

    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    Dim n As Long: n = 0
    Do While oEnum.MoveNext
        n = n + 1
    Loop
    CountGraphicalElements = n
End Function
