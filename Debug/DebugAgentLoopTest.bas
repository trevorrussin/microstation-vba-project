Option Explicit

' ============================================================
' WZTCBridge M5 AGENT-LOOP SMOKE TEST
' ------------------------------------------------------------
' Exercises HANDOFF, GET_JOURNAL, LIST_DEFERRED_HANDOFFS, and
' UNDO_LAST_OP through the same ExecuteOp path Python will use,
' same convention as DebugQueryTest.bas / DebugExecTest.bas.
'
' UNDO_LAST_OP is tested against a PLACE_CELL op (the most-proven
' primitive, from M1) rather than a new placement, specifically so
' this test also confirms PLACE_CELL's own elementId= field is
' recognized by the undo scan even though PLACE_CELL doesn't emit
' createdElementIds= (only the M5-era ops do). Run from the VBA
' IDE: cursor in TestAgentLoop, F5, watch Ctrl+G.
' ============================================================

Private failCount As Integer
Private checkCount As Integer

Public Sub TestAgentLoop()
    failCount = 0
    checkCount = 0
    Debug.Print "=== WZTCBridge M5 agent-loop smoke test ==="

    Call TestHandoff
    Call TestUndoLastOp
    Call TestGetJournal
    Call TestListDeferredHandoffs

    Debug.Print "assertions: " & checkCount
    If failCount = 0 Then
        Debug.Print "=== PASS - no failures ==="
    Else
        Debug.Print "=== FAIL - " & failCount & " failed assertion(s) ==="
    End If
End Sub

Private Sub TestHandoff()
    Dim resp As String
    resp = WZTCBridge.ExecuteOp("A1" & vbTab & "HANDOFF" & vbTab & "kind=dimension" & vbTab & _
                                "fromSta=100" & vbTab & "toSta=250" & vbTab & _
                                "reason=advance warning run length")
    Debug.Print "HANDOFF (dimension) -> " & resp
    Call Assert(InStr(resp, vbTab & "DEFERRED" & vbTab) > 0, "HANDOFF did not return DEFERRED: " & resp)

    Dim respMissing As String
    respMissing = WZTCBridge.ExecuteOp("A2" & vbTab & "HANDOFF")
    Debug.Print "HANDOFF (missing kind) -> " & respMissing
    Call Assert(InStr(respMissing, vbTab & "ERROR" & vbTab) > 0, _
                "HANDOFF with no kind should error, not: " & respMissing)
End Sub

' Places a cell (proven M1 primitive), confirms UNDO_LAST_OP removes THAT
' SPECIFIC element, then confirms a second undo attempt does NOT re-target
' the same op (idempotent).
'
' Checks the specific elementId rather than a raw rowCount/radius search at
' (5000,5000) — that fixture point can pick up harmless leftover cells from
' earlier manual re-runs of this same test (e.g. a run before this test
' file's journaling-bug fix, whose PLACE_CELL was never journaled and so
' could never be auto-undone). This test only asserts on the one element
' IT created, so it stays correct regardless of any such debris.
Private Sub TestUndoLastOp()
    Dim placeResp As String
    placeResp = WZTCBridge.ExecuteOp("A3" & vbTab & "PLACE_CELL" & vbTab & _
                                     "cellName=TWZAP_P" & vbTab & "ptX=5000" & vbTab & "ptY=5000")
    Debug.Print "PLACE_CELL (for undo test) -> " & placeResp
    Call Assert(InStr(placeResp, vbTab & "OK" & vbTab) > 0, "setup PLACE_CELL failed: " & placeResp)
    Dim placedID As String: placedID = ExtractStrParam(placeResp, "elementId")

    Dim findBefore As String
    findBefore = WZTCBridge.ExecuteOp("A4" & vbTab & "FIND_ELEMENTS_NEAR" & vbTab & _
                                      "x=5000" & vbTab & "y=5000" & vbTab & "radius=5")
    Call Assert(ResultFileContainsValue(ExtractStrParam(findBefore, "resultFile"), placedID), _
                "expected the just-placed cell (elementId=" & placedID & ") to be findable near (5000,5000): " & findBefore)

    Dim undoResp As String
    undoResp = WZTCBridge.ExecuteOp("A5" & vbTab & "UNDO_LAST_OP")
    Debug.Print "UNDO_LAST_OP -> " & undoResp
    Call Assert(InStr(undoResp, vbTab & "OK" & vbTab) > 0, "UNDO_LAST_OP did not return OK: " & undoResp)
    Call Assert(ExtractStrParam(undoResp, "undidReqId") = "A3", _
                "UNDO_LAST_OP undid the wrong reqId (expected A3): " & undoResp)

    Dim findAfter As String
    findAfter = WZTCBridge.ExecuteOp("A6" & vbTab & "FIND_ELEMENTS_NEAR" & vbTab & _
                                     "x=5000" & vbTab & "y=5000" & vbTab & "radius=5")
    Debug.Print "FIND_ELEMENTS_NEAR after undo -> " & findAfter
    Dim stillPresent As Boolean
    stillPresent = ResultFileContainsValue(ExtractStrParam(findAfter, "resultFile"), placedID)
    Call Assert(Not stillPresent, "the undone cell (elementId=" & placedID & _
                ") should no longer appear in FIND_ELEMENTS_NEAR results: " & findAfter)

    ' A second undo must NOT re-target A3 (already marked UNDONE) — it should
    ' either find an earlier op or gracefully report nothing left to undo.
    Dim undoResp2 As String
    undoResp2 = WZTCBridge.ExecuteOp("A7" & vbTab & "UNDO_LAST_OP")
    Debug.Print "UNDO_LAST_OP (again) -> " & undoResp2
    Call Assert(ExtractStrParam(undoResp2, "undidReqId") <> "A3", _
                "second UNDO_LAST_OP re-targeted an already-undone reqId: " & undoResp2)
End Sub

Private Sub TestGetJournal()
    Dim resp As String
    resp = WZTCBridge.ExecuteOp("A8" & vbTab & "GET_JOURNAL" & vbTab & "limit=10")
    Debug.Print "GET_JOURNAL -> " & resp
    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0, "GET_JOURNAL did not return OK: " & resp)
    Call Assert(ExtractIntParam(resp, "rowCount") >= 1, "GET_JOURNAL returned no rows: " & resp)
End Sub

Private Sub TestListDeferredHandoffs()
    Dim resp As String
    resp = WZTCBridge.ExecuteOp("A9" & vbTab & "LIST_DEFERRED_HANDOFFS")
    Debug.Print "LIST_DEFERRED_HANDOFFS -> " & resp
    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0, "LIST_DEFERRED_HANDOFFS did not return OK: " & resp)
    Call Assert(ExtractIntParam(resp, "rowCount") >= 1, _
                "expected at least the A1 dimension handoff queued earlier in this run: " & resp)
End Sub

' ============================================================
' PARSING HELPERS — identical to DebugQueryTest.bas
' ============================================================
Private Function ExtractStrParam(resp As String, key As String) As String
    Dim parts() As String: parts = Split(resp, vbTab)
    Dim i As Integer
    For i = 0 To UBound(parts)
        If Left(parts(i), Len(key) + 1) = key & "=" Then
            ExtractStrParam = Mid(parts(i), Len(key) + 2)
            Exit Function
        End If
    Next i
    ExtractStrParam = ""
End Function

Private Function ExtractIntParam(resp As String, key As String) As Long
    ExtractIntParam = CLng(Val(ExtractStrParam(resp, key)))
End Function

Private Function ResultFileContainsValue(path As String, val As String) As Boolean
    If path = "" Or Dir(path) = "" Then
        ResultFileContainsValue = False
        Exit Function
    End If
    Dim fnum As Integer: fnum = FreeFile
    Open path For Input As #fnum
    Dim found As Boolean: found = False
    Dim ln As String
    Do While Not EOF(fnum)
        Line Input #fnum, ln
        If InStr(ln, val) > 0 Then
            found = True
            Exit Do
        End If
    Loop
    Close #fnum
    ResultFileContainsValue = found
End Function

Private Sub Assert(cond As Boolean, msg As String)
    checkCount = checkCount + 1
    If Not cond Then
        failCount = failCount + 1
        Debug.Print "FAIL: " & msg
    End If
End Sub
