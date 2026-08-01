Option Explicit

' ============================================================
' WZTCBridge M6 COMMAND-REGISTRY SMOKE TEST
' ------------------------------------------------------------
' Focuses on the gate-refusal cases this layer exists to prevent:
' needs-testing and interactive-only-use-handoff must ERROR with
' a clear note (not silently no-op), and ownElementOnly must refuse
' an elementId that isn't in the journal.
'
' Live mutation/undo of MOVE/CHANGE_LEVEL/EDIT_TEXT still needs a
' design file with a freshly-placed element — those are covered as
' optional live checks at the end (skip cleanly if PLACE_CELL fails).
'
' Run from the VBA IDE: cursor in TestCommandRegistry, F5, watch Ctrl+G.
' Requires WZTCCommandRegistry.bas imported alongside the updated
' WZTCBridge.bas / WZTCExec.bas.
' ============================================================

Private failCount As Integer
Private checkCount As Integer

Public Sub TestCommandRegistry()
    failCount = 0
    checkCount = 0
    Debug.Print "=== WZTCBridge M6 command-registry smoke test ==="

    Call TestListAndDescribe
    Call TestGateRefusesNeedsTesting
    Call TestGateRefusesInteractiveHandoff
    Call TestRunVerifiedSettingsKeyin
    Call TestOwnElementOnlyRefusal
    Call TestDeleteNotUndoable
    Call TestLiveMoveAndUndo
    Call TestLiveChangeLevelAndUndo

    Debug.Print "assertions: " & checkCount
    If failCount = 0 Then
        Debug.Print "=== PASS - no failures ==="
    Else
        Debug.Print "=== FAIL - " & failCount & " failed assertion(s) ==="
    End If
End Sub

Private Sub TestListAndDescribe()
    Dim resp As String
    resp = WZTCBridge.ExecuteOp("R1" & vbTab & "LIST_REGISTRY_COMMANDS")
    Debug.Print "LIST_REGISTRY_COMMANDS -> " & resp
    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0, "LIST_REGISTRY_COMMANDS did not return OK: " & resp)
    Call Assert(ExtractIntParam(resp, "rowCount") >= 10, _
                "expected >= 10 seed rows, got rowCount=" & ExtractIntParam(resp, "rowCount"))

    Dim desc As String
    desc = WZTCBridge.ExecuteOp("R2" & vbTab & "DESCRIBE_REGISTRY_COMMAND" & vbTab & "opName=MOVE_ELEMENT")
    Debug.Print "DESCRIBE_REGISTRY_COMMAND (MOVE_ELEMENT) -> " & desc
    Call Assert(InStr(desc, vbTab & "OK" & vbTab) > 0, "DESCRIBE MOVE_ELEMENT failed: " & desc)
    Call Assert(InStr(desc, "safetyStatus=verified-headless-safe") > 0, _
                "MOVE_ELEMENT should be verified-headless-safe: " & desc)
    Call Assert(InStr(desc, "category=direct_api") > 0, _
                "MOVE_ELEMENT should be direct_api: " & desc)
End Sub

Private Sub TestGateRefusesNeedsTesting()
    ' The registry is actively promoted by the harvest pipeline (needs-testing
    ' -> verified-headless-safe), so a specific opName hardcoded here as "the"
    ' needs-testing example will eventually get promoted out from under this
    ' test. Ask the registry which row currently qualifies instead.
    Dim rows() As String
    rows = WZTCCommandRegistry.ListCommands("needs-testing")
    If UBound(rows) < 1 Then
        Debug.Print "SKIP TestGateRefusesNeedsTesting — no needs-testing rows currently in registry"
        Exit Sub
    End If
    Dim cols() As String: cols = Split(rows(1), vbTab)
    Dim opName As String: opName = cols(0)

    Dim resp As String
    resp = WZTCBridge.ExecuteOp("R3" & vbTab & "RUN_REGISTRY_COMMAND" & vbTab & _
                                "opName=" & opName)
    Debug.Print "RUN_REGISTRY_COMMAND (needs-testing: " & opName & ") -> " & resp
    Call Assert(InStr(resp, vbTab & "ERROR" & vbTab) > 0, _
                "needs-testing row must ERROR, not: " & resp)
    Call Assert(InStr(resp, "needs-testing") > 0, _
                "refusal note should mention needs-testing: " & resp)
End Sub

Private Sub TestGateRefusesInteractiveHandoff()
    Dim resp As String
    resp = WZTCBridge.ExecuteOp("R4" & vbTab & "RUN_REGISTRY_COMMAND" & vbTab & _
                                "opName=DIMENSION_SIZE_WITH_LINES")
    Debug.Print "RUN_REGISTRY_COMMAND (interactive-only) -> " & resp
    Call Assert(InStr(resp, vbTab & "ERROR" & vbTab) > 0, _
                "interactive-only row must ERROR, not: " & resp)
    Call Assert(InStr(LCase(resp), "handoff") > 0, _
                "refusal note should point at HANDOFF: " & resp)

    ' Close-out guard: even if somehow gated through, a COMMAND:-only
    ' recipe must be refused. DESCRIBE confirms the recipe has no datapoint.
    Dim desc As String
    desc = WZTCBridge.ExecuteOp("R5" & vbTab & "DESCRIBE_REGISTRY_COMMAND" & vbTab & _
                                "opName=DIMENSION_LINEAR_SIZE_ARROW")
    Call Assert(InStr(desc, "interactive-only-use-handoff") > 0, _
                "DIMENSION_LINEAR_SIZE_ARROW should be interactive-only: " & desc)
End Sub

Private Sub TestRunVerifiedSettingsKeyin()
    ' Settings-only recipe with no COMMAND: — should pass the gate and run.
    Dim resp As String
    resp = WZTCBridge.ExecuteOp("R6" & vbTab & "RUN_REGISTRY_COMMAND" & vbTab & _
                                "opName=ACTIVE_COLOR" & vbTab & "color=0")
    Debug.Print "RUN_REGISTRY_COMMAND (ACTIVE_COLOR) -> " & resp
    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0, _
                "verified settings keyin should OK: " & resp)

    ' direct_api via RUN_REGISTRY_COMMAND must refuse (use dedicated op)
    Dim refuse As String
    refuse = WZTCBridge.ExecuteOp("R7" & vbTab & "RUN_REGISTRY_COMMAND" & vbTab & _
                                  "opName=MOVE_ELEMENT" & vbTab & "elementId=1" & vbTab & _
                                  "deltaX=1" & vbTab & "deltaY=0")
    Debug.Print "RUN_REGISTRY_COMMAND (direct_api MOVE_ELEMENT) -> " & refuse
    Call Assert(InStr(refuse, vbTab & "ERROR" & vbTab) > 0, _
                "direct_api via RUN_REGISTRY_COMMAND must ERROR: " & refuse)
End Sub

Private Sub TestOwnElementOnlyRefusal()
    ' Fabricate an ID that cannot be in the journal
    Dim resp As String
    resp = WZTCBridge.ExecuteOp("R8" & vbTab & "MOVE_ELEMENT" & vbTab & _
                                "elementId=999999999" & vbTab & "deltaX=1" & vbTab & "deltaY=0")
    Debug.Print "MOVE_ELEMENT (ownElementOnly refuse) -> " & resp
    Call Assert(InStr(resp, vbTab & "ERROR" & vbTab) > 0, _
                "ownElementOnly must refuse unknown elementId: " & resp)
    Call Assert(InStr(resp, "ownElementOnly") > 0, _
                "refusal should mention ownElementOnly: " & resp)
End Sub

Private Sub TestDeleteNotUndoable()
    ' Place a cell, delete it, confirm UNDO_LAST_OP does not try to
    ' recreate it (skips the notUndoable=Y row and undoes the PLACE
    ' if somehow still present — or reports nothing undoable related
    ' to the delete itself).
    Dim placeResp As String
    placeResp = WZTCBridge.ExecuteOp("R9" & vbTab & "PLACE_CELL" & vbTab & _
                                     "cellName=TWZAP_P" & vbTab & "ptX=5100" & vbTab & "ptY=5100")
    Debug.Print "PLACE_CELL (for delete test) -> " & placeResp
    If InStr(placeResp, vbTab & "OK" & vbTab) = 0 Then
        Debug.Print "SKIP TestDeleteNotUndoable — PLACE_CELL failed (no design file?)"
        Exit Sub
    End If
    Dim placedID As String: placedID = ExtractStrParam(placeResp, "elementId")

    Dim delResp As String
    delResp = WZTCBridge.ExecuteOp("R10" & vbTab & "DELETE_ELEMENT" & vbTab & _
                                   "elementId=" & placedID)
    Debug.Print "DELETE_ELEMENT -> " & delResp
    Call Assert(InStr(delResp, vbTab & "OK" & vbTab) > 0, "DELETE_ELEMENT failed: " & delResp)
    Call Assert(InStr(delResp, "notUndoable=Y") > 0, _
                "DELETE_ELEMENT must declare notUndoable=Y: " & delResp)

    Dim undoResp As String
    undoResp = WZTCBridge.ExecuteOp("R11" & vbTab & "UNDO_LAST_OP")
    Debug.Print "UNDO_LAST_OP after DELETE -> " & undoResp
    ' Must NOT claim to have undone R10 (the delete). It may undo an
    ' earlier op (R9 is already deleted so its elementId delete is a
    ' no-op OK, or some older op) — but undidReqId must not be R10.
    If InStr(undoResp, vbTab & "OK" & vbTab) > 0 Then
        Call Assert(ExtractStrParam(undoResp, "undidReqId") <> "R10", _
                    "UNDO_LAST_OP must not target DELETE_ELEMENT reqId R10: " & undoResp)
    End If
End Sub

Private Sub TestLiveMoveAndUndo()
    Dim placeResp As String
    placeResp = WZTCBridge.ExecuteOp("R12" & vbTab & "PLACE_CELL" & vbTab & _
                                     "cellName=TWZAP_P" & vbTab & "ptX=5200" & vbTab & "ptY=5200")
    Debug.Print "PLACE_CELL (for move/undo) -> " & placeResp
    If InStr(placeResp, vbTab & "OK" & vbTab) = 0 Then
        Debug.Print "SKIP TestLiveMoveAndUndo — PLACE_CELL failed (no design file?)"
        Exit Sub
    End If
    Dim placedID As String: placedID = ExtractStrParam(placeResp, "elementId")

    Dim moveResp As String
    moveResp = WZTCBridge.ExecuteOp("R13" & vbTab & "MOVE_ELEMENT" & vbTab & _
                                    "elementId=" & placedID & vbTab & _
                                    "deltaX=10" & vbTab & "deltaY=0" & vbTab & _
                                    "reason=M6 live move test")
    Debug.Print "MOVE_ELEMENT -> " & moveResp
    Call Assert(InStr(moveResp, vbTab & "OK" & vbTab) > 0, "MOVE_ELEMENT failed: " & moveResp)
    Call Assert(InStr(moveResp, "priorDeltaX=-10") > 0, _
                "MOVE_ELEMENT should embed priorDeltaX=-10: " & moveResp)

    Dim undoResp As String
    undoResp = WZTCBridge.ExecuteOp("R14" & vbTab & "UNDO_LAST_OP")
    Debug.Print "UNDO_LAST_OP (move) -> " & undoResp
    Call Assert(InStr(undoResp, vbTab & "OK" & vbTab) > 0, "UNDO of MOVE failed: " & undoResp)
    Call Assert(ExtractStrParam(undoResp, "undidReqId") = "R13", _
                "UNDO should target R13 move, got: " & undoResp)

    ' Cleanup: undo the place. This is also the regression check for the
    ' undo ping-pong bug found live 2026-07-31 -- R14's own RESP line
    ' embeds elementId=/priorDeltaX= (the move-undo it performed), and
    ' without ExecUndoLastOp declaring notUndoable=Y on itself, this call
    ' would misread R14's response as itself a further-undoable move and
    ' redo R13 instead of reaching back to the original R12 PLACE_CELL.
    Dim cleanup As String
    cleanup = WZTCBridge.ExecuteOp("R15" & vbTab & "UNDO_LAST_OP")
    Debug.Print "UNDO_LAST_OP (cleanup place) -> " & cleanup
    Call Assert(InStr(cleanup, vbTab & "OK" & vbTab) > 0, "cleanup UNDO_LAST_OP failed: " & cleanup)
    Call Assert(ExtractStrParam(cleanup, "undidReqId") = "R12", _
                "cleanup UNDO_LAST_OP must reach back to R12 (PLACE_CELL), not redo R13/R14 " & _
                "(ping-pong regression): " & cleanup)
End Sub

Private Sub TestLiveChangeLevelAndUndo()
    ' Regression check for the el.Level.Name crash found live 2026-07-31
    ' (runtime error 91 reading a level back on a freshly re-scanned
    ' element -- ExecChangeElementLevelByID now guards that read the same
    ' way WZTCQuery.FindElementsNear already does).
    Dim placeResp As String
    placeResp = WZTCBridge.ExecuteOp("R16" & vbTab & "PLACE_CELL" & vbTab & _
                                     "cellName=TWZAP_P" & vbTab & "ptX=5400" & vbTab & "ptY=5400")
    Debug.Print "PLACE_CELL (for level-change test) -> " & placeResp
    If InStr(placeResp, vbTab & "OK" & vbTab) = 0 Then
        Debug.Print "SKIP TestLiveChangeLevelAndUndo — PLACE_CELL failed (no design file?)"
        Exit Sub
    End If
    Dim placedID As String: placedID = ExtractStrParam(placeResp, "elementId")

    Dim levelResp As String
    levelResp = WZTCBridge.ExecuteOp("R17" & vbTab & "CHANGE_ELEMENT_LEVEL" & vbTab & _
                                     "elementId=" & placedID & vbTab & "level=Default" & vbTab & _
                                     "reason=M6 live level-change regression test")
    Debug.Print "CHANGE_ELEMENT_LEVEL -> " & levelResp
    Call Assert(InStr(levelResp, vbTab & "OK" & vbTab) > 0, "CHANGE_ELEMENT_LEVEL failed: " & levelResp)

    ' Either the prior level was readable (normal path) or the function
    ' honestly declared it couldn't undo -- either is acceptable, silently
    ' crashing or silently lying about undoability is not.
    Dim gotPriorLevel As Boolean
    gotPriorLevel = InStr(levelResp, "priorLevel=") > 0
    Dim gotNotUndoable As Boolean
    gotNotUndoable = InStr(levelResp, "notUndoable=Y") > 0
    Call Assert(gotPriorLevel Or gotNotUndoable, _
                "CHANGE_ELEMENT_LEVEL must either report priorLevel or declare notUndoable=Y: " & levelResp)

    ' Cleanup regardless of which branch fired above.
    Dim cleanup As String
    cleanup = WZTCBridge.ExecuteOp("R18" & vbTab & "DELETE_ELEMENT" & vbTab & "elementId=" & placedID)
    Debug.Print "DELETE_ELEMENT (cleanup) -> " & cleanup
End Sub

' ============================================================
' PARSING HELPERS — identical to DebugAgentLoopTest.bas
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

Private Sub Assert(cond As Boolean, msg As String)
    checkCount = checkCount + 1
    If Not cond Then
        failCount = failCount + 1
        Debug.Print "FAIL: " & msg
    End If
End Sub
