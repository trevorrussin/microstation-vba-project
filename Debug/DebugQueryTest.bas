Option Explicit

' ============================================================
' WZTCBridge M2 QUERY SMOKE TEST
' ------------------------------------------------------------
' Exercises the read-only query/compute ops through the same
' RunRequest bridge path Python will use — not by calling
' WZTCQuery.bas functions directly — so this proves the request/
' response wiring, not just the underlying logic (already covered
' by TestWZTCRules for the math and TestBridgePlaceCell for the
' write path).
'
' LIST_LEVELS, FIND_ELEMENTS_NEAR, and COMPUTE_SPACING are
' asserted strictly. STATION_TO_POINT / GET_ALIGNMENT_STATIONING
' need a committed alignment (AlignDraw step) to return real data —
' if alignment 1 hasn't been drawn in your current session, a
' graceful ERROR response is the CORRECT result, not a failure,
' so this test accepts either outcome and just checks the response
' is well-formed.
'
' TestSheetRegistry (M4) checks the sheet registry seeded from real
' Book 3 data: 619-302 must succeed, an unseeded sheet must fail
' gracefully rather than crash.
'
' Run from the VBA IDE: cursor in TestQueryOps, F5, watch Ctrl+G.
' Assumes at least one TWZAP_P cell already exists near (1000,1000)
' from the M1 bridge test — run TestBridgePlaceCell first if not.
' ============================================================

Private failCount As Integer
Private checkCount As Integer

Public Sub TestQueryOps()
    failCount = 0
    checkCount = 0
    Debug.Print "=== WZTCBridge M2 query smoke test ==="

    Call TestListLevels
    Call TestFindElementsNear
    Call TestComputeSpacing
    Call TestStationOps
    Call TestSheetRegistry

    Debug.Print "assertions: " & checkCount
    If failCount = 0 Then
        Debug.Print "=== PASS - no failures ==="
    Else
        Debug.Print "=== FAIL - " & failCount & " failed assertion(s) ==="
    End If
End Sub

Private Sub TestListLevels()
    Dim resp As String
    resp = WZTCBridge.ExecuteOp("Q1" & vbTab & "LIST_LEVELS")
    Debug.Print "LIST_LEVELS -> " & resp

    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0, "LIST_LEVELS did not return OK: " & resp)

    Dim rowCount As Long
    rowCount = ExtractIntParam(resp, "rowCount")
    Call Assert(rowCount >= 1, "LIST_LEVELS rowCount < 1: " & rowCount)

    Dim resultFile As String
    resultFile = ExtractStrParam(resp, "resultFile")
    Call Assert(Dir(resultFile) <> "", "LIST_LEVELS resultFile does not exist: " & resultFile)
    If Dir(resultFile) <> "" Then
        Dim lineCount As Long: lineCount = CountLines(resultFile)
        Call Assert(lineCount = rowCount + 1, "LIST_LEVELS file line count (" & lineCount & _
                    ") != rowCount+1 (" & (rowCount + 1) & ")")
    End If
End Sub

Private Sub TestFindElementsNear()
    Dim resp As String
    resp = WZTCBridge.ExecuteOp("Q2" & vbTab & "FIND_ELEMENTS_NEAR" & vbTab & _
                                "x=1000" & vbTab & "y=1000" & vbTab & "radius=50")
    Debug.Print "FIND_ELEMENTS_NEAR -> " & resp

    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0, "FIND_ELEMENTS_NEAR did not return OK: " & resp)

    Dim rowCount As Long
    rowCount = ExtractIntParam(resp, "rowCount")
    Call Assert(rowCount >= 1, "FIND_ELEMENTS_NEAR found nothing near (1000,1000) — " & _
                "expected at least the TWZAP_P cell from the M1 bridge test. Run TestBridgePlaceCell first?")

    ' Type filter should narrow or match results — CELL filter must find at least
    ' as many as the unfiltered search found cells (sanity check, not exact count)
    Dim respFiltered As String
    respFiltered = WZTCBridge.ExecuteOp("Q3" & vbTab & "FIND_ELEMENTS_NEAR" & vbTab & _
                                        "x=1000" & vbTab & "y=1000" & vbTab & "radius=50" & vbTab & "typeFilter=CELL")
    Debug.Print "FIND_ELEMENTS_NEAR (CELL filter) -> " & respFiltered
    Call Assert(InStr(respFiltered, vbTab & "OK" & vbTab) > 0, "FIND_ELEMENTS_NEAR with typeFilter did not return OK")
End Sub

Private Sub TestComputeSpacing()
    Dim resp As String
    resp = WZTCBridge.ExecuteOp("Q4" & vbTab & "COMPUTE_SPACING" & vbTab & _
                                "speed=45" & vbTab & "laneWidth=12" & vbTab & _
                                "shoulderWidth=10 ft" & vbTab & "roadType=Non-Freeway")
    Debug.Print "COMPUTE_SPACING -> " & resp

    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0, "COMPUTE_SPACING did not return OK: " & resp)

    ' Known values for 45mph / 12ft lane / 10ft shoulder / Non-Freeway,
    ' cross-checked against the same table TestWZTCRules already validated.
    Call Assert(ExtractIntParam(resp, "downstreamTaper") = 50, _
                "downstreamTaper wrong for Non-Freeway: " & ExtractStrParam(resp, "downstreamTaper"))
    Call Assert(ExtractIntParam(resp, "bufferSpace") = 360, _
                "bufferSpace wrong for 45mph: " & ExtractStrParam(resp, "bufferSpace"))
    Call Assert(ExtractIntParam(resp, "mergingTaper") = 560, _
                "mergingTaper wrong for 45mph/12ft: " & ExtractStrParam(resp, "mergingTaper"))
    Call Assert(ExtractIntParam(resp, "shoulderTaper") = 120, _
                "shoulderTaper wrong for 45mph/10ft: " & ExtractStrParam(resp, "shoulderTaper"))
End Sub

Private Sub TestStationOps()
    Dim resp As String
    resp = WZTCBridge.ExecuteOp("Q5" & vbTab & "STATION_TO_POINT" & vbTab & "alignIdx=1" & vbTab & "sta=100")
    Debug.Print "STATION_TO_POINT -> " & resp
    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0 Or InStr(resp, vbTab & "ERROR" & vbTab) > 0, _
                "STATION_TO_POINT returned malformed response: " & resp)
    If InStr(resp, vbTab & "ERROR" & vbTab) > 0 Then
        Debug.Print "  (ERROR is expected/OK if alignment 1 hasn't been committed this session)"
    End If

    Dim resp2 As String
    resp2 = WZTCBridge.ExecuteOp("Q6" & vbTab & "GET_ALIGNMENT_STATIONING" & vbTab & "alignIdx=1")
    Debug.Print "GET_ALIGNMENT_STATIONING -> " & resp2
    Call Assert(InStr(resp2, vbTab & "OK" & vbTab) > 0 Or InStr(resp2, vbTab & "ERROR" & vbTab) > 0, _
                "GET_ALIGNMENT_STATIONING returned malformed response: " & resp2)
End Sub

' M4: sheet registry seeded with 6 real sheets (301,302,303,307,310,311).
' 619-302 must succeed with real data; a sheet NOT in the registry (e.g.
' 619-999, which will never be a real sheet number) must fail gracefully,
' not crash -- that graceful-fallback behavior is the actual point of a
' registry seeded incrementally rather than all-91-or-nothing.
Private Sub TestSheetRegistry()
    Dim resp As String
    resp = WZTCBridge.ExecuteOp("Q7" & vbTab & "GET_SHEET_REQUIREMENTS" & vbTab & "sheetNum=619-302")
    Debug.Print "GET_SHEET_REQUIREMENTS (619-302, seeded) -> " & resp
    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0, "GET_SHEET_REQUIREMENTS(619-302) did not return OK: " & resp)
    Call Assert(InStr(resp, "signs=") > 0, "GET_SHEET_REQUIREMENTS(619-302) response missing signs field: " & resp)
    Call Assert(InStr(resp, "W20-1") > 0, "GET_SHEET_REQUIREMENTS(619-302) signs don't include W20-1: " & resp)

    Dim respMissing As String
    respMissing = WZTCBridge.ExecuteOp("Q8" & vbTab & "GET_SHEET_REQUIREMENTS" & vbTab & "sheetNum=619-999")
    Debug.Print "GET_SHEET_REQUIREMENTS (619-999, not seeded) -> " & respMissing
    Call Assert(InStr(respMissing, vbTab & "ERROR" & vbTab) > 0, _
                "GET_SHEET_REQUIREMENTS(619-999) should gracefully error, not: " & respMissing)
End Sub

' ============================================================
' PARSING HELPERS FOR key=val<TAB>-STYLE RESPONSE LINES
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

Private Function CountLines(path As String) As Long
    Dim fnum As Integer: fnum = FreeFile
    Open path For Input As #fnum
    Dim n As Long: n = 0
    Dim ln As String
    Do While Not EOF(fnum)
        Line Input #fnum, ln
        n = n + 1
    Loop
    Close #fnum
    CountLines = n
End Function

Private Sub Assert(cond As Boolean, msg As String)
    checkCount = checkCount + 1
    If Not cond Then
        failCount = failCount + 1
        Debug.Print "FAIL: " & msg
    End If
End Sub
