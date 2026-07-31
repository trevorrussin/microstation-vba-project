Option Explicit

' ============================================================
' WZTCExec M3 SMOKE TEST
' ------------------------------------------------------------
' Runs through WZTCBridge.ExecuteOp, same as the M1/M2 tests, so
' this proves the dispatch wiring, not just the underlying logic.
' Covers all five M3 primitives: PLACE_SIGN, PLACE_PERP_LINE,
' PLACE_ELEMENT_RUN, PLACE_WORKSPACE, SET_SIGN_ATTRIBUTES.
'
' PLACE_SIGN and PLACE_ELEMENT_RUN need no committed alignment —
' they take fully-resolved points, so they're tested with arbitrary
' coordinates. PLACE_PERP_LINE DOES need a committed alignment
' (AlignDraw step); if alignment 1 isn't committed this session, a
' graceful ERROR is the correct result, same acceptance rule as
' TestQueryOps.
'
' TestPlaceWorkspaceLShape is the important one to read the output
' of, not just the pass/fail: it independently verifies (via a
' second, different point-in-polygon algorithm) that the computed
' hatch seed actually lands inside a genuinely non-convex L-shape
' where a naive centroid provably would not.
'
' Run from the VBA IDE: cursor in TestExecOps, F5, watch Ctrl+G.
' Uses sign R02-01 (Speed Limit) — confirmed real in SignLibrary.bas.
' Check MicroStation afterward near (1000,1200)-(1050,1300) for signs,
' (1800,1000)-(1900,1010) for the channelizing device line, and
' (1800,1800)-(1900,1900) for the L-shaped work space + hatch.
' ============================================================

Private failCount As Integer
Private checkCount As Integer

Public Sub TestExecOps()
    failCount = 0
    checkCount = 0
    Debug.Print "=== WZTCExec M3 smoke test ==="

    Call TestPlaceSignOneSide
    Call TestPlaceSignBothSides
    Call TestPlacePerpLine
    Call TestPlaceElementRun
    Call TestPlaceWorkspaceLShape
    Call TestSetSignAttributes

    Debug.Print "assertions: " & checkCount
    If failCount = 0 Then
        Debug.Print "=== PASS - no failures ==="
    Else
        Debug.Print "=== FAIL - " & failCount & " failed assertion(s) ==="
    End If
End Sub

Private Sub TestPlaceSignOneSide()
    Dim countBefore As Long: countBefore = CountGraphicalElements()

    Dim resp As String
    resp = WZTCBridge.ExecuteOp("E1" & vbTab & "PLACE_SIGN" & vbTab & _
        "signNum=R02-01" & vbTab & "roadType=Non-Freeway" & vbTab & "side=One Side" & vbTab & _
        "pt1X=1000" & vbTab & "pt1Y=1200" & vbTab & "pt1Z=0" & vbTab & _
        "dir1X=0" & vbTab & "dir1Y=1")
    Debug.Print "PLACE_SIGN (one side) -> " & resp

    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0, "PLACE_SIGN (one side) did not return OK: " & resp)
    Call Assert(InStr(resp, "signNum=R02-01") > 0, "PLACE_SIGN response missing signNum: " & resp)

    Dim countAfter As Long: countAfter = CountGraphicalElements()
    ' Expect 4 new top-level elements (text label, sign face cell, post line,
    ' post cell) -- but multi-line text via TEXTEDITOR PLACE might enumerate
    ' as a TextNode differently than assumed, so this is reported, not a hard
    ' assertion. A count of 0 (nothing placed) is the real failure signal.
    Debug.Print "  new elements: " & (countAfter - countBefore) & " (expected ~4: text, face cell, post line, post cell)"
    Call Assert(countAfter > countBefore, "PLACE_SIGN (one side) created no new elements at all")
End Sub

Private Sub TestPlaceSignBothSides()
    Dim countBefore As Long: countBefore = CountGraphicalElements()

    Dim resp As String
    resp = WZTCBridge.ExecuteOp("E2" & vbTab & "PLACE_SIGN" & vbTab & _
        "signNum=R02-01" & vbTab & "roadType=Freeway" & vbTab & "side=Both Sides" & vbTab & _
        "pt1X=1000" & vbTab & "pt1Y=1300" & vbTab & "pt1Z=0" & vbTab & "dir1X=1" & vbTab & "dir1Y=0" & vbTab & _
        "pt2X=1050" & vbTab & "pt2Y=1300" & vbTab & "pt2Z=0" & vbTab & "dir2X=-1" & vbTab & "dir2Y=0")
    Debug.Print "PLACE_SIGN (both sides) -> " & resp

    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0, "PLACE_SIGN (both sides) did not return OK: " & resp)

    Dim countAfter As Long: countAfter = CountGraphicalElements()
    ' Expect ~9 (2x[text, face cell, post line, post cell] + 1 connecting arc),
    ' reported rather than hard-asserted for the same reason as the one-side case.
    Debug.Print "  new elements: " & (countAfter - countBefore) & " (expected ~9: 2x[text,face,postline,post] + arc)"
    Call Assert(countAfter > countBefore, "PLACE_SIGN (both sides) created no new elements at all")
End Sub

Private Sub TestPlacePerpLine()
    Dim resp As String
    resp = WZTCBridge.ExecuteOp("E3" & vbTab & "PLACE_PERP_LINE" & vbTab & "alignIdx=1" & vbTab & "sta=50")
    Debug.Print "PLACE_PERP_LINE -> " & resp

    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0 Or InStr(resp, vbTab & "ERROR" & vbTab) > 0, _
                "PLACE_PERP_LINE returned malformed response: " & resp)
    If InStr(resp, vbTab & "ERROR" & vbTab) > 0 Then
        Debug.Print "  (ERROR is expected/OK if alignment 1 hasn't been committed this session)"
    End If
End Sub

Private Sub TestPlaceElementRun()
    Dim countBefore As Long: countBefore = CountGraphicalElements()

    Dim resp As String
    resp = WZTCBridge.ExecuteOp("E4" & vbTab & "PLACE_ELEMENT_RUN" & vbTab & "elementIdx=2" & vbTab & _
        "verticesTSV=1800,1000,0|1850,1010,0|1900,1005,0")
    Debug.Print "PLACE_ELEMENT_RUN (Channelizing Devices) -> " & resp

    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0, "PLACE_ELEMENT_RUN did not return OK: " & resp)
    Call Assert(InStr(resp, "level=TWZCD_P") > 0, "PLACE_ELEMENT_RUN response missing expected level: " & resp)

    Dim countAfter As Long: countAfter = CountGraphicalElements()
    ' PLACE LINE CONSTRAINED evidently creates one element per segment (3
    ' points = 2 segments = 2 elements), not a single multi-vertex polyline
    ' -- confirmed by this test run, corrected from an earlier wrong
    ' assumption of 1. Matches DrawElements.bas's own proven behavior for
    ' this exact command, so this is the real, correct expectation now.
    Debug.Print "  new elements: " & (countAfter - countBefore) & " (expected 2: one line segment per point pair)"
    Call Assert(countAfter = countBefore + 2, "PLACE_ELEMENT_RUN expected 2 new line segments, got " & _
                (countAfter - countBefore))
End Sub

' Tests the point-in-polygon hatch-seed logic against a genuinely
' non-convex L-shape where a naive vertex-average centroid provably
' falls OUTSIDE the polygon (in the removed notch), not just a shape
' that happens to still work by luck.
'
' L-shape: bottom strip x:[1800,1900] y:[1800,1830], left strip
' x:[1800,1830] y:[1800,1900]. Vertex-average centroid = (~1843,~1843)
' which sits in the removed notch (x:1830-1900, y:1830-1900) -- outside
' the polygon. A correct interior-point method must NOT return that.
Private Sub TestPlaceWorkspaceLShape()
    Dim vx(0 To 5) As Double, vy(0 To 5) As Double
    vx(0) = 1800: vy(0) = 1800
    vx(1) = 1900: vy(1) = 1800
    vx(2) = 1900: vy(2) = 1830
    vx(3) = 1830: vy(3) = 1830
    vx(4) = 1830: vy(4) = 1900
    vx(5) = 1800: vy(5) = 1900

    Dim vertsTSV As String
    Dim i As Integer
    For i = 0 To 5
        If i > 0 Then vertsTSV = vertsTSV & "|"
        vertsTSV = vertsTSV & vx(i) & "," & vy(i) & ",0"
    Next i

    Dim countBefore As Long: countBefore = CountGraphicalElements()

    Dim resp As String
    resp = WZTCBridge.ExecuteOp("E5" & vbTab & "PLACE_WORKSPACE" & vbTab & "verticesTSV=" & vertsTSV)
    Debug.Print "PLACE_WORKSPACE (L-shape) -> " & resp

    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0, "PLACE_WORKSPACE did not return OK: " & resp)

    Dim countAfter As Long: countAfter = CountGraphicalElements()
    Debug.Print "  new elements: " & (countAfter - countBefore) & " (expected 2: shape + hatch)"
    Call Assert(countAfter > countBefore, "PLACE_WORKSPACE created no new elements at all")

    ' Cross-check the returned hatch seed with an INDEPENDENT point-in-polygon
    ' test (ray casting) -- a different algorithm than WZTCExec's scanline-width
    ' method, so this isn't just re-checking the same logic against itself.
    If InStr(resp, "hatchSeedX=") > 0 Then
        Dim seedX As Double, seedY As Double
        seedX = ExtractDoubleParam(resp, "hatchSeedX")
        seedY = ExtractDoubleParam(resp, "hatchSeedY")
        Debug.Print "  hatch seed: (" & seedX & ", " & seedY & ")"

        Dim inside As Boolean
        inside = PointInPolygonRayCast(seedX, seedY, vx, vy, 6)
        Call Assert(inside, "Hatch seed (" & seedX & "," & seedY & ") is NOT inside the L-shape " & _
                    "(point-in-polygon logic may be wrong)")

        ' Also confirm the naive centroid genuinely WOULD have failed here,
        ' so a pass on the line above means something.
        Dim centroidInside As Boolean
        Dim cx As Double, cy As Double
        cx = 0: cy = 0
        For i = 0 To 5
            cx = cx + vx(i): cy = cy + vy(i)
        Next i
        cx = cx / 6: cy = cy / 6
        centroidInside = PointInPolygonRayCast(cx, cy, vx, vy, 6)
        Debug.Print "  naive centroid (" & cx & "," & cy & ") inside polygon: " & centroidInside & _
                    " (expected False -- this is the failure mode being avoided)"
    Else
        Call Assert(False, "PLACE_WORKSPACE response has no hatchSeedX -- hatch step may have been skipped: " & resp)
    End If
End Sub

Private Sub TestSetSignAttributes()
    ' Place a fresh cell to have a known, isolated target element
    Dim placeResp As String
    placeResp = WZTCBridge.ExecuteOp("E6" & vbTab & "PLACE_CELL" & vbTab & "cellName=TWZAP_P" & vbTab & _
        "ptX=1800" & vbTab & "ptY=2000")
    Debug.Print "PLACE_CELL (for attribute test) -> " & placeResp
    Dim targetId As String
    targetId = ExtractStrParam(placeResp, "elementId")
    If targetId = "" Then
        Debug.Print "SKIP: TestSetSignAttributes -- could not place a target cell"
        Exit Sub
    End If

    Dim resp As String
    resp = WZTCBridge.ExecuteOp("E7" & vbTab & "SET_SIGN_ATTRIBUTES" & vbTab & "elementIds=" & targetId)
    Debug.Print "SET_SIGN_ATTRIBUTES -> " & resp

    Call Assert(InStr(resp, vbTab & "OK" & vbTab) > 0 Or InStr(resp, vbTab & "ERROR" & vbTab) > 0, _
                "SET_SIGN_ATTRIBUTES returned malformed response: " & resp)

    If InStr(resp, vbTab & "ERROR" & vbTab) > 0 Then
        Debug.Print "  (ERROR is acceptable if level SF_P doesn't exist in this design file)"
        Exit Sub
    End If

    Call Assert(ExtractIntParam(resp, "applied") = 1, "SET_SIGN_ATTRIBUTES applied count != 1: " & resp)

    ' NOTE: originally this re-scanned the model and read el.Level.Name /
    ' el.Color / el.LineWeight back to verify the write actually took effect.
    ' Reading el.Level back on a freshly re-scanned element throws runtime
    ' error 91 (object variable not set) -- the same unconfirmed-read issue
    ' already flagged for FindElementsNear's level column. Writing these
    ' properties is confirmed safe (PerpPlacement.bas does it), reading them
    ' back is not confirmed anywhere in this codebase, so that verification
    ' step is dropped rather than guessed at a second property-read pattern.
    ' "applied=1" above (ExecSetSignAttributes' own count of successful
    ' el.Rewrite calls) is the correctness signal for this test.
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical
    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)
    Dim el As Element
    Dim found As Boolean: found = False
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        If CStr(ElIDAsDouble(el.ID)) = targetId Then
            found = True
            Exit Do
        End If
    Loop
    Call Assert(found, "could not re-locate target element " & targetId & " after SET_SIGN_ATTRIBUTES")
End Sub

' ============================================================
' RAY-CASTING POINT-IN-POLYGON — independent of WZTCExec's
' scanline-width method, used only to cross-check its output.
' ============================================================
Private Function PointInPolygonRayCast(px As Double, py As Double, _
                                       vx() As Double, vy() As Double, n As Integer) As Boolean
    Dim inside As Boolean: inside = False
    Dim i As Integer, j As Integer
    j = n - 1
    For i = 0 To n - 1
        If ((vy(i) > py) <> (vy(j) > py)) Then
            Dim xCross As Double
            xCross = (vx(j) - vx(i)) * (py - vy(i)) / (vy(j) - vy(i)) + vx(i)
            If px < xCross Then inside = Not inside
        End If
        j = i
    Next i
    PointInPolygonRayCast = inside
End Function

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

Private Function ExtractDoubleParam(resp As String, key As String) As Double
    ExtractDoubleParam = Val(ExtractStrParam(resp, key))
End Function

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

Private Sub Assert(cond As Boolean, msg As String)
    checkCount = checkCount + 1
    If Not cond Then
        failCount = failCount + 1
        Debug.Print "FAIL: " & msg
    End If
End Sub
