Option Explicit

' ============================================================
' WZTC RULES — DETERMINISTIC MUTCD NY / NYSDOT SPACING MATH
' ------------------------------------------------------------
' UI-free engineering rules extracted from WZTCDesigner.frm
' GenerateSpacingTable(). No form references, no MsgBox, no
' control reads — safe to call from a form, a test harness, or
' an automation bridge.
'
' Tables are transcribed verbatim from the form. Several entries
' break the pattern of their neighbours (see TABLE ANOMALIES
' below). They are reproduced exactly as they were; do not
' "correct" them without a NYSDOT reference, because changing
' one alters output geometry on every sheet that uses it.
' ============================================================

' ---- Computed spacing / clearance values for one configuration ----
Public Type WZTCSpacing
    DownstreamTaper       As Double   ' ft
    VehicleSpace          As Double   ' ft
    BufferSpace           As Double   ' ft
    MergingTaper          As Double   ' ft
    ShoulderTaper         As Double   ' ft
    AdvanceWarningSpacing As Double   ' ft
    RollAheadDistance     As Double   ' ft
    UpTaperBarrier        As Double   ' ft
    UpTaperBeam           As Double   ' ft
    FlareBarrier          As String   ' ratio, e.g. "11:1"
    FlareBeam             As String   ' ratio, e.g. "9:1"
    SkipMerge             As Integer
    SkipShoulder          As Integer
    SkipBuffer            As Integer
    SkipRollAhead         As Integer
    SkipTotal             As Integer
    ChanMerge             As Integer
    ChanShoulder          As Integer
    ChanTotal             As Integer
End Type

' ============================================================
' TABLE ANOMALIES — transcribed as-is, verified against the form
' ------------------------------------------------------------
' Merging taper, 65 mph / 12 ft lane : skip=19 chan=20
'   (neighbours step 16/17 then 18/19, so 20/21 would be the
'    pattern — the form says 19/20)
' Shoulder taper, 40 mph / "5-7 ft"  : taper=80 skip=1 chan=2
'   (every other taper=80 row uses skip=2 chan=3)
' Shoulder taper, 40 mph / "10 ft"   : taper=120 skip=2 chan=3
'   (every other taper=120 row uses skip=3 chan=4)
' ============================================================

' ============================================================
' COMPUTE ALL SPACING VALUES
' speedMph         : 25/30/35/40/45/50/55/65 (note: no 60)
' laneWidthFt      : 10/11/12
' shoulderWidthKey : display string — "<= 4 ft", "5-7 ft",
'                    "8 ft", "9 ft", "10 ft", "11 ft", "12 ft"
' roadType         : "Freeway" or "Non-Freeway"
' ============================================================
Public Function ComputeSpacing(speedMph As Integer, laneWidthFt As Integer, _
                               shoulderWidthKey As String, roadType As String) As WZTCSpacing
    Dim r As WZTCSpacing
    Dim speed As Integer, laneWidth As Integer
    speed = speedMph
    laneWidth = laneWidthFt

    If LCase(Trim(roadType)) = "non-freeway" Then
        r.DownstreamTaper = 50
    Else
        r.DownstreamTaper = 100
    End If
    r.VehicleSpace = 50

    Select Case speed
        Case 25: r.BufferSpace = 155
        Case 30: r.BufferSpace = 200
        Case 35: r.BufferSpace = 250
        Case 40: r.BufferSpace = 305
        Case 45: r.BufferSpace = 360
        Case 50: r.BufferSpace = 425
        Case 55: r.BufferSpace = 495
        Case 60: r.BufferSpace = 570   ' MUTCD 11th Ed. Table 6B-2 -- was missing; fell through
                                        ' to the Case Else formula (4200 ft, clearly wrong).
                                        ' Currently unreachable since the speed dropdown never
                                        ' offers 60 mph, but fixed now that it's confirmed wrong.
        Case 65: r.BufferSpace = 645
        Case Else: r.BufferSpace = speed * 70
    End Select

    Select Case speed
        Case 25: r.SkipBuffer = 4
        Case 30: r.SkipBuffer = 5
        Case 35: r.SkipBuffer = 6
        Case 40: r.SkipBuffer = 8
        Case 45: r.SkipBuffer = 9
        Case 50: r.SkipBuffer = 11
        Case 55: r.SkipBuffer = 13
        Case 65: r.SkipBuffer = 16
        Case Else: r.SkipBuffer = 0
    End Select

    ' ---- Merging / shifting taper: speed x lane width ----
    Select Case speed
        Case 25
            Select Case laneWidth
                Case 10: r.MergingTaper = 120: r.SkipMerge = 3: r.ChanMerge = 4
                Case 11: r.MergingTaper = 120: r.SkipMerge = 3: r.ChanMerge = 4
                Case 12: r.MergingTaper = 120: r.SkipMerge = 3: r.ChanMerge = 4
                Case Else: r.MergingTaper = 120: r.SkipMerge = 3: r.ChanMerge = 4
            End Select
        Case 30
            Select Case laneWidth
                Case 10: r.MergingTaper = 160: r.SkipMerge = 4: r.ChanMerge = 5
                Case 11: r.MergingTaper = 160: r.SkipMerge = 4: r.ChanMerge = 5
                Case 12: r.MergingTaper = 200: r.SkipMerge = 5: r.ChanMerge = 6
                Case Else: r.MergingTaper = 160: r.SkipMerge = 4: r.ChanMerge = 5
            End Select
        Case 35
            Select Case laneWidth
                Case 10: r.MergingTaper = 200: r.SkipMerge = 5: r.ChanMerge = 6
                Case 11: r.MergingTaper = 240: r.SkipMerge = 6: r.ChanMerge = 7
                Case 12: r.MergingTaper = 240: r.SkipMerge = 6: r.ChanMerge = 7
                Case Else: r.MergingTaper = 200: r.SkipMerge = 5: r.ChanMerge = 6
            End Select
        Case 40
            Select Case laneWidth
                Case 10: r.MergingTaper = 280: r.SkipMerge = 7: r.ChanMerge = 8
                Case 11: r.MergingTaper = 320: r.SkipMerge = 8: r.ChanMerge = 9
                Case 12: r.MergingTaper = 320: r.SkipMerge = 8: r.ChanMerge = 9
                Case Else: r.MergingTaper = 280: r.SkipMerge = 7: r.ChanMerge = 8
            End Select
        Case 45
            Select Case laneWidth
                Case 10: r.MergingTaper = 440: r.SkipMerge = 11: r.ChanMerge = 12
                Case 11: r.MergingTaper = 520: r.SkipMerge = 13: r.ChanMerge = 14
                Case 12: r.MergingTaper = 560: r.SkipMerge = 14: r.ChanMerge = 15
                Case Else: r.MergingTaper = 440: r.SkipMerge = 11: r.ChanMerge = 12
            End Select
        Case 50
            Select Case laneWidth
                Case 10: r.MergingTaper = 520: r.SkipMerge = 13: r.ChanMerge = 14
                Case 11: r.MergingTaper = 560: r.SkipMerge = 14: r.ChanMerge = 15
                Case 12: r.MergingTaper = 600: r.SkipMerge = 15: r.ChanMerge = 16
                Case Else: r.MergingTaper = 520: r.SkipMerge = 13: r.ChanMerge = 14
            End Select
        Case 55
            Select Case laneWidth
                Case 10: r.MergingTaper = 560: r.SkipMerge = 14: r.ChanMerge = 15
                Case 11: r.MergingTaper = 600: r.SkipMerge = 15: r.ChanMerge = 16
                Case 12: r.MergingTaper = 680: r.SkipMerge = 17: r.ChanMerge = 18
                Case Else: r.MergingTaper = 560: r.SkipMerge = 14: r.ChanMerge = 15
            End Select
        Case 65
            Select Case laneWidth
                Case 10: r.MergingTaper = 640: r.SkipMerge = 16: r.ChanMerge = 17
                Case 11: r.MergingTaper = 720: r.SkipMerge = 18: r.ChanMerge = 19
                Case 12: r.MergingTaper = 800: r.SkipMerge = 19: r.ChanMerge = 20
                Case Else: r.MergingTaper = 640: r.SkipMerge = 16: r.ChanMerge = 17
            End Select
        Case Else
            r.MergingTaper = (speed * (laneWidth) ^ 2) / 60: r.SkipMerge = 0: r.ChanMerge = 0
    End Select

    ' ---- Shoulder taper: speed x shoulder width (keyed on display string) ----
    Select Case speed
        Case 25
            Select Case shoulderWidthKey
                Case "<= 4 ft": r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "5-7 ft":  r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "8 ft":    r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "9 ft":    r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "10 ft":   r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "11 ft":   r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "12 ft":   r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
                Case Else:      r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
            End Select
        Case 30
            Select Case shoulderWidthKey
                Case "<= 4 ft": r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "5-7 ft":  r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "8 ft":    r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "9 ft":    r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "10 ft":   r.ShoulderTaper = 80: r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "11 ft":   r.ShoulderTaper = 80: r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "12 ft":   r.ShoulderTaper = 80: r.SkipShoulder = 2: r.ChanShoulder = 3
                Case Else:      r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
            End Select
        Case 35
            Select Case shoulderWidthKey
                Case "<= 4 ft": r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "5-7 ft":  r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "8 ft":    r.ShoulderTaper = 80: r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "9 ft":    r.ShoulderTaper = 80: r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "10 ft":   r.ShoulderTaper = 80: r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "11 ft":   r.ShoulderTaper = 80: r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "12 ft":   r.ShoulderTaper = 80: r.SkipShoulder = 2: r.ChanShoulder = 3
                Case Else:      r.ShoulderTaper = 40: r.SkipShoulder = 1: r.ChanShoulder = 2
            End Select
        Case 40
            Select Case shoulderWidthKey
                Case "<= 4 ft": r.ShoulderTaper = 40:  r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "5-7 ft":  r.ShoulderTaper = 80:  r.SkipShoulder = 1: r.ChanShoulder = 2
                Case "8 ft":    r.ShoulderTaper = 80:  r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "9 ft":    r.ShoulderTaper = 80:  r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "10 ft":   r.ShoulderTaper = 120: r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "11 ft":   r.ShoulderTaper = 120: r.SkipShoulder = 3: r.ChanShoulder = 4
                Case "12 ft":   r.ShoulderTaper = 120: r.SkipShoulder = 3: r.ChanShoulder = 4
                Case Else:      r.ShoulderTaper = 40:  r.SkipShoulder = 1: r.ChanShoulder = 2
            End Select
        Case 45
            Select Case shoulderWidthKey
                Case "<= 4 ft": r.ShoulderTaper = 80:  r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "5-7 ft":  r.ShoulderTaper = 80:  r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "8 ft":    r.ShoulderTaper = 120: r.SkipShoulder = 3: r.ChanShoulder = 4
                Case "9 ft":    r.ShoulderTaper = 120: r.SkipShoulder = 3: r.ChanShoulder = 4
                Case "10 ft":   r.ShoulderTaper = 120: r.SkipShoulder = 3: r.ChanShoulder = 4
                Case "11 ft":   r.ShoulderTaper = 120: r.SkipShoulder = 3: r.ChanShoulder = 4
                Case "12 ft":   r.ShoulderTaper = 160: r.SkipShoulder = 4: r.ChanShoulder = 5
                Case Else:      r.ShoulderTaper = 80:  r.SkipShoulder = 2: r.ChanShoulder = 3
            End Select
        Case 50
            Select Case shoulderWidthKey
                Case "<= 4 ft": r.ShoulderTaper = 80:  r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "5-7 ft":  r.ShoulderTaper = 120: r.SkipShoulder = 3: r.ChanShoulder = 4
                Case "8 ft":    r.ShoulderTaper = 160: r.SkipShoulder = 4: r.ChanShoulder = 5
                Case "9 ft":    r.ShoulderTaper = 160: r.SkipShoulder = 4: r.ChanShoulder = 5
                Case "10 ft":   r.ShoulderTaper = 160: r.SkipShoulder = 4: r.ChanShoulder = 5
                Case "11 ft":   r.ShoulderTaper = 160: r.SkipShoulder = 4: r.ChanShoulder = 5
                Case "12 ft":   r.ShoulderTaper = 160: r.SkipShoulder = 4: r.ChanShoulder = 5
                Case Else:      r.ShoulderTaper = 80:  r.SkipShoulder = 2: r.ChanShoulder = 3
            End Select
        Case 55
            Select Case shoulderWidthKey
                Case "<= 4 ft": r.ShoulderTaper = 80:  r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "5-7 ft":  r.ShoulderTaper = 120: r.SkipShoulder = 3: r.ChanShoulder = 4
                Case "8 ft":    r.ShoulderTaper = 160: r.SkipShoulder = 4: r.ChanShoulder = 5
                Case "9 ft":    r.ShoulderTaper = 160: r.SkipShoulder = 4: r.ChanShoulder = 5
                Case "10 ft":   r.ShoulderTaper = 160: r.SkipShoulder = 4: r.ChanShoulder = 5
                Case "11 ft":   r.ShoulderTaper = 200: r.SkipShoulder = 5: r.ChanShoulder = 6
                Case "12 ft":   r.ShoulderTaper = 200: r.SkipShoulder = 5: r.ChanShoulder = 6
                Case Else:      r.ShoulderTaper = 80:  r.SkipShoulder = 2: r.ChanShoulder = 3
            End Select
        Case 65
            Select Case shoulderWidthKey
                Case "<= 4 ft": r.ShoulderTaper = 80:  r.SkipShoulder = 2: r.ChanShoulder = 3
                Case "5-7 ft":  r.ShoulderTaper = 160: r.SkipShoulder = 4: r.ChanShoulder = 5
                Case "8 ft":    r.ShoulderTaper = 200: r.SkipShoulder = 5: r.ChanShoulder = 6
                Case "9 ft":    r.ShoulderTaper = 240: r.SkipShoulder = 6: r.ChanShoulder = 7
                Case "10 ft":   r.ShoulderTaper = 240: r.SkipShoulder = 6: r.ChanShoulder = 7
                Case "11 ft":   r.ShoulderTaper = 280: r.SkipShoulder = 7: r.ChanShoulder = 8
                Case "12 ft":   r.ShoulderTaper = 280: r.SkipShoulder = 7: r.ChanShoulder = 8
                Case Else:      r.ShoulderTaper = 80:  r.SkipShoulder = 2: r.ChanShoulder = 3
            End Select
        Case Else
            r.ShoulderTaper = speed * 0.8: r.SkipShoulder = 0: r.ChanShoulder = 0
    End Select

    Select Case speed
        Case 25: r.AdvanceWarningSpacing = 515
        Case 30: r.AdvanceWarningSpacing = 620
        Case 35: r.AdvanceWarningSpacing = 720
        Case 40: r.AdvanceWarningSpacing = 825
        Case 45: r.AdvanceWarningSpacing = 930
        Case 50: r.AdvanceWarningSpacing = 1030
        Case 55: r.AdvanceWarningSpacing = 1135
        Case 65: r.AdvanceWarningSpacing = 1365
        Case Else: r.AdvanceWarningSpacing = speed * 10
    End Select

    Select Case speed
        Case 25, 30, 35, 40: r.RollAheadDistance = 120
        Case 45, 50, 55:     r.RollAheadDistance = 160
        Case 65:             r.RollAheadDistance = 200
        Case Else:           r.RollAheadDistance = 120
    End Select

    Select Case speed
        Case 25, 30, 35, 40: r.SkipRollAhead = 3
        Case 45, 50, 55:     r.SkipRollAhead = 4
        Case 65:             r.SkipRollAhead = 5
        Case Else:           r.SkipRollAhead = 0
    End Select

    Select Case speed
        Case 25, 30, 35: r.FlareBarrier = "8:1"
        Case 40, 45:     r.FlareBarrier = "11:1"
        Case 50:         r.FlareBarrier = "14:1"
        Case 55:         r.FlareBarrier = "16:1"
        Case 65:         r.FlareBarrier = "20:1"
        Case Else:       r.FlareBarrier = ""
    End Select

    Select Case speed
        Case 25, 30, 35: r.FlareBeam = "7:1"
        Case 40, 45:     r.FlareBeam = "9:1"
        Case 50:         r.FlareBeam = "11:1"
        Case 55:         r.FlareBeam = "12:1"
        Case 65:         r.FlareBeam = "15:1"
        Case Else:       r.FlareBeam = ""
    End Select

    r.SkipTotal = r.SkipMerge + r.SkipShoulder + r.SkipBuffer + r.SkipRollAhead
    r.ChanTotal = r.ChanMerge + r.ChanShoulder

    r.UpTaperBarrier = ParseUpstreamTaper(r.FlareBarrier, laneWidth)
    r.UpTaperBeam = ParseUpstreamTaper(r.FlareBeam, laneWidth)

    ComputeSpacing = r
End Function

' ============================================================
' PARSE UPSTREAM TAPER — converts "X:Y" flare rate + lane width to taper length
' Formula: upstreamTaper = laneWidth x (X / Y)
' Example: "8:1" + 12ft lane -> 12 x 8 = 96 ft
' ============================================================
Public Function ParseUpstreamTaper(flareStr As String, laneWid As Integer) As Double
    On Error Resume Next
    ParseUpstreamTaper = 0
    If flareStr = "" Then Exit Function
    Dim parts() As String: parts = Split(flareStr, ":")
    If UBound(parts) < 1 Then Exit Function
    Dim num As Double: num = CDbl(Trim(parts(0)))
    Dim den As Double: den = CDbl(Trim(parts(1)))
    If den = 0 Then Exit Function
    ParseUpstreamTaper = laneWid * (num / den)
End Function

' ============================================================
' DEFAULT UPSTREAM ALIGNMENT ITEMS (alignment 1) — 7 spacing rows
' Fills the five parallel arrays RebuildAlignTable expects.
' Label strings must not change: SpacingBox.cls binds rows by
' display-label equality.
' ============================================================
Public Sub GetDefaultUpstreamItems(sp As WZTCSpacing, _
                                   ByRef rowTypes() As String, ByRef rowLabels() As String, _
                                   ByRef rowSpacings() As String, ByRef rowSizes() As String, _
                                   ByRef rowSides() As String, ByRef rowCount As Integer)
    rowCount = 7
    ReDim rowTypes(1 To 7)
    ReDim rowLabels(1 To 7)
    ReDim rowSpacings(1 To 7)
    ReDim rowSizes(1 To 7)
    ReDim rowSides(1 To 7)

    Dim i As Integer
    For i = 1 To 7
        rowTypes(i) = "Non-Sign"
        rowSizes(i) = ""
        rowSides(i) = "One Side"
    Next i

    rowLabels(1) = "Roll Ahead Distance":          rowSpacings(1) = Format(sp.RollAheadDistance, "0.0")
    rowLabels(2) = "Vehicle Space":                rowSpacings(2) = Format(sp.VehicleSpace, "0.0")
    rowLabels(3) = "Buffer Space":                 rowSpacings(3) = Format(sp.BufferSpace, "0.0")
    rowLabels(4) = "Merging/Shifting Taper":       rowSpacings(4) = Format(sp.MergingTaper, "0.0")
    rowLabels(5) = "Shoulder Taper":               rowSpacings(5) = Format(sp.ShoulderTaper, "0.0")
    rowLabels(6) = "Upstream Taper Temp Barrier":  rowSpacings(6) = Format(sp.UpTaperBarrier, "0.0")
    rowLabels(7) = "Upstream Taper Box/Corr Beam": rowSpacings(7) = Format(sp.UpTaperBeam, "0.0")
End Sub

' ============================================================
' DEFAULT DOWNSTREAM ALIGNMENT ITEMS (alignment 2) — 1 spacing row
' ============================================================
Public Sub GetDefaultDownstreamItems(sp As WZTCSpacing, _
                                     ByRef rowTypes() As String, ByRef rowLabels() As String, _
                                     ByRef rowSpacings() As String, ByRef rowSizes() As String, _
                                     ByRef rowSides() As String, ByRef rowCount As Integer)
    rowCount = 1
    ReDim rowTypes(1 To 1)
    ReDim rowLabels(1 To 1)
    ReDim rowSpacings(1 To 1)
    ReDim rowSizes(1 To 1)
    ReDim rowSides(1 To 1)

    rowTypes(1) = "Non-Sign"
    rowLabels(1) = "Downstream Taper"
    rowSpacings(1) = Format(sp.DownstreamTaper, "0.0")
    rowSizes(1) = ""
    rowSides(1) = "One Side"
End Sub
