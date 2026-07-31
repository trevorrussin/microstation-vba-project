Option Explicit

' ============================================================
' WZTCRules SMOKE TEST
' ------------------------------------------------------------
' Runs ComputeSpacing across the full input matrix that the
' WZTCDesigner dropdowns can produce (8 speeds x 3 lane widths
' x 7 shoulder widths x 2 road types = 336 combinations) and
' checks structural invariants.
'
' This does NOT re-verify the lookup table values — those were
' diffed mechanically against the pre-refactor GenerateSpacingTable.
' What it catches is what only the IDE can: compile errors, runtime
' errors, and any combination that silently yields no value.
'
' Run from the VBA IDE: place cursor in TestWZTCRules and press F5.
' Results go to the Immediate window (Ctrl+G).
' ============================================================

Private failCount As Long
Private checkCount As Long

Public Sub TestWZTCRules()
    failCount = 0
    checkCount = 0

    Dim speeds(1 To 8) As Integer
    speeds(1) = 25: speeds(2) = 30: speeds(3) = 35: speeds(4) = 40
    speeds(5) = 45: speeds(6) = 50: speeds(7) = 55: speeds(8) = 65

    Dim lanes(1 To 3) As Integer
    lanes(1) = 10: lanes(2) = 11: lanes(3) = 12

    Dim shoulders(1 To 7) As String
    shoulders(1) = "<= 4 ft": shoulders(2) = "5-7 ft": shoulders(3) = "8 ft"
    shoulders(4) = "9 ft": shoulders(5) = "10 ft": shoulders(6) = "11 ft"
    shoulders(7) = "12 ft"

    Dim roadTypes(1 To 2) As String
    roadTypes(1) = "Freeway": roadTypes(2) = "Non-Freeway"

    Debug.Print "=== WZTCRules smoke test ==="

    Dim si As Integer, li As Integer, wi As Integer, ri As Integer
    Dim combos As Long
    combos = 0

    For si = 1 To 8
        For li = 1 To 3
            For wi = 1 To 7
                For ri = 1 To 2
                    combos = combos + 1
                    Call CheckOneCombo(speeds(si), lanes(li), shoulders(wi), roadTypes(ri))
                Next ri
            Next wi
        Next li
    Next si

    Debug.Print "combinations tested: " & combos
    Debug.Print "assertions:          " & checkCount
    Call TestDefaultItems

    If failCount = 0 Then
        Debug.Print "=== PASS - no failures ==="
    Else
        Debug.Print "=== FAIL - " & failCount & " failed assertion(s) ==="
    End If
End Sub

Private Sub CheckOneCombo(speed As Integer, laneWidth As Integer, _
                          shoulderKey As String, roadType As String)
    Dim sp As WZTCSpacing
    Dim tag As String
    tag = speed & "mph/" & laneWidth & "ft/" & shoulderKey & "/" & roadType

    sp = WZTCRules.ComputeSpacing(speed, laneWidth, shoulderKey, roadType)

    ' Totals must equal the sum of their parts
    Call Assert(sp.SkipTotal = sp.SkipMerge + sp.SkipShoulder + sp.SkipBuffer + sp.SkipRollAhead, _
                tag & " SkipTotal mismatch: " & sp.SkipTotal & " <> " & _
                sp.SkipMerge & "+" & sp.SkipShoulder & "+" & sp.SkipBuffer & "+" & sp.SkipRollAhead)
    Call Assert(sp.ChanTotal = sp.ChanMerge + sp.ChanShoulder, _
                tag & " ChanTotal mismatch: " & sp.ChanTotal & " <> " & _
                sp.ChanMerge & "+" & sp.ChanShoulder)

    ' Every distance in the valid matrix must be populated
    Call Assert(sp.DownstreamTaper > 0, tag & " DownstreamTaper = 0")
    Call Assert(sp.VehicleSpace > 0, tag & " VehicleSpace = 0")
    Call Assert(sp.BufferSpace > 0, tag & " BufferSpace = 0")
    Call Assert(sp.MergingTaper > 0, tag & " MergingTaper = 0")
    Call Assert(sp.ShoulderTaper > 0, tag & " ShoulderTaper = 0")
    Call Assert(sp.AdvanceWarningSpacing > 0, tag & " AdvanceWarningSpacing = 0")
    Call Assert(sp.RollAheadDistance > 0, tag & " RollAheadDistance = 0")

    ' Flare ratios are defined for every speed in the list, and drive upstream taper
    Call Assert(Len(sp.FlareBarrier) > 0, tag & " FlareBarrier empty")
    Call Assert(Len(sp.FlareBeam) > 0, tag & " FlareBeam empty")
    Call Assert(sp.UpTaperBarrier > 0, tag & " UpTaperBarrier = 0")
    Call Assert(sp.UpTaperBeam > 0, tag & " UpTaperBeam = 0")

    ' Road type is the only thing that changes downstream taper
    If LCase(roadType) = "non-freeway" Then
        Call Assert(sp.DownstreamTaper = 50, tag & " Non-Freeway DownstreamTaper <> 50")
    Else
        Call Assert(sp.DownstreamTaper = 100, tag & " Freeway DownstreamTaper <> 100")
    End If
End Sub

' ============================================================
' DEFAULT ITEM BUILDERS — row counts, labels, and the exact
' label strings SpacingBox.cls binds against
' ============================================================
Private Sub TestDefaultItems()
    Dim sp As WZTCSpacing
    sp = WZTCRules.ComputeSpacing(45, 12, "10 ft", "Non-Freeway")

    Dim t() As String, l() As String, s() As String
    Dim z() As String, d() As String, n As Integer

    Call WZTCRules.GetDefaultUpstreamItems(sp, t, l, s, z, d, n)
    Call Assert(n = 7, "GetDefaultUpstreamItems count <> 7 (got " & n & ")")
    Call Assert(l(1) = "Roll Ahead Distance", "upstream label 1 changed: " & l(1))
    Call Assert(l(2) = "Vehicle Space", "upstream label 2 changed: " & l(2))
    Call Assert(l(3) = "Buffer Space", "upstream label 3 changed: " & l(3))
    Call Assert(l(4) = "Merging/Shifting Taper", "upstream label 4 changed: " & l(4))
    Call Assert(l(5) = "Shoulder Taper", "upstream label 5 changed: " & l(5))
    Call Assert(l(6) = "Upstream Taper Temp Barrier", "upstream label 6 changed: " & l(6))
    Call Assert(l(7) = "Upstream Taper Box/Corr Beam", "upstream label 7 changed: " & l(7))

    Dim i As Integer
    For i = 1 To n
        Call Assert(t(i) = "Non-Sign", "upstream row " & i & " type <> Non-Sign")
        Call Assert(Len(s(i)) > 0, "upstream row " & i & " spacing empty")
    Next i

    Call WZTCRules.GetDefaultDownstreamItems(sp, t, l, s, z, d, n)
    Call Assert(n = 1, "GetDefaultDownstreamItems count <> 1 (got " & n & ")")
    Call Assert(l(1) = "Downstream Taper", "downstream label 1 changed: " & l(1))
    Call Assert(t(1) = "Non-Sign", "downstream row 1 type <> Non-Sign")
End Sub

Private Sub Assert(cond As Boolean, msg As String)
    checkCount = checkCount + 1
    If Not cond Then
        failCount = failCount + 1
        Debug.Print "FAIL: " & msg
    End If
End Sub
