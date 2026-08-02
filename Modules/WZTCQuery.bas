Option Explicit

' ============================================================
' WZTC QUERY — READ-ONLY MODEL INTROSPECTION
' ------------------------------------------------------------
' Everything here is read-only: no element creation, no
' SharedState mutation, no CadInputQueue. Each function returns
' a String() array — element 1 is a tab-separated header row,
' subsequent elements are tab-separated data rows — so the
' bridge can write it straight to a results file, or a future
' VBA chat panel can format it directly.
'
' Element-proximity matching uses each element's Range bounding-
' box CENTER, not its true geometric centroid. A point near the
' end of a long line matches that line's midpoint, and large
' elements (sheet borders) can spuriously match. This is a known
' limitation, not a bug — flagged here and in every function that
' inherits it, per the plan (Layer 2 / WZTCQuery.bas).
' ============================================================

' ============================================================
' FIND ELEMENTS NEAR A POINT
' typeFilter: "" for all types, or one of CELL/LINE/ARC/SHAPE/
'             TEXT/TEXT_NODE/OTHER (case-insensitive)
' Returns all candidates within radius, closest first — not just
' the nearest one, so a caller can detect ambiguity.
' ============================================================
Public Function FindElementsNear(x As Double, y As Double, radius As Double, _
                                 typeFilter As String) As String()
    Dim rows() As String
    ReDim rows(0 To 0)
    rows(0) = "elementId" & vbTab & "type" & vbTab & "level" & vbTab & "cellName" & vbTab & _
              "cx" & vbTab & "cy" & vbTab & "distanceFt" & vbTab & _
              "rangeLowX" & vbTab & "rangeLowY" & vbTab & "rangeHighX" & vbTab & "rangeHighY"

    Dim wantType As String: wantType = UCase(Trim(typeFilter))

    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical

    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    ' Collect matches via Collection.Add, not ReDim Preserve in a loop --
    ' the same repeated-resize pattern caused an unexplained Subscript-
    ' out-of-range once 2+ items accumulated in WZTCExec.FindInteriorPoint.
    ' This was never actually exercised with 2+ matches before (every prior
    ' test only ever found exactly 1), so it was an untested latent risk,
    ' not a confirmed-safe pattern. Collections grow with no resizing at all.
    Dim matchEls As New Collection
    Dim matchDists As New Collection

    Dim el As Element
    Do While oEnum.MoveNext
        Set el = oEnum.Current

        Dim tName As String: tName = ElementTypeName(el.Type)
        If wantType <> "" And wantType <> tName Then GoTo NextEl

        On Error Resume Next
        Dim lo As Point3d, hi As Point3d
        lo = el.Range.Low: hi = el.Range.High
        If Err.Number <> 0 Then Err.Clear: GoTo NextEl
        On Error GoTo 0

        Dim cx As Double, cy As Double
        cx = (lo.X + hi.X) / 2: cy = (lo.Y + hi.Y) / 2
        Dim dist As Double
        dist = Sqr((cx - x) ^ 2 + (cy - y) ^ 2)

        If dist <= radius Then
            matchEls.Add el
            matchDists.Add dist
        End If
NextEl:
    Loop

    Dim matchCount As Integer: matchCount = matchEls.Count

    ' Copy into plain arrays sized exactly once (a single ReDim, not
    ' ReDim Preserve in a loop) for the sort step below.
    Dim matchEl() As Element
    Dim matchDist() As Double
    If matchCount > 0 Then
        ReDim matchEl(1 To matchCount)
        ReDim matchDist(1 To matchCount)
        Dim idx As Integer
        For idx = 1 To matchCount
            Set matchEl(idx) = matchEls(idx)
            matchDist(idx) = matchDists(idx)
        Next idx
    End If

    ' Insertion sort by distance ascending (matchCount is small — proximity radius)
    Dim i As Integer, j As Integer
    For i = 2 To matchCount
        Dim tmpD As Double: tmpD = matchDist(i)
        Dim tmpE As Element: Set tmpE = matchEl(i)
        j = i - 1
        ' Do While j >= 1 And matchDist(j) > tmpD -- VBA's And does NOT
        ' short-circuit, so once j reaches 0 the loop condition still
        ' evaluates matchDist(j) = matchDist(0), out of bounds for an
        ' array declared (1 To matchCount). Confirmed live (2026-08-01)
        ' as the actual "Subscript out of range" cause via checkpoint
        ' instrumentation: only triggers when an element sorts all the
        ' way to the front (needs j to reach 0), which no test before
        ' today happened to exercise. Split into an unconditional `j >= 1`
        ' loop guard plus an explicit Exit Do so matchDist(j) is only
        ' evaluated when j is already known to be in bounds.
        Do While j >= 1
            If matchDist(j) <= tmpD Then Exit Do
            matchDist(j + 1) = matchDist(j)
            Set matchEl(j + 1) = matchEl(j)
            j = j - 1
        Loop
        matchDist(j + 1) = tmpD
        Set matchEl(j + 1) = tmpE
    Next i

    ' Preserve, not a plain ReDim -- rows(0) already holds the header set
    ' at the top of this function. A plain ReDim here wiped it back to an
    ' empty string every time matchCount > 0, and every downstream
    ' Split(rows(0), vbTab) then produced a 1-element array instead of the
    ' 11 expected header columns -- confirmed live as the actual cause of
    ' "Subscript out of range" (2026-08-01): this path had never been
    ' exercised with any real match before today, since every prior
    ' find_elements_near call in this project's testing history queried an
    ' area with nothing in it.
    If matchCount > 0 Then ReDim Preserve rows(0 To matchCount)
    For i = 1 To matchCount
        Set el = matchEl(i)
        Dim lo2 As Point3d, hi2 As Point3d
        ' Same known-flaky Range read as the first loop above (already
        ' guarded there). Defensive: skip just this one row (left blank
        ' in rows()) rather than losing every other match if a second
        ' read of the same element's Range ever fails after the scan
        ' enumerator has moved on.
        On Error Resume Next
        lo2 = el.Range.Low: hi2 = el.Range.High
        If Err.Number <> 0 Then Err.Clear: GoTo NextRow
        On Error GoTo 0
        Dim cx2 As Double, cy2 As Double
        cx2 = (lo2.X + hi2.X) / 2: cy2 = (lo2.Y + hi2.Y) / 2

        Dim lvlName As String
        On Error Resume Next
        lvlName = el.Level.Name
        If Err.Number <> 0 Then lvlName = "": Err.Clear
        On Error GoTo 0

        Dim cellName As String: cellName = ""
        If el.Type = msdElementTypeCellHeader Then
            On Error Resume Next
            Dim ce As CellElement: Set ce = el
            cellName = ce.Name
            On Error GoTo 0
        End If

        rows(i) = CStr(ElIDAsDouble(el.ID)) & vbTab & ElementTypeName(el.Type) & vbTab & _
                  lvlName & vbTab & cellName & vbTab & _
                  Format(cx2, "0.00") & vbTab & Format(cy2, "0.00") & vbTab & Format(matchDist(i), "0.00") & vbTab & _
                  Format(lo2.X, "0.00") & vbTab & Format(lo2.Y, "0.00") & vbTab & _
                  Format(hi2.X, "0.00") & vbTab & Format(hi2.Y, "0.00")
NextRow:
    Next i

    FindElementsNear = rows
End Function

' ============================================================
' STATION -> POINT/TANGENT ON A COMMITTED ALIGNMENT
' Reuses PerpPlacement's arc-length path engine (same geometry
' PlacePerp already uses) rather than reimplementing it.
' Returns a 2-row array: header + one data row. Returns a single
' "error" row (no header) if the alignment isn't committed.
' ============================================================
Public Function StationToPoint(alignIdx As Integer, sta As Double) As String()
    Dim rows() As String
    Dim errMsg As String
    If Not AlignmentIsReady(alignIdx, errMsg) Then
        ReDim rows(0 To 1)
        rows(0) = "error"
        rows(1) = errMsg
        StationToPoint = rows
        Exit Function
    End If

    Call PerpPlacement.BuildAlignmentPath(alignIdx)

    Dim ptX As Double, ptY As Double, ptZ As Double, tanX As Double, tanY As Double
    Dim ok As Boolean
    ok = PerpPlacement.GetPointAndTangent(sta, ptX, ptY, ptZ, tanX, tanY)

    ReDim rows(0 To 1)
    If Not ok Then
        rows(0) = "error"
        rows(1) = "GetPointAndTangent failed for alignment " & alignIdx & " at station " & sta
    Else
        ' GetPointAndTangent silently CLAMPS an out-of-range station to the
        ' nearest path end rather than failing (PerpPlacement.bas:383-384) --
        ' confirmed live: requesting a station past the end of a real drawn
        ' alignment returned the endpoint with no indication anything was
        ' off. That's a real placement-accuracy risk (a caller could place a
        ' sign well short of where it actually belongs without knowing it),
        ' so surface it explicitly rather than making every caller remember
        ' to diff sta against totalPathLenFt themselves.
        Dim totalLen As Double: totalLen = PerpPlacement.GetTotalPathLength()
        Dim wasClamped As Boolean: wasClamped = (sta < 0 Or sta > totalLen)

        rows(0) = "ptX" & vbTab & "ptY" & vbTab & "ptZ" & vbTab & "tanX" & vbTab & "tanY" & vbTab & _
                  "totalPathLenFt" & vbTab & "clamped"
        rows(1) = Format(ptX, "0.00") & vbTab & Format(ptY, "0.00") & vbTab & Format(ptZ, "0.00") & vbTab & _
                  Format(tanX, "0.0000") & vbTab & Format(tanY, "0.0000") & vbTab & _
                  Format(totalLen, "0.00") & vbTab & CStr(wasClamped)
    End If
    StationToPoint = rows
End Function

' ============================================================
' PREDICTED STATIONING FOR EVERY CONFIGURED ITEM ON AN ALIGNMENT
' Sums configured spacings in table order — the same accumulation
' PlacePerp.PlaceLineForCurrentItem performs — so this reports
' where each item WOULD land if placed, without requiring
' PlacePerp to have actually run yet.
' ============================================================
Public Function GetAlignmentStationing(alignIdx As Integer) As String()
    Dim rows() As String
    Dim errMsg As String
    If Not AlignmentIsReady(alignIdx, errMsg) Then
        ReDim rows(0 To 1)
        rows(0) = "error"
        rows(1) = errMsg
        GetAlignmentStationing = rows
        Exit Function
    End If

    Dim rCount As Integer: rCount = wztcAlignRowCounts(alignIdx)
    ReDim rows(0 To rCount)
    rows(0) = "rowIndex" & vbTab & "type" & vbTab & "label" & vbTab & "spacingFt" & vbTab & "cumulativeStationFt"

    Dim cum As Double: cum = 0
    Dim r As Integer
    For r = 1 To rCount
        Dim spacing As Double: spacing = Val(wztcAlignRowSpacings(alignIdx, r))
        cum = cum + spacing
        rows(r) = r & vbTab & wztcAlignRowTypes(alignIdx, r) & vbTab & wztcAlignRowLabels(alignIdx, r) & vbTab & _
                  Format(spacing, "0.0") & vbTab & Format(cum, "0.0")
    Next r

    GetAlignmentStationing = rows
End Function

' ============================================================
' LIST ALL LEVELS IN THE ACTIVE DESIGN FILE
' ============================================================
' NOTE: lvl.Color is not a valid direct property (confirmed by compile
' error) — MicroStation VBA likely nests level symbology under something
' like lvl.ByLevelSymbology.Color rather than exposing it flat. Dropped
' rather than guessed a second time; level name/number/display state
' matter more to the agent than exact color. lvl.Number and lvl.IsDisplayed
' are still unconfirmed — if either fails next, same drill: tell me the
' exact highlighted token.
Public Function ListLevels() As String()
    Dim rows() As String
    Dim n As Integer: n = 0
    ReDim rows(0 To 0)
    rows(0) = "name" & vbTab & "number" & vbTab & "isDisplayed"

    Dim lvl As Level
    For Each lvl In ActiveDesignFile.Levels
        n = n + 1
        ReDim Preserve rows(0 To n)
        rows(n) = lvl.Name & vbTab & lvl.Number & vbTab & CStr(lvl.IsDisplayed)
    Next lvl

    ListLevels = rows
End Function

' ============================================================
' CLASSIFY SITE FEATURES NEAR A POINT
' Provisional heuristic: keyword match against level name and
' cell name. Placeholder until the M4 sheet-registry data file
' exists as the real level/cell -> category mapping — extend
' KEYWORD_MAP() below in the meantime, no logic changes needed.
' Anything that matches nothing comes back "unclassified" with
' its geometry intact, so a caller can still flag a raw
' obstruction it can't name.
' ============================================================
Public Function ClassifySiteFeatures(x As Double, y As Double, radius As Double) As String()
    Dim base() As String
    base = FindElementsNear(x, y, radius, "")

    Dim rows() As String
    Dim n As Integer: n = UBound(base)
    ReDim rows(0 To n)
    rows(0) = base(0) & vbTab & "classification"

    Dim i As Integer
    For i = 1 To n
        Dim fields() As String: fields = Split(base(i), vbTab)
        ' fields(2)=level, fields(3)=cellName
        Dim lvl As String: lvl = ""
        Dim cn As String: cn = ""
        If UBound(fields) >= 2 Then lvl = fields(2)
        If UBound(fields) >= 3 Then cn = fields(3)
        rows(i) = base(i) & vbTab & ClassifyByKeyword(lvl & " " & cn)
    Next i

    ClassifySiteFeatures = rows
End Function

Private Function ClassifyByKeyword(text As String) As String
    Dim t As String: t = UCase(text)
    Dim keywords() As String
    Dim labels() As String
    keywords = Split("POLE,UTIL,HYDRANT,TREE,DRAIN,CB,MH,GUARDRAIL,STRUCT", ",")
    labels = Split("Utility Pole,Utility,Hydrant,Tree,Drainage,Catch Basin,Manhole,Guardrail,Structure", ",")

    Dim i As Integer
    For i = 0 To UBound(keywords)
        If InStr(t, keywords(i)) > 0 Then
            ClassifyByKeyword = labels(i)
            Exit Function
        End If
    Next i
    ClassifyByKeyword = "unclassified"
End Function

' ============================================================
' HELPERS
' ============================================================
Public Function AlignmentIsReady(alignIdx As Integer, ByRef errMsg As String) As Boolean
    If alignIdx < 1 Or alignIdx > 10 Then
        errMsg = "alignIdx out of range: " & alignIdx
        AlignmentIsReady = False
        Exit Function
    End If
    If Not wztcAlignDrawn(alignIdx) Then
        errMsg = "alignment " & alignIdx & " has not been committed (AlignDraw step)"
        AlignmentIsReady = False
        Exit Function
    End If
    AlignmentIsReady = True
End Function

' NOTE: msdElementTypeCellHeader and msdElementTypeShape are NOT confirmed
' against a working reference in this codebase (unlike Line/Arc/Text/TextNode,
' which are already used in PerpPlacement.bas and the Bluebeam importer).
' If this module fails to compile, check these two names first.
Private Function ElementTypeName(t As MsdElementType) As String
    Select Case t
        Case msdElementTypeCellHeader:      ElementTypeName = "CELL"
        Case msdElementTypeLine:      ElementTypeName = "LINE"
        Case msdElementTypeArc:       ElementTypeName = "ARC"
        Case msdElementTypeShape:     ElementTypeName = "SHAPE"
        Case msdElementTypeText:      ElementTypeName = "TEXT"
        Case msdElementTypeTextNode:  ElementTypeName = "TEXT_NODE"
        Case Else:                    ElementTypeName = "OTHER"
    End Select
End Function

' ============================================================
' DESCRIBE DRAWING STATE — model/units/scale/level/symbology/ACS/
' view/reference/selection context, gathered once so an agent can
' orient itself before making any edits (rather than assuming feet,
' assuming annotation scale 1, assuming nothing is selected, etc.).
' Added 2026-08-02 after the sign-scale investigation showed a wrong
' assumption here (a fixed scale) causes silently-wrong placements.
' Each row is key<TAB>value; a value left blank means that property
' could not be read on this MicroStation version/session (guarded
' with On Error Resume Next per item so one missing property doesn't
' blank out the whole report).
' ============================================================
Public Function DescribeDrawingState() As String()
    Dim rows() As String
    Dim n As Integer: n = 0
    ReDim rows(0 To 200)
    rows(0) = "key" & vbTab & "value"
    n = 0

    Dim mr As ModelReference
    Set mr = ActiveModelReference

    ' --- Active model ---
    n = n + 1: rows(n) = "modelName" & vbTab & SafeStr(mr.Name)
    n = n + 1: rows(n) = "is3D" & vbTab & SafeStr(mr.Is3D)
    n = n + 1: rows(n) = "modelType" & vbTab & SafeStr(mr.Type)   ' raw MsdModelType enum value

    ' --- Units / resolution ---
    n = n + 1: rows(n) = "masterUnitLabel" & vbTab & SafeStr(mr.MasterUnit.Label)
    n = n + 1: rows(n) = "subUnitLabel" & vbTab & SafeStr(mr.SubUnit.Label)
    n = n + 1: rows(n) = "subUnitsPerMasterUnit" & vbTab & SafeStr(mr.SubUnitsPerMasterUnit)
    n = n + 1: rows(n) = "uorsPerMasterUnit" & vbTab & SafeStr(mr.UORsPerMasterUnit)

    ' --- Annotation scale (see 2026-08-02 sign-scale fix -- signs and
    ' other Annotation-class cells are auto-multiplied by this) ---
    On Error Resume Next
    Dim sd As SheetDefinition
    Set sd = mr.GetSheetDefinition
    On Error GoTo 0
    If Not sd Is Nothing Then
        n = n + 1: rows(n) = "annotationScaleFactor" & vbTab & SafeStr(sd.AnnotationScaleFactor)
        n = n + 1: rows(n) = "annotationScaleEnabled" & vbTab & SafeStr(sd.IsEnabled)
    Else
        n = n + 1: rows(n) = "annotationScaleFactor" & vbTab & ""
        n = n + 1: rows(n) = "annotationScaleEnabled" & vbTab & "False"
    End If

    ' --- Active level / symbology ---
    On Error Resume Next
    n = n + 1: rows(n) = "activeLevelName" & vbTab & SafeStr(ActiveSettings.Level.Name)
    n = n + 1: rows(n) = "activeLevelNumber" & vbTab & SafeStr(ActiveSettings.Level.Number)
    n = n + 1: rows(n) = "activeColor" & vbTab & SafeStr(ActiveSettings.Color)
    n = n + 1: rows(n) = "activeLineWeight" & vbTab & SafeStr(ActiveSettings.LineWeight)
    n = n + 1: rows(n) = "activeLineStyleName" & vbTab & SafeStr(ActiveSettings.LineStyleName)
    On Error GoTo 0

    ' --- Active ACS ---
    On Error Resume Next
    n = n + 1: rows(n) = "acsDefined" & vbTab & SafeStr(Application.ACSManager.IsDefined)
    n = n + 1: rows(n) = "acsName" & vbTab & SafeStr(Application.ACSManager.Name)
    On Error GoTo 0

    ' --- View info (active/open views, center, rotation) ---
    Dim v As View
    Dim i As Integer
    For i = 1 To 8
        On Error Resume Next
        Set v = ActiveDesignFile.Views(i)
        On Error GoTo 0
        If Not v Is Nothing Then
            If v.IsOpen Then
                n = n + 1: rows(n) = "view" & i & "Open" & vbTab & "True"
                n = n + 1: rows(n) = "view" & i & "Active" & vbTab & SafeStr(v.IsSelected)
                On Error Resume Next
                n = n + 1: rows(n) = "view" & i & "CenterX" & vbTab & SafeStr(v.Center.X)
                n = n + 1: rows(n) = "view" & i & "CenterY" & vbTab & SafeStr(v.Center.Y)
                n = n + 1: rows(n) = "view" & i & "RotationDeg" & vbTab & SafeStr(DrawSign.ViewRotationAngleDegrees(v))
                On Error GoTo 0
            End If
        End If
        Set v = Nothing
    Next i

    ' --- Reference attachments ---
    On Error Resume Next
    n = n + 1: rows(n) = "referenceAttachmentCount" & vbTab & SafeStr(mr.Attachments.Count)
    On Error GoTo 0

    ' --- Selected elements ---
    On Error Resume Next
    Dim selCount As Long: selCount = 0
    If mr.AnyElementsSelected Then
        Dim oSelEnum As ElementEnumerator
        Set oSelEnum = mr.GetSelectedElements()
        Do While oSelEnum.MoveNext
            selCount = selCount + 1
        Loop
    End If
    On Error GoTo 0
    n = n + 1: rows(n) = "selectedElementCount" & vbTab & CStr(selCount)

    ' --- File metadata ---
    On Error Resume Next
    n = n + 1: rows(n) = "fileName" & vbTab & SafeStr(ActiveDesignFile.Name)
    n = n + 1: rows(n) = "filePath" & vbTab & SafeStr(ActiveDesignFile.Path)
    On Error GoTo 0

    ReDim Preserve rows(0 To n)
    DescribeDrawingState = rows
End Function

' Converts any COM property read to a display string, swallowing
' errors so one unreadable property (e.g. no GCS on this file) blanks
' just that one row instead of aborting the whole report.
Private Function SafeStr(v As Variant) As String
    On Error GoTo Fail
    SafeStr = CStr(v)
    Exit Function
Fail:
    SafeStr = ""
End Function
