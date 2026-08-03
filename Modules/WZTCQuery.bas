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

' ============================================================
' FIND REFERENCE LINEWORK — locates connected line/line-string chains
' on a given level, in the active model and/or attached references,
' for auto-tracing an alignment or work-space boundary without clicks
' (agent-driven-8-step-wizard plan, Component 2a). Returns one row
' per disconnected candidate chain found — caller picks, usually the
' longest/highest totalLengthFt is the intended roadway. Each row's
' verticesTSV is the SAME "x,y,z|x,y,z|..." format place_workspace/
' place_polyline/place_arc already use, so a chosen candidate feeds
' straight into alignment/workspace creation with no re-encoding.
'
' Confirmed live (feasibility check, 2026-08-02): ActiveModelReference
' .Attachments exposes attached references; .AsAttachment gives a
' Scan()-able interface plus GetReferenceToMasterTransform() for
' converting attachment-space geometry into master coordinates —
' attachments are NOT always unit-scale (one test reference had a
' real 10x scale factor), so this transform is applied unconditionally
' whenever the source is an attachment, never skipped as an
' optimization.
'
' Arc elements are NOT included — el.AsVertexList only covers
' line-based geometry (Line/LineString/Shape/line-based Complex
' Chain). A road that includes true arc segments will come back
' broken into separate candidates at each arc rather than one
' continuous chain. This is flagged, not silently produced as a
' subtly-wrong single chain: if nothing plausible comes back, the
' caller should fall back to click-based tracing rather than guess.
'
' levelNameContains: required, case-insensitive substring (the level
'   name the engineer names when asked "what level has the roadway
'   centerline?").
' includeReferences: OFF by default -- scans only the active model.
'   Reference-attachment scanning is fully built and was confirmed
'   working via COM (feasibility check), but attachment reads have
'   their own failure modes independent of this function's own logic
'   (see [[feedback-reference-scan-hang-and-com-errors]] -- both of
'   DELETE.dgn's attachments were transiently unavailable right after
'   a MicroStation restart). Kept as an opt-in switch rather than
'   deleted, so the already-verified mechanism doesn't need rebuilding
'   once reference-tracing is wanted again -- just pass True.
' refNameContains: optional, case-insensitive substring narrowing to
'   attachments by reference Name (only used when includeReferences=True).
' ============================================================
Public Function FindReferenceLinework(levelNameContains As String, _
                                      Optional includeReferences As Boolean = False, _
                                      Optional refNameContains As String = "") As String()
    Dim needle As String: needle = UCase(Trim(levelNameContains))
    Dim refNeedle As String: refNeedle = UCase(Trim(refNameContains))

    If needle = "" Then
        Dim errRows(0 To 1) As String
        errRows(0) = "error"
        errRows(1) = "levelNameContains is required"
        FindReferenceLinework = errRows
        Exit Function
    End If

    ' segments(i) = a Point3d() array in MASTER coordinates (already
    ' transformed if it came from an attachment); segSource(i) =
    ' "active" or the attachment's Name, for the returned report.
    Dim segments() As Variant
    Dim segSource() As String
    Dim segCount As Long: segCount = 0
    ReDim segments(0 To 255)
    ReDim segSource(0 To 255)

    Call CollectLevelSegments(ActiveModelReference, Nothing, needle, "active", _
                              segments, segSource, segCount)

    ' Reference-attachment scanning is opt-in (includeReferences=True) --
    ' see the function header comment. Skipped entirely by default.
    If includeReferences Then
        ' Each attachment is handled independently -- one unavailable/
        ' missing/still-resolving reference (e.g. right after a MicroStation
        ' restart) must not abort the whole call, same reasoning as the
        ' per-element guards in CollectLevelSegments below. Every COM touch
        ' in this section (getting the collection, indexing it, converting
        ' to Attachment, reading .Name) is wrapped as ONE unit per
        ' attachment rather than guarded line-by-line -- an earlier,
        ' line-by-line version of this still let an unguarded .Name read
        ' through and produced an uncaught "design file unavailable" error
        ' (confirmed live 2026-08-02, right after a MicroStation restart).
        On Error Resume Next
        Dim atts As Object
        Set atts = ActiveModelReference.Attachments
        Dim attCount As Long: attCount = 0
        If Err.Number = 0 And Not atts Is Nothing Then attCount = atts.Count
        Err.Clear
        On Error GoTo 0

        Dim i As Long
        For i = 1 To attCount
            On Error Resume Next
            Dim att As Object: Set att = Nothing
            Dim attName As String: attName = ""
            Set att = atts.Item(i).AsAttachment
            If Not att Is Nothing Then attName = att.Name
            Dim attOk As Boolean: attOk = (Err.Number = 0 And Not att Is Nothing And attName <> "")
            Err.Clear
            On Error GoTo 0

            If attOk Then
                If refNeedle = "" Or InStr(1, UCase(attName), refNeedle) > 0 Then
                    Call CollectLevelSegments(att, att, needle, attName, segments, segSource, segCount)
                End If
            End If
        Next i
    End If

    If segCount = 0 Then
        Dim noneRows(0 To 1) As String
        noneRows(0) = "error"
        noneRows(1) = "no line/line-string elements found on a level matching " & levelNameContains
        FindReferenceLinework = noneRows
        Exit Function
    End If

    ' Hard cap before chaining, not just a nicety: the chaining pass
    ' below is worst-case O(segCount^3) (every growth step rescans all
    ' remaining segments) and VBA runs single-threaded on the same
    ' thread as MicroStation's UI message pump -- an uncapped call
    ' against a level with hundreds of matching elements can block the
    ' whole application for minutes, which is exactly what happened
    ' live (2026-08-02: an unfiltered "ROADBED" scan against a real
    ' reference model hung MicroStation badly enough that Ctrl+Break
    ' didn't recover it and the engineer had to restart the app). This
    ' cap turns that into an honest "narrow your search" error instead,
    ' matching list_cells/list_line_styles' existing convention rather
    ' than silently attempting unbounded work.
    Const MAX_LINEWORK_SEGMENTS As Long = 80
    If segCount > MAX_LINEWORK_SEGMENTS Then
        Dim tooManyRows(0 To 1) As String
        tooManyRows(0) = "error"
        tooManyRows(1) = segCount & " elements matched level " & levelNameContains & _
            " -- more than " & MAX_LINEWORK_SEGMENTS & " is refused (chaining cost grows " & _
            "cubically). Narrow levelNameContains and/or pass refNameContains to scope to " & _
            "one attachment."
        FindReferenceLinework = tooManyRows
        Exit Function
    End If

    ' -- Greedy nearest-endpoint chaining. Duplicate points at a join
    ' (where two segments' endpoints coincide within JOIN_TOL) are left
    ' in rather than de-duplicated -- harmless for place_polyline/
    ' place_workspace input, not worth the extra complexity here. --
    Const JOIN_TOL As Double = 0.5   ' ft
    Dim used() As Boolean
    ReDim used(0 To segCount - 1)

    Dim chainOrder() As Long, chainReversed() As Boolean
    ReDim chainOrder(0 To segCount - 1)
    ReDim chainReversed(0 To segCount - 1)

    Dim resultRows() As String
    ReDim resultRows(0 To segCount)
    resultRows(0) = "chainIdx" & vbTab & "source" & vbTab & "segmentCount" & vbTab & _
                   "vertexCount" & vbTab & "totalLengthFt" & vbTab & "verticesTSV"
    Dim chainIdx As Long: chainIdx = 0

    Dim s As Long
    For s = 0 To segCount - 1
        If Not used(s) Then
            Dim chainLen As Long
            chainOrder(0) = s: chainReversed(0) = False
            used(s) = True
            chainLen = 1

            Dim grew As Boolean
            Do
                grew = False
                Dim headPt As Point3d, tailPt As Point3d
                headPt = SegPoint(segments(chainOrder(0)), chainReversed(0), True)
                tailPt = SegPoint(segments(chainOrder(chainLen - 1)), chainReversed(chainLen - 1), False)

                Dim t As Long
                For t = 0 To segCount - 1
                    If Not used(t) Then
                        Dim tStart As Point3d, tEnd As Point3d
                        tStart = SegPoint(segments(t), False, True)
                        tEnd = SegPoint(segments(t), False, False)

                        If PointsClose(tailPt, tStart, JOIN_TOL) Then
                            chainOrder(chainLen) = t: chainReversed(chainLen) = False
                            used(t) = True: chainLen = chainLen + 1: grew = True
                            Exit For
                        ElseIf PointsClose(tailPt, tEnd, JOIN_TOL) Then
                            chainOrder(chainLen) = t: chainReversed(chainLen) = True
                            used(t) = True: chainLen = chainLen + 1: grew = True
                            Exit For
                        ElseIf PointsClose(headPt, tEnd, JOIN_TOL) Then
                            Dim m As Long
                            For m = chainLen To 1 Step -1
                                chainOrder(m) = chainOrder(m - 1)
                                chainReversed(m) = chainReversed(m - 1)
                            Next m
                            chainOrder(0) = t: chainReversed(0) = False
                            used(t) = True: chainLen = chainLen + 1: grew = True
                            Exit For
                        ElseIf PointsClose(headPt, tStart, JOIN_TOL) Then
                            Dim m2 As Long
                            For m2 = chainLen To 1 Step -1
                                chainOrder(m2) = chainOrder(m2 - 1)
                                chainReversed(m2) = chainReversed(m2 - 1)
                            Next m2
                            chainOrder(0) = t: chainReversed(0) = True
                            used(t) = True: chainLen = chainLen + 1: grew = True
                            Exit For
                        End If
                    End If
                Next t
            Loop While grew

            Dim totalVerts As Long: totalVerts = 0
            Dim ci As Long
            For ci = 0 To chainLen - 1
                Dim segPts() As Point3d
                segPts = segments(chainOrder(ci))
                totalVerts = totalVerts + (UBound(segPts) - LBound(segPts) + 1)
            Next ci

            Dim combined() As Point3d
            ReDim combined(0 To totalVerts - 1)
            Dim vi As Long: vi = 0
            Dim totalLen As Double: totalLen = 0
            For ci = 0 To chainLen - 1
                Dim pts() As Point3d
                pts = segments(chainOrder(ci))
                Dim lo As Long, hi As Long, stp As Long, p As Long
                If chainReversed(ci) Then
                    lo = UBound(pts): hi = LBound(pts): stp = -1
                Else
                    lo = LBound(pts): hi = UBound(pts): stp = 1
                End If
                p = lo
                Do
                    combined(vi) = pts(p)
                    If vi > 0 Then
                        Dim ddx As Double, ddy As Double, ddz As Double
                        ddx = combined(vi).X - combined(vi - 1).X
                        ddy = combined(vi).Y - combined(vi - 1).Y
                        ddz = combined(vi).Z - combined(vi - 1).Z
                        totalLen = totalLen + Sqr(ddx * ddx + ddy * ddy + ddz * ddz)
                    End If
                    vi = vi + 1
                    If p = hi Then Exit Do
                    p = p + stp
                Loop
            Next ci

            Dim vparts() As String
            ReDim vparts(0 To totalVerts - 1)
            Dim vv As Long
            For vv = 0 To totalVerts - 1
                vparts(vv) = Format(combined(vv).X, "0.0####") & "," & _
                            Format(combined(vv).Y, "0.0####") & "," & _
                            Format(combined(vv).Z, "0.0####")
            Next vv

            chainIdx = chainIdx + 1
            resultRows(chainIdx) = chainIdx & vbTab & segSource(chainOrder(0)) & vbTab & _
                                   chainLen & vbTab & totalVerts & vbTab & _
                                   Format(totalLen, "0.0") & vbTab & Join(vparts, "|")
        End If
    Next s

    Dim finalRows() As String
    ReDim finalRows(0 To chainIdx)
    Dim fr As Long
    For fr = 0 To chainIdx
        finalRows(fr) = resultRows(fr)
    Next fr
    FindReferenceLinework = finalRows
End Function

Private Function SegPoint(segVariant As Variant, reversed As Boolean, wantStart As Boolean) As Point3d
    Dim pts() As Point3d
    pts = segVariant
    If (wantStart And Not reversed) Or (Not wantStart And reversed) Then
        SegPoint = pts(LBound(pts))
    Else
        SegPoint = pts(UBound(pts))
    End If
End Function

Private Function PointsClose(a As Point3d, b As Point3d, tol As Double) As Boolean
    Dim dx As Double, dy As Double, dz As Double
    dx = a.X - b.X: dy = a.Y - b.Y: dz = a.Z - b.Z
    PointsClose = (Sqr(dx * dx + dy * dy + dz * dz) <= tol)
End Function

' Scans one scannable ref (ActiveModelReference, or an .AsAttachment
' object) for graphical elements on a level whose name contains
' needle, and appends each one's vertex-list geometry (transformed to
' master coordinates when att is not Nothing) into segments()/
' segSource(), doubling their capacity as needed.
Private Sub CollectLevelSegments(scanRef As Object, att As Object, needle As String, sourceName As String, _
                                 ByRef segments() As Variant, ByRef segSource() As String, ByRef segCount As Long)
    Dim oScan As New ElementScanCriteria
    oScan.ExcludeNonGraphical

    Dim oEnum As ElementEnumerator
    On Error Resume Next
    Set oEnum = scanRef.Scan(oScan)
    If Err.Number <> 0 Or oEnum Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Sub   ' this scanRef (active model or one attachment) isn't scannable right now -- skip it, don't abort the caller
    End If
    Err.Clear
    On Error GoTo 0

    ' Transform3d's exact declared type name isn't relied on here --
    ' Variant natively supports dot-field access on an Automation-
    ' compatible record return, confirmed live against this exact
    ' method (feasibility check, 2026-08-02). Using a guessed type
    ' name would risk a hard compile break if wrong; Variant does not.
    Dim xf As Variant
    Dim hasTransform As Boolean: hasTransform = False
    If Not att Is Nothing Then
        On Error Resume Next
        xf = att.GetReferenceToMasterTransform()
        hasTransform = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
    End If

    Dim el As Element
    Do While oEnum.MoveNext
        Set el = oEnum.Current

        ' Guard every per-element read -- a reference model can hold
        ' element types (cells, text, proxies, etc.) whose .AsVertexList/
        ' .GetVertices() throw Error 430 ("does not support Automation")
        ' rather than returning Nothing the way .AsLineElement does
        ' elsewhere in this codebase (confirmed live 2026-08-02: a bare
        ' scan of a real reference model hit this). One unreadable
        ' element must not abort the whole scan -- skip it and continue,
        ' same spirit as SafeStr()/the lvlName read below.
        On Error Resume Next

        Dim lvlName As String: lvlName = ""
        lvlName = el.Level.Name
        If Err.Number <> 0 Then Err.Clear: GoTo NextRefEl

        If InStr(1, UCase(lvlName), needle) > 0 Then
            Dim vl As VertexList
            Set vl = el.AsVertexList
            If Err.Number <> 0 Then Err.Clear: GoTo NextRefEl
            If Not vl Is Nothing Then
                Dim pts() As Point3d
                pts = vl.GetVertices()
                If Err.Number <> 0 Then Err.Clear: GoTo NextRefEl
                If UBound(pts) > LBound(pts) Then
                    If hasTransform Then
                        Dim k As Long
                        For k = LBound(pts) To UBound(pts)
                            pts(k) = TransformPoint(pts(k), xf)
                            If Err.Number <> 0 Then Err.Clear: GoTo NextRefEl
                        Next k
                    End If
                    If segCount > UBound(segments) Then
                        ReDim Preserve segments(0 To (UBound(segments) + 1) * 2 - 1)
                        ReDim Preserve segSource(0 To (UBound(segSource) + 1) * 2 - 1)
                    End If
                    segments(segCount) = pts
                    segSource(segCount) = sourceName
                    segCount = segCount + 1
                End If
            End If
        End If
NextRefEl:
        On Error GoTo 0
    Loop
End Sub

' Applies a reference->master transform to one point: master = Row*ref
' + Translation per axis (confirmed live via COM, feasibility check
' 2026-08-02 -- one test reference had a genuine 10x scale, so this is
' never skipped as an optimization when att is a real attachment).
Private Function TransformPoint(p As Point3d, xf As Variant) As Point3d
    Dim outPt As Point3d
    outPt.X = xf.RowX.X * p.X + xf.RowX.Y * p.Y + xf.RowX.Z * p.Z + xf.TranslationX
    outPt.Y = xf.RowY.X * p.X + xf.RowY.Y * p.Y + xf.RowY.Z * p.Z + xf.TranslationY
    outPt.Z = xf.RowZ.X * p.X + xf.RowZ.Y * p.Y + xf.RowZ.Z * p.Z + xf.TranslationZ
    TransformPoint = outPt
End Function
