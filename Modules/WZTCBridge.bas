Option Explicit

' ============================================================
' WZTC BRIDGE — M1 TRANSPORT PROOF
' ------------------------------------------------------------
' File-based request/response protocol so an external process
' can trigger MicroStation drawing operations without any
' CadInputQueue GetInput / user click.
'
' Protocol: tab-separated lines (see Bridge/README.md).
'   Request : reqId<TAB>OP_TYPE<TAB>key1=val1<TAB>key2=val2...
'   Response: reqId<TAB>OK|ERROR<TAB>key1=val1...
'
' M1 supports exactly one op, PLACE_CELL, deliberately kept this
' small — the goal is proving the transport (file trigger -> VBA
' executes -> file response, with a journal entry) before the
' full WZTCExec / WZTCQuery layers are built on top of it.
'
' M2-M6 have since layered query/compute/draw/handoff/undo and
' the command-registry edit vocabulary onto the same ExecuteOp
' dispatcher. PLACE_CELL remains the simplest end-to-end proof.
'
' Trigger from MicroStation Key-in: VBA RUN [ProjectName]WZTCBridge.RunRequest
' Trigger externally: an automation client sends the same keyin over COM.
' ============================================================

Private Const BRIDGE_DIR As String = "c:\repos\microstation-vba-project\Bridge\"
Private Const REQUEST_FILE As String = BRIDGE_DIR & "request.tsv"
Private Const RESPONSE_FILE As String = BRIDGE_DIR & "response.tsv"
Private Const JOURNAL_FILE As String = BRIDGE_DIR & "wztc-journal.tsv"
Private Const HANDOFF_FILE As String = BRIDGE_DIR & "deferred-handoffs.tsv"
Private Const CAPTURES_DIR As String = BRIDGE_DIR & "captures\"

' M7 (Stage 3) -- separate file pair for the chat-driver process
' (mcp-server/chat_driver.py), so it can hold its own COM connection to
' MicroStation alongside an existing Claude Code MCP session without both
' racing on the same request.tsv/response.tsv (each process's Python-side
' reqId counter independently starts at P1 -- see bridge_client.py). Shares
' ExecuteOp / the journal / everything else; only the request/response
' files differ.
Private Const CHAT_TOOL_REQUEST_FILE As String = BRIDGE_DIR & "chat-tool-request.tsv"
Private Const CHAT_TOOL_RESPONSE_FILE As String = BRIDGE_DIR & "chat-tool-response.tsv"

Private Const WZTC_CELL_LIB As String = "c:\pwworking\usny\d0119091\ny_plan_wztc.cel"

' Bounded pool size for results_slot<N>.tsv reuse -- see the comment above
' WriteResultRows for why. Module-level Const declarations must all live
' here in the General Declarations section; VBA raises "Only comments may
' appear after End Sub, End Function, or End Property" if a Const is
' declared between two procedures instead (confirmed live 2026-08-02).
Private Const RESULT_POOL_SIZE As Long = 8

' Bounded rotation for wztc-journal.tsv -- see RotateJournalIfOversized
' for the full rationale. Must live here (General Declarations), not
' next to RotateJournalIfOversized -- same rule as RESULT_POOL_SIZE
' above, learned the hard way just now.
Private Const JOURNAL_MAX_BYTES As Long = 2000000
Private Const JOURNAL_KEEP_LINES As Integer = 2000

' ============================================================
' ENTRY POINT — reads request.tsv, executes each line in order,
' writes response.tsv, appends every op to the journal.
' ============================================================
Public Sub RunRequest()
    Dim reqLines() As String
    Dim n As Integer
    n = ReadAllLines(REQUEST_FILE, reqLines)

    If n = 0 Then
        Call WriteFile(RESPONSE_FILE, "" & vbTab & "ERROR" & vbTab & "note=no request found or file empty")
        Exit Sub
    End If

    Dim responses() As String
    ReDim responses(1 To n)

    Dim i As Integer
    For i = 1 To n
        responses(i) = ExecuteOp(reqLines(i))
    Next i

    Call WriteLines(RESPONSE_FILE, responses, n)
End Sub

' ============================================================
' CHAT-TOOL ENTRY POINT (M7) -- identical to RunRequest, pointed at
' the chat-driver's own file pair instead. Trigger:
'   VBA RUN [ProjectName]WZTCBridge.RunChatToolRequest
' Every op still goes through the shared ExecuteOp dispatcher and the
' shared JOURNAL_FILE, so the audit trail is identical in shape
' regardless of which front end (this, or the file+keyin path any MCP
' client uses) issued the op.
' ============================================================
Public Sub RunChatToolRequest()
    Dim reqLines() As String
    Dim n As Integer
    n = ReadAllLines(CHAT_TOOL_REQUEST_FILE, reqLines)

    If n = 0 Then
        Call WriteFile(CHAT_TOOL_RESPONSE_FILE, "" & vbTab & "ERROR" & vbTab & "note=no request found or file empty")
        Exit Sub
    End If

    Dim responses() As String
    ReDim responses(1 To n)

    Dim i As Integer
    For i = 1 To n
        responses(i) = ExecuteOp(reqLines(i))
    Next i

    Call WriteLines(CHAT_TOOL_RESPONSE_FILE, responses, n)
End Sub

' ============================================================
' EXECUTE A SINGLE OP LINE
' Public so a future in-MicroStation chat panel can call this
' directly, bypassing the request file entirely. Journals every
' call right here (not in RunRequest) specifically so that a direct
' caller still produces an audit trail -- the whole point of the
' journal is to be reliable regardless of front end, and it was NOT
' being written at all when this was called outside RunRequest
' (confirmed by DebugAgentLoopTest.bas: UNDO_LAST_OP was scanning
' stale entries from unrelated past sessions instead of the op that
' had just run).
' ============================================================
Public Function ExecuteOp(opLine As String) As String
    Dim result As String
    result = ExecuteOpInner(opLine)
    Call AppendJournal(opLine, result)
    ExecuteOp = result
End Function

Private Function ExecuteOpInner(opLine As String) As String
    Dim reqId As String
    reqId = ""
    On Error GoTo OpError

    Dim parts() As String
    parts = Split(opLine, vbTab)
    If UBound(parts) < 1 Then
        ExecuteOpInner = "" & vbTab & "ERROR" & vbTab & "note=malformed line (need reqId<TAB>OPTYPE)"
        Exit Function
    End If

    Dim opType As String
    reqId = Trim(parts(0))
    opType = Trim(parts(1))

    Dim params As Object
    Set params = ParseParams(parts)

    Select Case opType
        Case "PLACE_CELL"
            ExecuteOpInner = ExecPlaceCell(reqId, params)
        Case "FIND_ELEMENTS_NEAR"
            ExecuteOpInner = ExecFindElementsNear(reqId, params)
        Case "STATION_TO_POINT"
            ExecuteOpInner = ExecStationToPoint(reqId, params)
        Case "GET_ALIGNMENT_STATIONING"
            ExecuteOpInner = ExecGetAlignmentStationing(reqId, params)
        Case "LIST_LEVELS"
            ExecuteOpInner = ExecListLevels(reqId, params)
        Case "DESCRIBE_DRAWING_STATE"
            ExecuteOpInner = ExecDescribeDrawingState(reqId, params)
        Case "CLASSIFY_SITE_FEATURES"
            ExecuteOpInner = ExecClassifySiteFeatures(reqId, params)
        Case "COMPUTE_SPACING"
            ExecuteOpInner = ExecComputeSpacing(reqId, params)
        Case "PLACE_PERP_LINE"
            ExecuteOpInner = BridgePlacePerpLine(reqId, params)
        Case "PLACE_SIGN"
            ExecuteOpInner = BridgePlaceSign(reqId, params)
        Case "PLACE_ELEMENT_RUN"
            ExecuteOpInner = BridgePlaceElementRun(reqId, params)
        Case "PLACE_WORKSPACE"
            ExecuteOpInner = BridgePlaceWorkspace(reqId, params)
        Case "SET_SIGN_ATTRIBUTES"
            ExecuteOpInner = BridgeSetSignAttributes(reqId, params)
        Case "GET_SHEET_REQUIREMENTS"
            ExecuteOpInner = ExecGetSheetRequirements(reqId, params)
        Case "HANDOFF"
            ExecuteOpInner = BridgeHandoff(reqId, params)
        Case "UNDO_LAST_OP"
            ExecuteOpInner = ExecUndoLastOp(reqId, params)
        Case "GET_JOURNAL"
            ExecuteOpInner = ExecGetJournal(reqId, params)
        Case "LIST_DEFERRED_HANDOFFS"
            ExecuteOpInner = ExecListDeferredHandoffs(reqId, params)
        Case "LIST_REGISTRY_COMMANDS"
            ExecuteOpInner = ExecListRegistryCommands(reqId, params)
        Case "DESCRIBE_REGISTRY_COMMAND"
            ExecuteOpInner = ExecDescribeRegistryCommand(reqId, params)
        Case "RUN_REGISTRY_COMMAND"
            ExecuteOpInner = ExecRunRegistryCommand(reqId, params, False)
        Case "TEST_REGISTRY_COMMAND"
            ' Promotion-only: bypasses needs-testing gate for exactly one
            ' manual IDE run. Never exposed in mcp-server/server.py.
            ExecuteOpInner = ExecRunRegistryCommand(reqId, params, True)
        Case "MOVE_ELEMENT"
            ExecuteOpInner = BridgeMoveElement(reqId, params)
        Case "COPY_ELEMENT"
            ExecuteOpInner = BridgeCopyElement(reqId, params)
        Case "ROTATE_ELEMENT"
            ExecuteOpInner = BridgeRotateElement(reqId, params)
        Case "SCALE_ELEMENT"
            ExecuteOpInner = BridgeScaleElement(reqId, params)
        Case "MIRROR_ELEMENT"
            ExecuteOpInner = BridgeMirrorElement(reqId, params)
        Case "ARRAY_ELEMENT"
            ExecuteOpInner = BridgeArrayElement(reqId, params)
        Case "CHANGE_ELEMENT_LEVEL"
            ExecuteOpInner = BridgeChangeElementLevel(reqId, params)
        Case "EDIT_TEXT_ELEMENT"
            ExecuteOpInner = BridgeEditTextElement(reqId, params)
        Case "DELETE_ELEMENT"
            ExecuteOpInner = BridgeDeleteElement(reqId, params)
        Case "HATCH_ELEMENT"
            ExecuteOpInner = BridgeHatchElement(reqId, params)
        Case "PLACE_ARC"
            ExecuteOpInner = BridgePlaceArc(reqId, params)
        Case "PLACE_TEXT_LABEL"
            ExecuteOpInner = BridgePlaceTextLabel(reqId, params)
        Case "PLACE_CIRCLE"
            ExecuteOpInner = BridgeGeomPlaceCircle(reqId, params)
        Case "PLACE_ELLIPSE"
            ExecuteOpInner = BridgeGeomPlaceEllipse(reqId, params)
        Case "PLACE_BLOCK"
            ExecuteOpInner = BridgeGeomPlaceBlock(reqId, params)
        Case "PLACE_POLYLINE"
            ExecuteOpInner = BridgeGeomPlacePolyline(reqId, params)
        Case "PLACE_POLYGON"
            ExecuteOpInner = BridgeGeomPlacePolygon(reqId, params)
        Case "CHANGE_ELEMENT_SYMBOLOGY"
            ExecuteOpInner = BridgeGeomChangeSymbology(reqId, params)
        Case "COPY_PARALLEL"
            ExecuteOpInner = BridgeGeomCopyParallel(reqId, params)
        Case "CROSSHATCH_ELEMENT"
            ExecuteOpInner = BridgeGeomCrossHatch(reqId, params)
        Case "REMOVE_HATCH"
            ExecuteOpInner = BridgeGeomRemoveHatch(reqId, params)
        Case "BREAK_LINE"
            ExecuteOpInner = BridgeGeomBreakLine(reqId, params)
        Case "EXTEND_LINE"
            ExecuteOpInner = BridgeGeomExtendLine(reqId, params)
        Case "FILLET_ELEMENTS"
            ExecuteOpInner = BridgeGeomFillet(reqId, params)
        Case "CREATE_COMPLEX_STRING"
            ExecuteOpInner = BridgeGeomComplexString(reqId, params)
        Case "PLACE_FENCE_BLOCK"
            ExecuteOpInner = BridgeGeomPlaceFence(reqId, params)
        Case "FENCE_UNDEFINE"
            ExecuteOpInner = BridgeGeomFenceUndefine(reqId, params)
        Case "FENCE_COPY_CONTENTS"
            ExecuteOpInner = BridgeGeomFenceCopy(reqId, params)
        Case "FENCE_MOVE_CONTENTS"
            ExecuteOpInner = BridgeGeomFenceMove(reqId, params)
        Case "FENCE_DELETE_CONTENTS"
            ExecuteOpInner = BridgeGeomFenceDelete(reqId, params)
        Case "SELECT_ELEMENT"
            ExecuteOpInner = BridgeGeomSelect(reqId, params)
        Case "CLEAR_SELECTION"
            ExecuteOpInner = BridgeGeomClearSelection(reqId, params)
        Case "CAPTURE_VIEW"
            ExecuteOpInner = ExecCaptureView(reqId, params)
        Case Else
            ExecuteOpInner = reqId & vbTab & "ERROR" & vbTab & "note=unknown op type: " & opType
    End Select
    Exit Function

OpError:
    ExecuteOpInner = reqId & vbTab & "ERROR" & vbTab & "note=runtime error: " & Err.Description
End Function

' ============================================================
' PLACE_CELL
' Required params: cellName, ptX, ptY
' Optional params: ptZ (default 0), angleDeg (default 0)
'
' Zero GetInput — coordinates come entirely from the request,
' matching the programmatic PLACE CELL ICON pattern already used
' in DrawSign.bas / BBMarkupProcessor.ExecuteAddCell.
' ============================================================
Private Function ExecPlaceCell(reqId As String, params As Object) As String
    On Error GoTo PlaceError

    If Not params.Exists("cellName") Then
        ExecPlaceCell = reqId & vbTab & "ERROR" & vbTab & "note=missing cellName"
        Exit Function
    End If
    If Not (params.Exists("ptX") And params.Exists("ptY")) Then
        ExecPlaceCell = reqId & vbTab & "ERROR" & vbTab & "note=missing ptX/ptY"
        Exit Function
    End If

    Dim cellName As String: cellName = params("cellName")
    Dim ptX As Double: ptX = CDbl(params("ptX"))
    Dim ptY As Double: ptY = CDbl(params("ptY"))
    Dim ptZ As Double: ptZ = 0
    If params.Exists("ptZ") Then ptZ = CDbl(params("ptZ"))
    Dim angleDeg As Double: angleDeg = 0
    If params.Exists("angleDeg") Then angleDeg = CDbl(params("angleDeg"))

    Dim pt As Point3d
    pt.X = ptX: pt.Y = ptY: pt.Z = ptZ

    CadInputQueue.SendKeyin "ACTIVE LEVEL Default"
    CadInputQueue.SendKeyin "ACTIVE COLOR 0"
    CadInputQueue.SendKeyin "ACTIVE WEIGHT 0"
    CadInputQueue.SendKeyin "ACTIVE ANGLE " & angleDeg
    CadInputQueue.SendCommand "ATTACH LIBRARY " & WZTC_CELL_LIB
    SetCExpressionValue "tcb->activeCellUtf16", cellName, ""
    CadInputQueue.SendCommand "PLACE CELL ICON"
    CadInputQueue.SendDataPoint pt, 1
    CadInputQueue.SendReset
    CommandState.StartDefaultCommand

    ' Identify the element just placed: highest element ID in the model.
    ' Reliable for a single synchronous placement; a busier bridge with
    ' concurrent writers would need a snapshot-based approach instead
    ' (same technique AlignmentTool.bas already uses for alignment IDs).
    Dim newID As Double
    newID = FindMaxElementID()

    ExecPlaceCell = reqId & vbTab & "OK" & vbTab & "elementId=" & CStr(newID) & vbTab & _
                    "note=placed " & cellName & " at " & ptX & "," & ptY
    Exit Function

PlaceError:
    ExecPlaceCell = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' QUERY OPS — all read-only, delegate to WZTCQuery.bas / WZTCRules.bas.
' Multi-row results go to a per-request results file (Bridge\results_<reqId>.tsv);
' single-row results embed directly as key=val pairs in the response line,
' matching the PLACE_CELL response style.
' ============================================================

Private Function ExecFindElementsNear(reqId As String, params As Object) As String
    On Error GoTo QErr
    If Not (params.Exists("x") And params.Exists("y") And params.Exists("radius")) Then
        ExecFindElementsNear = reqId & vbTab & "ERROR" & vbTab & "note=missing x/y/radius"
        Exit Function
    End If
    Dim typeFilter As String: typeFilter = ""
    If params.Exists("typeFilter") Then typeFilter = params("typeFilter")

    Dim rows() As String
    rows = WZTCQuery.FindElementsNear(CDbl(params("x")), CDbl(params("y")), CDbl(params("radius")), typeFilter)
    ExecFindElementsNear = WriteResultRows(reqId, rows)
    Exit Function
QErr:
    ExecFindElementsNear = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function ExecStationToPoint(reqId As String, params As Object) As String
    On Error GoTo QErr
    If Not (params.Exists("alignIdx") And params.Exists("sta")) Then
        ExecStationToPoint = reqId & vbTab & "ERROR" & vbTab & "note=missing alignIdx/sta"
        Exit Function
    End If

    Dim rows() As String
    rows = WZTCQuery.StationToPoint(CInt(params("alignIdx")), CDbl(params("sta")))
    If rows(0) = "error" Then
        ExecStationToPoint = reqId & vbTab & "ERROR" & vbTab & "note=" & rows(1)
        Exit Function
    End If

    ' Single data row — embed directly rather than writing a results file
    Dim hdr() As String: hdr = Split(rows(0), vbTab)
    Dim dat() As String: dat = Split(rows(1), vbTab)
    Dim kv As String: kv = ""
    Dim i As Integer
    For i = 0 To UBound(hdr)
        kv = kv & vbTab & hdr(i) & "=" & dat(i)
    Next i
    ExecStationToPoint = reqId & vbTab & "OK" & kv
    Exit Function
QErr:
    ExecStationToPoint = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function ExecGetAlignmentStationing(reqId As String, params As Object) As String
    On Error GoTo QErr
    If Not params.Exists("alignIdx") Then
        ExecGetAlignmentStationing = reqId & vbTab & "ERROR" & vbTab & "note=missing alignIdx"
        Exit Function
    End If

    Dim rows() As String
    rows = WZTCQuery.GetAlignmentStationing(CInt(params("alignIdx")))
    If rows(0) = "error" Then
        ExecGetAlignmentStationing = reqId & vbTab & "ERROR" & vbTab & "note=" & rows(1)
        Exit Function
    End If
    ExecGetAlignmentStationing = WriteResultRows(reqId, rows)
    Exit Function
QErr:
    ExecGetAlignmentStationing = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function ExecListLevels(reqId As String, params As Object) As String
    On Error GoTo QErr
    Dim rows() As String
    rows = WZTCQuery.ListLevels()
    ExecListLevels = WriteResultRows(reqId, rows)
    Exit Function
QErr:
    ExecListLevels = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function ExecDescribeDrawingState(reqId As String, params As Object) As String
    On Error GoTo QErr
    Dim rows() As String
    rows = WZTCQuery.DescribeDrawingState()
    ExecDescribeDrawingState = WriteResultRows(reqId, rows)
    Exit Function
QErr:
    ExecDescribeDrawingState = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' CAPTURE_VIEW -- PARKED. WZTCViewCapture.CaptureView is a no-op
' stub (see that module's header for why: MicroStation's own
' raster-export dialog chain couldn't be driven headlessly without
' guessing an unknown internal command, and a guess already hung
' the VBA thread once, requiring an external WM_CLOSE to recover).
' Always returns ERROR. Real capture moved to an OS-level
' screenshot in mcp-server/view_capture.py, which doesn't touch
' CadInputQueue at all and has no path to this op.
' ============================================================
Private Function ExecCaptureView(reqId As String, params As Object) As String
    On Error GoTo CapError

    Dim viewNum As Integer: viewNum = 1
    If params.Exists("view") Then viewNum = CInt(params("view"))

    Dim filePath As String
    filePath = CAPTURES_DIR & "capture_" & reqId & ".png"

    Dim ok As Boolean
    ok = WZTCViewCapture.CaptureView(viewNum, filePath)

    If Not ok Then
        ExecCaptureView = reqId & vbTab & "ERROR" & vbTab & "note=capture did not produce a file at " & filePath
        Exit Function
    End If

    ExecCaptureView = reqId & vbTab & "OK" & vbTab & "path=" & filePath
    Exit Function

CapError:
    ExecCaptureView = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function ExecClassifySiteFeatures(reqId As String, params As Object) As String
    On Error GoTo QErr
    If Not (params.Exists("x") And params.Exists("y") And params.Exists("radius")) Then
        ExecClassifySiteFeatures = reqId & vbTab & "ERROR" & vbTab & "note=missing x/y/radius"
        Exit Function
    End If

    Dim rows() As String
    rows = WZTCQuery.ClassifySiteFeatures(CDbl(params("x")), CDbl(params("y")), CDbl(params("radius")))
    ExecClassifySiteFeatures = WriteResultRows(reqId, rows)
    Exit Function
QErr:
    ExecClassifySiteFeatures = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' COMPUTE_SPACING — wraps WZTCRules.ComputeSpacing. Deterministic
' MUTCD math; the agent reads a value here, it never invents one.
' Required params: speed, laneWidth, shoulderWidth, roadType
' ============================================================
Private Function ExecComputeSpacing(reqId As String, params As Object) As String
    On Error GoTo QErr
    If Not (params.Exists("speed") And params.Exists("laneWidth") And _
            params.Exists("shoulderWidth") And params.Exists("roadType")) Then
        ExecComputeSpacing = reqId & vbTab & "ERROR" & vbTab & "note=missing speed/laneWidth/shoulderWidth/roadType"
        Exit Function
    End If

    Dim sp As WZTCSpacing
    sp = WZTCRules.ComputeSpacing(CInt(params("speed")), CInt(params("laneWidth")), _
                                   CStr(params("shoulderWidth")), CStr(params("roadType")))

    ExecComputeSpacing = reqId & vbTab & "OK" & vbTab & _
        "downstreamTaper=" & sp.DownstreamTaper & vbTab & _
        "vehicleSpace=" & sp.VehicleSpace & vbTab & _
        "bufferSpace=" & sp.BufferSpace & vbTab & _
        "mergingTaper=" & sp.MergingTaper & vbTab & _
        "shoulderTaper=" & sp.ShoulderTaper & vbTab & _
        "advanceWarningSpacing=" & sp.AdvanceWarningSpacing & vbTab & _
        "rollAheadDistance=" & sp.RollAheadDistance & vbTab & _
        "upTaperBarrier=" & sp.UpTaperBarrier & vbTab & _
        "upTaperBeam=" & sp.UpTaperBeam & vbTab & _
        "flareBarrier=" & sp.FlareBarrier & vbTab & _
        "flareBeam=" & sp.FlareBeam & vbTab & _
        "skipTotal=" & sp.SkipTotal & vbTab & _
        "chanTotal=" & sp.ChanTotal
    Exit Function
QErr:
    ExecComputeSpacing = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' Required params: sheetNum (e.g. "619-302")
Private Function ExecGetSheetRequirements(reqId As String, params As Object) As String
    On Error GoTo QErr
    If Not params.Exists("sheetNum") Then
        ExecGetSheetRequirements = reqId & vbTab & "ERROR" & vbTab & "note=missing sheetNum"
        Exit Function
    End If

    Dim rows() As String
    rows = WZTCSheetRegistry.GetSheetRequirements(CStr(params("sheetNum")))
    If rows(0) = "error" Then
        ExecGetSheetRequirements = reqId & vbTab & "ERROR" & vbTab & "note=" & rows(1)
        Exit Function
    End If

    ' Single data row -- embed directly, same convention as ExecStationToPoint
    Dim hdr() As String: hdr = Split(rows(0), vbTab)
    Dim dat() As String: dat = Split(rows(1), vbTab)
    Dim kv As String: kv = ""
    Dim i As Integer
    For i = 0 To UBound(hdr)
        Dim val As String
        If i <= UBound(dat) Then val = dat(i) Else val = ""
        kv = kv & vbTab & hdr(i) & "=" & val
    Next i
    ExecGetSheetRequirements = reqId & vbTab & "OK" & kv
    Exit Function
QErr:
    ExecGetSheetRequirements = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' M3 DRAW OPS — thin wrappers delegating to WZTCExec.bas.
' Required params:
'   PLACE_PERP_LINE : alignIdx, sta            optional: halfLen
'   PLACE_SIGN      : signNum, roadType, side, pt1X, pt1Y, pt1Z, dir1X, dir1Y
'                     optional (required only if side=Both Sides):
'                     pt2X, pt2Y, pt2Z, dir2X, dir2Y
' ============================================================
Private Function BridgePlacePerpLine(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("alignIdx") And params.Exists("sta")) Then
        BridgePlacePerpLine = reqId & vbTab & "ERROR" & vbTab & "note=missing alignIdx/sta"
        Exit Function
    End If

    Dim halfLen As Double: halfLen = 40
    If params.Exists("halfLen") Then halfLen = CDbl(params("halfLen"))

    Dim result As String
    result = WZTCExec.ExecPlacePerpLine(CInt(params("alignIdx")), CDbl(params("sta")), halfLen)
    BridgePlacePerpLine = reqId & vbTab & result
    Exit Function
WErr:
    BridgePlacePerpLine = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgePlaceSign(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("signNum") And params.Exists("roadType") And params.Exists("side") And _
            params.Exists("pt1X") And params.Exists("pt1Y") And params.Exists("pt1Z") And _
            params.Exists("dir1X") And params.Exists("dir1Y")) Then
        BridgePlaceSign = reqId & vbTab & "ERROR" & vbTab & _
            "note=missing required param (signNum/roadType/side/pt1X/pt1Y/pt1Z/dir1X/dir1Y)"
        Exit Function
    End If

    Dim pt2X As Double, pt2Y As Double, pt2Z As Double, dir2X As Double, dir2Y As Double
    pt2X = 0: pt2Y = 0: pt2Z = 0: dir2X = 0: dir2Y = 0
    If params.Exists("pt2X") Then pt2X = CDbl(params("pt2X"))
    If params.Exists("pt2Y") Then pt2Y = CDbl(params("pt2Y"))
    If params.Exists("pt2Z") Then pt2Z = CDbl(params("pt2Z"))
    If params.Exists("dir2X") Then dir2X = CDbl(params("dir2X"))
    If params.Exists("dir2Y") Then dir2Y = CDbl(params("dir2Y"))

    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecPlaceSign(CStr(params("signNum")), CStr(params("roadType")), CStr(params("side")), _
                                    CDbl(params("pt1X")), CDbl(params("pt1Y")), CDbl(params("pt1Z")), _
                                    CDbl(params("dir1X")), CDbl(params("dir1Y")), _
                                    pt2X, pt2Y, pt2Z, dir2X, dir2Y)
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgePlaceSign = reqId & vbTab & result
    Exit Function
WErr:
    BridgePlaceSign = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' Required params: elementIdx (2-5), verticesTSV (pipe-separated "x,y,z")
Private Function BridgePlaceElementRun(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("elementIdx") And params.Exists("verticesTSV")) Then
        BridgePlaceElementRun = reqId & vbTab & "ERROR" & vbTab & "note=missing elementIdx/verticesTSV"
        Exit Function
    End If
    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecPlaceElementRun(CInt(params("elementIdx")), CStr(params("verticesTSV")))
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgePlaceElementRun = reqId & vbTab & result
    Exit Function
WErr:
    BridgePlaceElementRun = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' Required params: verticesTSV (pipe-separated "x,y,z", >= 3 points)
Private Function BridgePlaceWorkspace(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not params.Exists("verticesTSV") Then
        BridgePlaceWorkspace = reqId & vbTab & "ERROR" & vbTab & "note=missing verticesTSV"
        Exit Function
    End If
    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecPlaceWorkspace(CStr(params("verticesTSV")))
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgePlaceWorkspace = reqId & vbTab & result
    Exit Function
WErr:
    BridgePlaceWorkspace = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' Required params: elementIds (comma-separated element IDs)
Private Function BridgeSetSignAttributes(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not params.Exists("elementIds") Then
        BridgeSetSignAttributes = reqId & vbTab & "ERROR" & vbTab & "note=missing elementIds"
        Exit Function
    End If
    Dim result As String
    result = WZTCExec.ExecSetSignAttributes(CStr(params("elementIds")))
    BridgeSetSignAttributes = reqId & vbTab & result
    Exit Function
WErr:
    BridgeSetSignAttributes = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' Required: elementId. Optional: spacing, angleDeg, ownElementOnly
Private Function BridgeHatchElement(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not params.Exists("elementId") Then
        BridgeHatchElement = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId"
        Exit Function
    End If

    Dim gateMsg As String
    gateMsg = WZTCCommandRegistry.CheckSafetyGate("HATCH_ELEMENT")
    If gateMsg <> "" Then
        BridgeHatchElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String
    gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then
        BridgeHatchElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gate
        Exit Function
    End If

    Dim spacing As Double: spacing = 10#
    Dim angleDeg As Double: angleDeg = 45#
    If params.Exists("spacing") Then spacing = CDbl(params("spacing"))
    If params.Exists("angleDeg") Then angleDeg = CDbl(params("angleDeg"))

    Dim result As String
    result = WZTCExec.ExecHatchClosedElementByID(CDbl(params("elementId")), spacing, angleDeg)
    BridgeHatchElement = reqId & vbTab & result
    Exit Function
WErr:
    BridgeHatchElement = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' Required: x1,y1,x2,y2,x3,y3. Optional: z
Private Function BridgePlaceArc(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("x1") And params.Exists("y1") And _
            params.Exists("x2") And params.Exists("y2") And _
            params.Exists("x3") And params.Exists("y3")) Then
        BridgePlaceArc = reqId & vbTab & "ERROR" & vbTab & "note=missing x1/y1/x2/y2/x3/y3"
        Exit Function
    End If

    Dim gateMsg As String
    gateMsg = WZTCCommandRegistry.CheckSafetyGate("PLACE_ARC")
    If gateMsg <> "" Then
        BridgePlaceArc = reqId & vbTab & "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim z As Double: z = 0
    If params.Exists("z") Then z = CDbl(params("z"))

    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecPlaceArc3Point(CDbl(params("x1")), CDbl(params("y1")), _
                                         CDbl(params("x2")), CDbl(params("y2")), _
                                         CDbl(params("x3")), CDbl(params("y3")), z)
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgePlaceArc = reqId & vbTab & result
    Exit Function
WErr:
    BridgePlaceArc = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' Required: text, x, y. Optional: z
Private Function BridgePlaceTextLabel(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("text") And params.Exists("x") And params.Exists("y")) Then
        BridgePlaceTextLabel = reqId & vbTab & "ERROR" & vbTab & "note=missing text/x/y"
        Exit Function
    End If

    Dim gateMsg As String
    gateMsg = WZTCCommandRegistry.CheckSafetyGate("PLACE_TEXT_LABEL")
    If gateMsg <> "" Then
        BridgePlaceTextLabel = reqId & vbTab & "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim z As Double: z = 0
    If params.Exists("z") Then z = CDbl(params("z"))

    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecPlaceTextLabel(CStr(params("text")), CDbl(params("x")), CDbl(params("y")), z)
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgePlaceTextLabel = reqId & vbTab & result
    Exit Function
WErr:
    BridgePlaceTextLabel = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGateOrEmpty(opName As String) As String
    BridgeGateOrEmpty = WZTCCommandRegistry.CheckSafetyGate(opName)
End Function

Private Function BridgeGeomPlaceCircle(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("cx") And params.Exists("cy") And params.Exists("radius")) Then
        BridgeGeomPlaceCircle = reqId & vbTab & "ERROR" & vbTab & "note=missing cx/cy/radius": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("PLACE_CIRCLE")
    If g <> "" Then BridgeGeomPlaceCircle = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim z As Double: If params.Exists("z") Then z = CDbl(params("z"))
    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecPlaceCircle(CDbl(params("cx")), CDbl(params("cy")), CDbl(params("radius")), z)
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgeGeomPlaceCircle = reqId & vbTab & result: Exit Function
WErr: BridgeGeomPlaceCircle = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomPlaceEllipse(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("cx") And params.Exists("cy") And params.Exists("primaryRadius") And params.Exists("secondaryRadius")) Then
        BridgeGeomPlaceEllipse = reqId & vbTab & "ERROR" & vbTab & "note=missing cx/cy/primaryRadius/secondaryRadius": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("PLACE_ELLIPSE")
    If g <> "" Then BridgeGeomPlaceEllipse = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim ang As Double, z As Double
    If params.Exists("angleDeg") Then ang = CDbl(params("angleDeg"))
    If params.Exists("z") Then z = CDbl(params("z"))
    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecPlaceEllipse(CDbl(params("cx")), CDbl(params("cy")), CDbl(params("primaryRadius")), CDbl(params("secondaryRadius")), ang, z)
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgeGeomPlaceEllipse = reqId & vbTab & result: Exit Function
WErr: BridgeGeomPlaceEllipse = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomPlaceBlock(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("x1") And params.Exists("y1") And params.Exists("x2") And params.Exists("y2")) Then
        BridgeGeomPlaceBlock = reqId & vbTab & "ERROR" & vbTab & "note=missing x1/y1/x2/y2": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("PLACE_BLOCK")
    If g <> "" Then BridgeGeomPlaceBlock = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim z As Double: If params.Exists("z") Then z = CDbl(params("z"))
    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecPlaceBlock(CDbl(params("x1")), CDbl(params("y1")), CDbl(params("x2")), CDbl(params("y2")), z)
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgeGeomPlaceBlock = reqId & vbTab & result: Exit Function
WErr: BridgeGeomPlaceBlock = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomPlacePolyline(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not params.Exists("verticesTSV") Then
        BridgeGeomPlacePolyline = reqId & vbTab & "ERROR" & vbTab & "note=missing verticesTSV": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("PLACE_POLYLINE")
    If g <> "" Then BridgeGeomPlacePolyline = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecPlacePolyline(CStr(params("verticesTSV")))
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgeGeomPlacePolyline = reqId & vbTab & result: Exit Function
WErr: BridgeGeomPlacePolyline = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomPlacePolygon(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("cx") And params.Exists("cy") And params.Exists("radius") And params.Exists("sides")) Then
        BridgeGeomPlacePolygon = reqId & vbTab & "ERROR" & vbTab & "note=missing cx/cy/radius/sides": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("PLACE_POLYGON")
    If g <> "" Then BridgeGeomPlacePolygon = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim z As Double: If params.Exists("z") Then z = CDbl(params("z"))
    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecPlacePolygon(CDbl(params("cx")), CDbl(params("cy")), CDbl(params("radius")), CInt(params("sides")), z)
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgeGeomPlacePolygon = reqId & vbTab & result: Exit Function
WErr: BridgeGeomPlacePolygon = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomChangeSymbology(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not params.Exists("elementId") Then
        BridgeGeomChangeSymbology = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("CHANGE_ELEMENT_SYMBOLOGY")
    If g <> "" Then BridgeGeomChangeSymbology = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String: gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then BridgeGeomChangeSymbology = reqId & vbTab & "ERROR" & vbTab & "note=" & gate: Exit Function
    Dim color As Long: color = -1
    Dim weight As Long: weight = -1
    Dim ls As Long: ls = -999
    If params.Exists("color") Then color = CLng(params("color"))
    If params.Exists("weight") Then weight = CLng(params("weight"))
    If params.Exists("lineStyleIndex") Then ls = CLng(params("lineStyleIndex"))
    BridgeGeomChangeSymbology = reqId & vbTab & WZTCExec.ExecChangeElementSymbology(CDbl(params("elementId")), color, weight, ls)
    Exit Function
WErr: BridgeGeomChangeSymbology = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomCopyParallel(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("elementId") And params.Exists("distance")) Then
        BridgeGeomCopyParallel = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId/distance": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("COPY_PARALLEL")
    If g <> "" Then BridgeGeomCopyParallel = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String: gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then BridgeGeomCopyParallel = reqId & vbTab & "ERROR" & vbTab & "note=" & gate: Exit Function
    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecCopyParallelLineByID(CDbl(params("elementId")), CDbl(params("distance")))
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgeGeomCopyParallel = reqId & vbTab & result: Exit Function
WErr: BridgeGeomCopyParallel = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomCrossHatch(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not params.Exists("elementId") Then
        BridgeGeomCrossHatch = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("CROSSHATCH_ELEMENT")
    If g <> "" Then BridgeGeomCrossHatch = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String: gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then BridgeGeomCrossHatch = reqId & vbTab & "ERROR" & vbTab & "note=" & gate: Exit Function
    Dim spacing As Double: spacing = 10#: Dim ang As Double: ang = 45#
    If params.Exists("spacing") Then spacing = CDbl(params("spacing"))
    If params.Exists("angleDeg") Then ang = CDbl(params("angleDeg"))
    BridgeGeomCrossHatch = reqId & vbTab & WZTCExec.ExecCrossHatchClosedElementByID(CDbl(params("elementId")), spacing, ang)
    Exit Function
WErr: BridgeGeomCrossHatch = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomRemoveHatch(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not params.Exists("elementId") Then
        BridgeGeomRemoveHatch = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("REMOVE_HATCH")
    If g <> "" Then BridgeGeomRemoveHatch = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String: gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then BridgeGeomRemoveHatch = reqId & vbTab & "ERROR" & vbTab & "note=" & gate: Exit Function
    BridgeGeomRemoveHatch = reqId & vbTab & WZTCExec.ExecRemoveHatchByID(CDbl(params("elementId")))
    Exit Function
WErr: BridgeGeomRemoveHatch = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomBreakLine(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("elementId") And params.Exists("x") And params.Exists("y")) Then
        BridgeGeomBreakLine = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId/x/y": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("BREAK_LINE")
    If g <> "" Then BridgeGeomBreakLine = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String: gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then BridgeGeomBreakLine = reqId & vbTab & "ERROR" & vbTab & "note=" & gate: Exit Function
    Dim z As Double: If params.Exists("z") Then z = CDbl(params("z"))
    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecBreakLineAtPoint(CDbl(params("elementId")), CDbl(params("x")), CDbl(params("y")), z)
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgeGeomBreakLine = reqId & vbTab & result: Exit Function
WErr: BridgeGeomBreakLine = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomExtendLine(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("elementId") And params.Exists("newLength")) Then
        BridgeGeomExtendLine = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId/newLength": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("EXTEND_LINE")
    If g <> "" Then BridgeGeomExtendLine = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String: gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then BridgeGeomExtendLine = reqId & vbTab & "ERROR" & vbTab & "note=" & gate: Exit Function
    BridgeGeomExtendLine = reqId & vbTab & WZTCExec.ExecExtendLineToLength(CDbl(params("elementId")), CDbl(params("newLength")))
    Exit Function
WErr: BridgeGeomExtendLine = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomFillet(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("elementId1") And params.Exists("elementId2") And params.Exists("radius") And params.Exists("pickX") And params.Exists("pickY")) Then
        BridgeGeomFillet = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId1/elementId2/radius/pickX/pickY": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("FILLET_ELEMENTS")
    If g <> "" Then BridgeGeomFillet = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String
    gate = CheckOwnElementGate(CStr(params("elementId1")), ownOnly)
    If gate = "" Then gate = CheckOwnElementGate(CStr(params("elementId2")), ownOnly)
    If gate <> "" Then BridgeGeomFillet = reqId & vbTab & "ERROR" & vbTab & "note=" & gate: Exit Function
    Dim pz As Double: If params.Exists("pickZ") Then pz = CDbl(params("pickZ"))
    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecFilletTwoElements(CDbl(params("elementId1")), CDbl(params("elementId2")), CDbl(params("radius")), CDbl(params("pickX")), CDbl(params("pickY")), pz)
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgeGeomFillet = reqId & vbTab & result: Exit Function
WErr: BridgeGeomFillet = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomComplexString(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not params.Exists("elementIds") Then
        BridgeGeomComplexString = reqId & vbTab & "ERROR" & vbTab & "note=missing elementIds": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("CREATE_COMPLEX_STRING")
    If g <> "" Then BridgeGeomComplexString = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecCreateComplexString(CStr(params("elementIds")))
    If Left(result, 2) = "OK" Then result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    BridgeGeomComplexString = reqId & vbTab & result: Exit Function
WErr: BridgeGeomComplexString = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomPlaceFence(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("x1") And params.Exists("y1") And params.Exists("x2") And params.Exists("y2")) Then
        BridgeGeomPlaceFence = reqId & vbTab & "ERROR" & vbTab & "note=missing x1/y1/x2/y2": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("PLACE_FENCE_BLOCK")
    If g <> "" Then BridgeGeomPlaceFence = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim z As Double: If params.Exists("z") Then z = CDbl(params("z"))
    Dim vn As Integer: vn = 1: If params.Exists("viewNum") Then vn = CInt(params("viewNum"))
    BridgeGeomPlaceFence = reqId & vbTab & WZTCExec.ExecPlaceFenceBlock(CDbl(params("x1")), CDbl(params("y1")), CDbl(params("x2")), CDbl(params("y2")), z, vn)
    Exit Function
WErr: BridgeGeomPlaceFence = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomFenceUndefine(reqId As String, params As Object) As String
    On Error GoTo WErr
    Dim g As String: g = BridgeGateOrEmpty("FENCE_UNDEFINE")
    If g <> "" Then BridgeGeomFenceUndefine = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    BridgeGeomFenceUndefine = reqId & vbTab & WZTCExec.ExecFenceUndefine()
    Exit Function
WErr: BridgeGeomFenceUndefine = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomFenceCopy(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("deltaX") And params.Exists("deltaY")) Then
        BridgeGeomFenceCopy = reqId & vbTab & "ERROR" & vbTab & "note=missing deltaX/deltaY": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("FENCE_COPY_CONTENTS")
    If g <> "" Then BridgeGeomFenceCopy = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim dz As Double: If params.Exists("deltaZ") Then dz = CDbl(params("deltaZ"))
    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecFenceCopyContents(CDbl(params("deltaX")), CDbl(params("deltaY")), dz)
    If Left(result, 2) = "OK" And InStr(result, "createdElementIds=") = 0 Then
        result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    End If
    BridgeGeomFenceCopy = reqId & vbTab & result: Exit Function
WErr: BridgeGeomFenceCopy = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomFenceMove(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("deltaX") And params.Exists("deltaY")) Then
        BridgeGeomFenceMove = reqId & vbTab & "ERROR" & vbTab & "note=missing deltaX/deltaY": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("FENCE_MOVE_CONTENTS")
    If g <> "" Then BridgeGeomFenceMove = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim dz As Double: If params.Exists("deltaZ") Then dz = CDbl(params("deltaZ"))
    BridgeGeomFenceMove = reqId & vbTab & WZTCExec.ExecFenceMoveContents(CDbl(params("deltaX")), CDbl(params("deltaY")), dz)
    Exit Function
WErr: BridgeGeomFenceMove = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomFenceDelete(reqId As String, params As Object) As String
    On Error GoTo WErr
    Dim g As String: g = BridgeGateOrEmpty("FENCE_DELETE_CONTENTS")
    If g <> "" Then BridgeGeomFenceDelete = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    BridgeGeomFenceDelete = reqId & vbTab & WZTCExec.ExecFenceDeleteContents()
    Exit Function
WErr: BridgeGeomFenceDelete = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomSelect(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not params.Exists("elementId") Then
        BridgeGeomSelect = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId": Exit Function
    End If
    Dim g As String: g = BridgeGateOrEmpty("SELECT_ELEMENT")
    If g <> "" Then BridgeGeomSelect = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    Dim clearFirst As Boolean: clearFirst = True
    If params.Exists("clearFirst") Then
        If UCase(CStr(params("clearFirst"))) = "N" Or CStr(params("clearFirst")) = "0" Then clearFirst = False
    End If
    BridgeGeomSelect = reqId & vbTab & WZTCExec.ExecSelectElementByID(CDbl(params("elementId")), clearFirst)
    Exit Function
WErr: BridgeGeomSelect = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeGeomClearSelection(reqId As String, params As Object) As String
    On Error GoTo WErr
    Dim g As String: g = BridgeGateOrEmpty("CLEAR_SELECTION")
    If g <> "" Then BridgeGeomClearSelection = reqId & vbTab & "ERROR" & vbTab & "note=" & g: Exit Function
    BridgeGeomClearSelection = reqId & vbTab & WZTCExec.ExecClearSelection()
    Exit Function
WErr: BridgeGeomClearSelection = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' WRITE A MULTI-ROW QUERY RESULT TO Bridge\results_slot<N>.tsv
' Response line points to it and reports the data row count
' (header row not counted).
'
' Results are read-once: the caller (bridge_client.py) reads resultFile
' immediately after getting the response and never reopens it. A unique
' filename per reqId used to mean one new file per query call forever
' (19+ leftover results_P*.tsv files observed after one session). Fixed
' by reusing a small bounded pool of filenames instead -- "Open ... For
' Output" below already truncates/overwrites, so no delete is needed
' anywhere, and the file count on disk never grows past RESULT_POOL_SIZE.
' Safe as long as a single batch (see WZTCBridge.RunRequest's loop over
' request.tsv lines) never has more than RESULT_POOL_SIZE ops that each
' produce a multi-row result -- true today (call_batch is only ever
' invoked with one op at a time from bridge_client.py); if real
' multi-op batching is added later, bump RESULT_POOL_SIZE well past the
' largest expected batch size.
' ============================================================
Private Function WriteResultRows(reqId As String, rows() As String) As String
    Dim resultPath As String
    resultPath = BRIDGE_DIR & "results_slot" & ResultSlotFor(reqId) & ".tsv"

    Dim fnum As Integer: fnum = FreeFile
    Open resultPath For Output As #fnum
    Dim i As Integer
    For i = 0 To UBound(rows)
        Print #fnum, rows(i)
    Next i
    Close #fnum

    WriteResultRows = reqId & vbTab & "OK" & vbTab & "rowCount=" & UBound(rows) & vbTab & "resultFile=" & resultPath
End Function

Private Function ResultSlotFor(reqId As String) As Long
    Dim digits As String
    Dim c As String
    Dim i As Integer
    For i = 1 To Len(reqId)
        c = Mid(reqId, i, 1)
        If c >= "0" And c <= "9" Then digits = digits & c
    Next i
    If Len(digits) = 0 Then
        ResultSlotFor = 0
        Exit Function
    End If
    ' Keep only the trailing digits before converting so a long-running
    ' reqId counter can never overflow CLng.
    If Len(digits) > 6 Then digits = Right(digits, 6)
    ResultSlotFor = CLng(digits) Mod RESULT_POOL_SIZE
End Function

' ============================================================
' MAX ELEMENT ID IN MODEL — same conversion SharedState.ElIDAsDouble
' and PerpPlacement.bas already use for element ID bookkeeping.
' ============================================================
Private Function FindMaxElementID() As Double
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical

    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    Dim maxID As Double: maxID = 0
    Dim el As Element
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        Dim idVal As Double: idVal = ElIDAsDouble(el.ID)
        If idVal > maxID Then maxID = idVal
    Loop
    FindMaxElementID = maxID
End Function

' ============================================================
' CAPTURE ELEMENT IDS CREATED BY A DRAW OP
' Snapshot FindMaxElementID() before calling the primitive, then
' scan again afterward and collect everything with an ID greater
' than the snapshot. Gives every draw op a uniform "what did this
' create" answer regardless of whether it places one element
' (a perp line) or several (a sign's post/face/text/arc), which is
' what UNDO_LAST_OP needs to delete exactly the right things.
' ============================================================
Private Function CaptureNewElementIDs(beforeMaxID As Double) As String
    Dim oScan As ElementScanCriteria
    Set oScan = New ElementScanCriteria
    oScan.ExcludeNonGraphical

    Dim oEnum As ElementEnumerator
    Set oEnum = ActiveModelReference.Scan(oScan)

    Dim ids As String: ids = ""
    Dim el As Element
    Do While oEnum.MoveNext
        Set el = oEnum.Current
        Dim idVal As Double: idVal = ElIDAsDouble(el.ID)
        If idVal > beforeMaxID Then
            If ids <> "" Then ids = ids & ","
            ids = ids & CStr(idVal)
        End If
    Loop
    CaptureNewElementIDs = ids
End Function

' ============================================================
' HANDOFF — dimensions and callouts have no programmatic
' CadInputQueue precedent anywhere in this repo (see WZTCExec.bas
' header and the plan's "red list"). Rather than faking success,
' HANDOFF queues the request to Bridge\deferred-handoffs.tsv and
' returns DEFERRED so the agent can tell the engineer exactly what
' still needs a few clicks through the existing interactive forms.
' Required params: kind (dimension|callout). Everything else the
' caller sends (fromSta, toSta, text, reason, ...) is passed through
' to the queue file untouched.
' ============================================================
Private Function BridgeHandoff(reqId As String, params As Object) As String
    On Error GoTo HErr
    If Not params.Exists("kind") Then
        BridgeHandoff = reqId & vbTab & "ERROR" & vbTab & "note=missing kind (dimension|callout)"
        Exit Function
    End If

    Dim kind As String: kind = params("kind")
    Dim detail As String: detail = ""
    Dim k As Variant
    For Each k In params.Keys
        If k <> "kind" Then detail = detail & vbTab & CStr(k) & "=" & CStr(params(k))
    Next k

    Dim fnum As Integer: fnum = FreeFile
    Open HANDOFF_FILE For Append As #fnum
    Print #fnum, Now & vbTab & reqId & vbTab & kind & detail
    Close #fnum

    BridgeHandoff = reqId & vbTab & "DEFERRED" & vbTab & _
        "note=" & kind & " queued for manual placement via the existing form" & detail
    Exit Function
HErr:
    BridgeHandoff = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' UNDO LAST OP
' Deliberately does NOT rely on MicroStation's own undo stack —
' its exact grouping behavior across a multi-element op (PLACE_SIGN
' creates 4 elements in one call) has not been verified in the IDE,
' and the plan explicitly flags the MARK keyin / API as unconfirmed.
' Instead this walks the journal backward for the most recent
' undoable op that isn't already undone:
'   - createdElementIds= / elementId=  -> delete those elements
'   - priorDeltaX/Y/Z=                 -> re-apply reverse move (M6)
'   - priorLevel=                      -> restore prior level (M6)
'   - priorText=                       -> restore prior text (M6)
' DELETE_ELEMENT is intentionally NOT undoable (no snapshot) —
' its response carries none of the above fields, so the walk
' skips it and keeps looking. Same honesty pattern as HANDOFF.
' ============================================================
Private Function ExecUndoLastOp(reqId As String, params As Object) As String
    On Error GoTo UErr

    Dim allLines() As String
    Dim n As Integer
    n = ReadAllLines(JOURNAL_FILE, allLines)

    ' Single backward pass, oldest-last. reqId strings are NOT globally
    ' unique across the whole journal history -- a caller (or, as found by
    ' testing, this repo's own DebugAgentLoopTest.bas) can reuse the same
    ' literal reqId across separate runs/sessions. An UNDONE marker must
    ' therefore only cancel RESP lines that are OLDER than it (i.e. already
    ' passed in this same backward walk), never a later occurrence of the
    ' same reqId text -- otherwise a brand-new op happens to share a reqId
    ' with some unrelated already-undone op from history and gets skipped.
    ' Tracking "undone so far" as we walk, rather than pre-scanning the
    ' whole file for every UNDONE marker regardless of position, is what
    ' makes that distinction correctly.
    Dim undoneSoFar As Object
    Set undoneSoFar = CreateObject("Scripting.Dictionary")

    Dim i As Integer
    For i = n To 1 Step -1
        Dim ln As String: ln = allLines(i)

        If InStr(ln, vbTab & "UNDONE" & vbTab) > 0 Then
            Dim uParts() As String: uParts = Split(ln, vbTab)
            If UBound(uParts) >= 2 Then undoneSoFar(uParts(2)) = True

        ElseIf InStr(ln, vbTab & "RESP" & vbTab) > 0 Then
            Dim parts() As String: parts = Split(ln, vbTab)
            ' timestamp<TAB>RESP<TAB>reqId<TAB>status<TAB>key=val...
            If UBound(parts) >= 3 Then
                Dim origReqId As String: origReqId = parts(2)
                If Not undoneSoFar.Exists(origReqId) Then
                    Dim undoResult As String
                    undoResult = TryUndoFromRespFields(parts)
                    If undoResult <> "" Then
                        Dim fnum As Integer: fnum = FreeFile
                        Open JOURNAL_FILE For Append As #fnum
                        Print #fnum, Now & vbTab & "UNDONE" & vbTab & origReqId & vbTab & undoResult
                        Close #fnum

                        ExecUndoLastOp = reqId & vbTab & "OK" & vbTab & "undidReqId=" & origReqId & _
                                        vbTab & "notUndoable=Y" & vbTab & "note=" & undoResult
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i

    ExecUndoLastOp = reqId & vbTab & "ERROR" & vbTab & _
        "note=no undoable op found in journal (nothing placed yet, or everything already undone)"
    Exit Function

UErr:
    ExecUndoLastOp = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' Returns the Exec* result string if this RESP line is undoable,
' or "" if it has no undoable fields (e.g. DELETE_ELEMENT, HANDOFF,
' a query, or an ERROR response).
Private Function TryUndoFromRespFields(parts() As String) As String
    ' Checked first, unconditionally: UNDO_LAST_OP's own RESP line embeds
    ' the underlying mutation's result verbatim (e.g. undoing a move emits
    ' elementId=/priorDeltaX= just like a fresh MOVE_ELEMENT would), so
    ' without this check a second undo_last_op call parses the first
    ' undo's own response as itself undoable and "undoes the undo" (a
    ' redo) instead of continuing further back through real history.
    ' Confirmed live 2026-07-31: a 3rd undo call ping-ponged a moved
    ' element back and forth instead of reaching the original PLACE_CELL.
    ' ExecUndoLastOp declares notUndoable=Y on itself for exactly this.
    Dim k As Integer
    For k = 4 To UBound(parts)
        If Left(parts(k), Len("notUndoable=")) = "notUndoable=" Then
            TryUndoFromRespFields = ""
            Exit Function
        End If
    Next k

    Dim idsField As String: idsField = ""
    Dim elId As String: elId = ""
    Dim priorDX As String: priorDX = ""
    Dim priorDY As String: priorDY = ""
    Dim priorDZ As String: priorDZ = "0"
    Dim priorLevel As String: priorLevel = ""
    Dim priorText As String: priorText = ""
    Dim hasPriorDelta As Boolean: hasPriorDelta = False
    Dim hasPriorLevel As Boolean: hasPriorLevel = False
    Dim hasPriorText As Boolean: hasPriorText = False

    Dim j As Integer
    For j = 4 To UBound(parts)
        If Left(parts(j), Len("createdElementIds=")) = "createdElementIds=" Then
            idsField = Mid(parts(j), Len("createdElementIds=") + 1)
        ElseIf Left(parts(j), Len("elementId=")) = "elementId=" Then
            If elId = "" Then elId = Mid(parts(j), Len("elementId=") + 1)
            If idsField = "" Then idsField = elId
        ElseIf Left(parts(j), Len("priorDeltaX=")) = "priorDeltaX=" Then
            priorDX = Mid(parts(j), Len("priorDeltaX=") + 1)
            hasPriorDelta = True
        ElseIf Left(parts(j), Len("priorDeltaY=")) = "priorDeltaY=" Then
            priorDY = Mid(parts(j), Len("priorDeltaY=") + 1)
            hasPriorDelta = True
        ElseIf Left(parts(j), Len("priorDeltaZ=")) = "priorDeltaZ=" Then
            priorDZ = Mid(parts(j), Len("priorDeltaZ=") + 1)
        ElseIf Left(parts(j), Len("priorLevel=")) = "priorLevel=" Then
            priorLevel = Mid(parts(j), Len("priorLevel=") + 1)
            hasPriorLevel = True
        ElseIf Left(parts(j), Len("priorText=")) = "priorText=" Then
            priorText = Mid(parts(j), Len("priorText=") + 1)
            hasPriorText = True
        End If
    Next j

    ' Mutation undos take priority over a bare elementId= delete when
    ' both are present (MOVE/CHANGE_LEVEL/EDIT_TEXT all emit elementId=
    ' plus their prior* field — undoing those means restoring state,
    ' not deleting the element).
    If hasPriorDelta And elId <> "" Then
        TryUndoFromRespFields = WZTCExec.ExecMoveElementByID(CDbl(elId), _
            CDbl(priorDX), CDbl(priorDY), CDbl(priorDZ))
        Exit Function
    End If
    If hasPriorLevel And elId <> "" Then
        TryUndoFromRespFields = WZTCExec.ExecChangeElementLevelByID(CDbl(elId), priorLevel)
        Exit Function
    End If
    If hasPriorText And elId <> "" Then
        TryUndoFromRespFields = WZTCExec.ExecEditTextByID(CDbl(elId), priorText)
        Exit Function
    End If
    If idsField <> "" And Not hasPriorDelta And Not hasPriorLevel And Not hasPriorText Then
        ' DELETE_ELEMENT responses carry elementId= of what was removed
        ' but recreating it is impossible -- already filtered out above
        ' via their notUndoable=Y flag, so anything reaching here is a
        ' genuine create (createdElementIds=/elementId= from a draw op).
        TryUndoFromRespFields = WZTCExec.ExecDeleteElementsByID(idsField)
        Exit Function
    End If

    TryUndoFromRespFields = ""
End Function

' ============================================================
' GET_JOURNAL — returns the last N raw journal lines (default 50)
' as a multi-row result. Read-only, for the agent to answer
' "what have we done so far" / "why is that sign there" — the
' reason= a caller passes on any op rides along in the params blob
' AppendJournal already logs verbatim, so no separate reason storage
' is needed.
' ============================================================
Private Function ExecGetJournal(reqId As String, params As Object) As String
    On Error GoTo QErr
    Dim limitN As Integer: limitN = 50
    If params.Exists("limit") Then limitN = CInt(params("limit"))

    Dim allLines() As String
    Dim n As Integer
    n = ReadAllLines(JOURNAL_FILE, allLines)

    Dim startIdx As Integer: startIdx = n - limitN + 1
    If startIdx < 1 Then startIdx = 1

    Dim rowCount As Integer: rowCount = n - startIdx + 1
    If rowCount < 0 Then rowCount = 0

    Dim rows() As String
    ReDim rows(0 To rowCount)
    rows(0) = "line"
    Dim i As Integer
    For i = startIdx To n
        rows(i - startIdx + 1) = allLines(i)
    Next i

    ExecGetJournal = WriteResultRows(reqId, rows)
    Exit Function
QErr:
    ExecGetJournal = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' LIST_DEFERRED_HANDOFFS — everything HANDOFF has queued so far.
' ============================================================
Private Function ExecListDeferredHandoffs(reqId As String, params As Object) As String
    On Error GoTo QErr
    Dim allLines() As String
    Dim n As Integer
    n = ReadAllLines(HANDOFF_FILE, allLines)

    Dim rows() As String
    ReDim rows(0 To n)
    rows(0) = "timestamp" & vbTab & "reqId" & vbTab & "kind" & vbTab & "detail"
    Dim i As Integer
    For i = 1 To n
        rows(i) = allLines(i)
    Next i

    ExecListDeferredHandoffs = WriteResultRows(reqId, rows)
    Exit Function
QErr:
    ExecListDeferredHandoffs = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' M6 — COMMAND REGISTRY OPS
' ============================================================

Private Function ExecListRegistryCommands(reqId As String, params As Object) As String
    On Error GoTo QErr
    Dim safetyFilter As String: safetyFilter = ""
    If params.Exists("safetyStatus") Then safetyFilter = CStr(params("safetyStatus"))
    Dim rows() As String
    rows = WZTCCommandRegistry.ListCommands(safetyFilter)
    ExecListRegistryCommands = WriteResultRows(reqId, rows)
    Exit Function
QErr:
    ExecListRegistryCommands = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function ExecDescribeRegistryCommand(reqId As String, params As Object) As String
    On Error GoTo QErr
    If Not params.Exists("opName") Then
        ExecDescribeRegistryCommand = reqId & vbTab & "ERROR" & vbTab & "note=missing opName"
        Exit Function
    End If

    Dim row As Object
    Set row = WZTCCommandRegistry.LookupCommand(CStr(params("opName")))
    If row Is Nothing Then
        ExecDescribeRegistryCommand = reqId & vbTab & "ERROR" & vbTab & _
            "note=op not in registry: " & params("opName")
        Exit Function
    End If

    Dim kv As String: kv = ""
    Dim k As Variant
    For Each k In row.Keys
        kv = kv & vbTab & CStr(k) & "=" & CStr(row(k))
    Next k
    ExecDescribeRegistryCommand = reqId & vbTab & "OK" & kv
    Exit Function
QErr:
    ExecDescribeRegistryCommand = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' allowNeedsTesting=True only for TEST_REGISTRY_COMMAND (manual IDE).
Private Function ExecRunRegistryCommand(reqId As String, params As Object, _
                                        allowNeedsTesting As Boolean) As String
    On Error GoTo WErr
    If Not params.Exists("opName") Then
        ExecRunRegistryCommand = reqId & vbTab & "ERROR" & vbTab & "note=missing opName"
        Exit Function
    End If

    Dim opName As String: opName = CStr(params("opName"))
    Dim gateMsg As String
    gateMsg = WZTCCommandRegistry.CheckSafetyGate(opName, allowNeedsTesting)
    If gateMsg <> "" Then
        ExecRunRegistryCommand = reqId & vbTab & "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim row As Object
    Set row = WZTCCommandRegistry.LookupCommand(opName)
    Dim creates As Boolean
    creates = False
    If Not row Is Nothing Then
        If UCase(Trim(row("createsElements"))) = "Y" Then creates = True
    End If

    Dim beforeMaxID As Double: beforeMaxID = 0
    If creates Then beforeMaxID = FindMaxElementID()

    Dim result As String
    result = WZTCCommandRegistry.ExecuteRecipe(opName, params, allowNeedsTesting)
    If Left(result, 2) = "OK" And creates Then
        result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    End If
    ExecRunRegistryCommand = reqId & vbTab & result
    Exit Function
WErr:
    ExecRunRegistryCommand = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' M6 — DIRECT-API EDIT OPS
' ownElementOnly defaults to Y: target must appear as a
' createdElementIds= / elementId= value in the journal, matching
' the plan's original M6 scope ("elements the agent itself created").
' ============================================================

Private Function BridgeMoveElement(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("elementId") And params.Exists("deltaX") And params.Exists("deltaY")) Then
        BridgeMoveElement = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId/deltaX/deltaY"
        Exit Function
    End If

    Dim gateMsg As String
    gateMsg = WZTCCommandRegistry.CheckSafetyGate("MOVE_ELEMENT")
    If gateMsg <> "" Then
        BridgeMoveElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String
    gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then
        BridgeMoveElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gate
        Exit Function
    End If

    Dim deltaZ As Double: deltaZ = 0
    If params.Exists("deltaZ") Then deltaZ = CDbl(params("deltaZ"))

    Dim result As String
    result = WZTCExec.ExecMoveElementByID(CDbl(params("elementId")), _
                                          CDbl(params("deltaX")), CDbl(params("deltaY")), deltaZ)
    BridgeMoveElement = reqId & vbTab & result
    Exit Function
WErr:
    BridgeMoveElement = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeCopyElement(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("elementId") And params.Exists("deltaX") And params.Exists("deltaY")) Then
        BridgeCopyElement = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId/deltaX/deltaY"
        Exit Function
    End If

    Dim gateMsg As String
    gateMsg = WZTCCommandRegistry.CheckSafetyGate("COPY_ELEMENT")
    If gateMsg <> "" Then
        BridgeCopyElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String
    gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then
        BridgeCopyElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gate
        Exit Function
    End If

    Dim deltaZ As Double: deltaZ = 0
    If params.Exists("deltaZ") Then deltaZ = CDbl(params("deltaZ"))

    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecCopyElementByID(CDbl(params("elementId")), _
                                          CDbl(params("deltaX")), CDbl(params("deltaY")), deltaZ)
    If Left(result, 2) = "OK" Then
        result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    End If
    BridgeCopyElement = reqId & vbTab & result
    Exit Function
WErr:
    BridgeCopyElement = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeRotateElement(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("elementId") And params.Exists("originX") And _
            params.Exists("originY") And params.Exists("angleDeg")) Then
        BridgeRotateElement = reqId & vbTab & "ERROR" & vbTab & _
            "note=missing elementId/originX/originY/angleDeg"
        Exit Function
    End If

    Dim gateMsg As String
    gateMsg = WZTCCommandRegistry.CheckSafetyGate("ROTATE_ELEMENT")
    If gateMsg <> "" Then
        BridgeRotateElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String
    gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then
        BridgeRotateElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gate
        Exit Function
    End If

    Dim originZ As Double: originZ = 0
    If params.Exists("originZ") Then originZ = CDbl(params("originZ"))

    Dim result As String
    result = WZTCExec.ExecRotateElementByID(CDbl(params("elementId")), _
                                            CDbl(params("originX")), CDbl(params("originY")), _
                                            CDbl(params("angleDeg")), originZ)
    BridgeRotateElement = reqId & vbTab & result
    Exit Function
WErr:
    BridgeRotateElement = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeScaleElement(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("elementId") And params.Exists("originX") And _
            params.Exists("originY") And params.Exists("scaleFactor")) Then
        BridgeScaleElement = reqId & vbTab & "ERROR" & vbTab & _
            "note=missing elementId/originX/originY/scaleFactor"
        Exit Function
    End If

    Dim gateMsg As String
    gateMsg = WZTCCommandRegistry.CheckSafetyGate("SCALE_ELEMENT")
    If gateMsg <> "" Then
        BridgeScaleElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String
    gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then
        BridgeScaleElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gate
        Exit Function
    End If

    Dim originZ As Double: originZ = 0
    If params.Exists("originZ") Then originZ = CDbl(params("originZ"))

    Dim result As String
    result = WZTCExec.ExecScaleElementByID(CDbl(params("elementId")), _
                                           CDbl(params("originX")), CDbl(params("originY")), _
                                           CDbl(params("scaleFactor")), originZ)
    BridgeScaleElement = reqId & vbTab & result
    Exit Function
WErr:
    BridgeScaleElement = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeMirrorElement(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("elementId") And params.Exists("x1") And params.Exists("y1") And _
            params.Exists("x2") And params.Exists("y2")) Then
        BridgeMirrorElement = reqId & vbTab & "ERROR" & vbTab & _
            "note=missing elementId/x1/y1/x2/y2"
        Exit Function
    End If

    Dim gateMsg As String
    gateMsg = WZTCCommandRegistry.CheckSafetyGate("MIRROR_ELEMENT")
    If gateMsg <> "" Then
        BridgeMirrorElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String
    gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then
        BridgeMirrorElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gate
        Exit Function
    End If

    Dim z1 As Double: z1 = 0
    Dim z2 As Double: z2 = 0
    If params.Exists("z1") Then z1 = CDbl(params("z1"))
    If params.Exists("z2") Then z2 = CDbl(params("z2"))

    Dim result As String
    result = WZTCExec.ExecMirrorElementByID(CDbl(params("elementId")), _
                                            CDbl(params("x1")), CDbl(params("y1")), _
                                            CDbl(params("x2")), CDbl(params("y2")), z1, z2)
    BridgeMirrorElement = reqId & vbTab & result
    Exit Function
WErr:
    BridgeMirrorElement = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeArrayElement(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("elementId") And params.Exists("count") And _
            params.Exists("spacingX") And params.Exists("spacingY")) Then
        BridgeArrayElement = reqId & vbTab & "ERROR" & vbTab & _
            "note=missing elementId/count/spacingX/spacingY"
        Exit Function
    End If

    Dim gateMsg As String
    gateMsg = WZTCCommandRegistry.CheckSafetyGate("ARRAY_ELEMENT")
    If gateMsg <> "" Then
        BridgeArrayElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String
    gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then
        BridgeArrayElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gate
        Exit Function
    End If

    Dim beforeMaxID As Double: beforeMaxID = FindMaxElementID()
    Dim result As String
    result = WZTCExec.ExecArrayElementByID(CDbl(params("elementId")), _
                                           CInt(params("count")), _
                                           CDbl(params("spacingX")), CDbl(params("spacingY")))
    If Left(result, 2) = "OK" Then
        result = result & vbTab & "createdElementIds=" & CaptureNewElementIDs(beforeMaxID)
    End If
    BridgeArrayElement = reqId & vbTab & result
    Exit Function
WErr:
    BridgeArrayElement = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeChangeElementLevel(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("elementId") And params.Exists("level")) Then
        BridgeChangeElementLevel = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId/level"
        Exit Function
    End If

    Dim gateMsg As String
    gateMsg = WZTCCommandRegistry.CheckSafetyGate("CHANGE_ELEMENT_LEVEL")
    If gateMsg <> "" Then
        BridgeChangeElementLevel = reqId & vbTab & "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String
    gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then
        BridgeChangeElementLevel = reqId & vbTab & "ERROR" & vbTab & "note=" & gate
        Exit Function
    End If

    Dim result As String
    result = WZTCExec.ExecChangeElementLevelByID(CDbl(params("elementId")), CStr(params("level")))
    BridgeChangeElementLevel = reqId & vbTab & result
    Exit Function
WErr:
    BridgeChangeElementLevel = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeEditTextElement(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not (params.Exists("elementId") And params.Exists("newText")) Then
        BridgeEditTextElement = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId/newText"
        Exit Function
    End If

    Dim gateMsg As String
    gateMsg = WZTCCommandRegistry.CheckSafetyGate("EDIT_TEXT_ELEMENT")
    If gateMsg <> "" Then
        BridgeEditTextElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String
    gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then
        BridgeEditTextElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gate
        Exit Function
    End If

    Dim result As String
    result = WZTCExec.ExecEditTextByID(CDbl(params("elementId")), CStr(params("newText")))
    BridgeEditTextElement = reqId & vbTab & result
    Exit Function
WErr:
    BridgeEditTextElement = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

Private Function BridgeDeleteElement(reqId As String, params As Object) As String
    On Error GoTo WErr
    If Not params.Exists("elementId") Then
        BridgeDeleteElement = reqId & vbTab & "ERROR" & vbTab & "note=missing elementId"
        Exit Function
    End If

    Dim gateMsg As String
    gateMsg = WZTCCommandRegistry.CheckSafetyGate("DELETE_ELEMENT")
    If gateMsg <> "" Then
        BridgeDeleteElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim ownOnly As Boolean: ownOnly = OwnElementOnlyFlag(params)
    Dim gate As String
    gate = CheckOwnElementGate(CStr(params("elementId")), ownOnly)
    If gate <> "" Then
        BridgeDeleteElement = reqId & vbTab & "ERROR" & vbTab & "note=" & gate
        Exit Function
    End If

    Dim result As String
    result = WZTCExec.ExecDeleteElementsByID(CStr(params("elementId")))
    ' Always declare not-undoable — no snapshot to restore. UNDO_LAST_OP
    ' skips RESP lines carrying notUndoable=Y.
    If Left(result, 2) = "OK" Then
        result = result & vbTab & "notUndoable=Y" & vbTab & _
                 "note=DELETE_ELEMENT is not undoable via UNDO_LAST_OP (no snapshot to restore)"
    End If
    BridgeDeleteElement = reqId & vbTab & result
    Exit Function
WErr:
    BridgeDeleteElement = reqId & vbTab & "ERROR" & vbTab & "note=" & Err.Description
End Function

' Defaults to True (ownElementOnly=Y) unless the caller explicitly
' passes ownElementOnly=N. Expanding to "edit anything in the DGN"
' must be an intentional opt-out, never the silent default.
Private Function OwnElementOnlyFlag(params As Object) As Boolean
    OwnElementOnlyFlag = True
    If params.Exists("ownElementOnly") Then
        Dim v As String: v = UCase(Trim(CStr(params("ownElementOnly"))))
        If v = "N" Or v = "0" Or v = "FALSE" Then OwnElementOnlyFlag = False
    End If
End Function

' Returns "" if the gate passes; otherwise a refusal reason.
Private Function CheckOwnElementGate(elementId As String, ownOnly As Boolean) As String
    If Not ownOnly Then
        CheckOwnElementGate = ""
        Exit Function
    End If
    If ElementIdInJournal(elementId) Then
        CheckOwnElementGate = ""
    Else
        CheckOwnElementGate = "ownElementOnly refused: elementId " & elementId & _
            " is not in the journal as a createdElementIds=/elementId= value " & _
            "(agent may only edit elements it itself created; pass ownElementOnly=N to override)"
    End If
End Function

Private Function ElementIdInJournal(elementId As String) As Boolean
    Dim allLines() As String
    Dim n As Integer
    n = ReadAllLines(JOURNAL_FILE, allLines)
    Dim want As String: want = Trim(elementId)
    Dim i As Integer
    For i = 1 To n
        Dim ln As String: ln = allLines(i)
        If InStr(ln, vbTab & "RESP" & vbTab) > 0 Then
            Dim parts() As String: parts = Split(ln, vbTab)
            Dim j As Integer
            For j = 0 To UBound(parts)
                If Left(parts(j), Len("createdElementIds=")) = "createdElementIds=" Then
                    If IdListContains(Mid(parts(j), Len("createdElementIds=") + 1), want) Then
                        ElementIdInJournal = True
                        Exit Function
                    End If
                ElseIf Left(parts(j), Len("elementId=")) = "elementId=" Then
                    If Trim(Mid(parts(j), Len("elementId=") + 1)) = want Then
                        ElementIdInJournal = True
                        Exit Function
                    End If
                End If
            Next j
        End If
    Next i
    ElementIdInJournal = False
End Function

Private Function IdListContains(idsCSV As String, want As String) As Boolean
    Dim ids() As String: ids = Split(idsCSV, ",")
    Dim i As Integer
    For i = 0 To UBound(ids)
        If Trim(ids(i)) = want Then
            IdListContains = True
            Exit Function
        End If
    Next i
    IdListContains = False
End Function

' ============================================================
' TSV PARAM PARSING — key=val<TAB>key=val... -> Dictionary
' Late-bound Scripting.Dictionary: available on stock Windows
' (scrrun.dll) in both 32- and 64-bit hosts, unlike MSScriptControl.
' Values may themselves contain '=' (e.g. newText=a=b) — only the
' first '=' splits key from value; the rest is rejoined.
' ============================================================
Private Function ParseParams(parts() As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Integer
    For i = 2 To UBound(parts)
        Dim eqPos As Integer
        eqPos = InStr(parts(i), "=")
        If eqPos > 0 Then
            d(Trim(Left(parts(i), eqPos - 1))) = Trim(Mid(parts(i), eqPos + 1))
        End If
    Next i
    Set ParseParams = d
End Function

' ============================================================
' FILE I/O HELPERS
' ============================================================
Private Function ReadAllLines(path As String, ByRef outLines() As String) As Integer
    On Error GoTo ReadErr
    Dim fnum As Integer: fnum = 0
    If Dir(path) = "" Then
        ReadAllLines = 0
        Exit Function
    End If

    fnum = FreeFile
    Open path For Input As #fnum

    Dim n As Integer: n = 0
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

Private Sub WriteFile(path As String, content As String)
    Dim fnum As Integer: fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, content
    Close #fnum
End Sub

Private Sub WriteLines(path As String, lines() As String, n As Integer)
    Dim fnum As Integer: fnum = FreeFile
    Open path For Output As #fnum
    Dim i As Integer
    For i = 1 To n
        Print #fnum, lines(i)
    Next i
    Close #fnum
End Sub

Private Sub AppendJournal(reqLine As String, respLine As String)
    Dim fnum As Integer: fnum = FreeFile
    Open JOURNAL_FILE For Append As #fnum
    Print #fnum, Now & vbTab & "REQ" & vbTab & reqLine
    Print #fnum, Now & vbTab & "RESP" & vbTab & respLine
    Close #fnum

    Call RotateJournalIfOversized
End Sub

' ============================================================
' KEEP wztc-journal.tsv FROM GROWING FOREVER
' Append-only since M1; nothing ever trimmed it. Both consumers
' (ExecUndoLastOp, ExecGetJournal) only ever need recent history --
' ExecUndoLastOp walks backward from the end, GET_JOURNAL returns just
' the last N lines -- so once the file passes JOURNAL_MAX_BYTES, the
' OLDER prefix (everything before the retained tail) is moved to
' Bridge\archive\ (never deleted) and the live file keeps only the
' most recent JOURNAL_KEEP_LINES lines. Runs on every AppendJournal
' call but is a no-op (one FileLen check) until the file actually
' crosses the threshold, so no meaningful per-op cost.
' ============================================================
Private Sub RotateJournalIfOversized()
    On Error GoTo RotateErr   ' rotation must never break the journal write that just succeeded
    If FileLen(JOURNAL_FILE) < JOURNAL_MAX_BYTES Then Exit Sub

    Dim allLines() As String
    Dim n As Integer
    n = ReadAllLines(JOURNAL_FILE, allLines)
    If n <= JOURNAL_KEEP_LINES Then Exit Sub

    Dim archiveDir As String
    archiveDir = BRIDGE_DIR & "archive\"
    If Dir(archiveDir, vbDirectory) = "" Then MkDir archiveDir

    Dim archivePath As String
    archivePath = archiveDir & "wztc-journal-" & Format(Now, "yyyy-mm-dd_hhnnss") & ".tsv"

    Dim fnum As Integer
    Dim i As Integer

    fnum = FreeFile
    Open archivePath For Output As #fnum
    For i = 1 To n - JOURNAL_KEEP_LINES
        Print #fnum, allLines(i)
    Next i
    Close #fnum

    fnum = FreeFile
    Open JOURNAL_FILE For Output As #fnum
    For i = n - JOURNAL_KEEP_LINES + 1 To n
        Print #fnum, allLines(i)
    Next i
    Close #fnum
    Exit Sub

RotateErr:
    ' Best-effort: leave the journal as-is and let it retry next call
    ' rather than risk leaving either file half-written.
End Sub
