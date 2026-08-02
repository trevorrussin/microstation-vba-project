Option Explicit

' ============================================================
' WZTC COMMAND REGISTRY — SAFETY-GATED MICROSTATION RECIPE CATALOG
' ------------------------------------------------------------
' Reads Data/command-registry.tsv. The unit of vetting is a named
' recipe (settings -> activation -> datapoints -> reset), never a
' bare command string — the same token (e.g. PLACE CELL ICON)
' appears at both headless-safe and GetInput-dependent call sites
' elsewhere in this repo.
'
' safetyStatus is the enforcement gate:
'   verified-headless-safe       — only status ExecuteRecipe will run
'   needs-testing                — catalogued, refused at execution
'   interactive-only-use-handoff — points caller at HANDOFF / forms
'   unsafe-blocked               — confirmed activate-and-abandon
'   external-app-blocked         — integrates with ProjectWise/DWG/IFC/
'                                   DMS or similar; policy-withheld from
'                                   every agent (2026-08-02), not a
'                                   technical finding -- still catalogued
'
' Zero GetInput calls anywhere in this module — same invariant as
' WZTCExec.bas. Structural close-out guard: any recipe with a
' COMMAND: step must eventually be followed by DATAPOINT: + RESET,
' regardless of safetyStatus — second defense against the
' BBMarkupProcessor.ExecuteAddDimension / ExecuteAddCallout
' activate-and-abandon anti-pattern.
' ============================================================

Private Const REGISTRY_FILE As String = "c:\repos\microstation-vba-project\Data\command-registry.tsv"

Private Const COL_OPNAME As Integer = 0
Private Const COL_CATEGORY As Integer = 1
Private Const COL_SAFETY As Integer = 2
Private Const COL_RECIPE As Integer = 3
Private Const COL_VBAFN As Integer = 4
Private Const COL_REQPARAMS As Integer = 5
Private Const COL_OPTPARAMS As Integer = 6
Private Const COL_CREATES As Integer = 7
Private Const COL_OWNONLY As Integer = 8
Private Const COL_SOURCEREFS As Integer = 9
Private Const COL_ADDEDDATE As Integer = 10
Private Const COL_PROMOTED As Integer = 11
Private Const COL_NOTES As Integer = 12

' ============================================================
' LOOK UP ONE REGISTRY ROW BY opName
' Returns a Dictionary of column-name -> value, or Nothing if
' the op is not in the registry / file unreadable.
' ============================================================
Public Function LookupCommand(opName As String) As Object
    Dim lines() As String
    Dim n As Integer
    n = ReadAllLines(REGISTRY_FILE, lines)
    If n < 2 Then
        Set LookupCommand = Nothing
        Exit Function
    End If

    Dim hdr() As String
    hdr = Split(lines(1), vbTab)

    Dim want As String: want = Trim(opName)
    Dim i As Integer
    For i = 2 To n
        Dim fields() As String
        fields = Split(lines(i), vbTab)
        If UBound(fields) >= COL_OPNAME Then
            If Trim(fields(COL_OPNAME)) = want Then
                Dim d As Object
                Set d = CreateObject("Scripting.Dictionary")
                Dim c As Integer
                For c = 0 To UBound(hdr)
                    Dim key As String: key = Trim(hdr(c))
                    Dim val As String: val = ""
                    If c <= UBound(fields) Then val = fields(c)
                    d(key) = val
                Next c
                Set LookupCommand = d
                Exit Function
            End If
        End If
    Next i
    Set LookupCommand = Nothing
End Function

' ============================================================
' LIST ALL REGISTRY ROWS (header + data) as a String array —
' same multi-row convention as WZTCQuery / WZTCSheetRegistry.
' ============================================================
Public Function ListCommands(Optional safetyFilter As String = "") As String()
    Dim lines() As String
    Dim n As Integer
    n = ReadAllLines(REGISTRY_FILE, lines)

    If n < 1 Then
        Dim emptyRows() As String
        ReDim emptyRows(0 To 0)
        emptyRows(0) = "opName" & vbTab & "category" & vbTab & "safetyStatus" & vbTab & _
                       "requiredParams" & vbTab & "createsElements" & vbTab & "ownElementOnly" & vbTab & "notes"
        ListCommands = emptyRows
        Exit Function
    End If

    Dim wantFilter As String: wantFilter = Trim(safetyFilter)
    Dim out() As String
    Dim outN As Integer: outN = 0
    ReDim out(0 To n)   ' upper bound trimmed below
    out(0) = "opName" & vbTab & "category" & vbTab & "safetyStatus" & vbTab & _
             "requiredParams" & vbTab & "optionalParams" & vbTab & _
             "createsElements" & vbTab & "ownElementOnly" & vbTab & "notes"
    outN = 1

    Dim i As Integer
    For i = 2 To n
        Dim fields() As String
        fields = Split(lines(i), vbTab)
        If UBound(fields) >= COL_SAFETY Then
            If wantFilter = "" Or Trim(fields(COL_SAFETY)) = wantFilter Then
                Dim notes As String: notes = ""
                If UBound(fields) >= COL_NOTES Then notes = fields(COL_NOTES)
                Dim optP As String: optP = ""
                If UBound(fields) >= COL_OPTPARAMS Then optP = fields(COL_OPTPARAMS)
                Dim creates As String: creates = ""
                If UBound(fields) >= COL_CREATES Then creates = fields(COL_CREATES)
                Dim ownOnly As String: ownOnly = ""
                If UBound(fields) >= COL_OWNONLY Then ownOnly = fields(COL_OWNONLY)
                Dim reqP As String: reqP = ""
                If UBound(fields) >= COL_REQPARAMS Then reqP = fields(COL_REQPARAMS)
                out(outN) = fields(COL_OPNAME) & vbTab & fields(COL_CATEGORY) & vbTab & _
                            fields(COL_SAFETY) & vbTab & reqP & vbTab & optP & vbTab & _
                            creates & vbTab & ownOnly & vbTab & notes
                outN = outN + 1
            End If
        End If
    Next i

    ReDim Preserve out(0 To outN - 1)
    ListCommands = out
End Function

' ============================================================
' SAFETY GATE — single enforcement point every registry-driven
' op runs through. Returns "" if the op may execute under the
' normal (non-TEST) path; otherwise a human-readable refusal
' reason the bridge embeds in an ERROR response.
'
' allowNeedsTesting=True is reserved for TEST_REGISTRY_COMMAND
' (manual IDE promotion only — never exposed via MCP).
' ============================================================
Public Function CheckSafetyGate(opName As String, _
                                Optional allowNeedsTesting As Boolean = False) As String
    Dim row As Object
    Set row = LookupCommand(opName)
    If row Is Nothing Then
        CheckSafetyGate = "op '" & opName & "' is not in the command registry"
        Exit Function
    End If

    Dim status As String: status = Trim(row("safetyStatus"))
    Select Case status
        Case "verified-headless-safe"
            CheckSafetyGate = ""
        Case "needs-testing"
            If allowNeedsTesting Then
                CheckSafetyGate = ""
            Else
                CheckSafetyGate = "op '" & opName & "' has safetyStatus=needs-testing — " & _
                    "refused. Promote after IDE verification (or use TEST_REGISTRY_COMMAND " & _
                    "by hand-editing Bridge/request.tsv). Not available to the agent."
            End If
        Case "interactive-only-use-handoff"
            CheckSafetyGate = "op '" & opName & "' is interactive-only — use HANDOFF " & _
                "(kind=dimension|callout) or the existing PlaceElements/PlaceCells form. " & _
                "See registry notes: " & row("notes")
        Case "unsafe-blocked"
            CheckSafetyGate = "op '" & opName & "' is unsafe-blocked (confirmed " & _
                "activate-and-abandon). Not promotable without a redesign. " & row("notes")
        Case "external-app-blocked"
            CheckSafetyGate = "op '" & opName & "' integrates with an external application, " & _
                "format, or document-management service (ProjectWise, DWG/AutoCAD interop, IFC, " & _
                "DMS, etc.) and is deliberately withheld from every agent for now -- a policy " & _
                "decision (2026-08-02), not a technical safety finding. Still catalogued in the " & _
                "registry for reference. See registry notes: " & row("notes")
        Case Else
            CheckSafetyGate = "op '" & opName & "' has unknown safetyStatus='" & status & "'"
    End Select
End Function

' ============================================================
' EXECUTE A keyin_recipe ROW
' Substitutes {param} placeholders from params Dictionary, runs
' the structural close-out guard, then interprets each pipe-
' separated recipe step. Returns "OK<TAB>..." or "ERROR<TAB>...".
' The safety gate is checked here AND in WZTCBridge.ExecRunRegistryCommand
' (belt-and-suspenders). Caller (WZTCBridge) is still responsible for
' wrapping createsElements=Y rows with CaptureNewElementIDs.
' ============================================================
Public Function ExecuteRecipe(opName As String, params As Object, _
                              Optional allowNeedsTesting As Boolean = False) As String
    On Error GoTo RecipeErr

    ' Gate is also checked by WZTCBridge.ExecRunRegistryCommand before this
    ' is called, but that's an external convention -- this function is the
    ' one that actually drives CadInputQueue, so it must not trust a future
    ' caller to have remembered the gate. Same "second defense" spirit as
    ' the close-out guard below.
    Dim gateMsg As String
    gateMsg = CheckSafetyGate(opName, allowNeedsTesting)
    If gateMsg <> "" Then
        ExecuteRecipe = "ERROR" & vbTab & "note=" & gateMsg
        Exit Function
    End If

    Dim row As Object
    Set row = LookupCommand(opName)
    If row Is Nothing Then
        ExecuteRecipe = "ERROR" & vbTab & "note=op not in registry: " & opName
        Exit Function
    End If

    If Trim(row("category")) <> "keyin_recipe" Then
        ExecuteRecipe = "ERROR" & vbTab & "note=op '" & opName & "' is category=" & _
            row("category") & " — call its dedicated bridge op (e.g. MOVE_ELEMENT), " & _
            "not RUN_REGISTRY_COMMAND"
        Exit Function
    End If

    Dim recipe As String: recipe = Trim(row("recipeLines"))
    If recipe = "" Then
        ExecuteRecipe = "ERROR" & vbTab & "note=op '" & opName & "' has empty recipeLines"
        Exit Function
    End If

    ' Required-param check
    Dim reqList As String: reqList = Trim(row("requiredParams"))
    If reqList <> "" Then
        Dim reqs() As String: reqs = Split(reqList, "|")
        Dim r As Integer
        For r = 0 To UBound(reqs)
            Dim reqName As String: reqName = Trim(reqs(r))
            If reqName <> "" Then
                If params Is Nothing Or Not params.Exists(reqName) Then
                    ExecuteRecipe = "ERROR" & vbTab & "note=missing required param: " & reqName
                    Exit Function
                End If
            End If
        Next r
    End If

    Dim expanded As String
    expanded = SubstituteParams(recipe, params)

    Dim guardMsg As String
    guardMsg = CheckCloseOutGuard(expanded)
    If guardMsg <> "" Then
        ExecuteRecipe = "ERROR" & vbTab & "note=close-out guard refused: " & guardMsg
        Exit Function
    End If

    Dim steps() As String
    steps = Split(expanded, "|")
    Dim s As Integer
    For s = 0 To UBound(steps)
        Dim stepText As String: stepText = Trim(steps(s))
        If stepText <> "" Then
            Dim stepErr As String
            stepErr = RunRecipeStep(stepText)
            If stepErr <> "" Then
                ExecuteRecipe = "ERROR" & vbTab & "note=step failed (" & stepText & "): " & stepErr
                Exit Function
            End If
        End If
    Next s

    ExecuteRecipe = "OK" & vbTab & "opName=" & opName & vbTab & _
                    "createsElements=" & Trim(row("createsElements")) & vbTab & _
                    "note=ran registry recipe " & opName
    Exit Function

RecipeErr:
    ExecuteRecipe = "ERROR" & vbTab & "note=" & Err.Description
End Function

' ============================================================
' STRUCTURAL CLOSE-OUT GUARD
' Any recipe containing a COMMAND: step must eventually include
' both a DATAPOINT: step and a RESET step. Prevents the
' activate-and-abandon pattern even if a row is mis-labelled
' verified-headless-safe.
' ============================================================
Private Function CheckCloseOutGuard(recipe As String) As String
    Dim steps() As String
    steps = Split(recipe, "|")
    Dim hasCommand As Boolean: hasCommand = False
    Dim hasDataPoint As Boolean: hasDataPoint = False
    Dim hasReset As Boolean: hasReset = False
    Dim i As Integer
    For i = 0 To UBound(steps)
        Dim t As String: t = UCase(Trim(steps(i)))
        If Left(t, 8) = "COMMAND:" Then hasCommand = True
        If Left(t, 10) = "DATAPOINT:" Then hasDataPoint = True
        If t = "RESET" Then hasReset = True
    Next i

    If hasCommand Then
        If Not hasDataPoint Or Not hasReset Then
            CheckCloseOutGuard = "recipe has COMMAND: without both DATAPOINT: and RESET " & _
                "(activate-and-abandon defense — see BBMarkupProcessor.ExecuteAddDimension)"
            Exit Function
        End If
    End If
    CheckCloseOutGuard = ""
End Function

' ============================================================
' SUBSTITUTE {param} PLACEHOLDERS FROM THE PARAMS DICTIONARY
' Unresolved placeholders are left as-is so a missing optional
' param surfaces as a bad keyin rather than a silent empty.
' ============================================================
Private Function SubstituteParams(recipe As String, params As Object) As String
    Dim result As String: result = recipe
    If params Is Nothing Then
        SubstituteParams = result
        Exit Function
    End If

    Dim k As Variant
    For Each k In params.Keys
        result = Replace(result, "{" & CStr(k) & "}", CStr(params(k)))
    Next k
    SubstituteParams = result
End Function

' ============================================================
' RUN ONE RECIPE STEP — returns "" on success, error text on fail
' ============================================================
Private Function RunRecipeStep(stepText As String) As String
    On Error GoTo StepErr
    Dim upper As String: upper = UCase(stepText)

    If Left(upper, 6) = "KEYIN:" Then
        CadInputQueue.SendKeyin Mid(stepText, 7)
        RunRecipeStep = ""
        Exit Function
    End If

    If Left(upper, 8) = "COMMAND:" Then
        CadInputQueue.SendCommand Mid(stepText, 9)
        RunRecipeStep = ""
        Exit Function
    End If

    If Left(upper, 8) = "SETCELL:" Then
        SetCExpressionValue "tcb->activeCellUtf16", Mid(stepText, 9), ""
        RunRecipeStep = ""
        Exit Function
    End If

    If Left(upper, 10) = "DATAPOINT:" Then
        Dim coords As String: coords = Mid(stepText, 11)
        Dim parts() As String: parts = Split(coords, ",")
        If UBound(parts) < 1 Then
            RunRecipeStep = "DATAPOINT needs at least x,y"
            Exit Function
        End If
        Dim pt As Point3d
        pt.X = CDbl(Trim(parts(0)))
        pt.Y = CDbl(Trim(parts(1)))
        If UBound(parts) >= 2 Then pt.Z = CDbl(Trim(parts(2))) Else pt.Z = 0
        CadInputQueue.SendDataPoint pt, 1
        RunRecipeStep = ""
        Exit Function
    End If

    If upper = "RESET" Then
        CadInputQueue.SendReset
        RunRecipeStep = ""
        Exit Function
    End If

    If upper = "DEFAULTCOMMAND" Then
        CommandState.StartDefaultCommand
        RunRecipeStep = ""
        Exit Function
    End If

    RunRecipeStep = "unknown recipe step: " & stepText
    Exit Function

StepErr:
    RunRecipeStep = Err.Description
End Function

' ============================================================
' FILE I/O — same pattern as WZTCSheetRegistry.ReadAllLines
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
