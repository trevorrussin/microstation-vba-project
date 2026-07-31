Option Explicit

' ============================================================
' WZTC SHEET REGISTRY — 619 SHEET NUMBER -> REQUIRED SIGNS/ELEMENTS
' ------------------------------------------------------------
' Reads Data/sheet-registry.tsv, an external file (not hardcoded
' VBA) so both this module and a future Python MCP server can
' read the same source, per the M4 plan design.
'
' Read-only, no CadInputQueue, no SharedState mutation -- same
' category of module as WZTCQuery.bas.
'
' Seeded incrementally: only 6 of 91 current 619 sheets are in
' the data file as of this writing (see Data/README.md). A sheet
' not found here is not an error -- GetSheetRequirements returns
' a clear "not in registry" result so the caller can fall back to
' manual entry, exactly like every other sheet already works today.
' ============================================================

Private Const REGISTRY_FILE As String = "c:\repos\microstation-vba-project\Data\sheet-registry.tsv"

' ============================================================
' GET REQUIREMENTS FOR ONE SHEET
' Returns a 2-row array: header + data row (same convention as
' WZTCQuery.StationToPoint), or a 2-row "error" array if the
' sheet isn't in the registry yet or the file can't be read.
' ============================================================
Public Function GetSheetRequirements(sheetNum As String) As String()
    Dim rows() As String
    Dim lines() As String
    Dim n As Integer
    n = ReadAllLines(REGISTRY_FILE, lines)

    If n < 2 Then
        ReDim rows(0 To 1)
        rows(0) = "error"
        rows(1) = "could not read sheet registry file (missing or empty): " & REGISTRY_FILE
        GetSheetRequirements = rows
        Exit Function
    End If

    Dim wantSheet As String: wantSheet = Trim(sheetNum)
    Dim i As Integer
    For i = 2 To n   ' line 1 is the header, data starts at line 2
        Dim fields() As String
        fields = Split(lines(i - 1), vbTab)
        If UBound(fields) >= 0 Then
            If Trim(fields(0)) = wantSheet Then
                ReDim rows(0 To 1)
                rows(0) = "sheetNum" & vbTab & "title" & vbTab & "roadType" & vbTab & _
                          "duration" & vbTab & "signs" & vbTab & "elements" & vbTab & "notes"
                rows(1) = lines(i - 1)
                GetSheetRequirements = rows
                Exit Function
            End If
        End If
    Next i

    ReDim rows(0 To 1)
    rows(0) = "error"
    rows(1) = "sheet " & wantSheet & " is not in the registry yet (seeded incrementally -- " & _
              "see Data/README.md). Fall back to manual sign/element entry for this sheet."
    GetSheetRequirements = rows
End Function

' ============================================================
' LIST ALL SHEET NUMBERS CURRENTLY IN THE REGISTRY
' Useful for a caller to check what's available before asking.
' ============================================================
Public Function ListRegisteredSheets() As String()
    Dim rows() As String
    Dim lines() As String
    Dim n As Integer
    n = ReadAllLines(REGISTRY_FILE, lines)

    If n < 2 Then
        ReDim rows(0 To 0)
        rows(0) = "sheetNum" & vbTab & "title"
        ListRegisteredSheets = rows
        Exit Function
    End If

    ReDim rows(0 To n - 1)
    rows(0) = "sheetNum" & vbTab & "title"
    Dim i As Integer
    For i = 2 To n
        Dim fields() As String
        fields = Split(lines(i - 1), vbTab)
        If UBound(fields) >= 1 Then
            rows(i - 1) = fields(0) & vbTab & fields(1)
        End If
    Next i
    ListRegisteredSheets = rows
End Function

' ============================================================
' FILE I/O HELPER — same pattern as WZTCBridge.ReadAllLines
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
