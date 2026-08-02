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
' Covers all 91 DesignerRef sheets (see Data/README.md). Some rows are
' stubs with empty signs when the sheet is not in the 2026 Book 3 PDF.
' A sheet not found here is not an error -- GetSheetRequirements returns
' a clear "not in registry" result so the caller can fall back to
' manual entry.
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

    ' lines() is 1-based: lines(1)=header, lines(2..n)=data rows.
    Dim wantSheet As String: wantSheet = Trim(sheetNum)
    Dim i As Integer
    For i = 2 To n
        Dim fields() As String
        fields = Split(lines(i), vbTab)
        If UBound(fields) >= 0 Then
            If Trim(fields(0)) = wantSheet Then
                ReDim rows(0 To 1)
                rows(0) = "sheetNum" & vbTab & "title" & vbTab & "roadType" & vbTab & _
                          "duration" & vbTab & "signs" & vbTab & "elements" & vbTab & "notes"
                rows(1) = lines(i)
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
        fields = Split(lines(i), vbTab)
        If UBound(fields) >= 1 Then
            rows(i - 1) = fields(0) & vbTab & fields(1)
        End If
    Next i
    ListRegisteredSheets = rows
End Function

' ============================================================
' FILE I/O HELPER — ADODB.Stream whole-file read
' ============================================================
Private Function ReadAllLines(path As String, ByRef outLines() As String) As Integer
    On Error GoTo ReadErr
    If Dir(path) = "" Then
        ReadAllLines = 0
        Exit Function
    End If

    ' Same approach as WZTCChatTimer: whole-file read via ADODB.Stream.
    ' Avoids Do While Not EOF + Line Input # quirks on this VBA host.
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    Dim content As String
    content = stream.ReadText(-1) ' adReadAll
    stream.Close

    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    Dim parts() As String
    parts = Split(content, vbLf)

    Dim n As Integer: n = 0
    Dim i As Integer
    For i = LBound(parts) To UBound(parts)
        Dim ln As String
        ln = parts(i)
        If Len(Trim(ln)) > 0 Then
            n = n + 1
            ReDim Preserve outLines(1 To n)
            outLines(n) = ln
        End If
    Next i
    ReadAllLines = n
    Exit Function

ReadErr:
    ReadAllLines = 0
End Function
