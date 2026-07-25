Option Explicit

' ============================================================
' BBFileParser.bas
' Parses Bluebeam markup export files into the bbMarkups() array.
'
' Supported formats:
'   .xml  — Bluebeam XML markup summary (standard export)
'   .csv  — Custom Bluebeam CSV export with optional PdfX/PdfY columns
'
' After parsing, call BBMarkupParser.ParseAllMarkups to classify
' each markup's operation type, then call
' BBCoordCalibrator.ConvertAllMarkups to fill MstnX/MstnY.
' ============================================================

' ============================================================
' ParseMarkupFile — top-level entry point
' Returns True on success, False on error (bbMarkupCount = 0)
' ============================================================
Public Function ParseMarkupFile(filePath As String) As Boolean
    On Error GoTo ParseErr

    bbMarkupCount = 0
    bbLoadedFilePath = ""

    If Len(filePath) = 0 Then
        ParseMarkupFile = False
        Exit Function
    End If

    Dim ext As String
    ext = LCase(Right(filePath, 4))

    If ext = ".xml" Then
        ParseXML filePath
    ElseIf ext = ".csv" Then
        ParseCSV filePath
    Else
        MsgBox "Unsupported file type. Please select a .xml or .csv file.", vbExclamation, "Bluebeam Importer"
        ParseMarkupFile = False
        Exit Function
    End If

    bbLoadedFilePath = filePath
    ParseMarkupFile = (bbMarkupCount > 0)
    Exit Function

ParseErr:
    MsgBox "Error parsing file: " & Err.Description, vbCritical, "Bluebeam Importer"
    ParseMarkupFile = False
End Function

' ============================================================
' ParseXML — Bluebeam standard XML markup summary
'
' Expected structure (Bluebeam Revu export):
'   <PDFMarkupReport>
'     <Markups>
'       <Markup>
'         <Page>1</Page>
'         <Author>...</Author>
'         <Type>Callout</Type>
'         <Subject>...</Subject>
'         <Comments>Delete this element</Comments>
'         <PdfX>234.5</PdfX>     (optional — not in standard export)
'         <PdfY>456.7</PdfY>     (optional)
'       </Markup>
'       ...
'     </Markups>
'   </PDFMarkupReport>
'
' Note: standard Bluebeam XML does not include X/Y coordinates.
' If coordinates are absent, HasCoord = False and the processor
' will prompt the user to click the location manually.
' ============================================================
Private Sub ParseXML(filePath As String)
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False

    If Not xmlDoc.Load(filePath) Then
        Dim parseError As Object
        Set parseError = xmlDoc.parseError
        MsgBox "XML parse error at line " & parseError.Line & ": " & parseError.reason, _
               vbCritical, "Bluebeam Importer"
        Exit Sub
    End If

    ' Try both <Markup> and <markup> (case variants)
    Dim nodes As Object
    Set nodes = xmlDoc.SelectNodes("//Markup")
    If nodes Is Nothing Then
        Set nodes = xmlDoc.SelectNodes("//markup")
    End If
    If nodes Is Nothing Or nodes.Length = 0 Then
        MsgBox "No <Markup> nodes found in the XML file. " & _
               "Export from Bluebeam via File > Export > Markup Summary > XML.", _
               vbExclamation, "Bluebeam Importer"
        Exit Sub
    End If

    ReDim bbMarkups(1 To nodes.Length)
    Dim i As Integer
    Dim n As Object

    For i = 0 To nodes.Length - 1
        Set n = nodes.Item(i)
        bbMarkupCount = bbMarkupCount + 1
        Dim m As BBMarkup

        m.MarkupID = "M" & Format(bbMarkupCount, "000")
        m.RawText = Trim(GetNodeText(n, "Comments"))
        If Len(m.RawText) = 0 Then
            m.RawText = Trim(GetNodeText(n, "Subject"))
        End If
        m.OpType = "UNKNOWN"
        m.Status = "Pending"
        m.HasCoord = False

        ' Try to read optional X/Y nodes
        Dim xStr As String, yStr As String
        xStr = GetNodeText(n, "PdfX")
        yStr = GetNodeText(n, "PdfY")
        If Len(xStr) > 0 And Len(yStr) > 0 Then
            On Error Resume Next
            m.PdfX = CDbl(xStr)
            m.PdfY = CDbl(yStr)
            If Err.Number = 0 Then m.HasCoord = True
            On Error GoTo 0
        End If

        ' Fallback: scan raw text for inline coordinates
        If Not m.HasCoord Then
            TryExtractInlineCoords m
        End If

        bbMarkups(bbMarkupCount) = m
    Next i
End Sub

' ============================================================
' ParseCSV — custom Bluebeam CSV export
'
' Expected column order (header row required):
'   Page, Author, Date, Type, Comments, Subject, PdfX, PdfY
'
' PdfX and PdfY columns are optional. The parser detects them
' by reading the header row and mapping column names.
' ============================================================
Private Sub ParseCSV(filePath As String)
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Input As #fNum

    If EOF(fNum) Then
        Close #fNum
        MsgBox "CSV file is empty.", vbExclamation, "Bluebeam Importer"
        Exit Sub
    End If

    ' Read header row to detect column positions
    Dim headerLine As String
    Line Input #fNum, headerLine
    Dim headers() As String
    headers = SplitCSVLine(headerLine)

    Dim colComments As Integer, colSubject As Integer
    Dim colPdfX As Integer, colPdfY As Integer
    colComments = -1: colSubject = -1: colPdfX = -1: colPdfY = -1

    Dim h As Integer
    For h = 0 To UBound(headers)
        Dim hName As String
        hName = LCase(Trim(headers(h)))
        Select Case hName
            Case "comments":  colComments = h
            Case "subject":   colSubject = h
            Case "pdfx":      colPdfX = h
            Case "pdfy":      colPdfY = h
            Case "x":         If colPdfX = -1 Then colPdfX = h
            Case "y":         If colPdfY = -1 Then colPdfY = h
        End Select
    Next h

    ' Count lines first for ReDim
    Dim lineCount As Integer
    lineCount = 0
    Do While Not EOF(fNum)
        Dim tmp As String
        Line Input #fNum, tmp
        If Len(Trim(tmp)) > 0 Then lineCount = lineCount + 1
    Loop
    Close #fNum

    If lineCount = 0 Then
        MsgBox "CSV has a header but no data rows.", vbExclamation, "Bluebeam Importer"
        Exit Sub
    End If

    ReDim bbMarkups(1 To lineCount)

    ' Re-open and read data rows
    fNum = FreeFile
    Open filePath For Input As #fNum
    Line Input #fNum, headerLine  ' skip header again

    Do While Not EOF(fNum)
        Dim dataLine As String
        Line Input #fNum, dataLine
        If Len(Trim(dataLine)) = 0 Then GoTo NextLine

        Dim cols() As String
        cols = SplitCSVLine(dataLine)

        bbMarkupCount = bbMarkupCount + 1
        Dim m As BBMarkup
        m.MarkupID = "M" & Format(bbMarkupCount, "000")
        m.Status = "Pending"
        m.OpType = "UNKNOWN"
        m.HasCoord = False

        If colComments >= 0 And colComments <= UBound(cols) Then
            m.RawText = Trim(cols(colComments))
        End If
        If Len(m.RawText) = 0 And colSubject >= 0 And colSubject <= UBound(cols) Then
            m.RawText = Trim(cols(colSubject))
        End If

        If colPdfX >= 0 And colPdfY >= 0 Then
            If colPdfX <= UBound(cols) And colPdfY <= UBound(cols) Then
                Dim xStr As String, yStr As String
                xStr = Trim(cols(colPdfX))
                yStr = Trim(cols(colPdfY))
                If Len(xStr) > 0 And Len(yStr) > 0 Then
                    On Error Resume Next
                    m.PdfX = CDbl(xStr)
                    m.PdfY = CDbl(yStr)
                    If Err.Number = 0 Then m.HasCoord = True
                    On Error GoTo 0
                End If
            End If
        End If

        If Not m.HasCoord Then TryExtractInlineCoords m

        bbMarkups(bbMarkupCount) = m
NextLine:
    Loop
    Close #fNum
End Sub

' ============================================================
' TryExtractInlineCoords
' Scans markup text for patterns like:
'   X: 1234.5, Y: 5678.9
'   X=1234 Y=5678
'   x:1234.5 y:5678.9
' If found, sets m.PdfX, m.PdfY, m.HasCoord = True
' ============================================================
Private Sub TryExtractInlineCoords(ByRef m As BBMarkup)
    Dim txt As String
    txt = m.RawText

    Dim xVal As Double, yVal As Double
    Dim found As Boolean
    found = False

    ' Pattern: X: <number> and Y: <number> (or X= Y=)
    Dim xPos As Integer, yPos As Integer
    xPos = InStr(1, LCase(txt), "x:")
    If xPos = 0 Then xPos = InStr(1, LCase(txt), "x=")
    yPos = InStr(1, LCase(txt), "y:")
    If yPos = 0 Then yPos = InStr(1, LCase(txt), "y=")

    If xPos > 0 And yPos > 0 Then
        On Error Resume Next
        Dim xStr As String, yStr As String
        xStr = ExtractNumberAfter(txt, xPos + 2)
        yStr = ExtractNumberAfter(txt, yPos + 2)
        If Len(xStr) > 0 And Len(yStr) > 0 Then
            xVal = CDbl(xStr)
            yVal = CDbl(yStr)
            If Err.Number = 0 Then found = True
        End If
        On Error GoTo 0
    End If

    If found Then
        m.PdfX = xVal
        m.PdfY = yVal
        m.HasCoord = True
    End If
End Sub

' ============================================================
' Helper: extract numeric string starting at position pos in txt
' Skips leading whitespace; reads digits, decimal point, minus sign
' ============================================================
Private Function ExtractNumberAfter(txt As String, startPos As Integer) As String
    Dim i As Integer
    Dim result As String
    result = ""
    i = startPos
    ' Skip whitespace
    Do While i <= Len(txt) And Mid(txt, i, 1) = " "
        i = i + 1
    Loop
    ' Read number chars
    Do While i <= Len(txt)
        Dim c As String
        c = Mid(txt, i, 1)
        If c >= "0" And c <= "9" Then
            result = result & c
        ElseIf c = "." Or c = "-" Then
            result = result & c
        Else
            Exit Do
        End If
        i = i + 1
    Loop
    ExtractNumberAfter = result
End Function

' ============================================================
' Helper: get text content of a named child node
' Returns "" if node not found
' ============================================================
Private Function GetNodeText(parent As Object, nodeName As String) As String
    On Error Resume Next
    Dim child As Object
    Set child = parent.SelectSingleNode(nodeName)
    If Not child Is Nothing Then
        GetNodeText = Trim(child.Text)
    Else
        GetNodeText = ""
    End If
    On Error GoTo 0
End Function

' ============================================================
' Helper: split a CSV line respecting quoted fields
' e.g. 1,"Hello, world","Test",42 → ["1","Hello, world","Test","42"]
' ============================================================
Private Function SplitCSVLine(line As String) As String()
    Dim result() As String
    Dim count As Integer
    count = 0
    ReDim result(0 To 50)

    Dim i As Integer
    Dim inQuote As Boolean
    Dim field As String
    inQuote = False
    field = ""

    For i = 1 To Len(line)
        Dim c As String
        c = Mid(line, i, 1)
        If c = Chr(34) Then         ' double-quote
            If inQuote And i < Len(line) And Mid(line, i + 1, 1) = Chr(34) Then
                field = field & Chr(34)   ' escaped quote: ""
                i = i + 1
            Else
                inQuote = Not inQuote
            End If
        ElseIf c = "," And Not inQuote Then
            If count > UBound(result) Then ReDim Preserve result(0 To count + 20)
            result(count) = field
            count = count + 1
            field = ""
        Else
            field = field & c
        End If
    Next i
    ' Last field
    If count > UBound(result) Then ReDim Preserve result(0 To count + 1)
    result(count) = field
    count = count + 1

    ReDim Preserve result(0 To count - 1)
    SplitCSVLine = result
End Function
