Option Explicit

' ============================================================
' BBMarkupParser.bas
' Keyword-based classification of markup comment text.
' Reads RawText from each BBMarkup, sets OpType + Param1 + Param2.
'
' Rules are evaluated top-to-bottom; first match wins.
' All matching is case-insensitive (text is uppercased before compare).
'
' Operation keywords:
'   DELETE  + CALLOUT → DELETE_CALLOUT
'   DELETE / REMOVE   → DELETE
'   MOVE / SHIFT      → MOVE       Param1=direction  Param2=distance(ft)
'   ADD CELL / PLACE CELL / INSERT CELL → ADD_CELL   Param1=cellName
'   LEVEL / CHANGE LEVEL / MOVE TO LEVEL → CHANGE_LEVEL  Param1=levelName
'   EDIT TEXT / UPDATE TEXT / CHANGE TEXT / "TEXT:" → EDIT_TEXT  Param1=newText
'   DIMENSION / DIM   → ADD_DIMENSION
'   ADD CALLOUT / NEW CALLOUT / CALLOUT → ADD_CALLOUT  Param1=calloutText
'   (none match)      → UNKNOWN
' ============================================================

' ============================================================
' ParseAllMarkups — classify every loaded markup
' ============================================================
Public Sub ParseAllMarkups()
    Dim i As Integer
    For i = 1 To bbMarkupCount
        ParseSingleMarkup i
    Next i
End Sub

' ============================================================
' ParseSingleMarkup — classify one markup by index
' ============================================================
Public Sub ParseSingleMarkup(idx As Integer)
    Dim m As BBMarkup
    m = bbMarkups(idx)

    If Len(Trim(m.RawText)) = 0 Then
        m.OpType = "UNKNOWN"
        bbMarkups(idx) = m
        Exit Sub
    End If

    Dim txt As String
    txt = UCase(m.RawText)   ' case-insensitive matching

    ' ---- Rule 1: DELETE CALLOUT (must test before plain DELETE) ----
    If (InStr(txt, "DELETE") > 0 Or InStr(txt, "REMOVE") > 0) And InStr(txt, "CALLOUT") > 0 Then
        m.OpType = "DELETE_CALLOUT"

    ' ---- Rule 2: DELETE / REMOVE ----
    ElseIf InStr(txt, "DELETE") > 0 Or InStr(txt, "REMOVE") > 0 Then
        m.OpType = "DELETE"

    ' ---- Rule 3: MOVE / SHIFT ----
    ElseIf InStr(txt, "MOVE") > 0 Or InStr(txt, "SHIFT") > 0 Then
        m.OpType = "MOVE"
        m.Param1 = ExtractDirection(m.RawText)   ' direction keyword
        m.Param2 = ExtractDistanceFt(m.RawText)  ' numeric distance in feet

    ' ---- Rule 4: ADD CELL / PLACE CELL / INSERT CELL ----
    ElseIf InStr(txt, "ADD CELL") > 0 Or InStr(txt, "PLACE CELL") > 0 Or _
           InStr(txt, "INSERT CELL") > 0 Or InStr(txt, "ADD_CELL") > 0 Then
        m.OpType = "ADD_CELL"
        m.Param1 = ExtractCellName(m.RawText)

    ' ---- Rule 5: CHANGE LEVEL / MOVE TO LEVEL / LEVEL ----
    ElseIf InStr(txt, "CHANGE LEVEL") > 0 Or InStr(txt, "MOVE TO LEVEL") > 0 Or _
           InStr(txt, "LEVEL:") > 0 Or InStr(txt, "LEVEL =") > 0 Or _
           InStr(txt, "LEVEL=") > 0 Then
        m.OpType = "CHANGE_LEVEL"
        m.Param1 = ExtractLevelName(m.RawText)

    ' LEVEL alone (looser — after the compound forms above)
    ElseIf InStr(txt, "LEVEL") > 0 Then
        m.OpType = "CHANGE_LEVEL"
        m.Param1 = ExtractLevelName(m.RawText)

    ' ---- Rule 6: EDIT TEXT / UPDATE TEXT / CHANGE TEXT / TEXT: ----
    ElseIf InStr(txt, "EDIT TEXT") > 0 Or InStr(txt, "UPDATE TEXT") > 0 Or _
           InStr(txt, "CHANGE TEXT") > 0 Or InStr(txt, "TEXT:") > 0 Then
        m.OpType = "EDIT_TEXT"
        m.Param1 = ExtractNewText(m.RawText)

    ' ---- Rule 7: DIMENSION / DIM ----
    ElseIf InStr(txt, "DIMENSION") > 0 Or InStr(txt, " DIM ") > 0 Or _
           Left(txt, 4) = "DIM " Or Right(txt, 4) = " DIM" Then
        m.OpType = "ADD_DIMENSION"

    ' ---- Rule 8: ADD CALLOUT / NEW CALLOUT / CALLOUT ----
    ElseIf InStr(txt, "ADD CALLOUT") > 0 Or InStr(txt, "NEW CALLOUT") > 0 Or _
           InStr(txt, "CALLOUT") > 0 Then
        m.OpType = "ADD_CALLOUT"
        m.Param1 = ExtractCalloutText(m.RawText)

    ' ---- Default ----
    Else
        m.OpType = "UNKNOWN"
    End If

    bbMarkups(idx) = m
End Sub

' ============================================================
' ExtractCellName
' Finds the word immediately after the keyword CELL.
' e.g. "Add cell TWZAP_P here" → "TWZAP_P"
' ============================================================
Private Function ExtractCellName(txt As String) As String
    Dim upper As String
    upper = UCase(txt)

    Dim pos As Integer
    pos = InStr(upper, "CELL")
    If pos = 0 Then
        ExtractCellName = ""
        Exit Function
    End If

    pos = pos + 4   ' move past "CELL"
    ExtractCellName = ExtractNextWord(txt, pos)
End Function

' ============================================================
' ExtractLevelName
' Finds the level name after LEVEL keyword.
' Handles: "level TWZBT_P", "level: TWZBT_P", "level=TWZBT_P",
'          "change level to TWZBT_P"
' ============================================================
Private Function ExtractLevelName(txt As String) As String
    Dim upper As String
    upper = UCase(txt)

    Dim pos As Integer
    pos = InStr(upper, "LEVEL")
    If pos = 0 Then
        ExtractLevelName = ""
        Exit Function
    End If

    pos = pos + 5   ' move past "LEVEL"

    ' Skip optional delimiter chars: ':', '=', ' ', "TO "
    Do While pos <= Len(txt)
        Dim c As String
        c = Mid(txt, pos, 1)
        If c = ":" Or c = "=" Or c = " " Then
            pos = pos + 1
        ElseIf UCase(Mid(txt, pos, 3)) = "TO " Then
            pos = pos + 3
        Else
            Exit Do
        End If
    Loop

    ExtractLevelName = ExtractNextWord(txt, pos)
End Function

' ============================================================
' ExtractNewText
' Returns the text to set on the element.
' Looks for content after "TEXT:", or between double quotes,
' or after the keyword phrase (EDIT TEXT, UPDATE TEXT, etc.)
' ============================================================
Private Function ExtractNewText(txt As String) As String
    ' Try: content between double quotes
    Dim q1 As Integer, q2 As Integer
    q1 = InStr(txt, Chr(34))
    If q1 > 0 Then
        q2 = InStr(q1 + 1, txt, Chr(34))
        If q2 > q1 Then
            ExtractNewText = Mid(txt, q1 + 1, q2 - q1 - 1)
            Exit Function
        End If
    End If

    ' Try: after "TEXT:"
    Dim upper As String
    upper = UCase(txt)
    Dim pos As Integer
    pos = InStr(upper, "TEXT:")
    If pos > 0 Then
        pos = pos + 5
        Do While pos <= Len(txt) And Mid(txt, pos, 1) = " "
            pos = pos + 1
        Loop
        ExtractNewText = Trim(Mid(txt, pos))
        Exit Function
    End If

    ' Fallback: grab everything after the keyword phrase
    Dim kwEnd As Integer
    kwEnd = 0
    Dim phrases(4) As String
    phrases(0) = "EDIT TEXT"
    phrases(1) = "UPDATE TEXT"
    phrases(2) = "CHANGE TEXT"
    Dim p As Integer
    For p = 0 To 2
        pos = InStr(upper, phrases(p))
        If pos > 0 Then
            kwEnd = pos + Len(phrases(p))
            Exit For
        End If
    Next p

    If kwEnd > 0 Then
        Do While kwEnd <= Len(txt) And (Mid(txt, kwEnd, 1) = " " Or Mid(txt, kwEnd, 1) = ":")
            kwEnd = kwEnd + 1
        Loop
        ExtractNewText = Trim(Mid(txt, kwEnd))
    Else
        ExtractNewText = ""
    End If
End Function

' ============================================================
' ExtractCalloutText
' Returns the text to place in a new callout.
' Looks for content after "CALLOUT:" or between quotes.
' ============================================================
Private Function ExtractCalloutText(txt As String) As String
    ' Try quotes first
    Dim q1 As Integer, q2 As Integer
    q1 = InStr(txt, Chr(34))
    If q1 > 0 Then
        q2 = InStr(q1 + 1, txt, Chr(34))
        If q2 > q1 Then
            ExtractCalloutText = Mid(txt, q1 + 1, q2 - q1 - 1)
            Exit Function
        End If
    End If

    ' Try after "CALLOUT:"
    Dim upper As String
    upper = UCase(txt)
    Dim pos As Integer
    pos = InStr(upper, "CALLOUT:")
    If pos > 0 Then
        pos = pos + 8
        Do While pos <= Len(txt) And Mid(txt, pos, 1) = " "
            pos = pos + 1
        Loop
        ExtractCalloutText = Trim(Mid(txt, pos))
        Exit Function
    End If

    ExtractCalloutText = ""
End Function

' ============================================================
' ExtractDirection
' Returns N, S, E, W, LEFT, RIGHT, UP, or DOWN from the text.
' Used for MOVE operations.
' ============================================================
Private Function ExtractDirection(txt As String) As String
    Dim upper As String
    upper = UCase(txt)

    Dim dirs(7) As String
    dirs(0) = "NORTH":  dirs(1) = "SOUTH": dirs(2) = "EAST":  dirs(3) = "WEST"
    dirs(4) = "LEFT":   dirs(5) = "RIGHT": dirs(6) = "UP":    dirs(7) = "DOWN"

    Dim short(7) As String
    short(0) = "N": short(1) = "S": short(2) = "E": short(3) = "W"
    short(4) = "LEFT": short(5) = "RIGHT": short(6) = "UP": short(7) = "DOWN"

    Dim d As Integer
    For d = 0 To 7
        If InStr(upper, dirs(d)) > 0 Then
            ExtractDirection = short(d)
            Exit Function
        End If
    Next d

    ExtractDirection = ""
End Function

' ============================================================
' ExtractDistanceFt
' Finds the first numeric value in the text.
' e.g. "Move right 10 feet", "Shift 5.5 ft left" → "10" / "5.5"
' ============================================================
Private Function ExtractDistanceFt(txt As String) As String
    Dim i As Integer
    Dim numStr As String
    numStr = ""
    Dim inNum As Boolean
    inNum = False

    For i = 1 To Len(txt)
        Dim c As String
        c = Mid(txt, i, 1)
        If (c >= "0" And c <= "9") Or (c = "." And inNum) Then
            numStr = numStr & c
            inNum = True
        ElseIf c = "-" And Not inNum Then
            numStr = numStr & c
        ElseIf inNum Then
            Exit For   ' end of number
        End If
    Next i

    ExtractDistanceFt = numStr
End Function

' ============================================================
' ExtractNextWord
' Returns the non-whitespace word starting at position pos in txt.
' Skips leading spaces. Stops at first whitespace after word.
' ============================================================
Private Function ExtractNextWord(txt As String, startPos As Integer) As String
    Dim pos As Integer
    pos = startPos

    ' Skip leading whitespace
    Do While pos <= Len(txt) And Mid(txt, pos, 1) = " "
        pos = pos + 1
    Loop

    ' Read word characters (letters, digits, underscore, hyphen)
    Dim word As String
    word = ""
    Do While pos <= Len(txt)
        Dim c As String
        c = Mid(txt, pos, 1)
        If c = " " Or c = "," Or c = "." Or c = Chr(13) Or c = Chr(10) Then
            Exit Do
        End If
        word = word & c
        pos = pos + 1
    Loop

    ExtractNextWord = word
End Function
