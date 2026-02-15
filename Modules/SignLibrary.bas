Attribute VB_Name = "Module3"
' ============================================================
' SIGN LIBRARY MANAGER - REVISED
' Stores and retrieves sign properties and associated text
'
' IMPORTANT: This must be a STANDARD MODULE
' Uses a workaround to avoid VBA Type parameter restrictions
' ============================================================

Option Explicit

' Array storage for signs (avoids Variant/UDT coercion with Collection)
Private signLibrary() As signData
Private signLibraryCount As Long
Private signLibraryInitialized As Boolean

' ============================================================
' INITIALIZATION
' ============================================================
Public Sub InitializeSignLibrary()
    ReDim signLibrary(1 To 64)
    signLibraryCount = 0
    signLibraryInitialized = True
    Call LoadDefaultSigns
End Sub

' ============================================================
' LOAD DEFAULT SIGNS
' Add all standard MUTCD and NY signs here
' ============================================================
Private Sub LoadDefaultSigns()
    
    ' Example Sign 1: Road Work Ahead (R02-10sNY)
    Call AddSign("R02-10sNY", "Road Work Ahead", "R02-10sNY", _
                 "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel", _
                 "NYR09-11", "48"" x 48""", 48, 48, _
                 "TWZSGN_P", "c:\pwworking\usny\d0119091\ny_plan_wztc.cel", 500)
    
    ' Example Sign 2: Lane Closed Ahead (W20-5)
    Call AddSign("W20-5", "Lane Closed Ahead", "W20-5", _
                 "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel", _
                 "NYR09-11", "48"" x 48""", 48, 48, _
                 "TWZSGN_P", "c:\pwworking\usny\d0119091\ny_plan_wztc.cel", 500)
    
    ' Example Sign 3: Right Lane Closed (W20-5aL)
    Call AddSign("W20-5aL", "Right Lane Closed", "W20-5aL", _
                 "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel", _
                 "NYR09-11", "36"" x 36""", 36, 36, _
                 "TWZSGN_P", "c:\pwworking\usny\d0119091\ny_plan_wztc.cel", 350)
    
    ' Example Sign 4: Road Closed Ahead (W20-3)
    Call AddSign("W20-3", "Road Closed Ahead", "W20-3", _
                 "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel", _
                 "NYR09-11", "48"" x 48""", 48, 48, _
                 "TWZSGN_P", "c:\pwworking\usny\d0119091\ny_plan_wztc.cel", 500)
    
    ' Example Sign 5: Flagger Ahead (W20-7a)
    Call AddSign("W20-7a", "Flagger Ahead", "W20-7a", _
                 "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel", _
                 "NYR09-11", "36"" x 36""", 36, 36, _
                 "TWZSGN_P", "c:\pwworking\usny\d0119091\ny_plan_wztc.cel", 350)
    
    ' Example Sign 6: Be Prepared To Stop (W3-4)
    Call AddSign("W3-4", "Be Prepared To Stop", "W3-4", _
                 "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel", _
                 "NYR09-11", "48"" x 30""", 48, 30, _
                 "TWZSGN_P", "c:\pwworking\usny\d0119091\ny_plan_wztc.cel", 350)
    
    ' Example Sign 7: Speed Limit 25 (R2-1-25)
    Call AddSign("R2-1-25", "Speed Limit 25", "R2-1-25", _
                 "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel", _
                 "NYR09-11", "30"" x 36""", 30, 36, _
                 "TWZSGN_P", "c:\pwworking\usny\d0119091\ny_plan_wztc.cel", 300)
    
    ' Example Sign 8: Workers Present (W21-103)
    Call AddSign("W21-103", "Workers Present", "W21-103", _
                 "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel", _
                 "NYR09-11", "48"" x 48""", 48, 48, _
                 "TWZSGN_P", "c:\pwworking\usny\d0119091\ny_plan_wztc.cel", 500)
    
    ' Example Sign 9: End Road Work (G20-2)
    Call AddSign("G20-2", "End Road Work", "G20-2", _
                 "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel", _
                 "NYR09-11", "36"" x 24""", 36, 24, _
                 "TWZSGN_P", "c:\pwworking\usny\d0119091\ny_plan_wztc.cel", 250)
    
    ' Example Sign 10: Detour (M4-8)
    Call AddSign("M4-8", "Detour", "M4-8", _
                 "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel", _
                 "NYR09-11", "36"" x 18""", 36, 18, _
                 "TWZSGN_P", "c:\pwworking\usny\d0119091\ny_plan_wztc.cel", 300)
    
    ' Add more signs as needed...
    
End Sub

' ============================================================
' ADD SIGN TO LIBRARY (PRIVATE - INTERNAL USE ONLY)
' Uses individual parameters instead of SignData type
' ============================================================
Private Sub AddSign(SignNumber As String, _
                    Description As String, _
                    CellName As String, _
                    CellLibPath As String, _
                    TextLabel As String, _
                    TextLine2 As String, _
                    widthIn As Double, _
                    heightIn As Double, _
                    PostType As String, _
                    PostLibPath As String, _
                    spacing As Double)
    
    Dim sign As signData
    
    ' Build the sign data structure
    sign.SignNumber = SignNumber
    sign.Description = Description
    sign.CellName = CellName
    sign.CellLibraryPath = CellLibPath
    sign.TextLabel = TextLabel
    sign.TextLine2 = TextLine2
    sign.WidthInches = widthIn
    sign.HeightInches = heightIn
    sign.PostType = PostType
    sign.PostLibraryPath = PostLibPath
    sign.DefaultSpacing = spacing
    
    ' Add to array (check duplicate first)
    Dim i As Long
    For i = 1 To signLibraryCount
        If signLibrary(i).SignNumber = sign.SignNumber Then
            Debug.Print "Warning: Sign " & sign.SignNumber & " already exists in library"
            Exit Sub
        End If
    Next i
    signLibraryCount = signLibraryCount + 1
    If signLibraryCount > UBound(signLibrary) Then
        ReDim Preserve signLibrary(1 To UBound(signLibrary) + 64)
    End If
    signLibrary(signLibraryCount) = sign
End Sub

' ============================================================
' GET SIGN FROM LIBRARY
' Returns sign data for a given sign number (UDT-to-UDT, no Variant)
' ============================================================
Public Function GetSignData(SignNumber As String) As signData
    Dim sign As signData
    Dim i As Long
    
    If Not signLibraryInitialized Then
        Call InitializeSignLibrary
    End If
    
    For i = 1 To signLibraryCount
        If signLibrary(i).SignNumber = SignNumber Then
            sign = signLibrary(i)
            GetSignData = sign
            Exit Function
        End If
    Next i
    
    ' Not found
    sign.SignNumber = ""
    sign.Description = "Sign not found"
    Debug.Print "Warning: Sign " & SignNumber & " not found in library"
    GetSignData = sign
End Function

' ============================================================
' CHECK IF SIGN EXISTS
' ============================================================
Public Function SignExists(SignNumber As String) As Boolean
    Dim i As Long
    
    If Not signLibraryInitialized Then
        Call InitializeSignLibrary
    End If
    
    For i = 1 To signLibraryCount
        If signLibrary(i).SignNumber = SignNumber Then
            SignExists = True
            Exit Function
        End If
    Next i
    SignExists = False
End Function

' ============================================================
' GET ALL SIGN NUMBERS
' Returns an array of all sign numbers in the library
' ============================================================
Public Function GetAllSignNumbers() As String()
    Dim signNumbers() As String
    Dim i As Long
    
    If Not signLibraryInitialized Then
        Call InitializeSignLibrary
    End If
    
    If signLibraryCount = 0 Then
        ReDim signNumbers(0)
        signNumbers(0) = ""
        GetAllSignNumbers = signNumbers
        Exit Function
    End If
    
    ReDim signNumbers(1 To signLibraryCount)
    For i = 1 To signLibraryCount
        signNumbers(i) = signLibrary(i).SignNumber
    Next i
    GetAllSignNumbers = signNumbers
End Function

' ============================================================
' GET LIBRARY COUNT
' ============================================================
Public Function GetSignCount() As Integer
    If Not signLibraryInitialized Then
        Call InitializeSignLibrary
    End If
    GetSignCount = CInt(signLibraryCount)
End Function

' ============================================================
' ADD CUSTOM SIGN (PUBLIC - for runtime additions)
' ============================================================
Public Sub AddCustomSign(SignNumber As String, _
                        Description As String, _
                        CellName As String, _
                        CellLibPath As String, _
                        TextLabel As String, _
                        TextLine2 As String, _
                        widthIn As Double, _
                        heightIn As Double, _
                        PostType As String, _
                        PostLibPath As String, _
                        spacing As Double)
    
    If Not signLibraryInitialized Then
        Call InitializeSignLibrary
    End If
    
    ' Use the private AddSign method
    Call AddSign(SignNumber, Description, CellName, CellLibPath, _
                 TextLabel, TextLine2, widthIn, heightIn, _
                 PostType, PostLibPath, spacing)
End Sub

' ============================================================
' EXPORT LIBRARY TO TEXT FILE (for backup/reference)
' ============================================================
Public Sub ExportLibraryToFile(filePath As String)
    Dim fileNum As Integer
    Dim i As Long
    
    If Not signLibraryInitialized Then
        Call InitializeSignLibrary
    End If
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    Print #fileNum, "SIGN LIBRARY EXPORT"
    Print #fileNum, "===================="
    Print #fileNum, ""
    Print #fileNum, "Total Signs: " & signLibraryCount
    Print #fileNum, ""
    
    For i = 1 To signLibraryCount
        With signLibrary(i)
            Print #fileNum, "Sign Number: " & .SignNumber
            Print #fileNum, "Description: " & .Description
            Print #fileNum, "Cell Name: " & .CellName
            Print #fileNum, "Cell Library: " & .CellLibraryPath
            Print #fileNum, "Text Label: " & .TextLabel
            Print #fileNum, "Text Line 2: " & .TextLine2
            Print #fileNum, "Dimensions: " & .WidthInches & """W x " & .HeightInches & """H"
            Print #fileNum, "Post Type: " & .PostType
            Print #fileNum, "Post Library: " & .PostLibraryPath
            Print #fileNum, "Default Spacing: " & .DefaultSpacing & " ft"
        End With
        Print #fileNum, "--------------------"
    Next i
    
    Close #fileNum
    
    Debug.Print "Library exported to: " & filePath
End Sub

' ============================================================
' CLEAR LIBRARY (for reinitialization)
' ============================================================
Public Sub ClearLibrary()
    signLibraryCount = 0
    signLibraryInitialized = True
    ReDim signLibrary(1 To 1)
End Sub

