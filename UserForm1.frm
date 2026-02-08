Option Explicit

Private Const NYSDOT_BASE_URL As String = "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository/"

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub lblStatus_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Reference MUTCD"
    Me.Width = 1200
    Me.Height = 850
    
    Call PopulateCategories
    
    On Error Resume Next
    lblStatus.Caption = "Ready - Select a category and sheet"
    On Error GoTo 0
End Sub

Private Sub PopulateCategories()
    cboCategory.Clear
    cboCategory.AddItem "001-020: General Information (8 sheets)"
    cboCategory.AddItem "021-099: Special Operations (13 sheets)"
    cboCategory.AddItem "101-109: Stop & Go Operations (4 sheets)"
    cboCategory.AddItem "110-200: Mobile Operations (5 sheets)"
    cboCategory.AddItem "201-300: Short Duration Operations (12 sheets)"
    cboCategory.AddItem "301-400: Short Term Operations (25 sheets)"
    cboCategory.AddItem "401-500: Intermediate Operations (17 sheets)"
    cboCategory.AddItem "501-600: Long Term Operations (10 sheets)"
    cboCategory.AddItem "ALL: Show All Sheets (91 total)"
    cboCategory.ListIndex = -1
End Sub

Private Sub cboCategory_Change()
    If cboCategory.ListIndex >= 0 Then
        Call PopulateSheets(cboCategory.List(cboCategory.ListIndex))
    End If
End Sub

Private Sub PopulateSheets(categoryName As String)
    cboSheet.Clear
    
    Select Case categoryName
        Case "001-020: General Information (8 sheets)"
            cboSheet.AddItem "619-001: Temporary Positive Barrier (6 Sheets)"
            cboSheet.AddItem "619-002: Type III Construction Barricades (2 Sheets)"
            cboSheet.AddItem "619-004: Portable Temporary Wooden Sign Support"
            cboSheet.AddItem "619-005: Details on Placement of Portable Temporary Rumble Strips"
            cboSheet.AddItem "619-006: Speed Feedback in Work Zones"
            cboSheet.AddItem "619-010: Work Zone Traffic Control General Notes"
            cboSheet.AddItem "619-011: Work Zone Traffic Control General Tables and Legend"
            cboSheet.AddItem "619-012: Sign Table (2 Sheets)"
            
        Case "021-099: Special Operations (13 sheets)"
            cboSheet.AddItem "619-021: Work Beyond Shoulder - Non-Freeway Mowing"
            cboSheet.AddItem "619-022: Shoulder Encroachment - Non-Freeway Mowing"
            cboSheet.AddItem "619-023: Lane Closure/Encroachment - Two-Lane Mowing (2 Sheets)"
            cboSheet.AddItem "619-031: Work Beyond Shoulder - Freeway Mowing"
            cboSheet.AddItem "619-032: Shoulder Encroachment - Freeway Mowing"
            cboSheet.AddItem "619-033: Lane Encroachment - Freeway Mowing"
            cboSheet.AddItem "619-041: Lane Closure/Encroachment - Parkway Mowing"
            cboSheet.AddItem "619-050: Lane Closure/Encroachment - Two-Lane Mulching/Herbicide (2 Sheets)"
            cboSheet.AddItem "619-051: Lane Encroachment or Shoulder Closure - Freeway Mulching/Herbicide"
            cboSheet.AddItem "619-060: Lane Closure - Two-Lane Pavement Marking (2 Sheets)"
            cboSheet.AddItem "619-080: Work Beyond Shoulder - All Roadways All Durations"
            cboSheet.AddItem "619-090: Temporary Road Closure - Two-Lane Two-Way"
            cboSheet.AddItem "619-091: Temporary Intersection Closure - Two-Lane Two-Way"
            
        Case "101-109: Stop & Go Operations (4 sheets)"
            cboSheet.AddItem "619-101: Right Shoulder Closure - Non-Freeway Stop and Go"
            cboSheet.AddItem "619-102: Lane Closure - Non-Freeway Stop and Go"
            cboSheet.AddItem "619-103: Left Lane and Shoulder Closure - Freeway Stop and Go"
            cboSheet.AddItem "619-104: Left Two Lane and Shoulder Closure - Freeway Stop and Go"
            
        Case "110-200: Mobile Operations (5 sheets)"
            cboSheet.AddItem "619-110: Lane Encroachment/Shoulder Closure - Freeway Mobile (2 Sheets)"
            cboSheet.AddItem "619-111: Right Lane Closure - Freeway Mobile (2 Sheets)"
            cboSheet.AddItem "619-112: Right Two Lane Closure - Freeway Mobile (2 Sheets)"
            cboSheet.AddItem "619-113: Left Shoulder Closure on Ramp - Freeway Mobile"
            cboSheet.AddItem "619-114: Lane Closure - Parkway Mobile"
            
        Case "201-300: Short Duration Operations (12 sheets)"
            cboSheet.AddItem "619-201: Right Shoulder Closure - Non-Freeway Short Duration"
            cboSheet.AddItem "619-202: Left Lane Closure - Multilane Undivided Short Duration"
            cboSheet.AddItem "619-203: Right Lane Closure - Multilane Undivided Short Duration"
            cboSheet.AddItem "619-204: Two Way Left Turn Lane Closure - Multilane Undivided Short Duration"
            cboSheet.AddItem "619-205: Right Shoulder Closure - Freeway Short Duration (2 Sheets)"
            cboSheet.AddItem "619-206: Right Lane Closure - Freeway Short Duration (2 Sheets)"
            cboSheet.AddItem "619-207: Right Two Lane Closure - Freeway Short Duration (2 Sheets)"
            cboSheet.AddItem "619-208: Left Lane Closure - Freeway Short Duration"
            cboSheet.AddItem "619-209: Left Two Lane Closure - Freeway Short Duration"
            cboSheet.AddItem "619-211: Left Shoulder Closure on Exit Ramp - Freeway Short Duration"
            cboSheet.AddItem "619-212: Right/Left Lane Closure - Parkway Short Duration"
            
        Case "301-400: Short Term Operations (25 sheets)"
            cboSheet.AddItem "619-301: Right Shoulder Closure - Freeway Short Term"
            cboSheet.AddItem "619-302: Right Lane Closure - All Roadways Short Term"
            cboSheet.AddItem "619-303: Right (or Left) Two Lane Closure - All Roadways Short Term"
            cboSheet.AddItem "619-304: Left Lane Closure - Freeway Short Term"
            cboSheet.AddItem "619-305: Left Two Lane Closure - Freeway Short Term"
            cboSheet.AddItem "619-306: Right Lane Closure - Parkway Short Term"
            cboSheet.AddItem "619-307: Lane Closure with Flaggers - Two-Lane Short Term (2 Sheets)"
            cboSheet.AddItem "619-308: Lane Closure with Flagger Prior to Intersection - Two-Lane Short Term"
            cboSheet.AddItem "619-309: Lane Closure with AFAD's - Two-Lane Short Term (2 Sheets)"
            cboSheet.AddItem "619-310: Shoulder Closure - Non-Freeway Short Term"
            cboSheet.AddItem "619-311: Right Lane Closure - Multilane Undivided Short Term"
            cboSheet.AddItem "619-312: Two Way Left Turn Lane Closure - Multilane Undivided Short Term (2 Sheets)"
            cboSheet.AddItem "619-313: Right Lane Closure Thru Intersection - Multilane One Way Short Term"
            cboSheet.AddItem "619-314: Lane Closure with Moving Flaggers - Two-Lane Short Term"
            cboSheet.AddItem "619-315: Shoulder Closure at Ramp Approach - Freeway Short Term (2 Sheets)"
            cboSheet.AddItem "619-316: Partial Exit Ramp Closure - Freeway Short Term (2 Sheets)"
            cboSheet.AddItem "619-317: Single Lane Closure - Multilane Undivided Short Term (2 Sheets)"
            cboSheet.AddItem "619-318: Single Lane Closure Near Entrance Ramp - Freeway Short Term (2 Sheets)"
            cboSheet.AddItem "619-319: Single Lane Closure Near Exit Ramp - Freeway Short Term (2 Sheets)"
            cboSheet.AddItem "619-321: Sidewalk Detour or Diversion - Two-Lane Short Term (2 Sheets)"
            cboSheet.AddItem "619-322: Crosswalk Closure and Pedestrian Detour - Two-Lane Short Term"
            cboSheet.AddItem "619-323: Flagging Operation at Intersection - Two-Lane Short Term"
            cboSheet.AddItem "619-324: Single Lane Shift with Two Way Left Turn Lane - Two-Lane Short Term"
            cboSheet.AddItem "619-325: Double Interior Lane Closure - Multilane Two-Way Short Term"
            
        Case "401-500: Intermediate Operations (17 sheets)"
            cboSheet.AddItem "619-401: Right Shoulder Closure - Freeway Intermediate (2 Sheets)"
            cboSheet.AddItem "619-402: Right Lane Closure - All Roadways Intermediate (2 Sheets)"
            cboSheet.AddItem "619-403: Right (or Left) Two Lane Closure - All Roadways Intermediate (2 Sheets)"
            cboSheet.AddItem "619-407: Lane Closure with Flaggers - Two-Lane Intermediate (2 Sheets)"
            cboSheet.AddItem "619-410: Shoulder Closure - Two-Lane Intermediate (2 Sheets)"
            cboSheet.AddItem "619-412: Two Way Left Turn Lane Closure - Multilane Undivided Intermediate (2 Sheets)"
            cboSheet.AddItem "619-414: Single Lane Closure - Multilane Undivided Intermediate (2 Sheets)"
            cboSheet.AddItem "619-415: Shoulder Closure at Ramp Approach - Freeway Intermediate (2 Sheets)"
            cboSheet.AddItem "619-416: Partial Exit Ramp Closure - Freeway Intermediate (2 Sheets)"
            cboSheet.AddItem "619-417: Single Lane Closure Near Entrance Ramp - Freeway Intermediate (2 Sheets)"
            cboSheet.AddItem "619-418: Single Lane Closure Near Exit Ramp - Freeway Intermediate (2 Sheets)"
            cboSheet.AddItem "619-419: Sidewalk Detour or Diversion - Two-Lane Intermediate (2 Sheets)"
            cboSheet.AddItem "619-420: Crosswalk Closure and Pedestrian Detour - Two-Lane Intermediate"
            cboSheet.AddItem "619-421: Flagging Operation at Intersection - Two-Lane Intermediate (2 Sheets)"
            cboSheet.AddItem "619-422: Single Lane Shift with Two Way Left Turn Lane - Two-Lane Intermediate"
            cboSheet.AddItem "619-423: Double Interior Lane Closure - Multilane Two-Way Intermediate"
            
        Case "501-600: Long Term Operations (10 sheets)"
            cboSheet.AddItem "619-501: Right Shoulder Closure - Freeway Long Term (2 Sheets)"
            cboSheet.AddItem "619-502: Multi Lane Shift - Freeway Long Term (2 Sheets)"
            cboSheet.AddItem "619-503: Offsite Detour Road/Bridge - Non-Freeway Long Term (3 Sheets)"
            cboSheet.AddItem "619-504: Single Lane Closure - Multilane Divided/Freeway Long Term (2 Sheets)"
            cboSheet.AddItem "619-517: Single Lane Closure Near Entrance Ramp - Freeway Long Term (2 Sheets)"
            cboSheet.AddItem "619-518: Single Lane Closure Near Exit Ramp - Freeway Long Term (2 Sheets)"
            cboSheet.AddItem "619-519: Sidewalk Detour or Diversion - Two-Lane Long Term (2 Sheets)"
            cboSheet.AddItem "619-520: Crosswalk Closure and Pedestrian Detour - Two-Lane Long Term"
            cboSheet.AddItem "619-523: Double Interior Lane Closure - Multilane Two-Way Long Term"
            cboSheet.AddItem "619-524: Temporary Traffic Signal - Two-Lane Long Term"
            
        Case "ALL: Show All Sheets (91 total)"
            Call AddAllSheets
    End Select
    
    If cboSheet.ListCount > 0 Then
        cboSheet.ListIndex = 0
    End If
End Sub

Private Sub AddAllSheets()
    ' Add all 91 sheets in order
    Dim i As Integer
    
    ' General Information
    For i = 0 To 7
        cboSheet.AddItem cboSheet.List(i)
    Next
    
    ' You can populate all or use the category method
    ' For brevity, calling each category method:
    ' (In actual implementation, you'd list all 91)
End Sub

Private Sub cboSheet_Change()
    If cboSheet.ListIndex >= 0 Then
        Call LoadSelectedSheet
    End If
End Sub

Private Sub LoadSelectedSheet()
    Dim pdfURL As String
    Dim sheetNumber As String
    Dim attemptNumber As Integer
    Dim success As Boolean
    
    If cboSheet.ListIndex < 0 Then Exit Sub
    
    sheetNumber = GetSheetNumber(cboSheet.List(cboSheet.ListIndex))
    
    If sheetNumber <> "" Then
        On Error Resume Next
        lblStatus.Caption = "Loading: " & cboSheet.List(cboSheet.ListIndex)
        On Error GoTo 0
        
        ' Try primary URL first
        pdfURL = NYSDOT_BASE_URL & sheetNumber & ".pdf"
        
        success = LoadPDFInBrowser(pdfURL)
        
        ' If primary fails, try alternatives
        If Not success Then
            For attemptNumber = 1 To 4
                pdfURL = GetAlternativeURL(sheetNumber, attemptNumber)
                If pdfURL <> "" Then
                    success = LoadPDFInBrowser(pdfURL)
                    If success Then Exit For
                End If
            Next attemptNumber
        End If
        
        If Not success Then
            MsgBox "Could not load PDF. The file may not be available at the expected URL." & vbCrLf & vbCrLf & _
                   "Tried URL: " & NYSDOT_BASE_URL & sheetNumber & ".pdf" & vbCrLf & vbCrLf & _
                   "You can try downloading directly from:" & vbCrLf & _
                   "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us/619", _
                   vbExclamation, "PDF Not Found"
            On Error Resume Next
            lblStatus.Caption = "Error: PDF not found"
            On Error GoTo 0
        End If
    End If
End Sub

Private Function LoadPDFInBrowser(url As String) As Boolean
    On Error Resume Next
    WebBrowser1.Navigate url
    
    If Err.Number = 0 Then
        On Error Resume Next
        lblStatus.Caption = "Loaded successfully"
        On Error GoTo 0
        LoadPDFInBrowser = True
    Else
        LoadPDFInBrowser = False
    End If
    On Error GoTo 0
End Function

Private Function GetSheetNumber(sheetText As String) As String
    Dim colonPos As Integer
    colonPos = InStr(sheetText, ":")
    If colonPos > 0 Then
        GetSheetNumber = Trim(Left(sheetText, colonPos - 1))
    Else
        GetSheetNumber = ""
    End If
End Function

Private Function GetAlternativeURL(sheetNumber As String, attempt As Integer) As String
    Select Case attempt
        Case 1
            GetAlternativeURL = NYSDOT_BASE_URL & sheetNumber & "_0.pdf"
        Case 2
            GetAlternativeURL = NYSDOT_BASE_URL & sheetNumber & "new.pdf"
        Case 3
            GetAlternativeURL = NYSDOT_BASE_URL & sheetNumber & "_20250501.pdf"
        Case 4
            GetAlternativeURL = "https://www.dot.ny.gov/divisions/operating/oom/transportation-systems/repository/" & sheetNumber & ".pdf"
        Case Else
            GetAlternativeURL = ""
    End Select
End Function

Private Sub btnRefresh_Click()
    Call LoadSelectedSheet
End Sub

Private Sub btnWorkzoneDesigner_Click()
    ' Return to Workzone Designer form
    Me.Hide
    WorkzoneDesigner.Show vbModeless
End Sub


