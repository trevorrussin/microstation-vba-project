Attribute VB_Name = "DebugWorkflowSequence"
Option Explicit

' ============================================================
' DEBUG - Verify Full WZTC Workflow Form Sequence
' ------------------------------------------------------------
' Run TestWorkflowSequence() from the MicroStation VBA IDE
' Immediate Window.  This test instantiates every form in the
' 6-step WZTC workflow and confirms that each one can load
' its Initialize event without error.
'
' Workflow order:
'   1. WZTCDesigner   (configuration)
'   2. AlignDraw      (alignment drawing)
'   3. PlacePerp      (perpendicular line placement)
'   4. PlaceSign      (sign placement)
'   5. PlaceElements  (element drawing)
'   6. PlaceCells     (cell library placement)
'
' Expected output:  PASS/FAIL for each form, then a summary.
' ============================================================

Private passCount As Integer
Private failCount As Integer

Sub TestWorkflowSequence()
    passCount = 0
    failCount = 0

    Debug.Print "============================================"
    Debug.Print " WZTC Workflow Form Sequence Test"
    Debug.Print " " & Now()
    Debug.Print "============================================"
    Debug.Print ""

    ' --- Step 1: WZTCDesigner ---
    TestFormLoad "WZTCDesigner", "Step 1 - Configuration"

    ' --- Step 2: AlignDraw ---
    TestFormLoad "AlignDraw", "Step 2 - Alignment Drawing"

    ' --- Step 3: PlacePerp ---
    TestFormLoad "PlacePerp", "Step 3 - Perpendicular Placement"

    ' --- Step 4: PlaceSign ---
    TestFormLoad "PlaceSign", "Step 4 - Sign Placement"

    ' --- Step 5: PlaceElements ---
    TestFormLoad "PlaceElements", "Step 5 - Element Drawing"

    ' --- Step 6: PlaceCells ---
    TestFormLoad "PlaceCells", "Step 6 - Cell Library"

    ' --- Bonus: SheetViewer ---
    TestFormLoad "SheetViewer", "Bonus  - NYSDOT Sheet Viewer"

    Debug.Print ""
    Debug.Print "============================================"
    Debug.Print " RESULTS:  " & passCount & " passed, " & failCount & " failed  (of 7 forms)"
    If failCount = 0 Then
        Debug.Print " STATUS:   ALL FORMS LOAD SUCCESSFULLY"
    Else
        Debug.Print " STATUS:   " & failCount & " FORM(S) FAILED TO LOAD"
        Debug.Print " ACTION:   Check Immediate Window output above for error details"
    End If
    Debug.Print "============================================"
End Sub

' ============================================================
' HELPER - Attempt to load a single form by name
' ============================================================
Private Sub TestFormLoad(formName As String, stepLabel As String)
    Dim frm As Object
    On Error GoTo FormErr

    Debug.Print "  Loading " & formName & " (" & stepLabel & ")..."

    ' UserForms.Add requires the form name registered in the VBA project.
    ' We use VBA.UserForms.Add to create without showing.
    Set frm = VBA.UserForms.Add(formName)

    Debug.Print "  PASS  " & formName & "  - Caption: """ & frm.Caption & """"
    passCount = passCount + 1

    Unload frm
    Set frm = Nothing
    Exit Sub

FormErr:
    Debug.Print "  FAIL  " & formName & "  - Error " & Err.Number & ": " & Err.Description
    failCount = failCount + 1
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    Set frm = Nothing
    On Error GoTo 0
End Sub