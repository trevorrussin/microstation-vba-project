Attribute VB_Name = "DebugDesignerLoad"
Option Explicit

' ============================================================
' DEBUG - Verify WZTCDesigner Form Loads Correctly
' ------------------------------------------------------------
' Run TestDesignerFormLoad() from the MicroStation VBA IDE
' Immediate Window to verify that WZTCDesigner.frm has all
' required controls and initialises without errors.
'
' Expected output:  PASS/FAIL for each control, then a summary.
' ============================================================

Private passCount As Integer
Private failCount As Integer

Sub TestDesignerFormLoad()
    passCount = 0
    failCount = 0

    Debug.Print "============================================"
    Debug.Print " WZTCDesigner Form Load Test"
    Debug.Print " " & Now()
    Debug.Print "============================================"

    ' ---------- Create the form (triggers UserForm_Initialize) ----------
    Dim frm As WZTCDesigner
    On Error GoTo LoadFailed
    Set frm = New WZTCDesigner

    Debug.Print ""
    Debug.Print "--- Input Controls ---"
    AssertControl frm, "lblCategory"
    AssertControl frm, "cboCategory"
    AssertControl frm, "lblSheet"
    AssertControl frm, "cboSheet"
    AssertControl frm, "lblRoadSpeed"
    AssertControl frm, "cboRoadSpeed"
    AssertControl frm, "lblRoadType"
    AssertControl frm, "cboRoadType"
    AssertControl frm, "lblLaneWidth"
    AssertControl frm, "cboLaneWidth"
    AssertControl frm, "lblShoulderWidth"
    AssertControl frm, "cboShoulderWidth"

    Debug.Print ""
    Debug.Print "--- Frames ---"
    AssertControl frm, "frameSpacingValues"
    AssertControl frm, "frameSignTable"
    AssertControl frm, "frameWZTCOrder"

    Debug.Print ""
    Debug.Print "--- Sign Table & Order ---"
    AssertControl frm, "lblSignTable"
    AssertControl frm, "lstWZTCOrder"
    AssertControl frm, "btnOrderUp"
    AssertControl frm, "btnOrderDown"
    AssertControl frm, "btnOrderDelete"
    AssertControl frm, "btnRefreshOrder"

    Debug.Print ""
    Debug.Print "--- Action Buttons ---"
    AssertControl frm, "btnAddRow"
    AssertControl frm, "btnRemoveRow"
    AssertControl frm, "btnReference"
    AssertControl frm, "btnSubmit"
    AssertControl frm, "lblStatus"

    Debug.Print ""
    Debug.Print "============================================"
    Debug.Print " RESULTS:  " & passCount & " passed, " & failCount & " failed"
    If failCount = 0 Then
        Debug.Print " STATUS:   ALL CONTROLS PRESENT"
    Else
        Debug.Print " STATUS:   " & failCount & " CONTROL(S) MISSING"
        Debug.Print " ACTION:   Add missing controls in the VBA IDE form designer"
    End If
    Debug.Print "============================================"

    Unload frm
    Exit Sub

LoadFailed:
    Debug.Print "FATAL: Could not create WZTCDesigner form."
    Debug.Print "       Error " & Err.Number & ": " & Err.Description
    Debug.Print "============================================"
End Sub

' ============================================================
' HELPER - Check whether a named control exists on the form
' ============================================================
Private Sub AssertControl(frm As UserForm, ctrlName As String)
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = frm.Controls(ctrlName)
    On Error GoTo 0

    If Not ctrl Is Nothing Then
        Debug.Print "  PASS  " & ctrlName
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL  " & ctrlName & "  << MISSING"
        failCount = failCount + 1
    End If
End Sub