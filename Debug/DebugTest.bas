Option Explicit

' MINIMAL TEST VERSION - Use this to test if basic form loads
' Copy this code into your UserForm to test

Private Sub UserForm_Initialize()
    On Error GoTo InitError

    Me.Caption = "TEST - Minimal Load"

    ' Test each control exists
    If Not TestControl("lblCategory") Then Exit Sub
    If Not TestControl("cboCategory") Then Exit Sub
    If Not TestControl("lblSheet") Then Exit Sub
    If Not TestControl("cboSheet") Then Exit Sub
    If Not TestControl("lblRoadSpeed") Then Exit Sub
    If Not TestControl("cboRoadSpeed") Then Exit Sub
    If Not TestControl("lblRoadType") Then Exit Sub
    If Not TestControl("cboRoadType") Then Exit Sub
    If Not TestControl("frameSpacingValues") Then Exit Sub
    If Not TestControl("frameSignTable") Then Exit Sub
    If Not TestControl("lblSignTable") Then Exit Sub
    If Not TestControl("btnAddRow") Then Exit Sub
    If Not TestControl("btnReference") Then Exit Sub
    If Not TestControl("btnSubmit") Then Exit Sub
    If Not TestControl("lblStatus") Then Exit Sub

    MsgBox "SUCCESS! All controls exist. Form loaded correctly.", vbInformation

    Exit Sub
InitError:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function TestControl(controlName As String) As Boolean
    On Error GoTo NotFound
    Dim ctrl As Object
    Set ctrl = Me.Controls(controlName)
    Debug.Print "✓ Found: " & controlName
    TestControl = True
    Exit Function
NotFound:
    MsgBox "MISSING CONTROL: " & controlName & vbCrLf & "Please add this control to the form!", vbCritical
    Debug.Print "✗ MISSING: " & controlName
    TestControl = False
End Function
