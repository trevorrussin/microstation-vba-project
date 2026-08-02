Option Explicit

' ============================================================
' CAPTURE SCREEN MODAL HANDLER -- PARKED, NOT CURRENTLY USED
' Would auto-fill and close MicroStation's "Capture Screen Output"
' dialog (opened by the old CAPTURE VIEW keyin), but live testing
' (2026-08-01) found this MicroStation install's CAPTURE VIEW never
' opens that dialog at all -- see WZTCViewCapture.bas's header for
' what was actually observed and why this is parked. Kept in the
' project (not deleted) in case a future session finds the right
' command sequence; not called from anywhere right now.
' ============================================================

Private m_strFilePath As String

Implements IModalDialogEvents

Private Sub IModalDialogEvents_OnDialogClosed(ByVal DialogBoxName As String, ByVal DialogResult As MsdDialogBoxResult)
End Sub

Public Property Let FilePath(s As String)
    m_strFilePath = s
End Property

Private Sub IModalDialogEvents_OnDialogOpened(ByVal DialogBoxName As String, DialogResult As MsdDialogBoxResult)
    If DialogBoxName = "Capture Screen Output" Then

        Const Separator As String = "\"
        Dim path As String, file As String, ext As String
        Dim pos As Integer

        pos = InStrRev(m_strFilePath, Separator)
        If (0 < pos) Then
            path = Left(m_strFilePath, pos)
            file = Mid(m_strFilePath, pos + 1)
        End If

        pos = InStrRev(file, ".")
        If (0 < pos) Then
            ext = Mid(file, pos + 1)
        Else
            ext = "png"
        End If

        CadInputQueue.SendCommand "MDL COMMAND MGDSHOOK,fileList_setFilterCmd *." & ext
        CadInputQueue.SendCommand "MDL COMMAND MGDSHOOK,fileList_setDirectoryCmd " & path
        CadInputQueue.SendCommand "MDL COMMAND MGDSHOOK,fileList_setFileNameCmd " & file

        DialogResult = msdDialogBoxResultOK
    End If
End Sub
