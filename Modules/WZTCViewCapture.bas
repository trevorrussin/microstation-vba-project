Option Explicit

' ============================================================
' VIEW CAPTURE -- PARKED, DO NOT CALL
' ------------------------------------------------------------
' Investigated driving MicroStation's own raster-export dialog
' chain headlessly (CAPTURE VIEW / DIALOG SAVEIMAGE), matching
' the confirmed-working Bentley/LA Solutions KB pattern for the
' old "Capture Screen Output" dialog. Live-tested against this
' install (MicroStation CONNECT, 2026-08-01) and found it doesn't
' apply here:
'   - "capture view" + "selview n" never opened any dialog at all
'     (IModalDialogEvents never fired -- confirmed via logging).
'   - "dialog saveimage" DOES open a real dialog, but it's named
'     "Save Image" (not "Capture Screen Output"), and per Bentley
'     docs its own "Save" setting opens a SECOND dialog
'     ("Save Image As") to actually pick the file -- an unknown
'     internal command to trigger programmatically. Left
'     unhandled, this dialog blocks the VBA thread waiting for a
'     real user click; that happened live and required an
'     external WM_CLOSE (Win32) to un-stick the session.
'
' Per CLAUDE.md ("Do not guess CadInputQueue sequences"), this
' should not be wired back up on a further guess at the "Save"
' command. Capture moved to an OS-level screenshot approach
' instead (mcp-server/view_capture.py) that never touches
' CadInputQueue and can't hang MicroStation. WZTCBridge's
' CAPTURE_VIEW op still exists but returns ERROR without calling
' anything below -- see WZTCBridge.ExecCaptureView.
'
' If a future session wants to revisit this: use MicroStation's
' key-in browser/macro recorder to observe the actual command the
' "Save" button on the Save Image dialog sends, rather than
' guessing.
' ============================================================

Public Function CaptureView(viewNum As Integer, filePath As String) As Boolean
    CaptureView = False
End Function
