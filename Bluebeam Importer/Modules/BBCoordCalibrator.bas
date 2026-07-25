Option Explicit

' ============================================================
' BBCoordCalibrator.bas
' Two-point affine transform: PDF page coordinates → MicroStation design units.
'
' The user provides two reference points that exist on both the
' exported PDF and in the MicroStation DGN file (e.g., survey
' corners, station markers, or benchmark locations).
'
' Math:
'   scale  = MicroStation distance / PDF distance  (uniform, same X and Y)
'   offset = mstnRef1 - scale * pdfRef1
'   convert: mstnX = scale * pdfX + offsetX
'            mstnY = scale * pdfY + offsetY
'
' Note: This assumes no rotation between PDF and drawing (the PDF
' was exported from MicroStation with North up). If there is
' rotation, the two-point calibration will still produce
' reasonable results near the reference points, but accuracy
' degrades far from them.
' ============================================================

' ============================================================
' ComputeCalibration
' Call after bbPdfRef1X/Y, bbPdfRef2X/Y, bbMstnRef1X/Y,
' bbMstnRef2X/Y have all been set.
' Sets bbScaleX, bbScaleY, bbOffsetX, bbOffsetY, bbCalibrated.
' ============================================================
Public Sub ComputeCalibration()
    On Error GoTo CalcErr

    ' Distance between the two reference points in each space
    Dim pdfDx As Double, pdfDy As Double
    pdfDx = bbPdfRef2X - bbPdfRef1X
    pdfDy = bbPdfRef2Y - bbPdfRef1Y
    Dim pdfDist As Double
    pdfDist = Sqr(pdfDx * pdfDx + pdfDy * pdfDy)

    Dim mstnDx As Double, mstnDy As Double
    mstnDx = bbMstnRef2X - bbMstnRef1X
    mstnDy = bbMstnRef2Y - bbMstnRef1Y
    Dim mstnDist As Double
    mstnDist = Sqr(mstnDx * mstnDx + mstnDy * mstnDy)

    If pdfDist < 0.0001 Then
        MsgBox "PDF reference points are too close together. " & _
               "Choose two points that are far apart on the sheet.", _
               vbExclamation, "Bluebeam Importer"
        Exit Sub
    End If

    ' Uniform scale (same ratio for X and Y — assumes no differential scaling)
    bbScaleX = mstnDist / pdfDist
    bbScaleY = bbScaleX

    ' Translation offset: mstn = scale * pdf + offset
    ' → offset = mstn1 - scale * pdf1
    bbOffsetX = bbMstnRef1X - bbScaleX * bbPdfRef1X
    bbOffsetY = bbMstnRef1Y - bbScaleY * bbPdfRef1Y

    bbCalibrated = True

    ' Immediately convert all already-parsed markups that have coordinates
    ConvertAllMarkups

    Exit Sub
CalcErr:
    MsgBox "Calibration error: " & Err.Description, vbCritical, "Bluebeam Importer"
End Sub

' ============================================================
' ConvertPdfToMstn
' Convert a single PDF coordinate pair to MicroStation units.
' Requires calibration to have been run first.
' ============================================================
Public Sub ConvertPdfToMstn(pdfX As Double, pdfY As Double, _
                              ByRef mstnX As Double, ByRef mstnY As Double)
    If Not bbCalibrated Then
        mstnX = 0
        mstnY = 0
        Exit Sub
    End If
    mstnX = bbScaleX * pdfX + bbOffsetX
    mstnY = bbScaleY * pdfY + bbOffsetY
End Sub

' ============================================================
' ConvertAllMarkups
' Converts PdfX/PdfY → MstnX/MstnY for every markup where
' HasCoord = True. Called automatically by ComputeCalibration.
' Safe to call multiple times (idempotent).
' ============================================================
Public Sub ConvertAllMarkups()
    If Not bbCalibrated Then Exit Sub
    If bbMarkupCount = 0 Then Exit Sub

    Dim i As Integer
    For i = 1 To bbMarkupCount
        If bbMarkups(i).HasCoord Then
            ConvertPdfToMstn bbMarkups(i).PdfX, bbMarkups(i).PdfY, _
                             bbMarkups(i).MstnX, bbMarkups(i).MstnY
        End If
    Next i
End Sub

' ============================================================
' CalibrationSummary
' Returns a human-readable string for the status label.
' e.g. "Calibrated — scale: 12.05 ft/pt"
' ============================================================
Public Function CalibrationSummary() As String
    If Not bbCalibrated Then
        CalibrationSummary = "Not calibrated"
        Exit Function
    End If
    CalibrationSummary = "Calibrated  scale: " & Format(bbScaleX, "0.000") & " ft/pt"
End Function
