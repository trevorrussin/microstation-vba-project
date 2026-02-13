Option Explicit

' ============================================================
' WZTC STORED DATA
' Public variables to persist workzone designer selections
' after frmWorkzoneDesigner is unloaded
' ============================================================

' Spacing & Clearances
Public wztcDownstreamTaper As String
Public wztcRollAhead As String
Public wztcVehicleSpace As String
Public wztcBufferSpace As String
Public wztcMergingTaper As String
Public wztcShoulderTapers As String
Public wztcAdvancedWarningSpacing As String
Public wztcSkipLines As String
Public wztcChannelizing As String
Public wztcFlareBarrier As String
Public wztcFlareBeam As String

' User Selections
Public wztcCategory As String
Public wztcSheet As String
Public wztcSpeed As String
Public wztcRoadType As String
Public wztcLaneWidth As String
Public wztcShoulderWidth As String

' Sign Table
Public wztcSignCount As Integer
Public wztcSignNumbers() As String
Public wztcSignSpacings() As String
Public wztcSignSizes() As String
Public wztcSignSides() As String

' Sign perpendicular line geometry (populated during alignment placement step).
' One entry per sign item where the user clicked "Place Line".
' Used by the sign drawing step to constrain post clicks to the correct perp line.
Public wztcPlacedSignCount As Integer
Public wztcPlacedSignNums() As String    ' sign number (matches wztcSignNumbers)
Public wztcPlacedSignPtX() As Double     ' alignment point X (midpoint of perp line)
Public wztcPlacedSignPtY() As Double     ' alignment point Y
Public wztcPlacedSignPtZ() As Double     ' alignment point Z
Public wztcPlacedSignPerpX() As Double   ' unit perpendicular vector X
Public wztcPlacedSignPerpY() As Double   ' unit perpendicular vector Y
Public wztcPlacedSignSide() As String    ' "One Side" or "Both Sides"
Public wztcPlacedSignSize() As String    ' size string from sign table

' WZTC Order table (parameter labels + sign labels in display order after Submit & Draw)
Public wztcOrderLabelCount As Integer
Public wztcOrderLabels() As String

' Alignment drawing tracking (set by Module6 before/during alignment drawing)
' wztcAlignmentStartMaxID: max element ID in model before user starts drawing
' wztcAlignmentFirstPoint*: first click point of the drawn alignment
Public wztcAlignmentStartMaxID As Double   ' stored as Double; use ElIDAsDouble() to convert el.ID
Public wztcAlignmentFirstPointX As Double
Public wztcAlignmentFirstPointY As Double
Public wztcAlignmentFirstPointZ As Double

' ============================================================
' UTILITY: Convert a MicroStation DLong element ID to Double
' so it can be stored and compared with normal operators.
' ============================================================
Public Function ElIDAsDouble(dl As DLong) As Double
    Dim hi As Double, lo As Double
    hi = CDbl(dl.High)
    lo = CDbl(dl.Low)
    If hi < 0 Then hi = hi + 4294967296#
    If lo < 0 Then lo = lo + 4294967296#
    ElIDAsDouble = hi * 4294967296# + lo
End Function
