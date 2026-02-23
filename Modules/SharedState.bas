Option Explicit

' ============================================================
' WZTC STORED DATA
' Public variables to persist workzone designer selections
' after WZTCDesigner is unloaded
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
Public wztcUpTaperBarrier As String
Public wztcUpTaperBeam As String

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

' WZTC Cell Library placement counts (1-based, indexed by GetCellCatalogue() position)
Public wztcCellPlacementCounts(1 To 16) As Integer

' Alignment drawing tracking (set by AlignmentTool before/during alignment drawing)
' wztcAlignmentStartMaxID: max element ID in model before user starts drawing
' wztcAlignmentFirstPoint*: first click point of the drawn alignment
Public wztcAlignmentStartMaxID As Double   ' stored as Double; use ElIDAsDouble() to convert el.ID
Public wztcAlignmentFirstPointX As Double
Public wztcAlignmentFirstPointY As Double
Public wztcAlignmentFirstPointZ As Double

' Perpendicular line element IDs – one entry per line placed during alignment placement.
' Used by PlaceCells Finish to selectively delete only those reference lines.
Public wztcPerpLineIDCount As Integer
Public wztcPerpLineIDs() As Double

' Multi-alignment combined table data (set on Submit & Draw).
' Index 1 = Upstream alignment; index 2 = Downstream; 3+ = additional alignments.
' wztcOrderLabels / wztcSignNumbers etc. remain populated from alignment 1 for
' downstream compatibility (AlignDraw → PlacePerp → PlaceSign unchanged).
Public wztcAlignCount As Integer
Public wztcAlignNames(1 To 10) As String
Public wztcAlignRowCounts(1 To 10) As Integer
Public wztcAlignRowTypes(1 To 10, 1 To 50) As String     ' "Sign" or "Non-Sign"
Public wztcAlignRowLabels(1 To 10, 1 To 50) As String    ' sign number or description
Public wztcAlignRowSpacings(1 To 10, 1 To 50) As String  ' spacing value in ft
Public wztcAlignRowSizes(1 To 10, 1 To 50) As String     ' sign size string (Signs only)
Public wztcAlignRowSides(1 To 10, 1 To 50) As String     ' "One Side"/"Both Sides" (Signs only)

' Per-alignment drawing tracking (set by AlignmentTool during multi-alignment drawing step).
' wztcAlignMaxIDSnapshot: max element ID before user starts drawing that alignment
' wztcAlignGraphicGroup: graphic group number assigned to that alignment's elements on Commit
' wztcAlignDrawn: True once CommitCurrentAlignment() has been called for that alignment
' wztcCurrentAlignDrawIdx: which alignment the user is currently drawing (1-based)
Public wztcAlignMaxIDSnapshot(1 To 10) As Double
Public wztcAlignGraphicGroup(1 To 10) As Integer
Public wztcAlignDrawn(1 To 10) As Boolean
Public wztcCurrentAlignDrawIdx As Integer
Public wztcAlignFirstPtX(1 To 10) As Double  ' first click X for each alignment (for path orientation)
Public wztcAlignFirstPtY(1 To 10) As Double
Public wztcAlignFirstPtZ(1 To 10) As Double

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
