Option Explicit

' ============================================================
' BBSharedState.bas
' All public state for the Bluebeam Markup Importer.
' Mirrors the role of SharedState.bas in the WZTC project.
' No logic here — only declarations.
' ============================================================

' ============================================================
' BBMarkup — one parsed markup from the Bluebeam export file
' ============================================================
Public Type BBMarkup
    MarkupID    As String    ' sequential label: "M001", "M002", etc.
    RawText     As String    ' full unmodified comment string from export
    OpType      As String    ' "DELETE" | "MOVE" | "ADD_CELL" | "CHANGE_LEVEL" |
                             ' "EDIT_TEXT" | "ADD_DIMENSION" | "ADD_CALLOUT" |
                             ' "DELETE_CALLOUT" | "UNKNOWN"
    Param1      As String    ' primary param depending on OpType:
                             '   ADD_CELL      → cell name (e.g. "TWZAP_P")
                             '   CHANGE_LEVEL  → level name (e.g. "TWZBT_P")
                             '   EDIT_TEXT     → new text string
                             '   ADD_CALLOUT   → callout text
                             '   MOVE          → direction keyword ("N","S","E","W","LEFT","RIGHT","UP","DOWN")
    Param2      As String    ' secondary param:
                             '   MOVE          → distance in feet (numeric string)
    HasCoord    As Boolean   ' True if PDF X/Y coordinates were found in export or inline text
    PdfX        As Double    ' raw PDF page coordinate X (in PDF units, typically points)
    PdfY        As Double    ' raw PDF page coordinate Y
    MstnX       As Double    ' converted MicroStation design coordinate X (0 if not calibrated)
    MstnY       As Double    ' converted MicroStation design coordinate Y
    Status      As String    ' "Pending" | "Done" | "Skipped" | "Error"
    StatusNote  As String    ' short description of result or error
End Type

' ============================================================
' Calibration state — two-point affine PDF → MicroStation
' ============================================================
Public bbCalibrated     As Boolean   ' True once ComputeCalibration() has been called
Public bbPdfRef1X       As Double    ' PDF coordinate of reference point 1
Public bbPdfRef1Y       As Double
Public bbPdfRef2X       As Double    ' PDF coordinate of reference point 2
Public bbPdfRef2Y       As Double
Public bbMstnRef1X      As Double    ' MicroStation coordinate of reference point 1 (user click)
Public bbMstnRef1Y      As Double
Public bbMstnRef2X      As Double    ' MicroStation coordinate of reference point 2 (user click)
Public bbMstnRef2Y      As Double
Public bbScaleX         As Double    ' MicroStation units per PDF unit (X axis)
Public bbScaleY         As Double    ' MicroStation units per PDF unit (Y axis; same as X for uniform scale)
Public bbOffsetX        As Double    ' translation offset X
Public bbOffsetY        As Double    ' translation offset Y

' ============================================================
' Markup array — populated by BBFileParser
' ============================================================
Public bbMarkups()      As BBMarkup
Public bbMarkupCount    As Integer   ' number of valid markups loaded
Public bbCurrentIdx     As Integer   ' index of markup currently being processed (1-based)

' ============================================================
' Session state
' ============================================================
Public bbLoadedFilePath As String    ' path of the last successfully loaded markup file
