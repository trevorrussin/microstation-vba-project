
' ============================================================
' SIGN LIBRARY MANAGER - REVISED
' Stores and retrieves sign properties and associated text.
' Contains the signData type and all library operations.
'
' IMPORTANT: This must be a STANDARD MODULE
' Uses a workaround to avoid VBA Type parameter restrictions
' ============================================================

Option Explicit

' Sign data structure - must be in a standard module to be public
Public Type signData
    SignNumber As String        ' e.g., "R01-01", "W20-05"
    Description As String       ' e.g., "Stop", "Lane Closed Ahead"
    CellName As String          ' Cell name in library: must match .cel exactly
    CellLibraryPath As String   ' Full path to .cel file
    TextLabel As String         ' Text label (sign code for display)
    TextLine2 As String         ' Resolved size string for current road type (set by GetSignData)
    TextLine2NonFreeway As String ' Size string for Non-Freeway (e.g. "30"" x 30""")
    TextLine2Freeway As String  ' Size string for Freeway (e.g. "36"" x 36""")
    WidthInches As Double       ' Sign width in inches (Non-Freeway)
    HeightInches As Double      ' Sign height in inches (Non-Freeway)
    PostType As String          ' e.g., "TWZSGN_P"
    PostLibraryPath As String   ' Path to post cell library
    DefaultSpacing As Double    ' Default spacing in feet
End Type

' Array storage for signs (avoids Variant/UDT coercion with Collection)
Private signLibrary() As signData
Private signLibraryCount As Long
Private signLibraryInitialized As Boolean

Private Const celPath    As String = "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel"
Private Const postPath   As String = "c:\pwworking\usny\d0119091\ny_plan_wztc.cel"
Private Const postType   As String = "TWZSGN_P"
Private Const defSpacing As Double = 350

' ============================================================
' INITIALIZATION
' ============================================================
Public Sub InitializeSignLibrary()
    ReDim signLibrary(1 To 256)
    signLibraryCount = 0
    signLibraryInitialized = True
    Call LoadDefaultSigns
End Sub

' ============================================================
' LOAD DEFAULT SIGNS
' MUTCD sign table: Description, SignNumber/CellName, Size Non-Freeway, Size Freeway.
' Paths and post type shared; default spacing 350 ft.
' ============================================================
Private Sub LoadDefaultSigns()
    Call LoadSigns_Part1
    Call LoadSigns_Part2
    Call LoadSigns_Part3
    Call LoadSigns_Part4
End Sub

Private Sub LoadSigns_Part1()
Call AddSign("G20-01", "Road Work Next XX Miles", "G20-01", celPath, "G20-01", "36"" x 18""", "48"" x 24""", 36, 18, postType, postPath, defSpacing)
Call AddSign("G20-02", "End Road Work", "G20-02", celPath, "G20-02", "36"" x 18""", "48"" x 24""", 36, 18, postType, postPath, defSpacing)
Call AddSign("G20-04", "Pilot Car Follow Me", "G20-04", celPath, "G20-04", "36"" x 18""", "36"" x 18""", 36, 18, postType, postPath, defSpacing)

Call AddSign("E05-02",  "Exit Open",  "E05-02",  celPath, "E05-02",  "48"" x 36""", "48"" x 36""", 48, 36, postType, postPath, defSpacing)
Call AddSign("E05-02a", "Exit Closed", "E05-02a", celPath, "E05-02a", "48"" x 36""", "48"" x 36""", 48, 36, postType, postPath, defSpacing)
Call AddSign("E05-03",  "Exit Only",  "E05-03",  celPath, "E05-03",  "48"" x 36""", "48"" x 36""", 48, 36, postType, postPath, defSpacing)

Call AddSign("R01-01", "Stop", "R01-01", celPath, "R01-01", "30"" x 30""", "30"" x 30""", 30, 30, postType, postPath, defSpacing)
Call AddSign("R01-01", "Stop (on Stop/Slow Paddle)", "R01-01", celPath, "R01-01", "18"" x 18""", "18"" x 18""", 18, 18, postType, postPath, defSpacing)
Call AddSign("R01-02", "Yield", "R01-02", celPath, "R01-02", "36"" x 36""", "36"" x 36""", 36, 36, postType, postPath, defSpacing)
Call AddSign("R01-02aP", "To Oncoming Traffic (plaque)", "R01-02aP", celPath, "R01-02aP", "36"" x 30""", "48"" x 36""", 36, 30, postType, postPath, defSpacing)
Call AddSign("R01-07", "Wait on Stop", "R01-07", celPath, "R01-07", "24"" x 30""", "24"" x 30""", 24, 30, postType, postPath, defSpacing)
Call AddSign("R01-07a", "Wait on Stop - Go on Slow", "R01-07a", celPath, "R01-07a", "30"" x 36""", "30"" x 36""", 30, 36, postType, postPath, defSpacing)
Call AddSign("R01-08", "Go on Slow", "R01-08", celPath, "R01-08", "24"" x 30""", "24"" x 30""", 24, 30, postType, postPath, defSpacing)

Call AddSign("R02-01",  "Speed Limit",                 "R02-01",  celPath, "R02-01",  "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
Call AddSign("R02-06P",  "Fines Higher (plaque)",       "R02-06P",  celPath, "R02-06P",  "24"" x 18""", "36"" x 24""", 24, 18, postType, postPath, defSpacing)
Call AddSign("R02-06aP", "Fines Double (plaque)",       "R02-06aP", celPath, "R02-06aP", "24"" x 18""", "36"" x 24""", 24, 18, postType, postPath, defSpacing)
Call AddSign("R02-06bP", "$XX Fine (plaque)",           "R02-06bP", celPath, "R02-06bP", "24"" x 18""", "36"" x 24""", 24, 18, postType, postPath, defSpacing)
Call AddSign("R02-10",   "Begin Higher Fines Zone",     "R02-10",   celPath, "R02-10",   "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
Call AddSign("R02-11",   "End Higher Fines Zone",       "R02-11",   celPath, "R02-11",   "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
Call AddSign("R02-12",   "End Work Zone Speed Limit",   "R02-12",   celPath, "R02-12",   "24"" x 36""", "36"" x 54""", 24, 36, postType, postPath, defSpacing)

Call AddSign("R03-01", "Movement Prohibition", "R03-01", celPath, "R03-01", "24"" x 24""", "36"" x 36""", 24, 24, postType, postPath, defSpacing)
Call AddSign("R03-05", "Mandatory Movement Lane Control - Turn Only", "R03-05", celPath, "R03-05", "30"" x 36""", "30"" x 36""", 30, 36, postType, postPath, defSpacing)
Call AddSign("R03-06", "Optional Movement Lane Control - Thru and Turn", "R03-06", celPath, "R03-06", "30"" x 36""", "30"" x 36""", 30, 36, postType, postPath, defSpacing)
Call AddSign("R03-07", "Right (Left) Lane Must Turn Right (Left)", "R03-07", celPath, "R03-07", "30"" x 30""", "30"" x 30""", 30, 30, postType, postPath, defSpacing)
Call AddSign("R03-08", "Advance Intersection Lane Control (2 lanes)", "R03-08", celPath, "R03-08", "30"" x 30""", "30"" x 30""", 30, 30, postType, postPath, defSpacing)
Call AddSign("R03-18", "Movement Prohibition - No U or Left Turn", "R03-18", celPath, "R03-18", "24"" x 24""", "36"" x 36""", 24, 24, postType, postPath, defSpacing)
Call AddSign("R03-27", "Movement Prohibition - No Straight Through", "R03-27", celPath, "R03-27", "24"" x 24""", "36"" x 36""", 24, 24, postType, postPath, defSpacing)

Call AddSign("R04-01", "Do Not Pass", "R04-01", celPath, "R04-01", "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
Call AddSign("R04-02", "Pass With Care", "R04-02", celPath, "R04-02", "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
Call AddSign("R04-07", "Keep Right", "R04-07", celPath, "R04-07", "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
Call AddSign("R04-07c", "Narrow Keep Right", "R04-07c", celPath, "R04-07c", "18"" x 30""", "18"" x 30""", 18, 30, postType, postPath, defSpacing)
Call AddSign("R04-09", "Stay in Lane", "R04-09", celPath, "R04-09", "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
Call AddSign("R04-09a", "Stay In Lane To Merge Point", "R04-09a", celPath, "R04-09a", "36"" x 48""", "36"" x 48""", 36, 48, postType, postPath, defSpacing)

Call AddSign("R05-01",  "Do Not Enter", "R05-01",  celPath, "R05-01",  "30"" x 30""", "36"" x 36""", 30, 30, postType, postPath, defSpacing)
Call AddSign("R05-01a", "Wrong Way",    "R05-01a", celPath, "R05-01a", "36"" x 24""", "42"" x 30""", 36, 24, postType, postPath, defSpacing)
Call AddSign("R06-01",  "One Way",      "R06-01",  celPath, "R06-01",  "36"" x 12""", "48"" x 18""", 36, 12, postType, postPath, defSpacing)
Call AddSign("R06-02",  "One Way",      "R06-02",  celPath, "R06-02",  "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)

Call AddSign("R08-03", "No Parking (symbol)", "R08-03", celPath, "R08-03", "24"" x 24""", "36"" x 36""", 24, 24, postType, postPath, defSpacing)

Call AddSign("R09-08", "Pedestrian Crosswalk",                "R09-08", celPath, "R09-08", "36"" x 18""", "36"" x 18""", 36, 18, postType, postPath, defSpacing)
Call AddSign("R09-09", "Sidewalk Closed",                     "R09-09", celPath, "R09-09", "24"" x 12""", "24"" x 12""", 24, 12, postType, postPath, defSpacing)
Call AddSign("R09-10", "Sidewalk Closed, Use Other Side",     "R09-10", celPath, "R09-10", "24"" x 12""", "24"" x 12""", 24, 12, postType, postPath, defSpacing)
Call AddSign("R09-11", "Sidewalk Closed Ahead, Cross Here",   "R09-11", celPath, "R09-11", "24"" x 18""", "24"" x 18""", 24, 18, postType, postPath, defSpacing)
Call AddSign("R09-11a","Sidewalk Closed, Cross Here",         "R09-11a",celPath, "R09-11a","24"" x 12""", "24"" x 12""", 24, 12, postType, postPath, defSpacing)
Call AddSign("R09-12", "Bike Lane Closed",                    "R09-12", celPath, "R09-12", "24"" x 12""", "24"" x 12""", 24, 12, postType, postPath, defSpacing)

Call AddSign("R10-06", "Stop Here on Red", "R10-06", celPath, "R10-06", "24"" x 36""", "24"" x 36""", 24, 36, postType, postPath, defSpacing)

Call AddSign("R11-02", "Road Closed",                         "R11-02", celPath, "R11-02", "48"" x 30""", "48"" x 30""", 48, 30, postType, postPath, defSpacing)
Call AddSign("R11-03", "Road Closed - Local Traffic Only",    "R11-03", celPath, "R11-03", "60"" x 30""", "60"" x 30""", 60, 30, postType, postPath, defSpacing)

' R12-01,02 -> add both codes individually
Call AddSign("R12-01", "Weight Limit", "R12-01", celPath, "R12-01", "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
Call AddSign("R12-02", "Weight Limit", "R12-02", celPath, "R12-02", "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
Call AddSign("R12-05", "Weight Limit", "R12-05", celPath, "R12-05", "24"" x 36""", "36"" x 48""", 24, 36, postType, postPath, defSpacing)

Call AddSign("R22-02", "Turn Off 2-Way Radio and Cell Phone", "R22-02", celPath, "R22-02", "42"" x 36""", "42"" x 36""", 42, 36, postType, postPath, defSpacing)
Call AddSign("G20-05aP", "Work Zone (plaque)", "G20-05aP", celPath, "G20-05aP", "24"" x 18""", "30"" x 24""", 24, 18, postType, postPath, defSpacing)
End Sub

Private Sub LoadSigns_Part2()
' ------ M-series auxiliary/detour/arrow signs ------
Call AddSign("M02-01",  "Auxiliary (JCT)",                    "M02-01",  celPath, "M02-01",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M02-01c", "Auxiliary (JCT) (Colored)",          "M02-01c", celPath, "M02-01c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M03-01",  "Auxiliary (North)",                  "M03-01",  celPath, "M03-01",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M03-01c", "Auxiliary (North) (Colored)",        "M03-01c", celPath, "M03-01c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M03-02",  "Auxiliary (East)",                   "M03-02",  celPath, "M03-02",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M03-02c", "Auxiliary (East) (Colored)",         "M03-02c", celPath, "M03-02c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M03-03",  "Auxiliary (South)",                  "M03-03",  celPath, "M03-03",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M03-03c", "Auxiliary (South) (Colored)",        "M03-03c", celPath, "M03-03c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M03-04",  "Auxiliary (West)",                   "M03-04",  celPath, "M03-04",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M03-04c", "Auxiliary (West) (Colored)",         "M03-04c", celPath, "M03-04c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M04-01",  "Auxiliary (Alternate)",              "M04-01",  celPath, "M04-01",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-01a", "Auxiliary (ALT)",                    "M04-01a", celPath, "M04-01a", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-01ac","Auxiliary (ALT) (Colored)",          "M04-01ac",celPath, "M04-01ac","36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-01c", "Auxiliary (Alternate) (Colored)",    "M04-01c", celPath, "M04-01c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M04-02",  "Auxiliary (By-Pass)",                "M04-02",  celPath, "M04-02",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-02c", "Auxiliary (By-Pass) (Colored)",      "M04-02c", celPath, "M04-02c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-03",  "Auxiliary (Business)",               "M04-03",  celPath, "M04-03",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-03c", "Auxiliary (Business) (Colored)",     "M04-03c", celPath, "M04-03c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M04-04",  "Auxiliary (Truck)",                  "M04-04",  celPath, "M04-04",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-04c", "Auxiliary (Truck) (Colored)",        "M04-04c", celPath, "M04-04c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-05",  "Auxiliary (To)",                     "M04-05",  celPath, "M04-05",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-05c", "Auxiliary (To) (Colored)",           "M04-05c", celPath, "M04-05c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M04-06",  "Auxiliary (End)",                    "M04-06",  celPath, "M04-06",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-06c", "Auxiliary (End) (Colored)",          "M04-06c", celPath, "M04-06c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-06z", "Auxiliary (End) (Scenic Byway)",     "M04-06z", celPath, "M04-06z", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-06za","Auxiliary (End) (Scenic Byway) (Within Adirondack Park)", "M04-06za", celPath, "M04-06za", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M04-07",  "Auxiliary (Temporary)",              "M04-07",  celPath, "M04-07",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-07a", "Auxiliary (TEMP)",                   "M04-07a", celPath, "M04-07a", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-07ac","Auxiliary (TEMP) (Colored)",         "M04-07ac",celPath, "M04-07ac","36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-07c", "Auxiliary (Temporary) (Colored)",    "M04-07c", celPath, "M04-07c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M04-08",  "Auxiliary (Detour)",                 "M04-08",  celPath, "M04-08",  "24"" x 12""", "30"" x 15""", 24, 12, postType, postPath, defSpacing)
Call AddSign("M04-08a", "Auxiliary (End Detour)",             "M04-08a", celPath, "M04-08a", "24"" x 18""", "24"" x 18""", 24, 18, postType, postPath, defSpacing)
Call AddSign("M04-08b", "Auxiliary (End) (Detour)",           "M04-08b", celPath, "M04-08b", "24"" x 12""", "24"" x 12""", 24, 12, postType, postPath, defSpacing)

Call AddSign("M04-09A",  "Detour (Ahead)",                           "M04-09A",  celPath, "M04-09A",  "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09aA", "Detour (Bicycle and Pedestrian) (Ahead)",   "M04-09aA", celPath, "M04-09aA", "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09aL", "Detour (Bicycle and Pedestrian) (Left)",    "M04-09aL", celPath, "M04-09aL", "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09aR", "Detour (Bicycle and Pedestrian) (Right)",   "M04-09aR", celPath, "M04-09aR", "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09bA", "Detour (Pedestrian) (Ahead)",               "M04-09bA", celPath, "M04-09bA", "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09bL", "Detour (Pedestrian) (Left)",                "M04-09bL", celPath, "M04-09bL", "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09bR", "Detour (Pedestrian) (Right)",               "M04-09bR", celPath, "M04-09bR", "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09cA", "Detour (Bicycle) (Ahead)",                  "M04-09cA", celPath, "M04-09cA", "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09cL", "Detour (Bicycle) (Left)",                   "M04-09cL", celPath, "M04-09cL", "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09cR", "Detour (Bicycle) (Right)",                  "M04-09cR", celPath, "M04-09cR", "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)

Call AddSign("M04-09L",  "Detour (Left)",                "M04-09L",  celPath, "M04-09L",  "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09LBE", "Detour (Left) (Bent)",        "M04-09LBE",celPath, "M04-09LBE","30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09LBR", "Detour (Left) (Broken)",      "M04-09LBR",celPath, "M04-09LBR","30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09LD",  "Detour (Left) (Diagonal)",    "M04-09LD", celPath, "M04-09LD", "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09R",   "Detour (Right)",              "M04-09R",  celPath, "M04-09R",  "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)

' ---- Block 2 of 5 — Signs 101–200 ----

Call AddSign("M04-09RBE", "Detour (Right) (Bent)",      "M04-09RBE", celPath, "M04-09RBE", "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09RBR", "Detour (Right) (Broken)",    "M04-09RBR", celPath, "M04-09RBR", "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
Call AddSign("M04-09RD",  "Detour (Right) (Diagonal)",  "M04-09RD",  celPath, "M04-09RD",  "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)

Call AddSign("M04-10L", "Detour (Left) (Inside Arrow)",  "M04-10L", celPath, "M04-10L", "48"" x 18""", "48"" x 18""", 48, 18, postType, postPath, defSpacing)
Call AddSign("M04-10R", "Detour (Right) (Inside Arrow)", "M04-10R", celPath, "M04-10R", "48"" x 18""", "48"" x 18""", 48, 18, postType, postPath, defSpacing)

Call AddSign("M04-14",  "Auxiliary (Begin)",                                  "M04-14",  celPath, "M04-14",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-14c", "Auxiliary (Begin) (Colored)",                        "M04-14c", celPath, "M04-14c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-14z", "Auxiliary (Begin) (Scenic Byway)",                   "M04-14z", celPath, "M04-14z", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-14za","Auxiliary (Begin) (Scenic Byway) (Within Adirondack Park)", "M04-14za", celPath, "M04-14za", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M04-15", "Auxiliary (Toll)",         "M04-15", celPath, "M04-15", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-16", "Auxiliary (No Cash)",      "M04-16", celPath, "M04-16", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-17", "Auxiliary (Toll Collector)","M04-17", celPath, "M04-17", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-18", "Auxiliary (Exact Change)", "M04-18", celPath, "M04-18", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M04-20",     "Auxiliary (ETC Only)",                         "M04-20",     celPath, "M04-20",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-20oNY",  "Auxiliary (Name) (One Line) (Colored)",        "M04-20oNY",  celPath, "M04-20oNY",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-20tNY",  "Auxiliary (Name) (Two Line) (Colored)",        "M04-20tNY",  celPath, "M04-20tNY",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-20zSBNY","Auxiliary (Scenic Byway)",                     "M04-20zSBNY",celPath, "M04-20zSBNY","36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M04-20zSBNYa","Auxiliary (Scenic Byway) (Within Adirondack Park)","M04-20zSBNYa",celPath,"M04-20zSBNYa","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)

Call AddSign("M05-01L",  "Arrow (Broken) (Left)",              "M05-01L",  celPath, "M05-01L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-01Lc", "Arrow (Broken) (Left) (Colored)",    "M05-01Lc", celPath, "M05-01Lc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-01R",  "Arrow (Broken) (Right)",             "M05-01R",  celPath, "M05-01R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-01Rc", "Arrow (Broken) (Right) (Colored)",   "M05-01Rc", celPath, "M05-01Rc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M05-02L",  "Arrow (Bent) (Left)",                "M05-02L",  celPath, "M05-02L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-02Lc", "Arrow (Bent) (Left) (Colored)",      "M05-02Lc", celPath, "M05-02Lc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-02R",  "Arrow (Bent) (Right)",               "M05-02R",  celPath, "M05-02R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-02Rc", "Arrow (Bent) (Right) (Colored)",     "M05-02Rc", celPath, "M05-02Rc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M05-03",   "Arrow (Advance) (Roundabout)",       "M05-03",   celPath, "M05-03",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-03c",  "Arrow (Advance) (Roundabout) (Colored)", "M05-03c", celPath, "M05-03c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M05-03NE",  "Arrow (Next) (Exit)",               "M05-03NE",  celPath, "M05-03NE",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-03NEc", "Arrow (Next) (Exit) (Colored)",     "M05-03NEc", celPath, "M05-03NEc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-03NL",  "Arrow (Next) (Left)",               "M05-03NL",  celPath, "M05-03NL",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-03NLc", "Arrow (Next) (Left) (Colored)",     "M05-03NLc", celPath, "M05-03NLc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-03NR",  "Arrow (Next) (Right)",              "M05-03NR",  celPath, "M05-03NR",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-03NRc", "Arrow (Next) (Right) (Colored)",    "M05-03NRc", celPath, "M05-03NRc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M05-03SE",  "Arrow (Second) (Exit)",             "M05-03SE",  celPath, "M05-03SE",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-03SEc", "Arrow (Second) (Exit) (Colored)",   "M05-03SEc", celPath, "M05-03SEc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-03SL",  "Arrow (Second) (Left)",             "M05-03SL",  celPath, "M05-03SL",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-03SLc", "Arrow (Second) (Left) (Colored)",   "M05-03SLc", celPath, "M05-03SLc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-03SR",  "Arrow (Second) (Right)",            "M05-03SR",  celPath, "M05-03SR",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-03SRc", "Arrow (Second) (Right) (Colored)",  "M05-03SRc", celPath, "M05-03SRc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M05-04",  "Auxiliary (Left Lane)",        "M05-04",  celPath, "M05-04",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-04c", "Auxiliary (Left Lane) (Colored)","M05-04c", celPath, "M05-04c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M05-05",  "Auxiliary (Center Lane)",       "M05-05",  celPath, "M05-05",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-05c", "Auxiliary (Center Lane) (Colored)","M05-05c", celPath, "M05-05c","36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M05-06",  "Auxiliary (Right Lane)",        "M05-06",  celPath, "M05-06",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M05-06c", "Auxiliary (Right Lane) (Colored)","M05-06c", celPath, "M05-06c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M06-01L",  "Arrow (Left)",                                "M06-01L",  celPath, "M06-01L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-01Lc", "Arrow (Left) (Colored)",                      "M06-01Lc", celPath, "M06-01Lc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-01Lz", "Arrow (Left) (Scenic Byway)",                 "M06-01Lz", celPath, "M06-01Lz", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-01Lza","Arrow (Left) (Scenic Byway) (Within Adirondack Park)", "M06-01Lza", celPath, "M06-01Lza", "36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M06-01R",  "Arrow (Right)",                               "M06-01R",  celPath, "M06-01R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-01Rc", "Arrow (Right) (Colored)",                     "M06-01Rc", celPath, "M06-01Rc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-01Rz", "Arrow (Right) (Scenic Byway)",                "M06-01Rz", celPath, "M06-01Rz", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-01Rza","Arrow (Right) (Scenic Byway) (Within Adirondack Park)", "M06-01Rza", celPath, "M06-01Rza","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)

Call AddSign("M06-02aL",  "Arrow (Diagonal) (Down) (Left)",             "M06-02aL",  celPath, "M06-02aL",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-02aLc", "Arrow (Diagonal) (Down) (Left) (Colored)",   "M06-02aLc", celPath, "M06-02aLc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-02aR",  "Arrow (Diagonal) (Down) (Right)",            "M06-02aR",  celPath, "M06-02aR",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-02aRc", "Arrow (Diagonal) (Down) (Right) (Colored)",  "M06-02aRc", celPath, "M06-02aRc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M06-02L",  "Arrow (Diagonal) (Left)",                     "M06-02L",  celPath, "M06-02L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-02Lc", "Arrow (Diagonal) (Left) (Colored)",           "M06-02Lc", celPath, "M06-02Lc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-02R",  "Arrow (Diagonal) (Right)",                    "M06-02R",  celPath, "M06-02R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-02Rc", "Arrow (Diagonal) (Right) (Colored)",          "M06-02Rc", celPath, "M06-02Rc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M06-03",  "Arrow (Up)",                                   "M06-03",  celPath, "M06-03",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-03c", "Arrow (Up) (Colored)",                         "M06-03c", celPath, "M06-03c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-03z", "Arrow (Up) (Scenic Byway)",                    "M06-03z", celPath, "M06-03z", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-03za","Arrow (Up) (Scenic Byway) (Within Adirondack Park)", "M06-03za", celPath, "M06-03za", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M06-04",  "Arrow (Left and Right)",                       "M06-04",  celPath, "M06-04",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-04c", "Arrow (Left and Right) (Colored)",             "M06-04c", celPath, "M06-04c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M06-05L",  "Arrow (Diagonal) (Left and Right)",           "M06-05L",  celPath, "M06-05L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-05Lc", "Arrow (Diagonal) (Left and Right) (Colored)", "M06-05Lc", celPath, "M06-05Lc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-05R",  "Arrow (Diagonal) (Right and Left)",           "M06-05R",  celPath, "M06-05R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-05Rc", "Arrow (Diagonal) (Right and Left) (Colored)", "M06-05Rc", celPath, "M06-05Rc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M06-06L",  "Arrow (Left and Up)",                         "M06-06L",  celPath, "M06-06L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-06Lc", "Arrow (Left and Up) (Colored)",               "M06-06Lc", celPath, "M06-06Lc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-06R",  "Arrow (Right and Up)",                        "M06-06R",  celPath, "M06-06R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-06Rc", "Arrow (Right and Up) (Colored)",              "M06-06Rc", celPath, "M06-06Rc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M06-07L",  "Arrow (Diagonal Left and Up)",                "M06-07L",  celPath, "M06-07L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-07Lc", "Arrow (Diagonal Left and Up) (Colored)",      "M06-07Lc", celPath, "M06-07Lc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-07R",  "Arrow (Diagonal Right and Up)",               "M06-07R",  celPath, "M06-07R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-07Rc", "Arrow (Diagonal Right and Up) (Colored)",     "M06-07Rc", celPath, "M06-07Rc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M06-08L",  "Arrow (Stemmed) (Left and Up)",               "M06-08L",  celPath, "M06-08L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-08Lc", "Arrow (Stemmed) (Left and Up) (Colored)",     "M06-08Lc", celPath, "M06-08Lc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-08R",  "Arrow (Stemmed) (Right and Up)",              "M06-08R",  celPath, "M06-08R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-08Rc", "Arrow (Stemmed) (Right and Up) (Colored)",    "M06-08Rc", celPath, "M06-08Rc", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("M06-09",  "Arrow (Stemmed) (Left and Right)",             "M06-09",  celPath, "M06-09",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("M06-09c", "Arrow (Stemmed) (Left and Right) (Colored)",   "M06-09c", celPath, "M06-09c", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
End Sub

Private Sub LoadSigns_Part3()
' ---- Begin W-series (first group in this block) ----
Call AddSign("W01-01aL", "Turn (With Speed) (Left)",   "W01-01aL", celPath, "W01-01aL", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("W01-01aR", "Turn (With Speed) (Right)",  "W01-01aR", celPath, "W01-01aR", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("W01-01L",  "Turn (Left)",                "W01-01L",  celPath, "W01-01L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("W01-01R",  "Turn (Right)",               "W01-01R",  celPath, "W01-01R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("W01-02aL", "Curve (With Speed) (Left)",  "W01-02aL", celPath, "W01-02aL", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("W01-02aR", "Curve (With Speed) (Right)", "W01-02aR", celPath, "W01-02aR", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("W01-02L",  "Curve (Left)",               "W01-02L",  celPath, "W01-02L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("W01-02R",  "Curve (Right)",              "W01-02R",  celPath, "W01-02R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("W01-03L",  "Turn (Revese) (Left)",       "W01-03L",  celPath, "W01-03L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("W01-03R",  "Turn (Revese) (Right)",      "W01-03R",  celPath, "W01-03R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

Call AddSign("W01-04bL", "Curve (Reverse) (2 Lanes) (Left)",  "W01-04bL", celPath, "W01-04bL", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("W01-04bR", "Curve (Reverse) (2 Lanes) (Right)", "W01-04bR", celPath, "W01-04bR", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("W01-04cL", "Curve (Reverse) (3 Lanes) (Left)",  "W01-04cL", celPath, "W01-04cL", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
Call AddSign("W01-04cR", "Curve (Reverse) (3 Lanes) (Right)", "W01-04cR", celPath, "W01-04cR", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)

' ---- Block 3 of 5 — Signs 201–300 ----

Call AddSign("W01-04L",  "Curve (Reverse) (Left)",                     "W01-04L",  celPath, "W01-04L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)   '1
Call AddSign("W01-04R",  "Curve (Reverse) (Right)",                    "W01-04R",  celPath, "W01-04R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)  '2
Call AddSign("W01-05L",  "Curve (Winding Road) (Left)",                "W01-05L",  celPath, "W01-05L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)  '3
Call AddSign("W01-05R",  "Curve (Winding Road) (Right)",               "W01-05R",  celPath, "W01-05R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '4
Call AddSign("W01-06L",  "Arrow (Large) (Left)",                       "W01-06L",  celPath, "W01-06L",  "48"" x 24""", "60"" x 30""", 48, 24, postType, postPath, defSpacing) '5
Call AddSign("W01-06R",  "Arrow (Large) (Right)",                      "W01-06R",  celPath, "W01-06R",  "48"" x 24""", "60"" x 30""", 48, 24, postType, postPath, defSpacing) '6
Call AddSign("W01-07",   "Arrow (Large) (2 Direction)",                 "W01-07",   celPath, "W01-07",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)  '7
Call AddSign("W01-08L",  "Chevron (Left)",                              "W01-08L",  celPath, "W01-08L",  "18"" x 24""", "30"" x 36""", 18, 24, postType, postPath, defSpacing) '8
Call AddSign("W01-08R",  "Chevron (Right)",                             "W01-08R",  celPath, "W01-08R",  "18"" x 24""", "30"" x 36""", 18, 24, postType, postPath, defSpacing) '9

Call AddSign("W01-10CaL","Curve (Left) (With Side Road) (Ahead)",       "W01-10CaL",celPath,"W01-10CaL","36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing)   '10
Call AddSign("W01-10CaR","Curve (Right) (With Side Road) (Ahead)",      "W01-10CaR",celPath,"W01-10CaR","36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing)  '11
Call AddSign("W01-10CbL","Curve (Left) (With Side Road) (Diagonal Righft)","W01-10CbL",celPath,"W01-10CbL","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)     '12
Call AddSign("W01-10CbR","Curve (Right) (With Side Road) (Diagonal Left)","W01-10CbR",celPath,"W01-10CbR","36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing) '13
Call AddSign("W01-10CcL","Curve (Left) (With Side Road) (Back Right)",  "W01-10CcL",celPath,"W01-10CcL","36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing)   '14
Call AddSign("W01-10CcR","Curve (Right) (With Side Road) (Back Left)",  "W01-10CcR",celPath,"W01-10CcR","36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing)   '15
Call AddSign("W01-10CdL","Curve (Left) (With Side Road) (Ahead and Back Right)","W01-10CdL",celPath,"W01-10CdL","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing) '16
Call AddSign("W01-10CdR","Curve (Right) (With Side Road) (Ahead and Back Left)","W01-10CdR",celPath,"W01-10CdR","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing) '17
Call AddSign("W01-10CeL","Curve (Left) (With Cross Road)",              "W01-10CeL",celPath,"W01-10CeL","36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing)  '18
Call AddSign("W01-10CeR","Curve (Right) (With Cross Road)",             "W01-10CeR",celPath,"W01-10CeR","36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing) '19
Call AddSign("W01-10CfL","Curve (Left) (With Side Road) (Diagonal Left)","W01-10CfL",celPath,"W01-10CfL","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)        '20
Call AddSign("W01-10CfR","Curve (Right) (With Side Road) (Diagonal Right)","W01-10CfR",celPath,"W01-10CfR","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)     '21
Call AddSign("W01-10d",  "Curve (Reverse) (Right) (With Side Road) (Left)","W01-10d",celPath,"W01-10d","36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing)     '22
Call AddSign("W01-10e",  "Curve (Reverse) (Right) (With Cross Road)",   "W01-10e",  celPath,"W01-10e",  "36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing)   '23
Call AddSign("W01-10TaL","Turn (Left) (With Side Road) (Ahead)",        "W01-10TaL",celPath,"W01-10TaL","36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing)  '24
Call AddSign("W01-10TaR","Turn (Right) (With Side Road) (Ahead)",       "W01-10TaR",celPath,"W01-10TaR","36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing)  '25
Call AddSign("W01-10TbL","Turn (Left) (With Side Road) (Diagonal Right)","W01-10TbL",celPath,"W01-10TbL","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)        '26
Call AddSign("W01-10TbR","Turn (Right) (With Side Road) (Diagonal Left)","W01-10TbR",celPath,"W01-10TbR","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)       '27
Call AddSign("W01-10TcL","Turn (Left) (With Side Road) (Right)",        "W01-10TcL",celPath,"W01-10TcL","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)        '28
Call AddSign("W01-10TcR","Turn (Right) (With Side Road) (Left)",        "W01-10TcR",celPath,"W01-10TcR","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)        '29
Call AddSign("W01-10TdL","Turn (Left) (With Side Road) (Ahead and Right)","W01-10TdL",celPath,"W01-10TdL","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)      '30
Call AddSign("W01-10TdR","Turn (Right) (With Side Road) (Ahead and Left)","W01-10TdR",celPath,"W01-10TdR","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)      '31
Call AddSign("W01-10TeL","Turn (Left) (With Cross Road)",               "W01-10TeL",celPath,"W01-10TeL","36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing)  '32
Call AddSign("W01-10TeR","Turn (Right) (With Cross Road)",              "W01-10TeR",celPath,"W01-10TeR","36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing) '33
Call AddSign("W01-10TfL","Turn (Left) (With Side Road) (Diagonal Left)", "W01-10TfL",celPath,"W01-10TfL","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)       '34
Call AddSign("W01-10TfR","Turn (Right) (With Side Road) (Diagonal Right)","W01-10TfR",celPath,"W01-10TfR","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)      '35

Call AddSign("W01-11L",  "Curve (Hairpin) (Left)",                      "W01-11L",  celPath, "W01-11L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '36
Call AddSign("W01-11R",  "Curve (Hairpin) (Right)",                     "W01-11R",  celPath, "W01-11R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '37
Call AddSign("W01-13L",  "Curve (Truck Rollover) (Left)",               "W01-13L",  celPath, "W01-13L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '38
Call AddSign("W01-13R",  "Curve (Truck Rollover) (Right)",              "W01-13R",  celPath, "W01-13R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '39
Call AddSign("W01-15L",  "Curve (270 Degree) (Left)",                   "W01-15L",  celPath, "W01-15L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '40
Call AddSign("W01-15R",  "Curve (270 Degree) (Right)",                  "W01-15R",  celPath, "W01-15R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '41

Call AddSign("W02-01",   "Intersection (Cross Road)",                   "W02-01",   celPath, "W02-01",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '42
Call AddSign("W02-02L",  "Intersection (Side Road) (Left)",             "W02-02L",  celPath, "W02-02L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '43
Call AddSign("W02-02R",  "Intersection (Side Road) (Right)",            "W02-02R",  celPath, "W02-02R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '44
Call AddSign("W02-03L",  "Intersection (Side Road) (Diagonal) (Left)",  "W02-03L",  celPath, "W02-03L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '45
Call AddSign("W02-03R",  "Intersection (Side Road) (Diagonal) (Right)", "W02-03R",  celPath, "W02-03R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '46
Call AddSign("W02-04",   "Intersection (T)",                            "W02-04",   celPath, "W02-04",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '47
Call AddSign("W02-05",   "Intersection (Y)",                            "W02-05",   celPath, "W02-05",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '48
Call AddSign("W02-06",   "Intersection (Circular)",                     "W02-06",   celPath, "W02-06",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '49
Call AddSign("W02-072LB","Intersection (Offset Roads) (2 Left Bottom Right)","W02-072LB",celPath,"W02-072LB","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)   '50
Call AddSign("W02-072LM","Intersection (Offset Roads) (2 Left Middle Right)","W02-072LM",celPath,"W02-072LM","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)   '51
Call AddSign("W02-072LT","Intersection (Offset Roads) (2 Left Top Right)","W02-072LT",celPath,"W02-072LT","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)      '52
Call AddSign("W02-072RB","Intersection (Offset Roads) (2 Right Bottom Left)","W02-072RB",celPath,"W02-072RB","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)  '53
Call AddSign("W02-072RM","Intersection (Offset Roads) (2 Right Middle Left)","W02-072RM",celPath,"W02-072RM","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)  '54
Call AddSign("W02-072RT","Intersection (Offset Roads) (2 Right Top Left)","W02-072RT",celPath,"W02-072RT","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)     '55
Call AddSign("W02-07L",  "Intersection (Offset Roads) (Left then Right)","W02-07L",  celPath, "W02-07L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)  '56
Call AddSign("W02-07R",  "Intersection (Offset Roads) (Right then Left)","W02-07R",  celPath, "W02-07R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '57
Call AddSign("W02-08L",  "Intersection (2 Left)",                       "W02-08L",  celPath, "W02-08L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '58
Call AddSign("W02-08R",  "Intersection (2 Right)",                      "W02-08R",  celPath, "W02-08R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '59

Call AddSign("W03-01",   "Stop Ahead",                                  "W03-01",   celPath, "W03-01",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '60
Call AddSign("W03-02",   "Yield Ahead",                                 "W03-02",   celPath, "W03-02",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '61
Call AddSign("W03-03",   "Signal Ahead",                                "W03-03",   celPath, "W03-03",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '62
Call AddSign("W03-04",   "Be Prepared to Stop",                         "W03-04",   celPath, "W03-04",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '63
Call AddSign("W03-05",   "Speed Limit Ahead",                           "W03-05",   celPath, "W03-05",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '64
Call AddSign("W03-06",   "Draw Bridge Ahead",                           "W03-06",   celPath, "W03-06",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '65
Call AddSign("W03-07",   "Ramp Meter Ahead",                            "W03-07",   celPath, "W03-07",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '66
Call AddSign("W03-08",   "Ramp Metered When Flashing",                  "W03-08",   celPath, "W03-08",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '67

Call AddSign("W04-01L",  "Merge (Left)",                                "W04-01L",  celPath, "W04-01L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '68
Call AddSign("W04-01R",  "Merge (Right)",                               "W04-01R",  celPath, "W04-01R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '69
Call AddSign("W04-02L",  "Lane (Ends) (Symbol) (Left)",                 "W04-02L",  celPath, "W04-02L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '70
Call AddSign("W04-02NY", "Lane (Ends) (Symbol)",                        "W04-02NY", celPath, "W04-02NY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '71
Call AddSign("W04-02R",  "Lane (Ends) (Symbol) (Right)",                "W04-02R",  celPath, "W04-02R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '72
Call AddSign("W04-03L",  "Lane (Added) (Left)",                         "W04-03L",  celPath, "W04-03L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '73
Call AddSign("W04-03R",  "Lane (Added) (Right)",                        "W04-03R",  celPath, "W04-03R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '74
Call AddSign("W04-04aLP","Auxiliary (Traffic) (Does Not Stop) (From Left)","W04-04aLP",celPath,"W04-04aLP","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)      '75
Call AddSign("W04-04aRP","Auxiliary (Traffic) (Does Not Stop) (From Right)","W04-04aRP",celPath,"W04-04aRP","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)    '76
Call AddSign("W04-04bP", "Auxiliary (Traffic) (Does Not Stop) (Oncoming)","W04-04bP", celPath,"W04-04bP","36"" x 36""","48"" x 48""", 36, 36, postType, postPath, defSpacing)   '77
Call AddSign("W04-04P",  "Auxiliary (Traffic) (Does Not Stop) (Cross)", "W04-04P",  celPath, "W04-04P",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)  '78
Call AddSign("W04-05L",  "Merge (Entering Radway) (Left)",              "W04-05L",  celPath, "W04-05L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '79
Call AddSign("W04-05P",  "Auxiliary (No Merge Area)",                   "W04-05P",  celPath, "W04-05P",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '80
Call AddSign("W04-05R",  "Merge (Entering Radway) (Right)",             "W04-05R",  celPath, "W04-05R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '81
Call AddSign("W04-06L",  "Lane (Added) (Entering Roadway) (Left)",      "W04-06L",  celPath, "W04-06L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '82
Call AddSign("W04-06R",  "Lane (Added) (Entering Roadway) (Right)",     "W04-06R",  celPath, "W04-06R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '83

Call AddSign("W05-01",   "Road Narrows",                                "W05-01",   celPath, "W05-01",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '84
Call AddSign("W05-02",   "Bridge (Narrow)",                             "W05-02",   celPath, "W05-02",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '85
Call AddSign("W05-03",   "Bridge (1 Lane)",                              "W05-03",   celPath, "W05-03",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '86
Call AddSign("W05-03NY", "Road (One Lane)",                             "W05-03NY", celPath, "W05-03NY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '87
Call AddSign("W05-04",   "Ramp Narrows",                                 "W05-04",   celPath, "W05-04",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '88
Call AddSign("W05-04a",  "Bikeway Narrows",                              "W05-04a",  celPath, "W05-04a",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '89

Call AddSign("W06-01",   "Divided Highway",                              "W06-01",   celPath, "W06-01",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '90
Call AddSign("W06-02",   "Divided Highway Ends",                         "W06-02",   celPath, "W06-02",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '91
Call AddSign("W06-03",   "Traffic (2 Way)",                              "W06-03",   celPath, "W06-03",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '92
Call AddSign("W06-04",   "Traffic (2 Way) (Median) (Construction)",      "W06-04",   celPath, "W06-04",   "12"" x 18""", "12"" x 18""", 12, 18, postType, postPath, defSpacing) '93

Call AddSign("W07-01",   "Hill (Symbol)",                                "W07-01",   celPath, "W07-01",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '94
Call AddSign("W07-01a",  "Hill (With Grade) (Symbol)",                   "W07-01a",  celPath, "W07-01a",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '95
Call AddSign("W07-02bP", "Auxiliary (Hill) (Use Lower Gear) (Trucks)",   "W07-02bP", celPath, "W07-02bP", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '96
Call AddSign("W07-02P",  "Auxiliary (Hill) (Use Low Gear)",              "W07-02P",  celPath, "W07-02P",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '97
Call AddSign("W07-03aP", "Auxiliary (Hill) (Next Miles)",                "W07-03aP", celPath, "W07-03aP", "24"" x 18""", "36"" x 30""", 24, 18, postType, postPath, defSpacing) '98
Call AddSign("W07-03bP", "Auxiliary (Hill) (Grade and Miles)",           "W07-03bP", celPath, "W07-03bP", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '99
Call AddSign("W07-03P",  "Auxiliary (Hill) (Grade)",                     "W07-03P",  celPath, "W07-03P",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '100

' ---- Block 4 of 5 — Signs 301–400 ----

Call AddSign("W07-04",   "Truck (Escape Ramp) (Advance)",                 "W07-04",   celPath, "W07-04",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '301
Call AddSign("W07-04b",  "Truck (Escape Ramp) (Directional)",             "W07-04b",  celPath, "W07-04b",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '302
Call AddSign("W07-04c",  "Truck (Escape Ramp)",                           "W07-04c",  celPath, "W07-04c",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '303
Call AddSign("W07-04dP", "Auxiliary (Escape Ramp) (Sand)",                "W07-04dP", celPath, "W07-04dP", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '304
Call AddSign("W07-04eP", "Auxiliary (Escape Ramp) (Gravel)",              "W07-04eP", celPath, "W07-04eP", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '305
Call AddSign("W07-04fP", "Auxiliary (Escape Ramp) (Paved)",               "W07-04fP", celPath, "W07-04fP", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '306
Call AddSign("W07-05",   "Hill (Bicycle)",                                "W07-05",   celPath, "W07-05",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '307
Call AddSign("W07-06",   "Hill Blocks View",                              "W07-06",   celPath, "W07-06",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '308

Call AddSign("W08-01",   "Bump",                                          "W08-01",   celPath, "W08-01",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '309
Call AddSign("W08-02",   "Dip",                                           "W08-02",   celPath, "W08-02",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '310
Call AddSign("W08-03",   "Pavement Ends",                                 "W08-03",   celPath, "W08-03",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '311
Call AddSign("W08-04",   "Shoulder (Soft)",                               "W08-04",   celPath, "W08-04",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '312
Call AddSign("W08-05",   "Slippery",                                      "W08-05",   celPath, "W08-05",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '313
Call AddSign("W08-05aP", "Auxiliary (Slippery) (Ice)",                    "W08-05aP", celPath, "W08-05aP", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '314
Call AddSign("W08-05bP", "Auxiliary (Slippery) (Steel Deck)",             "W08-05bP", celPath, "W08-05bP", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '315
Call AddSign("W08-05cP", "Auxiliary (Slippery) (Excess Oil)",             "W08-05cP", celPath, "W08-05cP", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '316
Call AddSign("W08-05P",  "Auxiliary (Slippery) (When Wet)",               "W08-05P",  celPath, "W08-05P",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '317
Call AddSign("W08-07",   "Loose Gravel",                                  "W08-07",   celPath, "W08-07",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '318
Call AddSign("W08-08",   "Road (Rough)",                                  "W08-08",   celPath, "W08-08",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '319
Call AddSign("W08-09",   "Shoulder (Low)",                                "W08-09",   celPath, "W08-09",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '320
Call AddSign("W08-10",   "Slippery (Bicycle)",                            "W08-10",   celPath, "W08-10",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '321
Call AddSign("W08-11",   "Lanes (Uneven)",                                "W08-11",   celPath, "W08-11",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '322
Call AddSign("W08-12",   "No Center Stripe",                              "W08-12",   celPath, "W08-12",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '323
Call AddSign("W08-13",   "Bridge (Ices Before Road)",                     "W08-13",   celPath, "W08-13",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '324
Call AddSign("W08-14",   "Fallen Rocks",                                  "W08-14",   celPath, "W08-14",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '325
Call AddSign("W08-14SNY","Slides",                                        "W08-14SNY",celPath, "W08-14SNY","36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '326
Call AddSign("W08-15",   "Pavement (Grooved)",                            "W08-15",   celPath, "W08-15",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '327
Call AddSign("W08-15B",  "Pavement (Brick)",                              "W08-15B",  celPath, "W08-15B",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '328
Call AddSign("W08-15NY", "Rumble Strips -",                               "W08-15NY", celPath, "W08-15NY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '329
Call AddSign("W08-15P",  "Auxiliary (Motorcycle)",                        "W08-15P",  celPath, "W08-15P",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '330
Call AddSign("W08-15T",  "Pavement (Textured)",                           "W08-15T",  celPath, "W08-15T",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '331
Call AddSign("W08-16",   "Bridge (Metal Deck)",                           "W08-16",   celPath, "W08-16",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '332
Call AddSign("W08-17L",  "Shoulder (Drop Off) (Symbol) (Left)",           "W08-17L",  celPath, "W08-17L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '333
Call AddSign("W08-17P",  "Auxiliary (Shoulder Drop Off)",                 "W08-17P",  celPath, "W08-17P",  "24"" x 18""", "30"" x 24""", 24, 18, postType, postPath, defSpacing)      '334
Call AddSign("W08-17R",  "Shoulder (Drop Off) (Symbol) (Right)",          "W08-17R",  celPath, "W08-17R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '335
Call AddSign("W08-18",   "Flood (Road May)",                              "W08-18",   celPath, "W08-18",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '336
Call AddSign("W08-19",   "Flood (Guage)",                                 "W08-19",   celPath, "W08-19",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '337
Call AddSign("W08-21",   "Gusty Winds Area",                              "W08-21",   celPath, "W08-21",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '338
Call AddSign("W08-22",   "Fog Area",                                      "W08-22",   celPath, "W08-22",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '339
Call AddSign("W08-22NY", "Smoke",                                         "W08-22NY", celPath, "W08-22NY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '340
Call AddSign("W08-23",   "Shoulders (No)",                                "W08-23",   celPath, "W08-23",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '341
Call AddSign("W08-24",   "Pavement (Steel Plate on)",                     "W08-24",   celPath, "W08-24",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '342
Call AddSign("W08-25",   "Shoulder (Ends)",                               "W08-25",   celPath, "W08-25",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '343
End Sub

Private Sub LoadSigns_Part4()
Call AddSign("W09-01L",  "Lane (Ends) (Left)",                            "W09-01L",  celPath, "W09-01L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '344
Call AddSign("W09-01NY", "Lane (Single)",                                 "W09-01NY", celPath, "W09-01NY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '345
Call AddSign("W09-01R",  "Lane (Ends) (Right)",                           "W09-01R",  celPath, "W09-01R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '346
Call AddSign("W09-03",   "Closed (Lane) (Ahead) (Center)",                "W09-03",   celPath, "W09-03",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '347
Call AddSign("W09-04",   "Toll Road Advance",                             "W09-04",   celPath, "W09-04",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '348
Call AddSign("W09-05",   "Toll Road Begins",                              "W09-05",   celPath, "W09-05",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '349
Call AddSign("W09-06",   "Pay Toll Advance - Cars Price",                 "W09-06",   celPath, "W09-06",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '350
Call AddSign("W09-06a",  "Pay Toll Ahead - Cars Price",                   "W09-06a",  celPath, "W09-06a",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '351
Call AddSign("W09-06aP", "Auxiliary (Pay Toll Ahead)",                    "W09-06aP", celPath, "W09-06aP", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '352
Call AddSign("W09-06P",  "Auxiliary (Pay Toll Advance - Cars Price)",     "W09-06P",  celPath, "W09-06P",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '353
Call AddSign("W09-07L",  "Lane Exit Only Ahead (Left)",                   "W09-07L",  celPath, "W09-07L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '354
Call AddSign("W09-07R",  "Lane Exit Only Ahead (Right)",                  "W09-07R",  celPath, "W09-07R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '355

Call AddSign("W10-01",   "Rail Road (Crossing) (Advance)",                "W10-01",   celPath, "W10-01",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '356
Call AddSign("W10-01aP", "Auxiliary (Exempt) (Train) (Warning)",          "W10-01aP", celPath, "W10-01aP", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '357
Call AddSign("W10-01NY", "Auxiliary (Train) (When Flashing)",             "W10-01NY", celPath, "W10-01NY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '358
Call AddSign("W10-02L",  "Rail Road (Crossing) (Left) (at Cross Road)",   "W10-02L",  celPath, "W10-02L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '359
Call AddSign("W10-02R",  "Rail Road (Crossing) (Right) (at Cross Road)",  "W10-02R",  celPath, "W10-02R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '360
Call AddSign("W10-03L",  "Rail Road (Crossing) (Left) (at Side Road Left)","W10-03L", celPath, "W10-03L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '361
Call AddSign("W10-03R",  "Rail Road (Crossing) (Right) (at Side Road Right)","W10-03R",celPath,"W10-03R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '362
Call AddSign("W10-04L",  "Rail Road (Crossing) (Left) (at T Intersection)","W10-04L", celPath, "W10-04L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '363
Call AddSign("W10-04R",  "Rail Road (Crossing) (Right) (at T Intersection)","W10-04R",celPath,"W10-04R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '364
Call AddSign("W10-05",   "Rail Road (Low Ground Clearance)",              "W10-05",   celPath, "W10-05",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '365
Call AddSign("W10-05P",  "Auxiliary (Train) (Low Ground Clearance)",      "W10-05P",  celPath, "W10-05P",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '366
Call AddSign("W10-07",   "Rail Road (Light) (Blank Out)",                 "W10-07",   celPath, "W10-07",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '367
Call AddSign("W10-08",   "Rail Road (Trains May Exceed XX MPH)",          "W10-08",   celPath, "W10-08",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '368
Call AddSign("W10-09",   "Rail Road (No Train Horn)",                     "W10-09",   celPath, "W10-09",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '369
Call AddSign("W10-09P",  "Auxiliary (Train) (No Horn)",                   "W10-09P",  celPath, "W10-09P",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '370

Call AddSign("W10-11a",  "Auxiliary (Train) (Storage Space)",             "W10-11a",  celPath, "W10-11a",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '371
Call AddSign("W10-11b",  "Auxiliary (Train) (Storage Space) (Behind)",    "W10-11b",  celPath, "W10-11b",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '372
Call AddSign("W10-11C",  "Rail Road (Storage Space) (at Cross Road)",     "W10-11C",  celPath, "W10-11C",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '373
Call AddSign("W10-11L",  "Rail Road (Storage Space) (at Side Road Left)", "W10-11L",  celPath, "W10-11L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '374
Call AddSign("W10-11R",  "Rail Road (Storage Space) (at Side Road Right)","W10-11R",  celPath, "W10-11R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '375
Call AddSign("W10-11T",  "Rail Road (Storage Space) (at T-Intersection)", "W10-11T",  celPath, "W10-11T",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '376
Call AddSign("W10-12L",  "Rail Road (Crossing) (Skewed Left)",            "W10-12L",  celPath, "W10-12L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '377
Call AddSign("W10-12R",  "Rail Road (Crossing) (Skewed Right)",           "W10-12R",  celPath, "W10-12R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '378
Call AddSign("W10-13P",  "Auxiliary (Train) (No Gates or Lights)",        "W10-13P",  celPath, "W10-13P",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '379
Call AddSign("W10-14aP", "Auxiliary (Train) (Use Next Crossing)",         "W10-14aP", celPath, "W10-14aP", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '380
Call AddSign("W10-14P",  "Auxiliary (Train) (Next Crossing)",             "W10-14P",  celPath, "W10-14P",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '381
Call AddSign("W10-15P",  "Auxiliary (Train) (Rough Crossing)",            "W10-15P",  celPath, "W10-15P",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '382

Call AddSign("W11-01",   "Traffic (Bicycle)",                             "W11-01",   celPath, "W11-01",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '383
Call AddSign("W11-01aNY","Auxiliary (Bicycle) (In Lane)",                 "W11-01aNY",celPath, "W11-01aNY","36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '384
Call AddSign("W11-02L",  "Traffic (Pedestrian) (Left)",                   "W11-02L",  celPath, "W11-02L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '385
Call AddSign("W11-02NY", "Traffic (Pedestrian) (Safety Zone)",            "W11-02NY", celPath, "W11-02NY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '386
Call AddSign("W11-02R",  "Traffic (Pedestrian) (Right)",                  "W11-02R",  celPath, "W11-02R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '387
Call AddSign("W11-03",   "Traffic (Deer)",                                "W11-03",   celPath, "W11-03",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '388
Call AddSign("W11-04",   "Traffic (Cattle)",                              "W11-04",   celPath, "W11-04",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '389
Call AddSign("W11-05",   "Traffic (Farm Machinery)",                      "W11-05",   celPath, "W11-05",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '390
Call AddSign("W11-05a",  "Traffic (Farm Machinery) (Modern)",             "W11-05a",  celPath, "W11-05a",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '391
Call AddSign("W11-06",   "Traffic (Snowmobile)",                          "W11-06",   celPath, "W11-06",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '392
Call AddSign("W11-06NY", "Traffic (Motorcycle)",                          "W11-06NY", celPath, "W11-06NY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '393
Call AddSign("W11-07",   "Traffic (Equestrian)",                          "W11-07",   celPath, "W11-07",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '394
Call AddSign("W11-08",   "Traffic (Emergency Vehicle)",                   "W11-08",   celPath, "W11-08",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '395
Call AddSign("W11-08NY", "Traffic (Snowplow Turn)",                       "W11-08NY", celPath, "W11-08NY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '396
Call AddSign("W11-09",   "Traffic (Handicapped)",                         "W11-09",   celPath, "W11-09",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '397
Call AddSign("W11-10",   "Traffic (Truck)",                               "W11-10",   celPath, "W11-10",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '398
Call AddSign("W11-10BNY","Traffic (Buses)",                               "W11-10BNY",celPath,"W11-10BNY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '399
Call AddSign("W11-10BTNY","Traffic (Bus Turn)",                           "W11-10BTNY",celPath,"W11-10BTNY","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)        '400

' ---- Block 5 — continuing after Block 4 ----

Call AddSign("W11-10SMBNY", "Traffic (Slow Moving Buses)", "W11-10SMBNY", celPath, "W11-10SMBNY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)        '401
Call AddSign("W11-11",     "Traffic (Golf Cart)",           "W11-11",      celPath, "W11-11",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '402
Call AddSign("W11-11NY",   "Traffic (All Teraine Vehicle)", "W11-11NY",    celPath, "W11-11NY",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)         '403
Call AddSign("W11-12P",    "Auxiliary (Signal) (Emergency) (Ahead)", "W11-12P", celPath, "W11-12P", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)        '404
Call AddSign("W11-14",     "Traffic (Horsedrawn Vehicle)",  "W11-14",      celPath, "W11-14",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '405
Call AddSign("W11-14NY",   "Traffic (Low Flying Planes)",   "W11-14NY",    celPath, "W11-14NY",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)         '406
Call AddSign("W11-15L",    "Traffic (Trail) (Left)",        "W11-15L",     celPath, "W11-15L",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '407
Call AddSign("W11-15P",    "Auxiliary (Trail X-ing)",       "W11-15P",     celPath, "W11-15P",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '408
Call AddSign("W11-15R",    "Traffic (Trail) (Right)",       "W11-15R",     celPath, "W11-15R",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '409
Call AddSign("W11-16",     "Traffic (Bear)",                "W11-16",      celPath, "W11-16",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '410
Call AddSign("W11-16NY",   "Traffic (Bird Nesting Area)",   "W11-16NY",    celPath, "W11-16NY",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)         '411
Call AddSign("W11-17",     "Traffic (Sheep)",               "W11-17",      celPath, "W11-17",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '412
Call AddSign("W11-18",     "Traffic (Mountain Goat)",       "W11-18",      celPath, "W11-18",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '413
Call AddSign("W11-19",     "Traffic (Donkey)",              "W11-19",      celPath, "W11-19",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '414
Call AddSign("W11-20",     "Traffic (Elk)",                 "W11-20",      celPath, "W11-20",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)             '415
Call AddSign("W11-21",     "Traffic (Moose)",               "W11-21",      celPath, "W11-21",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '416
Call AddSign("W11-22",     "Traffic (Horse)",               "W11-22",      celPath, "W11-22",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '417

Call AddSign("W12-01",     "Pass Left or Right",            "W12-01",      celPath, "W12-01",      "30"" x 30""", "36"" x 36""", 30, 30, postType, postPath, defSpacing)             '418
Call AddSign("W12-01NY",   "Pass Left or Right (Advance)",  "W12-01NY",    celPath, "W12-01NY",    "30"" x 30""", "36"" x 36""", 30, 30, postType, postPath, defSpacing)           '419
Call AddSign("W12-02",     "Restriction (Clearance) (Bridge) (Advance)", "W12-02", celPath, "W12-02", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)        '420
Call AddSign("W12-02vNY",  "Underpass",                    "W12-02vNY",    celPath, "W12-02vNY",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '421
Call AddSign("W12-02wNY",  "Restriction (Weight Limit) (Bridge)", "W12-02wNY", celPath, "W12-02wNY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)        '422
Call AddSign("W12-02xNY",  "Restriction (Weight Limit) (Warning)", "W12-02xNY", celPath, "W12-02xNY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)      '423
Call AddSign("W12-02yNY",  "Restriction (Weight Limit) (No R Permit Trucks)", "W12-02yNY", celPath, "W12-02yNY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '424
Call AddSign("W12-02zNY",  "Restriction (Weight Limit) (No R Permit Trucks) (W)", "W12-02zNY", celPath, "W12-02zNY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '425

Call AddSign("W13-01",     "Auxiliary (Speed) (Advisory)",  "W13-01",      celPath, "W13-01",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '426
Call AddSign("W13-02",     "Speed (Advisory) (Exit)",       "W13-02",      celPath, "W13-02",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '427
Call AddSign("W13-03",     "Speed (Advisory) (Ramp)",       "W13-03",      celPath, "W13-03",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '428
Call AddSign("W13-04LLP",  "Auxiliary (Left Lane)",         "W13-04LLP",   celPath, "W13-04LLP",   "36"" x 36""", "36"" x 36""", 36, 36, postType, postPath, defSpacing)            '429
Call AddSign("W13-04P",    "Auxiliary (On Ramp)",           "W13-04P",     celPath, "W13-04P",     "36"" x 36""", "36"" x 36""", 36, 36, postType, postPath, defSpacing)            '430
Call AddSign("W13-04RLP",  "Auxiliary (Right Lane)",        "W13-04RLP",   celPath, "W13-04RLP",   "36"" x 36""", "36"" x 36""", 36, 36, postType, postPath, defSpacing)            '431
Call AddSign("W13-06",     "Speed (Advisory) (Exit) (With Arrow)", "W13-06", celPath, "W13-06", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)              '432
Call AddSign("W13-07",     "Speed (Advisory) (Ramp) (With Arrow)", "W13-07", celPath, "W13-07", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)              '433

Call AddSign("W14-01",     "Dead End",                      "W14-01",      celPath, "W14-01",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '434
Call AddSign("W14-01aL",   "Dead End (Left)",               "W14-01aL",    celPath, "W14-01aL",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '435
Call AddSign("W14-01aR",   "Dead End (Right)",              "W14-01aR",    celPath, "W14-01aR",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '436
Call AddSign("W14-02",     "No Outlet",                     "W14-02",      celPath, "W14-02",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '437
Call AddSign("W14-02aL",   "No Outlet (Left)",              "W14-02aL",    celPath, "W14-02aL",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '438
Call AddSign("W14-02aR",   "No Outlet (Right)",             "W14-02aR",    celPath, "W14-02aR",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '439
Call AddSign("W14-03",     "No Passing Zone",               "W14-03",      celPath, "W14-03",      "48"" x 48""", "64"" x 64""", 48, 48, postType, postPath, defSpacing)            '440
Call AddSign("W14-03NY",   "Accident Ahead",                "W14-03NY",    celPath, "W14-03NY",    "48"" x 48""", "64"" x 64""", 48, 48, postType, postPath, defSpacing)          '441
Call AddSign("W14-04NY",   "Emergency Scene Ahead",         "W14-04NY",    celPath, "W14-04NY",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '442

Call AddSign("W15-01",     "Playground",                    "W15-01",      celPath, "W15-01",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '443
Call AddSign("W15-01xNY",  "Children at Play",              "W15-01xNY",   celPath, "W15-01xNY",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)         '444
Call AddSign("W15-01yNY",  "Deaf Child Area",               "W15-01yNY",   celPath, "W15-01yNY",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)         '445
Call AddSign("W15-01zNY",  "Blind Child Area",              "W15-01zNY",   celPath, "W15-01zNY",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)         '446

Call AddSign("W16-01P",    "Auxiliary (Share the Road)",    "W16-01P",     celPath, "W16-01P",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '447
Call AddSign("W16-02aP",   "Auxiliary (Distance) (Feet) (1 Line) (Warning)", "W16-02aP", celPath, "W16-02aP", "24"" x 18""", "30"" x 24""", 24, 18, postType, postPath, defSpacing) '448
Call AddSign("W16-02P",    "Auxiliary (Distance) (Feet) (2 Lines) (Warning)","W16-02P",  celPath, "W16-02P",  "24"" x 18""", "30"" x 24""", 24, 18, postType, postPath, defSpacing) '449
Call AddSign("W16-03aP",   "Auxiliary (Distance) (Mile) (1 Line) (Warning)", "W16-03aP", celPath, "W16-03aP", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '450
Call AddSign("W16-03P",    "Auxiliary (Distance) (Mile) (2 Lines) (Warning)","W16-03P",  celPath, "W16-03P",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing) '451
Call AddSign("W16-04P",    "Auxiliary (Distance) (Next Feet) (2 Lines) (Warning)","W16-04P",celPath,"W16-04P","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)    '452
Call AddSign("W16-05PL",   "Arrow (Left) (Warning)",        "W16-05PL",    celPath, "W16-05PL",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '453
Call AddSign("W16-05PR",   "Arrow (Right) (Warning)",       "W16-05PR",    celPath, "W16-05PR",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '454
Call AddSign("W16-06PL",   "Arrow (Broken Left) (Warning)", "W16-06PL",    celPath, "W16-06PL",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '455
Call AddSign("W16-06PR",   "Arrow (Broken Right) (Warning)","W16-06PR",    celPath, "W16-06PR",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '456
Call AddSign("W16-07PL",   "Arrow (Diagonal Left) (Warning)","W16-07PL",   celPath, "W16-07PL",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '457
Call AddSign("W16-07PR",   "Arrow (Diagonal Right) (Warning)","W16-07PR",  celPath, "W16-07PR",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '458
Call AddSign("W16-09P",    "Auxiliary (Ahead) (Warning)",   "W16-09P",     celPath, "W16-09P",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '459
Call AddSign("W16-10P",    "Auxiliary (Photo Enforced) (Symbol) (Warning)","W16-10P", celPath, "W16-10P", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '460
Call AddSign("W16-11LP",   "Auxiliary (HOV) (Symbol) (Lane)","W16-11LP",  celPath, "W16-11LP",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '461
Call AddSign("W16-11OP",   "Auxiliary (HOV) (Symbol) (Only)","W16-11OP",  celPath, "W16-11OP",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '462
Call AddSign("W16-11P",    "Auxiliary (HOV) (Symbol)",      "W16-11P",     celPath, "W16-11P",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '463
Call AddSign("W16-12P",    "Auxiliary (Traffic Circle) (Warning)","W16-12P", celPath,"W16-12P",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '464
Call AddSign("W16-12wPNY", "Auxiliary (Driveways) (1 Line)","W16-12wPNY",  celPath, "W16-12wPNY",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '465
Call AddSign("W16-12xPNY", "Auxiliary (Driveways) (2 Lines)","W16-12xPNY", celPath, "W16-12xPNY",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '466
Call AddSign("W16-12zNY",  "Increased Enforcement Area",    "W16-12zNY",   celPath, "W16-12zNY",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '467
Call AddSign("W16-13P",    "Auxiliary (When Flashing) (Warning)","W16-13P", celPath, "W16-13P",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '468
Call AddSign("W16-15P",    "Auxiliary (New) (Warning)",     "W16-15P",     celPath, "W16-15P",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '469
Call AddSign("W16-16P",    "Auxiliary (Last Exit Before Toll)","W16-16P",  celPath, "W16-16P",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '470
Call AddSign("W16-17P",    "Auxiliary (Roundabout)",        "W16-17P",     celPath, "W16-17P",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '471

Call AddSign("W17-01B",    "Speed Bump",                    "W17-01B",     celPath, "W17-01B",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '472
Call AddSign("W17-01C",    "Raised Crosswalk",              "W17-01C",     celPath, "W17-01C",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '473
Call AddSign("W17-01H",    "Speed Hump",                    "W17-01H",     celPath, "W17-01H",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '474
Call AddSign("W17-01I",    "Raised Intersection",           "W17-01I",     celPath, "W17-01I",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '475

Call AddSign("W18-01",     "No Traffic Signs",              "W18-01",      celPath, "W18-01",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '476
Call AddSign("W18-01NY",   "Minimum Maintenance Road",     "W18-01NY",    celPath, "W18-01NY",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)         '477

Call AddSign("W19-01",     "Freeway Ends Advance",         "W19-01",       celPath, "W19-01",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '478
Call AddSign("W19-02",     "Expressway Ends Advance",      "W19-02",       celPath, "W19-02",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '479
Call AddSign("W19-03",     "Freeway Ends",                 "W19-03",       celPath, "W19-03",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)             '480
Call AddSign("W19-04",     "Expressway Ends",              "W19-04",       celPath, "W19-04",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '481
Call AddSign("W19-05",     "All Traffic Must Exit",        "W19-05",       celPath, "W19-05",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '482

Call AddSign("W20-01RA",   "Work (Road) (Ahead)",          "W20-01RA",     celPath, "W20-01RA",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '483
Call AddSign("W20-01RF",   "Work (Road) (Feet)",           "W20-01RF",     celPath, "W20-01RF",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '484
Call AddSign("W20-01RM",   "Work (Road) (Mile)",           "W20-01RM",     celPath, "W20-01RM",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '485
Call AddSign("W20-01RPM",  "Work (Road) (Part Mile)",      "W20-01RPM",    celPath, "W20-01RPM",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '486
Call AddSign("W20-01SA",   "Work (Street) (Ahead)",        "W20-01SA",     celPath, "W20-01SA",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '487
Call AddSign("W20-01SF",   "Work (Street) (Feet)",         "W20-01SF",     celPath, "W20-01SF",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '488
Call AddSign("W20-01SM",   "Work (Street) (Mile)",         "W20-01SM",     celPath, "W20-01SM",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '489
Call AddSign("W20-01SPM",  "Work (Street) (Part Mile)",    "W20-01SPM",    celPath, "W20-01SPM",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '490

Call AddSign("W20-02A",    "Detour (Ahead)",               "W20-02A",      celPath, "W20-02A",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '491
Call AddSign("W20-02F",    "Detour (Feet)",                "W20-02F",      celPath, "W20-02F",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '492
Call AddSign("W20-02M",    "Detour (Mile)",                "W20-02M",      celPath, "W20-02M",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '493
Call AddSign("W20-02PM",   "Detour (Part Mile)",           "W20-02PM",     celPath, "W20-02PM",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '494

Call AddSign("W20-03BANY", "Closed (Bridge) (Ahead)",      "W20-03BANY",   celPath, "W20-03BANY",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '495
Call AddSign("W20-03BFNY", "Closed (Bridge) (Feet)",       "W20-03BFNY",   celPath, "W20-03BFNY",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '496
Call AddSign("W20-03BMNY", "Closed (Bridge) (Mile)",       "W20-03BMNY",   celPath, "W20-03BMNY",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '497
Call AddSign("W20-03BPMNY","Closed (Bridge) (Part Mile)",  "W20-03BPMNY",  celPath, "W20-03BPMNY", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)         '498
Call AddSign("W20-03RA",   "Closed (Road) (Ahead)",        "W20-03RA",     celPath, "W20-03RA",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '499
Call AddSign("W20-03RF",   "Closed (Road) (Feet)",         "W20-03RF",     celPath, "W20-03RF",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '500
Call AddSign("W20-03RM",   "Closed (Road) (Mile)",         "W20-03RM",     celPath, "W20-03RM",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '501
Call AddSign("W20-03RPM",  "Closed (Road) (Part Mile)",    "W20-03RPM",    celPath, "W20-03RPM",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '502
Call AddSign("W20-03SA",   "Closed (Street) (Ahead)",      "W20-03SA",     celPath, "W20-03SA",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '503
Call AddSign("W20-03SF",   "Closed (Street) (Feet)",       "W20-03SF",     celPath, "W20-03SF",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '504
Call AddSign("W20-03SM",   "Closed (Street) (Mile)",       "W20-03SM",     celPath, "W20-03SM",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '505
Call AddSign("W20-03SPM",  "Closed (Street) (Part Mile)",  "W20-03SPM",    celPath, "W20-03SPM",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '506

Call AddSign("W20-04A",    "Lane (One) (Ahead)",           "W20-04A",      celPath, "W20-04A",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '507
Call AddSign("W20-04F",    "Lane (One) (Feet)",            "W20-04F",      celPath, "W20-04F",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '508
Call AddSign("W20-04M",    "Lane (One) (Mile)",            "W20-04M",      celPath, "W20-04M",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '509
Call AddSign("W20-04PM",   "Lane (One) (Part Mile)",       "W20-04PM",     celPath, "W20-04PM",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '510

Call AddSign("W20-05aLA",  "Closed (Lane) (Left) (Two) (Ahead)",  "W20-05aLA", celPath, "W20-05aLA", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)        '511
Call AddSign("W20-05aLF",  "Closed (Lane) (Left) (Two) (Feet)",   "W20-05aLF", celPath, "W20-05aLF", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)        '512
Call AddSign("W20-05aLM",  "Closed (Lane) (Left) (Two) (Mile)",   "W20-05aLM", celPath, "W20-05aLM", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)        '513
Call AddSign("W20-05aLPM", "Closed (Lane) (Left) (Two) (Part Mile)","W20-05aLPM", celPath, "W20-05aLPM","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)          '514
Call AddSign("W20-05aRA",  "Closed (Lane) (Right) (Two) (Ahead)", "W20-05aRA", celPath, "W20-05aRA", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)        '515
Call AddSign("W20-05aRF",  "Closed (Lane) (Right) (Two) (Feet)",  "W20-05aRF", celPath, "W20-05aRF", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)        '516
Call AddSign("W20-05aRM",  "Closed (Lane) (Right) (Two) (Mile)",  "W20-05aRM", celPath, "W20-05aRM", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)        '517
Call AddSign("W20-05aRPM", "Closed (Lane) (Right) (Two) (Part Mile)","W20-05aRPM", celPath,"W20-05aRPM","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)         '518

Call AddSign("W20-05LA",   "Closed (Lane) (Left) (Ahead)",  "W20-05LA",    celPath, "W20-05LA",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '519
Call AddSign("W20-05LF",   "Closed (Lane) (Left) (Feet)",   "W20-05LF",    celPath, "W20-05LF",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '520
Call AddSign("W20-05LM",   "Closed (Lane) (Left) (Mile)",   "W20-05LM",    celPath, "W20-05LM",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '521
Call AddSign("W20-05LPM",  "Closed (Lane) (Left) (Part Mile)","W20-05LPM", celPath, "W20-05LPM",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '522
Call AddSign("W20-05RA",   "Closed (Lane) (Right) (Ahead)", "W20-05RA",    celPath, "W20-05RA",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '523
Call AddSign("W20-05RF",   "Closed (Lane) (Right) (Feet)",  "W20-05RF",    celPath, "W20-05RF",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '524
Call AddSign("W20-05RM",   "Closed (Lane) (Right) (Mile)",  "W20-05RM",    celPath, "W20-05RM",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '525
Call AddSign("W20-05RPM",  "Closed (Lane) (Right) (Part Mile)","W20-05RPM", celPath, "W20-05RPM",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)         '526

Call AddSign("W20-07",     "Flagger (Symbol)",             "W20-07",       celPath, "W20-07",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '527
Call AddSign("W20-08",     "Slow (Flagger Paddle)",        "W20-08",       celPath, "W20-08",      "18"" x 18""", "18"" x 18""", 18, 18, postType, postPath, defSpacing)             '528

Call AddSign("W21-01",     "Workers (Symbol)",             "W21-01",       celPath, "W21-01",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '529
Call AddSign("W21-02DNY",  "Wet Paint (Diamond)",          "W21-02DNY",    celPath, "W21-02DNY",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '530
Call AddSign("W21-02NY",   "Wet Paint",                    "W21-02NY",     celPath, "W21-02NY",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '531
Call AddSign("W21-02O",    "Fresh Oil",                    "W21-02O",      celPath, "W21-02O",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '532
Call AddSign("W21-02T",    "Fresh Tar",                    "W21-02T",      celPath, "W21-02T",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '533

Call AddSign("W21-03A",    "Road Machinery (Ahead)",       "W21-03A",      celPath, "W21-03A",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '534
Call AddSign("W21-03F",    "Road Machinery (Feet)",        "W21-03F",      celPath, "W21-03F",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '535
Call AddSign("W21-03M",    "Road Machinery (Mile)",        "W21-03M",      celPath, "W21-03M",     "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '536
Call AddSign("W21-03PM",   "Road Machinery (Part Mile)",   "W21-03PM",     celPath, "W21-03PM",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '537

Call AddSign("W21-04",     "Slow Moving Vehicle",          "W21-04",       celPath, "W21-04",      "36"" x 18""", "36"" x 18""", 36, 18, postType, postPath, defSpacing)             '538

Call AddSign("W21-05",     "Shoulder (Work)",              "W21-05",       celPath, "W21-05",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '539
Call AddSign("W21-05aL",   "Closed (Shoulder) (Left)",     "W21-05aL",     celPath, "W21-05aL",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '540
Call AddSign("W21-05aR",   "Closed (Shoulder) (Right)",    "W21-05aR",     celPath, "W21-05aR",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)           '541

Call AddSign("W21-05bLA",  "Closed (Shoulder) (Left) (Ahead)",  "W21-05bLA", celPath, "W21-05bLA", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '542
Call AddSign("W21-05bLF",  "Closed (Shoulder) (Left) (Feet)",   "W21-05bLF", celPath, "W21-05bLF", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '543
Call AddSign("W21-05bLM",  "Closed (Shoulder) (Left) (Mile)",   "W21-05bLM", celPath, "W21-05bLM", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '544
Call AddSign("W21-05bLPM", "Closed (Shoulder) (Left) (Part Mile)","W21-05bLPM", celPath, "W21-05bLPM","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)            '545
Call AddSign("W21-05bRA",  "Closed (Shoulder) (Right) (Ahead)", "W21-05bRA", celPath, "W21-05bRA", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)         '546
Call AddSign("W21-05bRF",  "Closed (Shoulder) (Right) (Feet)",  "W21-05bRF", celPath, "W21-05bRF", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)         '547
Call AddSign("W21-05bRM",  "Closed (Shoulder) (Right) (Mile)",  "W21-05bRM", celPath, "W21-05bRM", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)         '548
Call AddSign("W21-05bRPM", "Closed (Shoulder) (Right) (Part Mile)","W21-05bRPM", celPath,"W21-05bRPM","36"" x 36""","48"" x 48""",36,36,postType,postPath,defSpacing)           '549

Call AddSign("W21-06",     "Survey Crew",                  "W21-06",       celPath, "W21-06",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '550
Call AddSign("W21-07",     "Work (Ahead) (Uility)",        "W21-07",       celPath, "W21-07",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '551
Call AddSign("W21-08",     "Mowing Ahead",                 "W21-08",       celPath, "W21-08",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '552
Call AddSign("W21-08NY",   "Sandblasting",                 "W21-08NY",     celPath, "W21-08NY",    "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)          '553

Call AddSign("W22-01",     "Blasting Zone (Ahead)",        "W22-01",       celPath, "W22-01",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '554
Call AddSign("W22-02",     "Blasting Zone (Turn Off Radios and Cell Phone)", "W22-02", celPath, "W22-02","36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)   '555
Call AddSign("W22-03",     "Blasting Zone (End)",          "W22-03",       celPath, "W22-03",      "42"" x 36""", "42"" x 36""", 42, 36, postType, postPath, defSpacing)            '556

Call AddSign("W23-01",     "Slow Traffic Ahead",           "W23-01",       celPath, "W23-01",      "48"" x 24""", "48"" x 24""", 48, 24, postType, postPath, defSpacing)            '557
Call AddSign("W23-01wDNY", "Stay in Lane (Diamond)",       "W23-01wDNY",   celPath, "W23-01wDNY",  "48"" x 24""", "48"" x 24""", 48, 24, postType, postPath, defSpacing)          '558
Call AddSign("W23-01wNY",  "Stay in Lane",                 "W23-01wNY",    celPath, "W23-01wNY",   "48"" x 24""", "48"" x 24""", 48, 24, postType, postPath, defSpacing)           '559
Call AddSign("W23-01xDNY", "Do Not Pass (Diamond)",        "W23-01xDNY",   celPath, "W23-01xDNY",  "48"" x 24""", "48"" x 24""", 48, 24, postType, postPath, defSpacing)          '560
Call AddSign("W23-01xNY",  "Do Not Pass",                  "W23-01xNY",    celPath, "W23-01xNY",   "48"" x 24""", "48"" x 24""", 48, 24, postType, postPath, defSpacing)           '561
Call AddSign("W23-01yDNY", "Lane Closed (Diamond)",        "W23-01yDNY",   celPath, "W23-01yDNY",  "48"" x 24""", "48"" x 24""", 48, 24, postType, postPath, defSpacing)          '562
Call AddSign("W23-01yNY",  "Lane Closed",                  "W23-01yNY",    celPath, "W23-01yNY",   "48"" x 24""", "48"" x 24""", 48, 24, postType, postPath, defSpacing)           '563
Call AddSign("W23-02",     "New Traffic Pattern Ahead",    "W23-02",       celPath, "W23-02",      "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)            '564

Call AddSign("W24-01aL",   "Curve (Reverse) (Double) (2 Lanes) (Left)",  "W24-01aL", celPath, "W24-01aL", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '565
Call AddSign("W24-01aR",   "Curve (Reverse) (Double) (2 Lanes) (Right)", "W24-01aR", celPath, "W24-01aR", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '566
Call AddSign("W24-01bL",   "Curve (Reverse) (Double) (3 Lanes) (Left)",  "W24-01bL", celPath, "W24-01bL", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '567
Call AddSign("W24-01bR",   "Curve (Reverse) (Double) (3 Lanes) (Right)", "W24-01bR", celPath, "W24-01bR", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)    '568
Call AddSign("W24-01L",    "Curve (Reverse) (Double) (Left)",            "W24-01L",  celPath, "W24-01L",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '569
Call AddSign("W24-01R",    "Curve (Reverse) (Double) (Right)",           "W24-01R",  celPath, "W24-01R",  "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '570

Call AddSign("W25-01",     "Oncoming Traffic Has Extended Green",        "W25-01",   celPath, "W25-01",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '571
Call AddSign("W25-02",     "Oncoming Traffic May Have Extended Green",   "W25-02",   celPath, "W25-02",   "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)     '572

End Sub

' ============================================================
' ADD SIGN TO LIBRARY (PRIVATE - INTERNAL USE ONLY)
' Uses individual parameters instead of SignData type
' ============================================================
Private Sub AddSign(SignNumber As String, _
                    Description As String, _
                    CellName As String, _
                    CellLibPath As String, _
                    TextLabel As String, _
                    TextLine2NonFreeway As String, _
                    TextLine2Freeway As String, _
                    widthIn As Double, _
                    heightIn As Double, _
                    PostType As String, _
                    PostLibPath As String, _
                    spacing As Double)
    
    Dim sign As signData
    
    ' Build the sign data structure
    sign.SignNumber = SignNumber
    sign.Description = Description
    sign.CellName = CellName
    sign.CellLibraryPath = CellLibPath
    sign.TextLabel = TextLabel
    sign.TextLine2NonFreeway = TextLine2NonFreeway
    sign.TextLine2Freeway = TextLine2Freeway
    sign.TextLine2 = TextLine2NonFreeway   ' default; GetSignData overwrites by road type
    sign.WidthInches = widthIn
    sign.HeightInches = heightIn
    sign.PostType = PostType
    sign.PostLibraryPath = PostLibPath
    sign.DefaultSpacing = spacing
    
    ' Add to array (check duplicate first)
    Dim i As Long
    For i = 1 To signLibraryCount
        If signLibrary(i).SignNumber = sign.SignNumber Then
            Debug.Print "Warning: Sign " & sign.SignNumber & " already exists in library"
            Exit Sub
        End If
    Next i
    signLibraryCount = signLibraryCount + 1
    If signLibraryCount > UBound(signLibrary) Then
        ReDim Preserve signLibrary(1 To UBound(signLibrary) + 64)
    End If
    signLibrary(signLibraryCount) = sign
End Sub

' ============================================================
' GET SIGN FROM LIBRARY
' Returns sign data for a given sign number. Optional RoadType
' ("Freeway" or "Non-Freeway") sets TextLine2 to the matching size.
' ============================================================
Public Function GetSignData(SignNumber As String, Optional RoadType As String = "Non-Freeway") As signData
    Dim sign As signData
    Dim i As Long
    
    If Not signLibraryInitialized Then
        Call InitializeSignLibrary
    End If
    
    For i = 1 To signLibraryCount
        If signLibrary(i).SignNumber = SignNumber Then
            sign = signLibrary(i)
            If LCase(Trim(RoadType)) = "freeway" Then
                sign.TextLine2 = sign.TextLine2Freeway
            Else
                sign.TextLine2 = sign.TextLine2NonFreeway
            End If
            GetSignData = sign
            Exit Function
        End If
    Next i
    
    ' Not found
    sign.SignNumber = ""
    sign.Description = "Sign not found"
    Debug.Print "Warning: Sign " & SignNumber & " not found in library"
    GetSignData = sign
End Function

' ============================================================
' CHECK IF SIGN EXISTS
' ============================================================
Public Function SignExists(SignNumber As String) As Boolean
    Dim i As Long
    
    If Not signLibraryInitialized Then
        Call InitializeSignLibrary
    End If
    
    For i = 1 To signLibraryCount
        If signLibrary(i).SignNumber = SignNumber Then
            SignExists = True
            Exit Function
        End If
    Next i
    SignExists = False
End Function

' ============================================================
' GET ALL SIGN NUMBERS
' Returns an array of all sign numbers in the library
' ============================================================
Public Function GetAllSignNumbers() As String()
    Dim signNumbers() As String
    Dim i As Long
    
    If Not signLibraryInitialized Then
        Call InitializeSignLibrary
    End If
    
    If signLibraryCount = 0 Then
        ReDim signNumbers(0)
        signNumbers(0) = ""
        GetAllSignNumbers = signNumbers
        Exit Function
    End If
    
    ReDim signNumbers(1 To signLibraryCount)
    For i = 1 To signLibraryCount
        signNumbers(i) = signLibrary(i).SignNumber
    Next i
    GetAllSignNumbers = signNumbers
End Function

' ============================================================
' GET LIBRARY COUNT
' ============================================================
Public Function GetSignCount() As Integer
    If Not signLibraryInitialized Then
        Call InitializeSignLibrary
    End If
    GetSignCount = CInt(signLibraryCount)
End Function

' ============================================================
' ADD CUSTOM SIGN (PUBLIC - for runtime additions)
' ============================================================
Public Sub AddCustomSign(SignNumber As String, _
                        Description As String, _
                        CellName As String, _
                        CellLibPath As String, _
                        TextLabel As String, _
                        TextLine2NonFreeway As String, _
                        TextLine2Freeway As String, _
                        widthIn As Double, _
                        heightIn As Double, _
                        PostType As String, _
                        PostLibPath As String, _
                        spacing As Double)
    
    If Not signLibraryInitialized Then
        Call InitializeSignLibrary
    End If
    
    Call AddSign(SignNumber, Description, CellName, CellLibPath, _
                 TextLabel, TextLine2NonFreeway, TextLine2Freeway, _
                 widthIn, heightIn, PostType, PostLibPath, spacing)
End Sub

' ============================================================
' EXPORT LIBRARY TO TEXT FILE (for backup/reference)
' ============================================================
Public Sub ExportLibraryToFile(filePath As String)
    Dim fileNum As Integer
    Dim i As Long
    
    If Not signLibraryInitialized Then
        Call InitializeSignLibrary
    End If
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    Print #fileNum, "SIGN LIBRARY EXPORT"
    Print #fileNum, "===================="
    Print #fileNum, ""
    Print #fileNum, "Total Signs: " & signLibraryCount
    Print #fileNum, ""
    
    For i = 1 To signLibraryCount
        With signLibrary(i)
            Print #fileNum, "Sign Number: " & .SignNumber
            Print #fileNum, "Description: " & .Description
            Print #fileNum, "Cell Name: " & .CellName
            Print #fileNum, "Cell Library: " & .CellLibraryPath
            Print #fileNum, "Text Label: " & .TextLabel
            Print #fileNum, "Size Non-Freeway: " & .TextLine2NonFreeway & "  Freeway: " & .TextLine2Freeway
            Print #fileNum, "Dimensions: " & .WidthInches & """W x " & .HeightInches & """H"
            Print #fileNum, "Post Type: " & .PostType
            Print #fileNum, "Post Library: " & .PostLibraryPath
            Print #fileNum, "Default Spacing: " & .DefaultSpacing & " ft"
        End With
        Print #fileNum, "--------------------"
    Next i
    
    Close #fileNum
    
    Debug.Print "Library exported to: " & filePath
End Sub

' ============================================================
' CLEAR LIBRARY (for reinitialization)
' ============================================================
Public Sub ClearLibrary()
    signLibraryCount = 0
    signLibraryInitialized = True
    ReDim signLibrary(1 To 1)
End Sub

