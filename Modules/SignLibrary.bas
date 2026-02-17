
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
    Const celPath As String = "c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel"
    Const postPath As String = "c:\pwworking\usny\d0119091\ny_plan_wztc.cel"
    Const postType As String = "TWZSGN_P"
    Const defSpacing As Double = 350
    
    Call AddSign("R01-01", "Stop", "R01-01", celPath, "R01-01", "30"" x 30""", "30"" x 30""", 30, 30, postType, postPath, defSpacing)
    Call AddSign("R01-01-P", "Stop (on Stop/Slow Paddle)", "R01-01", celPath, "R01-01-P", "18"" x 18""", "18"" x 18""", 18, 18, postType, postPath, defSpacing)
    Call AddSign("R01-02", "Yield", "R01-02", celPath, "R01-02", "36"" x 36""", "36"" x 36""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("R01-02aP", "To Oncoming Traffic (plaque)", "R01-02aP", celPath, "R01-02aP", "36"" x 30""", "48"" x 36""", 36, 30, postType, postPath, defSpacing)
    Call AddSign("R01-07", "Wait on Stop", "R01-07", celPath, "R01-07", "24"" x 30""", "24"" x 30""", 24, 30, postType, postPath, defSpacing)
    Call AddSign("R01-07a", "Wait on Stop - Go on Slow", "R01-07a", celPath, "R01-07a", "30"" x 36""", "30"" x 36""", 30, 36, postType, postPath, defSpacing)
    Call AddSign("R01-08", "Go on Slow", "R01-08", celPath, "R01-08", "24"" x 30""", "24"" x 30""", 24, 30, postType, postPath, defSpacing)
    Call AddSign("R02-01", "Speed Limit", "R02-01", celPath, "R02-01", "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
    Call AddSign("R02-06P", "Fines Higher (plaque)", "R02-06P", celPath, "R02-06P", "24"" x 18""", "36"" x 24""", 24, 18, postType, postPath, defSpacing)
    Call AddSign("R02-06aP", "Fines Double (plaque)", "R02-06aP", celPath, "R02-06aP", "24"" x 18""", "36"" x 24""", 24, 18, postType, postPath, defSpacing)
    Call AddSign("R02-06bP", "$XX Fine (plaque)", "R02-06bP", celPath, "R02-06bP", "24"" x 18""", "36"" x 24""", 24, 18, postType, postPath, defSpacing)
    Call AddSign("R02-10", "Begin Higher Fines Zone", "R02-10", celPath, "R02-10", "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
    Call AddSign("R02-11", "End Higher Fines Zone", "R02-11", celPath, "R02-11", "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
    Call AddSign("R02-12", "End Work Zone Speed Limit", "R02-12", celPath, "R02-12", "24"" x 36""", "36"" x 54""", 24, 36, postType, postPath, defSpacing)
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
    Call AddSign("R05-01", "Do Not Enter", "R05-01", celPath, "R05-01", "30"" x 30""", "36"" x 36""", 30, 30, postType, postPath, defSpacing)
    Call AddSign("R05-01a", "Wrong Way", "R05-01a", celPath, "R05-01a", "36"" x 24""", "42"" x 30""", 36, 24, postType, postPath, defSpacing)
    Call AddSign("R06-01", "One Way", "R06-01", celPath, "R06-01", "36"" x 12""", "48"" x 18""", 36, 12, postType, postPath, defSpacing)
    Call AddSign("R06-02", "One Way", "R06-02", celPath, "R06-02", "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
    Call AddSign("R08-03", "No Parking (symbol)", "R08-03", celPath, "R08-03", "24"" x 24""", "36"" x 36""", 24, 24, postType, postPath, defSpacing)
    Call AddSign("R09-08", "Pedestrian Crosswalk", "R09-08", celPath, "R09-08", "36"" x 18""", "36"" x 18""", 36, 18, postType, postPath, defSpacing)
    Call AddSign("R09-09", "Sidewalk Closed", "R09-09", celPath, "R09-09", "24"" x 12""", "24"" x 12""", 24, 12, postType, postPath, defSpacing)
    Call AddSign("R09-10", "Sidewalk Closed, Use Other Side", "R09-10", celPath, "R09-10", "24"" x 12""", "24"" x 12""", 24, 12, postType, postPath, defSpacing)
    Call AddSign("R09-11", "Sidewalk Closed Ahead, Cross Here", "R09-11", celPath, "R09-11", "24"" x 18""", "24"" x 18""", 24, 18, postType, postPath, defSpacing)
    Call AddSign("R09-11a", "Sidewalk Closed, Cross Here", "R09-11a", celPath, "R09-11a", "24"" x 12""", "24"" x 12""", 24, 12, postType, postPath, defSpacing)
    Call AddSign("R09-12", "Bike Lane Closed", "R09-12", celPath, "R09-12", "24"" x 12""", "24"" x 12""", 24, 12, postType, postPath, defSpacing)
    Call AddSign("R10-06", "Stop Here on Red", "R10-06", celPath, "R10-06", "24"" x 36""", "24"" x 36""", 24, 36, postType, postPath, defSpacing)
    Call AddSign("R11-02", "Road Closed", "R11-02", celPath, "R11-02", "48"" x 30""", "48"" x 30""", 48, 30, postType, postPath, defSpacing)
    Call AddSign("R11-03", "Road Closed - Local Traffic Only", "R11-03", celPath, "R11-03", "60"" x 30""", "60"" x 30""", 60, 30, postType, postPath, defSpacing)
    Call AddSign("R12-01,02", "Weight Limit", "R12-01,02", celPath, "R12-01,02", "24"" x 30""", "36"" x 48""", 24, 30, postType, postPath, defSpacing)
    Call AddSign("R12-05", "Weight Limit", "R12-05", celPath, "R12-05", "24"" x 36""", "36"" x 48""", 24, 36, postType, postPath, defSpacing)
    Call AddSign("R22-02", "Turn Off 2-Way Radio and Cell Phone", "R22-02", celPath, "R22-02", "42"" x 36""", "42"" x 36""", 42, 36, postType, postPath, defSpacing)
    Call AddSign("G20-05aP", "Work Zone (plaque)", "G20-05aP", celPath, "G20-05aP", "24"" x 18""", "30"" x 24""", 24, 18, postType, postPath, defSpacing)
    Call AddSign("W01-01", "Turn and Curve Signs", "W01-01", celPath, "W01-01", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W01-04b", "Reverse Curve (2 or more lanes)", "W01-04b", celPath, "W01-04b", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W01-06", "Large Arrow (1-direction)", "W01-06", celPath, "W01-06", "48"" x 24""", "60"" x 30""", 48, 24, postType, postPath, defSpacing)
    Call AddSign("W01-08", "Chevron Alignment", "W01-08", celPath, "W01-08", "18"" x 24""", "30"" x 36""", 18, 24, postType, postPath, defSpacing)
    Call AddSign("W03-01", "Stop Ahead", "W03-01", celPath, "W03-01", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W03-02", "Yield Ahead", "W03-02", celPath, "W03-02", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W03-03", "Signal Ahead", "W03-03", celPath, "W03-03", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W03-04", "Be Prepared to Stop", "W03-04", celPath, "W03-04", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W03-05", "Reduced Speed Limit Ahead", "W03-05", celPath, "W03-05", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W03-05a", "XX MPH Speed Zone Ahead", "W03-05a", celPath, "W03-05a", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W04-01", "Merging Traffic", "W04-01", celPath, "W04-01", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W04-02", "Lane Ends", "W04-02", celPath, "W04-02", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W04-03", "Added Lane", "W04-03", celPath, "W04-03", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W04-05aP", "No Merge Area (plaque)", "W04-05aP", celPath, "W04-05aP", "18"" x 24""", "24"" x 30""", 18, 24, postType, postPath, defSpacing)
    Call AddSign("W05-01", "Road Narrows", "W05-01", celPath, "W05-01", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W05-02", "Narrow Bridge", "W05-02", celPath, "W05-02", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W05-03", "One Lane Bridge", "W05-03", celPath, "W05-03", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W05-04", "Ramp Narrows", "W05-04", celPath, "W05-04", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W06-01", "Divided Highway", "W06-01", celPath, "W06-01", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W06-02", "Divided Highway Ends", "W06-02", celPath, "W06-02", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W06-03", "Two-Way Traffic", "W06-03", celPath, "W06-03", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W06-04", "Narrow Two-Way Traffic", "W06-04", celPath, "W06-04", "12"" x 18""", "12"" x 18""", 12, 18, postType, postPath, defSpacing)
    Call AddSign("W07-01", "Hill", "W07-01", celPath, "W07-01", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W07-03aP", "Next XX Miles (plaque)", "W07-03aP", celPath, "W07-03aP", "24"" x 18""", "36"" x 30""", 24, 18, postType, postPath, defSpacing)
    Call AddSign("W08-01", "Bump", "W08-01", celPath, "W08-01", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-02", "Dip", "W08-02", celPath, "W08-02", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-03", "Pavement Ends", "W08-03", celPath, "W08-03", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-04", "Soft Shoulder", "W08-04", celPath, "W08-04", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-05", "Slippery When Wet", "W08-05", celPath, "W08-05", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-06", "Truck Crossing", "W08-06", celPath, "W08-06", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-07", "Loose Gravel", "W08-07", celPath, "W08-07", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-08", "Rough Road", "W08-08", celPath, "W08-08", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-09", "Low Shoulder", "W08-09", celPath, "W08-09", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-11", "Uneven Lanes", "W08-11", celPath, "W08-11", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-12", "No Center Line", "W08-12", celPath, "W08-12", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-14", "Fallen Rocks", "W08-14", celPath, "W08-14", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-15", "Grooved Pavement", "W08-15", celPath, "W08-15", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-15aP", "Motorcycle (plaque)", "W08-15aP", celPath, "W08-15aP", "24"" x 18""", "30"" x 24""", 24, 18, postType, postPath, defSpacing)
    Call AddSign("W08-16", "Metal Bridge Deck", "W08-16", celPath, "W08-16", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-17", "Shoulder Drop Off (symbol)", "W08-17", celPath, "W08-17", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-17P", "Shoulder Drop-Off (plaque)", "W08-17P", celPath, "W08-17P", "24"" x 18""", "30"" x 24""", 24, 18, postType, postPath, defSpacing)
    Call AddSign("W08-18", "Road May Flood", "W08-18", celPath, "W08-18", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-23", "No Shoulder", "W08-23", celPath, "W08-23", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-24", "Steel Plate Ahead", "W08-24", celPath, "W08-24", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W08-25", "Shoulder Ends", "W08-25", celPath, "W08-25", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W09-01,02", "Lane Ends", "W09-01,02", celPath, "W09-01,02", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W09-02a", "Merge Here Take Turns", "W09-02a", celPath, "W09-02a", "36"" x 48""", "36"" x 48""", 36, 48, postType, postPath, defSpacing)
    Call AddSign("W09-03", "Interior Lane Shift Ahead", "W09-03", celPath, "W09-03", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W09-05a", "Bicycles Merging", "W09-05a", celPath, "W09-05a", "30"" x 30""", "30"" x 30""", 30, 30, postType, postPath, defSpacing)
    Call AddSign("W10-01", "Grade Crossing Advance Warning", "W10-01", celPath, "W10-01", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W11-10", "Truck", "W11-10", celPath, "W11-10", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W12-01", "Double Arrow", "W12-01", celPath, "W12-01", "30"" x 30""", "36"" x 36""", 30, 30, postType, postPath, defSpacing)
    Call AddSign("W12-02", "Low Clearance", "W12-02", celPath, "W12-02", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W13-01P", "Advisory Speed (plaque)", "W13-01P", celPath, "W13-01P", "18"" x 18""", "24"" x 24""", 18, 18, postType, postPath, defSpacing)
    Call AddSign("W13-04P", "On Ramp (plaque)", "W13-04P", celPath, "W13-04P", "36"" x 36""", "36"" x 36""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W14-03", "No Passing Zone (pennant)", "W14-03", celPath, "W14-03", "48"" x 48""", "64"" x 64""", 48, 48, postType, postPath, defSpacing)
    Call AddSign("W16-02P", "XX Feet (2-line plaque)", "W16-02P", celPath, "W16-02P", "24"" x 18""", "30"" x 24""", 24, 18, postType, postPath, defSpacing)
    Call AddSign("W20-01", "Road Work (with distance)", "W20-01", celPath, "W20-01", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W20-01b", "Path Work (with distance)", "W20-01b", celPath, "W20-01b", "36"" x 36""", "36"" x 36""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W20-02", "Detour (with distance)", "W20-02", celPath, "W20-02", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W20-02a", "Bike Detour (with distance)", "W20-02a", celPath, "W20-02a", "36"" x 36""", "36"" x 36""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W20-02b", "Bike Diversion (with distance)", "W20-02b", celPath, "W20-02b", "36"" x 36""", "36"" x 36""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W20-03", "Road Closed (with distance)", "W20-03", celPath, "W20-03", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W20-03a", "Path Closed (with distance)", "W20-03a", celPath, "W20-03a", "36"" x 36""", "36"" x 36""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W20-04", "One Lane Road (with distance)", "W20-04", celPath, "W20-04", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W20-05", "Lane(s) Closed (with distance)", "W20-05", celPath, "W20-05", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W20-05b", "Bike Lane Closed (with distance)", "W20-05b", celPath, "W20-05b", "36"" x 36""", "36"" x 36""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W20-07", "Flagger (symbol)", "W20-07", celPath, "W20-07", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W20-07a", "Flagger", "W20-07a", celPath, "W20-07a", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W20-08", "Slow (on Stop/Slow Paddle)", "W20-08", celPath, "W20-08", "18"" x 18""", "18"" x 18""", 18, 18, postType, postPath, defSpacing)
    Call AddSign("W21-01", "Workers", "W21-01", celPath, "W21-01", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W21-02", "Fresh Oil", "W21-02", celPath, "W21-02", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W21-03", "Road Machinery Ahead", "W21-03", celPath, "W21-03", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W21-04", "Slow Moving Vehicle", "W21-04", celPath, "W21-04", "36"" x 18""", "36"" x 18""", 36, 18, postType, postPath, defSpacing)
    Call AddSign("W21-05", "Shoulder Work", "W21-05", celPath, "W21-05", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W21-05a", "Shoulder Closed", "W21-05a", celPath, "W21-05a", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W21-05b", "Shoulder Closed (with distance)", "W21-05b", celPath, "W21-05b", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W21-06", "Survey Crew", "W21-06", celPath, "W21-06", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W21-07", "Utility Work (with distance)", "W21-07", celPath, "W21-07", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W21-08", "Mowing Ahead", "W21-08", celPath, "W21-08", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W22-01", "Blasting Zone Ahead", "W22-01", celPath, "W22-01", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W22-03", "End Blasting Zone", "W22-03", celPath, "W22-03", "42"" x 36""", "42"" x 36""", 42, 36, postType, postPath, defSpacing)
    Call AddSign("W23-01", "Slow Traffic Ahead", "W23-01", celPath, "W23-01", "48"" x 24""", "48"" x 24""", 48, 24, postType, postPath, defSpacing)
    Call AddSign("W23-02", "New Traffic Pattern Ahead", "W23-02", celPath, "W23-02", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W24-01", "Double Reverse Curve (1 lane)", "W24-01", celPath, "W24-01", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W24-01a", "Double Reverse Curve (2 lanes)", "W24-01a", celPath, "W24-01a", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W24-01b", "Double Reverse Curve (3 lanes)", "W24-01b", celPath, "W24-01b", "36"" x 36""", "48"" x 48""", 36, 36, postType, postPath, defSpacing)
    Call AddSign("W24-01cP", "All Lanes (plaque)", "W24-01cP", celPath, "W24-01cP", "24"" x 18""", "30"" x 24""", 24, 18, postType, postPath, defSpacing)
    Call AddSign("G20-01", "Road Work Next XX Miles", "G20-01", celPath, "G20-01", "36"" x 18""", "48"" x 24""", 36, 18, postType, postPath, defSpacing)
    Call AddSign("G20-02", "End Road Work", "G20-02", celPath, "G20-02", "36"" x 18""", "48"" x 24""", 36, 18, postType, postPath, defSpacing)
    Call AddSign("G20-04", "Pilot Car Follow Me", "G20-04", celPath, "G20-04", "36"" x 18""", "36"" x 18""", 36, 18, postType, postPath, defSpacing)
    Call AddSign("E05-02", "Exit Open", "E05-02", celPath, "E05-02", "48"" x 36""", "48"" x 36""", 48, 36, postType, postPath, defSpacing)
    Call AddSign("E05-02a", "Exit Closed", "E05-02a", celPath, "E05-02a", "48"" x 36""", "48"" x 36""", 48, 36, postType, postPath, defSpacing)
    Call AddSign("E05-03", "Exit Only", "E05-03", celPath, "E05-03", "48"" x 36""", "48"" x 36""", 48, 36, postType, postPath, defSpacing)
    Call AddSign("M04-08P", "Detour", "M04-08P", celPath, "M04-08P", "24"" x 12""", "30"" x 15""", 24, 12, postType, postPath, defSpacing)
    Call AddSign("M04-08a", "End Detour", "M04-08a", celPath, "M04-08a", "24"" x 18""", "24"" x 18""", 24, 18, postType, postPath, defSpacing)
    Call AddSign("M04-08bP", "End (plaque)", "M04-08bP", celPath, "M04-08bP", "24"" x 12""", "24"" x 12""", 24, 12, postType, postPath, defSpacing)
    Call AddSign("M04-09", "Detour", "M04-09", celPath, "M04-09", "30"" x 24""", "48"" x 36""", 30, 24, postType, postPath, defSpacing)
    Call AddSign("M04-09a", "Bike/Pedestrian Detour", "M04-09a", celPath, "M04-09a", "30"" x 24""", "30"" x 24""", 30, 24, postType, postPath, defSpacing)
    Call AddSign("M04-09b", "Pedestrian Detour", "M04-09b", celPath, "M04-09b", "30"" x 24""", "30"" x 24""", 30, 24, postType, postPath, defSpacing)
    Call AddSign("M04-09c", "Bike Detour (with arrow)", "M04-09c", celPath, "M04-09c", "30"" x 24""", "30"" x 24""", 30, 24, postType, postPath, defSpacing)
    Call AddSign("M04-10", "Detour", "M04-10", celPath, "M04-10", "48"" x 18""", "48"" x 18""", 48, 18, postType, postPath, defSpacing)
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

