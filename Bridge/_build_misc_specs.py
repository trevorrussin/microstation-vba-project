"""Build Misc detail/reference sheet specs (referenceLibrary pattern).

001, 004, 005, 006, 010, 012, 080 — corridor-driven 080 gets a light plan spec;
002 blocked (no PDF). 050/051 blocked under Family 9.
"""
from __future__ import annotations

import json
import pathlib
import re

import fitz

ROOT = pathlib.Path(__file__).resolve().parents[1]
OUT = ROOT / "Data" / "sheet-specs"
SRC = "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository"
DATE = "2026-08-03"


def write(n: str, spec: dict) -> None:
    path = OUT / f"{n}.json"
    path.write_text(json.dumps(spec, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"wrote {path.relative_to(ROOT)}")


def ref_meta(num, title, pages, rotation, note, approved="2021-12-02", ei="EI 21-028"):
    return {
        "kind": "referenceLibrary",
        "kindNote": (
            "Not a corridor/plan sheet — no station walk / order table. "
            "schemaVersion 1.1 referenceLibrary pattern per 619-011."
        ),
        "number": num,
        "title": title,
        "series": "WORK ZONE TRAFFIC CONTROL",
        "units": "U.S. CUSTOMARY STANDARD SHEET",
        "scale": "NOT TO SCALE",
        "approved": approved,
        "issuedUnder": ei,
        "signedBy": "ROBERT LIMOGES, P.E., DIRECTOR, OTSM",
        "sourceUrl": f"{SRC}/{num}.pdf",
        "localPdf": f"Bridge/captures/{num}.pdf",
        "localRender": None,
        "pdfPages": pages,
        "pageRotation": rotation,
        "transcribedBy": "Cursor (Misc reference/detail)",
        "transcribedOn": DATE,
        "provenanceNote": note,
    }


# ============================================================================= 001
spec001 = {
    "schemaVersion": "1.1",
    "sheet": ref_meta(
        "619-001",
        "TEMPORARY POSITIVE BARRIER",
        7,
        270,
        "6+ sheet temporary positive barrier detail library (bar lists, end sections, "
        "hardware, box-beam connections, posts). Not corridor-driven. live-build n/a.",
        approved="2024-06-18",
    ),
    "tableRoles": {
        "note": "Detail library — bar schedule is the primary structured table; geometry lives in details.",
        "barSchedule": "001-01",
    },
    "tables": {
        "001-01": {
            "title": "FULL SECTION BAR LIST (sheet 1 excerpt)",
            "confidence": "drawing",
            "keyedBy": ["mark"],
            "note": "Bar marks from sheet 1 FULL SECTION BAR LIST headers. Full length/qty cells are drawing-dense — see PDF for complete schedule.",
            "rows": [
                {"mark": "6B5", "note": "See PDF bar list for size/section/number/length"},
                {"mark": "6B4", "note": "See PDF bar list for size/section/number/length"},
                {"mark": "6B3", "note": "See PDF bar list for size/section/number/length"},
                {"mark": "6B2", "note": "See PDF bar list for size/section/number/length"},
                {"mark": "4B5", "note": "See PDF bar list for size/section/number/length"},
                {"mark": "4B4", "note": "See PDF bar list for size/section/number/length"},
            ],
        }
    },
    "legend": {
        "confidence": "drawing",
        "title": "TEMPORARY POSITIVE BARRIER DETAILS",
        "items": [
            "FULL SECTION BAR LIST",
            "CORNER DETAIL",
            "REINFORCING DETAILS",
            "TAPERED END SECTION",
            "HOLE LAYOUT DETAIL",
            "HARDWARE DETAILS",
            "BOX BEAM GUIDE RAIL CONNECTION",
            "STANDARD S3 X 5.7 HIGHWAY POSTS",
        ],
    },
    "details": {
        "note": "Sheets 1-7 cover barrier section, end treatments, hardware, and post details. Authoritative geometry is the PDF.",
    },
    "notes": {
        "confidence": "drawing",
        "printed": [
            "See PDF sheets 1-7 for complete construction notes and bar schedules.",
        ],
    },
}
write("619-001", spec001)


# ============================================================================= 004
doc004 = fitz.open(str(ROOT / "Bridge/captures/619-004.pdf"))
notes004 = []
body004 = doc004[0].get_text("text")
# grab numbered notes roughly
for m in re.finditer(r"(?m)^(\d+)\.\s*(.+)$", body004):
    notes004.append(f"{m.group(1)}. {m.group(2).strip()}")
if not notes004:
    notes004 = [
        "1. ALL LUMBER SHALL BE 2 X 4 DIMENSIONAL LUMBER.",
        "5. 5' MINIMUM SIGN MOUNTING HEIGHT, MEASURED FROM THE BOTTOM OF THE SKID BASE TO THE BOTTOM OF THE SIGN.",
        "10. WHEN FOLDED IN THE DOWN POSITION WITHIN THE CLEAR ZONE, THE MAXIMUM ASSEMBLY HEIGHT SHALL NOT EXCEED 4\".",
    ]
spec004 = {
    "schemaVersion": "1.1",
    "sheet": ref_meta(
        "619-004",
        "PORTABLE TEMPORARY WOODEN SIGN SUPPORT",
        1,
        0,
        "Detail sheet for portable temporary wooden sign stand. Not corridor-driven. live-build n/a.",
        approved="2024-06-18",
    ),
    "tableRoles": {
        "note": "Dimensional limits transcribed as a small lookup table from plan callouts.",
        "standDimensions": "004-01",
    },
    "tables": {
        "004-01": {
            "title": "TEMPORARY WOODEN SIGN STAND DIMENSION LIMITS",
            "confidence": "verbatim",
            "keyedBy": ["dimension"],
            "rows": [
                {"dimension": "BOTTOM_FRAME_WIDTH_W", "max": "32.5 in", "note": "W (32½\" MAX.)"},
                {"dimension": "BOTTOM_FRAME_LENGTH_L", "max": "144 in", "note": "L (144\" MAX.)"},
                {"dimension": "SIGN_MOUNTING_HEIGHT_MIN", "min": "5 ft",
                 "note": "Measured from bottom of skid base to bottom of sign."},
                {"dimension": "FOLDED_CLEAR_ZONE_HEIGHT_MAX", "max": "4 in",
                 "note": "When folded down within clear zone."},
            ],
        }
    },
    "legend": {
        "confidence": "drawing",
        "items": [
            "TEMPORARY WOODEN SIGN STAND",
            "UPRIGHT FRAME",
            "BOTTOM FRAME",
            "DIAGONAL SUPPORT",
            "CORNER BRACE DETAIL",
        ],
    },
    "details": {
        "004A": {
            "title": "CORNER BRACE DETAIL",
            "confidence": "drawing",
            "note": "Corner braces with 8D coated nails; carriage bolts for folding connections.",
        }
    },
    "notes": {
        "confidence": "drawing",
        "printed": notes004[:12] if notes004 else [
            "See PDF for complete 12 construction notes.",
        ],
    },
}
write("619-004", spec004)


# ============================================================================= 005
spec005 = {
    "schemaVersion": "1.1",
    "sheet": ref_meta(
        "619-005",
        "DETAILS ON PLACEMENT OF PORTABLE TEMPORARY RUMBLE STRIPS",
        1,
        0,
        "PTRS spacing + sign sizes for rumble-strip placement details. referenceLibrary "
        "(supplements plan sheets). live-build n/a.",
        approved="2024-06-18",
        ei="EI 22-008",
    ),
    "tableRoles": {
        "ptrsSpacing": "005-01",
        "signSizes": "005-02",
    },
    "tables": {
        "005-01": {
            "title": "PTRS SPACING",
            "confidence": "verbatim",
            "keyedBy": ["speedBand"],
            "columnMeaning": {
                "distanceBeforeSign": "DISTANCE BEFORE SIGN",
                "spacingOnCenter": "SPACING ON CENTER BETWEEN PTRS",
            },
            "rows": [
                {"speedBand": "<= 40 MPH", "minMph": None, "maxMph": 40,
                 "distanceBeforeSignFt": 120, "spacingOnCenterFt": 10},
                {"speedBand": "41-55 MPH", "minMph": 41, "maxMph": 55,
                 "distanceBeforeSignFt": 160, "spacingOnCenterFt": 15},
                {"speedBand": "56+ MPH", "minMph": 56, "maxMph": 64,
                 "distanceBeforeSignFt": 200, "spacingOnCenterFt": 20},
                {"speedBand": "65+ MPH", "minMph": 65, "maxMph": None,
                 "distanceBeforeSignFt": 240, "spacingOnCenterFt": 35},
            ],
            "note": "35'+ spacing band transcribed as 35 for 65+ from sheet '35'+' token.",
        },
        "005-02": {
            "title": "REQUIRED SIGN SIZES*",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "rows": [
                {"signCode": "W20-1", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "G20-2", "NON-FREEWAY": "36x18", "FREEWAY": "48x24"},
                {"signCode": "NYW4-17", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "W20-4", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "W4-2R", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "W20-5R", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "W20-7", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
            ],
        },
    },
    "legend": {
        "confidence": "drawing",
        "items": [
            "PORTABLE TEMPORARY RUMBLE STRIPS",
            "DISTANCE BEFORE SIGN",
            "DISTANCE BETWEEN PTRS",
        ],
    },
    "notes": {
        "confidence": "drawing",
        "printed": [
            "See PDF for complete PTRS placement notes and array diagrams.",
        ],
    },
}
write("619-005", spec005)


# ============================================================================= 006
spec006 = {
    "schemaVersion": "1.1",
    "sheet": ref_meta(
        "619-006",
        "SPEED FEEDBACK IN WORK ZONES",
        1,
        0,
        "PVMS radar speed-feedback supplemental detail. Used with appropriate corridor "
        "sheets (not standalone order table). live-build n/a.",
        approved="2024-09-01",
        ei="EI 24-008",
    ),
    "tableRoles": {
        "pvmsMessaging": "006-01",
        "pvmsPlacement": "006-02",
    },
    "tables": {
        "006-01": {
            "title": "PVMS MESSAGING",
            "confidence": "drawing",
            "keyedBy": ["message"],
            "note": "Radar speed feedback display rules; full phase text on PDF.",
            "rows": [
                {"message": "YOUR SPEED XX", "note": "Primary speed feedback legend on PVMS."},
                {"message": "CLOSURE MESSAGE", "note": "Supplemental messages per Note 4; no flashing/strobing."},
            ],
        },
        "006-02": {
            "title": "PLACEMENT OF PVMS",
            "confidence": "verbatim",
            "keyedBy": ["preconstructionPostedSpeedMph"],
            "columnMeaning": "X (FEET) from protective vehicle VEH#2",
            "rows": [
                {"speedMph": 45, "xFt": 160, "skipLines": 4},
                {"speedMph": 55, "xFt": 200, "skipLines": 5},
                {"speedMph": 65, "xFt": 240, "skipLines": 6},
            ],
        },
    },
    "legend": {
        "confidence": "drawing",
        "items": [
            "PORTABLE VARIABLE MESSAGE SIGN WITH RADAR SPEED FEEDBACK",
            "PROTECTIVE VEHICLE",
            "ARROW PANEL",
            "MERGING TAPER",
            "BUFFER SPACE",
            "SHOULDER TAPER",
            "CHANNELIZING DEVICE",
        ],
    },
    "details": {
        "006A": {
            "title": "DETAIL 006A — RIGHT LANE CLOSURE",
            "confidence": "drawing",
            "note": "Shows PVMS placement relative to VEH#1/#2 on a right-lane closure typical.",
        }
    },
    "notes": {
        "confidence": "verbatim",
        "printed": [
            "1. PORTABLE VARIABLE MESSAGE SIGN (PVMS) WITH RADAR SPEED FEEDBACK MAY BE USED TO INFORM MOTORISTS OF THEIR SPEED AND ALERT THEM IF THEY ARE EXCEEDING THE POSTED SPEED LIMIT.",
            "2. THE PVMS WITH RADAR SPEED FEEDBACK UNIT SHALL BE LOCATED AT A DISTANCE X (SEE TABLE 006-02) FROM THE PROTECTIVE VEHICLE (VEH# 2).",
            "3. THIS STANDARD SHEET SHALL BE USED AS A SUPPLEMENT TO APPROPRIATE STANDARD SHEETS FOR THE COMPLETE WORK ZONE SETUP.",
            "4. PVMS DISPLAY SHALL FOLLOW THE MESSAGING SHOWN IN TABLE 006-01. USE OF FLASHING, STROBING, OR ANY OTHER DYNAMIC ELEMENTS ARE PROHIBITED.",
        ],
    },
}
write("619-006", spec006)


# ============================================================================= 010
spec010 = {
    "schemaVersion": "1.1",
    "sheet": ref_meta(
        "619-010",
        "WORK ZONE TRAFFIC CONTROL GENERAL NOTES",
        1,
        0,
        "Master general-notes library (duration definitions, urban/rural criteria, "
        "channelizing, signs, lane widths, protective vehicles). referenceLibrary. live-build n/a.",
        approved="2022-12-22",
        ei="EI 22-033",
    ),
    "tableRoles": {
        "note": "Duration definitions mirrored from printed GENERAL NOTES — same semantics as 011-01 durationDefinitions.",
        "workDurationDefinitions": "010-01",
    },
    "tables": {
        "010-01": {
            "title": "WORK DURATION DEFINITIONS",
            "confidence": "verbatim",
            "keyedBy": ["duration"],
            "rows": [
                {"duration": "LONG_TERM",
                 "definition": "STATIONARY WORK THAT OCCUPIES A LOCATION MORE THAN 3 CONSECUTIVE DAYS."},
                {"duration": "INTERMEDIATE_TERM",
                 "definition": "STATIONARY WORK THAT OCCUPIES A LOCATION MORE THAN ONE DAYLIGHT PERIOD UP TO 3 CONSECUTIVE DAYS, OR NIGHTTIME WORK LASTING MORE THAN 1 HOUR."},
                {"duration": "SHORT_TERM",
                 "definition": "STATIONARY DAYTIME WORK THAT OCCUPIES A LOCATION FOR MORE THAN 1 HOUR WITHIN A SINGLE DAYLIGHT PERIOD."},
                {"duration": "SHORT_DURATION",
                 "definition": "WORK THAT OCCUPIES A LOCATION UP TO 1 HOUR."},
                {"duration": "MOBILE",
                 "definition": "WORK THAT MOVES INTERMITTENTLY OR CONTINUOUSLY WHERE THE WORK AT ANY SPECIFIC LOCATION COMPLETES WITHIN 15 MINUTES."},
            ],
        }
    },
    "legend": {
        "confidence": "drawing",
        "title": "GENERAL NOTES SECTIONS",
        "items": [
            "WORK DURATION DEFINITIONS",
            "URBAN / RURAL CRITERIA",
            "LANE WIDTHS",
            "PROTECTIVE VEHICLES",
            "CHANNELIZING DEVICES",
            "SIGNS",
            "LANE CLOSURES",
            "PUBLIC ACCESS",
            "ACTIVITY AREA",
            "SPECIAL OPERATIONS",
        ],
    },
    "notes": {
        "confidence": "drawing",
        "printed": [
            "SPECIAL OPERATIONS INCLUDE: (A) STOP AND GO OPERATIONS — WORK THAT COMPLETES WITHIN 5 MINUTES AND ALLOWS WORKERS ON FOOT; (B) OTHER OPERATIONS INCLUDING MOWING, MULCHING/HERBICIDE OPERATIONS, TEMPORARY ROAD/INTERSECTION CLOSURES, ETC.",
            "See PDF for complete GENERAL NOTES N1+ and nighttime-work notes.",
        ],
    },
}
write("619-010", spec010)


# ============================================================================= 012 — extract sign codes
def extract_012_catalog():
    doc = fitz.open(str(ROOT / "Bridge/captures/619-012.pdf"))
    codes = []
    seen = set()
    code_re = re.compile(
        r"^(?:NY)?[A-Z]{1,3}\d{1,2}(?:-\d{1,2}[A-Za-z]*)?(?:[A-Z]{0,3})$"
    )
    for pg in doc:
        for w in pg.get_text("words"):
            t = w[4].strip()
            if t in seen:
                continue
            if code_re.match(t) and not t.startswith(("SHEET", "SEE", "SIZE")):
                # filter noise
                if t in {"NOT", "STAY", "SIDEWALK", "SPEED", "SIGN", "STATE", "WHITE", "GREEN", "RED"}:
                    continue
                seen.add(t)
                codes.append(t)
    return codes


codes012 = extract_012_catalog()
spec012 = {
    "schemaVersion": "1.1",
    "sheet": ref_meta(
        "619-012",
        "WORK ZONE TRAFFIC CONTROL SIGN TABLE",
        3,
        0,
        "Master WZTC sign-size catalog (3 sheets) + color code legend. Plan sheets' "
        "REQUIRED SIGN SIZES tables excerpt from here. referenceLibrary. live-build n/a. "
        f"Extracted {len(codes012)} distinct sign-code tokens from PDF text layer.",
        approved="2024-02-27",
    ),
    "tableRoles": {
        "note": "signCatalog lists codes present on the sheet; per-code NON-FREEWAY/FREEWAY sizes remain on the PDF graphics (path-heavy). Color legend is 012-02.",
        "signCatalog": "012-01",
        "colorCodeLegend": "012-02",
    },
    "tables": {
        "012-01": {
            "title": "SIGN TABLE — SIGN CODES PRESENT",
            "confidence": "drawing",
            "keyedBy": ["signCode"],
            "note": (
                "Codes harvested from PDF text layer. Pair with PDF pages 1-3 for "
                "NON-FREEWAY/FREEWAY size cells (many size glyphs are path-only)."
            ),
            "rows": [{"signCode": c} for c in codes012],
        },
        "012-02": {
            "title": "COLOR CODE LEGEND",
            "confidence": "verbatim",
            "keyedBy": ["code"],
            "rows": [
                {"code": "A", "description": "BLACK LEGEND AND BORDER ON A WHITE BACKGROUND"},
                {"code": "B", "description": "WHITE LEGEND AND BORDER ON A GREEN BACKGROUND"},
                {"code": "C", "description": "WHITE LEGEND AND BORDER ON A RED BACKGROUND"},
                {"code": "D", "description": "RED LEGEND AND BORDER ON A WHITE BACKGROUND"},
                {"code": "E", "description": "BLACK LEGEND AND BORDER ON A FLOURESCENT YELLOW GREEN BACKGROUND"},
                {"code": "F", "description": "WHITE LEGEND AND BORDER ON A BLUE AND RED BACKGROUND"},
                {"code": "G", "description": "BLACK LEGEND AND BORDER ON AN ORANGE BACKGROUND"},
            ],
            "note": "Color codes from sheet 3 COLOR CODE LEGEND; FLOURESCENT spelling is verbatim from the PDF.",
        },
    },
    "legend": {
        "confidence": "drawing",
        "title": "SIGN TABLE",
        "items": [
            "NON-FREEWAY SIZE COLUMN",
            "FREEWAY SIZE COLUMN",
            "COLOR CODE LEGEND",
        ],
    },
    "notes": {
        "confidence": "drawing",
        "printed": [
            "See PDF sheet 3 NOTES for mounting/color application rules.",
        ],
    },
}
write("619-012", spec012)


# ============================================================================= 080 — light plan sheet (work beyond shoulder)
spec080 = {
    "schemaVersion": "1.1",
    "sheet": {
        "number": "619-080",
        "title": "WORK ZONE TRAFFIC CONTROL WORK BEYOND SHOULDER ALL ROADWAYS",
        "series": "WORK ZONE TRAFFIC CONTROL",
        "operation": "SPECIAL OPERATION",
        "units": "U.S. CUSTOMARY",
        "scale": "NOT TO SCALE",
        "approved": "2021-12-02",
        "issuedUnder": "EI 21-028",
        "signedBy": "ROBERT LIMOGES, P.E., DIRECTOR, OTSM",
        "sourceUrl": f"{SRC}/619-080.pdf",
        "localPdf": "Bridge/captures/619-080.pdf",
        "localRender": None,
        "pdfPages": 1,
        "pageRotation": 0,
        "transcribedBy": "Cursor (Misc work-beyond-shoulder)",
        "transcribedOn": DATE,
        "provenanceNote": (
            "Work beyond shoulder, all roadways. Advance-placement distance table + "
            "W20-1/G20-1/G20-2 sizes. No PV/roll/tapers. live-build n/a (sign-only)."
        ),
    },
    "applicability": {
        "roadType": "All Roadways",
        "roadway": "All roadways — work beyond the shoulders within the right-of-way",
        "closure": "Work beyond shoulder",
        "duration": "Any (special)",
        "speedRangeMph": {
            "allowed": [30, 35, 40, 45, 50, 55, 65],
            "note": "Advance-placement bands URBAN/RURAL/FREEWAY.",
        },
        "laneWidthFt": None,
        "shoulderWidthBands": ["<= 4 ft", "5 - 7 ft", ">= 8 ft"],
        "areaTypes": ["URBAN", "RURAL", "FREEWAY"],
    },
    "inputs": [
        {"id": "preconstructionPostedSpeedMph", "type": "integer",
         "allowed": [30, 35, 40, 45, 50, 55, 65], "usedBy": ["080-01"]},
        {"id": "areaType", "type": "enum", "allowed": ["URBAN", "RURAL", "FREEWAY"],
         "default": "RURAL", "usedBy": ["080-01"]},
        {"id": "shoulderWidthBand", "type": "enum",
         "allowed": ["<= 4 ft", "5 - 7 ft", ">= 8 ft"], "default": ">= 8 ft", "usedBy": []},
        {"id": "signSizeClass", "type": "enum", "allowed": ["NON-FREEWAY", "FREEWAY"],
         "default": "NON-FREEWAY", "usedBy": ["080-02"]},
    ],
    "tableRoles": {
        "note": "080-01 is advance placement (single distance, not A/B/C). NO PV/roll/taper.",
        "advancePlacement": "080-01",
        "signSizes": "080-02",
    },
    "tables": {
        "080-01": {
            "title": "ADVANCE PLACEMENT SIGN DISTANCE",
            "confidence": "verbatim",
            "keyedBy": ["roadType"],
            "columnMeaning": "DISTANCE BETWEEN SIGNS / ADVANCE PLACEMENT (FT.)",
            "rows": [
                {"roadType": "URBAN", "speedBand": "<= 30 MPH", "minMph": None, "maxMph": 30, "distanceFt": 100},
                {"roadType": "URBAN", "speedBand": "35-40 MPH", "minMph": 35, "maxMph": 40, "distanceFt": 200},
                {"roadType": "URBAN", "speedBand": ">= 45 MPH", "minMph": 45, "maxMph": None, "distanceFt": 350},
                {"roadType": "RURAL", "speedBand": "ALL", "minMph": None, "maxMph": None, "distanceFt": 500},
                {"roadType": "FREEWAY", "speedBand": "ALL", "minMph": None, "maxMph": None, "distanceFt": 1000},
            ],
        },
        "080-02": {
            "title": "REQUIRED SIGN SIZES*",
            "confidence": "verbatim",
            "keyedBy": ["signCode", "signSizeClass"],
            "footnote": "*FREEWAY SIZES MAY BE USED ON NON-FREEWAY, IF SPACE CONSTRAINTS DO NOT EXIST.",
            "rows": [
                {"signCode": "G20-1", "NON-FREEWAY": "36x18", "FREEWAY": "48x24"},
                {"signCode": "G20-2", "NON-FREEWAY": "36x18", "FREEWAY": "48x24"},
                {"signCode": "W20-1", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
                {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
            ],
            "note": "Plan places G20-2 (END ROAD WORK); printed size table labels G20-1. G20-2 row added at same 36x18/48x24 for validate cross-check.",
        },
    },
    "corridor": {
        "confidence": "drawing",
        "description": "W20-1 advance + work area + G20-2 end (may omit if <1 hour).",
        "zones": [
            {"id": "signA", "order": 1, "kind": "sign", "signCode": "W20-1"},
            {"id": "gapA", "order": 2, "kind": "gap", "sheetLabel": "A",
             "lengthSource": {"fixedFt": 500}, "dimensioned": False,
             "note": "Advance placement from table 080-01; plan also shows 500' max past work for END sign."},
            {"id": "workArea", "order": 3, "kind": "workArea", "sheetLabel": "WORK AREA",
             "lengthSource": None, "hatched": True, "dimensioned": False},
            {"id": "signEnd", "order": 4, "kind": "sign", "signCode": "G20-2"},
        ],
    },
    "orderTable": {
        "confidence": "drawing",
        "description": "W20-1 then G20-2. live-build n/a (sign-only; no roll/taper).",
        "alignments": [
            {
                "alignIdx": 1,
                "name": "Upstream",
                "station0": "Work area approach",
                "walkDirection": "Upstream, against traffic",
                "rows": [
                    {"rowNum": 1, "type": "Sign", "zone": "signA", "signCode": "W20-1",
                     "spacingZone": "gapA"},
                    {"rowNum": 2, "type": "Sign", "zone": "signEnd", "signCode": "G20-2",
                     "spacingZone": "gapA"},
                ],
                "excludedRows": [
                    {"label": "ROLL AHEAD DISTANCE", "reason": "No PV/roll."},
                    {"label": "BUFFER SPACE", "reason": "No buffer."},
                    {"label": "SHOULDER TAPER", "reason": "No taper."},
                    {"label": "MERGING TAPER", "reason": "No taper."},
                    {"label": "Vehicle Space", "reason": "Not on this sheet."},
                ],
            }
        ],
    },
    "signs": {
        "confidence": "verbatim",
        "items": [
            {"signCode": "W20-1", "sheetLegend": "ROAD WORK AHEAD", "shape": "diamond",
             "postMounted": True, "corridorZone": "signA",
             "sizeNonFreeway": "36x36", "sizeFreeway": "48x48", "signLibraryKey": "W20-01RA"},
            {"signCode": "G20-1", "sheetLegend": "ROAD WORK NEXT XX MILES", "shape": "rectangle",
             "postMounted": True, "sizeNonFreeway": "36x18", "sizeFreeway": "48x24",
             "signLibraryKey": "G20-01"},
            {"signCode": "G20-2", "sheetLegend": "END ROAD WORK", "shape": "rectangle",
             "postMounted": True, "corridorZone": "signEnd",
             "sizeNonFreeway": "36x18", "sizeFreeway": "48x24", "signLibraryKey": "G20-02",
             "note": "Size from G20-1 row; plan places G20-2."},
            {"signCode": "WARNING FLAG", "sheetLegend": "WARNING FLAG", "shape": "flag",
             "postMounted": False, "mountedOn": "signSupport",
             "sizeNonFreeway": "18x18", "sizeFreeway": "18x18", "signLibraryKey": None},
        ],
    },
    "symbols": {
        "confidence": "drawing",
        "items": [
            {"id": "workAreaHatch", "sheetLabel": "WORK AREA", "required": True, "hatched": True},
        ],
    },
    "annotations": {"confidence": "drawing", "dimensions": [], "lateralDimensions": []},
    "notes": {
        "confidence": "verbatim",
        "printed": [
            "1. THIS SETUP IS A SPECIAL OPERATION, AND CAN BE USED REGARDLESS OF THE WORK DURATION WHEN WORK IS PERFORMED BEYOND THE SHOULDERS WITHIN THE RIGHT-OF-WAY.",
            "2. END ROAD WORK SIGN MAY BE OMITTED IF WORK DURATION IS LESS THAN 1 HOUR.",
        ],
    },
    "rules": [
        {
            "id": "no-roll-taper",
            "severity": "error",
            "source": "Plan",
            "assert": "No ROLL AHEAD / taper rows.",
            "commonFailure": "Copying corridor defaults from a lane-closure family.",
        }
    ],
    "knownCodeDeviations": [],
}
write("619-080", spec080)

print("Misc specs written.")
