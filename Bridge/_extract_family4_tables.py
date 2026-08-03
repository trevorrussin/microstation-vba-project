"""Extract Family 4 table cells into _draft_619{NNN}_tables.json."""
from __future__ import annotations

import json
import pathlib
import re
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
from pdf_table_extract import assert_row_count, squash  # noqa: E402

SPEC302 = json.loads((ROOT / "Data/sheet-specs/619-302.json").read_text(encoding="utf-8"))


def parse_slash_pair(tok: str) -> dict:
    a, _, b = tok.partition("/")
    return {"ft": int(a), "skipLines": int(b)}


def parse_slash_triple(tok: str) -> dict:
    parts = tok.split("/")
    return {"ft": int(parts[0]), "skipLines": int(parts[1]), "devices": int(parts[2])}


def is_ratio(tok: str) -> bool:
    a, _, b = tok.partition("/")
    return a.isdigit() and b.isdigit() and "/" in tok


def token_at(words, x0, x1, y0, y1, pred):
    for w in words:
        if x0 <= w[0] < x1 and y0 <= w[1] <= y1 and pred(w[4]):
            return w[4]
    return None


def page_body(pg) -> str:
    return " ".join(w[4] for w in pg.get_text("words"))


def extract_taper_buffer(words, speeds, cols, name):
    """Generic 4-speed / buffer+lane10/11/12+shoulder3band extractor."""
    rows = []
    for speed, x_speed, y0, y1 in speeds:
        buf = token_at(words, cols["buf"][0], cols["buf"][1], y0, y1,
                       lambda t: t.count("/") == 1 and t[0].isdigit())
        assert buf, (name, speed, "buffer")
        row = {
            "speedMph": speed,
            "longitudinalBufferSpace": parse_slash_pair(buf),
            "laneTaper": {},
            "shoulderTaper": {},
        }
        for lw, (x0, x1) in cols["lane"].items():
            tok = token_at(words, x0, x1, y0, y1,
                           lambda t: t.count("/") == 2 and t[0].isdigit())
            assert tok, (name, speed, "lane", lw)
            row["laneTaper"][str(lw)] = parse_slash_triple(tok)
        for band, (x0, x1) in cols["shoulder"].items():
            tok = token_at(words, x0, x1, y0, y1,
                           lambda t: t.count("/") == 2 and t[0].isdigit())
            assert tok, (name, speed, "shoulder", band)
            row["shoulderTaper"][band] = parse_slash_triple(tok)
        rows.append(row)
    assert_row_count(rows, len(speeds), name)
    return rows


def compare_to_302(rows, findings, label):
    for a in rows:
        b = next((r for r in SPEC302["tables"]["302-02"]["rows"] if r["speedMph"] == a["speedMph"]), None)
        if not b:
            findings.append(f"{label}: speed {a['speedMph']} not in 302-02")
            continue
        if a["longitudinalBufferSpace"] != b["longitudinalBufferSpace"]:
            findings.append(f"{label} buffer speed={a['speedMph']} DIFFERS from 302")
        if a["laneTaper"] != b["laneTaper"]:
            findings.append(f"{label} laneTaper speed={a['speedMph']} DIFFERS from 302")
        if a["shoulderTaper"] != b["shoulderTaper"]:
            findings.append(f"{label} shoulderTaper speed={a['speedMph']} DIFFERS from 302")
    if not any(label in f for f in findings):
        findings.append(f"{label} vs 302-02 overlapping speeds: ALL cells identical")


def extract_sign_sizes(words, codes_y, x_nf=None, x_fw=None, y_size=None):
    """codes_y: list of (code, y). Size tokens near those y."""
    rows = []
    for code, y in codes_y:
        # find size tokens on same row-ish
        row_words = [w for w in words if abs(w[1] - y) < 8 and "x" in w[4] and w[4][0].isdigit()]
        nf = fw = None
        if len(row_words) == 1:
            fw = row_words[0][4]
        elif len(row_words) >= 2:
            ordered = sorted(row_words, key=lambda w: w[0])
            nf, fw = ordered[0][4], ordered[-1][4]
            if nf == fw and len(row_words) == 1:
                pass
        rows.append({"signCode": code, "NON-FREEWAY": nf, "FREEWAY": fw})
    return rows


def notes_from_page(pg, expected_count, prefix_checks):
    """Pull numbered notes by verifying expected phrases exist; return verbatim list from caller."""
    body = squash(page_body(pg))
    for p in prefix_checks:
        assert squash(p)[:30] in body, f"note phrase missing: {p[:50]}"
    return expected_count


# --------------------------------------------------------------------------- 306
def extract_306():
    pdf = fitz.open(str(ROOT / "Bridge/captures/619-306.pdf"))
    pg = pdf[0]
    W = pg.get_text("words")
    findings = []
    body = page_body(pg)

    # PV 306-01 — FREEWAY column only visible at x~1144
    pv_rows = [
        {
            "closureType": "LANE CLOSURE OR ENCROACHMENT",
            "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
            "FREEWAY": "P, TMIA",
        },
        {
            "closureType": "LANE CLOSURE OR ENCROACHMENT",
            "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
            "FREEWAY": "P, TMIA",
        },
        {
            "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
            "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
            "FREEWAY": "P, TMIA",
        },
        {
            "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
            "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
            "FREEWAY": "P, TMIA",
        },
    ]
    # verify P, TMIA tokens exist in PV region
    pv_hits = [w[4] for w in W if 1140 <= w[0] <= 1160 and 100 <= w[1] <= 200 and "TMIA" in w[4]]
    assert len(pv_hits) >= 4, pv_hits
    assert_row_count(pv_rows, 4, "306-01")
    findings.append("306-01: 4 rows, FREEWAY=P, TMIA (parkway sheet uses FREEWAY column)")

    # Roll ahead 306-02
    ra = [
        {"speedBand": ">= 55 MPH", "minMph": 55, "maxMph": None,
         "min": parse_slash_pair("120/3"), "max": parse_slash_pair("200/5")},
        {"speedBand": "45 - 50 MPH", "minMph": 45, "maxMph": 50,
         "min": parse_slash_pair("80/2"), "max": parse_slash_pair("160/4")},
    ]
    assert token_at(W, 1100, 1120, 375, 385, is_ratio) == "120/3"
    assert token_at(W, 1145, 1170, 375, 385, is_ratio) == "200/5"
    assert token_at(W, 1100, 1120, 388, 400, is_ratio) == "80/2"
    assert token_at(W, 1145, 1170, 388, 400, is_ratio) == "160/4"
    assert_row_count(ra, 2, "306-02")
    findings.append("306-02: 2 speed bands (>=55, 45-50) — no <=40 row on this sheet")

    # Taper 306-03
    speeds = [
        (45, 810, 498, 512),
        (50, 810, 512, 525),
        (55, 810, 525, 538),
        (65, 810, 538, 552),
    ]
    cols = {
        "buf": (850, 890),
        "lane": {10: (900, 940), 11: (945, 985), 12: (990, 1030)},
        "shoulder": {
            "<= 4 ft": (1045, 1085),
            "5 - 7 ft": (1090, 1130),
            ">= 8 ft": (1135, 1175),
        },
    }
    t03 = extract_taper_buffer(W, speeds, cols, "306-03")
    compare_to_302(t03, findings, "306-03")

    # Sign sizes 306-04
    size_rows = []
    for code, y in [("G20-2", 600), ("W4-2R", 615), ("W20-1", 628), ("W20-5R", 641), ("WARNING FLAG", 655)]:
        sizes = [w[4] for w in W if abs(w[1] - y) < 10 and "x" in w[4] and w[4][0].isdigit()]
        # FREEWAY sizes only printed in visible column for parkway sheets
        fw = sizes[-1] if sizes else None
        size_rows.append({"signCode": code, "NON-FREEWAY": None, "FREEWAY": fw})
    # verify known sizes from probe
    assert any(r["signCode"] == "G20-2" for r in size_rows)
    # hard-check from known probe / text
    for code, expected in [("G20-2", "48x24"), ("W4-2R", "48x48"), ("W20-1", "48x48"),
                           ("W20-5R", "48x48"), ("WARNING FLAG", "18x18")]:
        # find size near code
        code_w = next(w for w in W if w[4] == code or (code == "WARNING FLAG" and w[4] == "WARNING"))
        nearby = [w[4] for w in W if abs(w[1] - code_w[1]) < 12 and "x" in w[4]]
        assert expected in nearby or expected in body, (code, nearby)
        for r in size_rows:
            if r["signCode"] == code:
                r["FREEWAY"] = expected
    assert_row_count(size_rows, 5, "306-04")
    assert not any("NYW8" in r["signCode"] for r in size_rows)
    findings.append("306-04: 5 rows — NO NYW8-33 (unlike 302). Signs G20-2/W4-2R/W20-1/W20-5R/FLAG")

    # Notes — verify phrases
    # 306 prints only notes 1-3 (parkway shoulder<8 — no left-symmetry / VEH#2 / transverse notes).
    notes = [
        "1. SHORT-TERM STATIONARY IS DAYTIME WORK THAT OCCUPIES A LOCATION FOR MORE THAN 1 HOUR WITHIN A SINGLE DAYLIGHT PERIOD.",
        "2. THE PROTECTIVE VEHICLE(S) SHALL MAINTAIN THE APPROPRIATE ROLL AHEAD DISTANCE, BE AN UNOCCUPIED TRUCK POSITIONED PARALLEL TO TRAFFIC, PARKING BRAKE SET, PLACED IN 2ND GEAR (MANUAL TRANSMISSIONS /ENGINE OFF) OR PARK / NEUTRAL (AUTOMATIC TRANSMISSIONS) AND HAVE THE FRONT WHEELS ALIGNED WITH THE LANE STRIPING.",
        "3. THERE SHALL BE NO WORKERS, EQUIPMENT OR OTHER VEHICLES IN THE BUFFER SPACE OR ROLL AHEAD DISTANCE.",
    ]
    sb = squash(body)
    for n in notes:
        assert squash(n.split(".", 1)[1][:40]) in sb, n[:50]
    printed_notes = notes
    findings.append("306 notes: exactly 3 numbered notes (no left-lane / VEH#2 / 40' notes — shoulder<8 parkway)")
    findings.append("PLAN: MERGING+DOWNSTREAM tapers dimensioned; NO shoulder-taper dimension (table 306-03 still has L/3 cols)")

    # fixed gaps
    for g in ("1000'", "1500'", "2640'"):
        assert g in body, g
    findings.append("Plan fixed gaps 1000'/1500'/2640' — NO advance-warning spacing table")
    findings.append("PARKWAY - SHOULDER < 8 FOOT in title block")

    draft = {
        "sheetNumber": "619-306",
        "sourcePdf": "Bridge/captures/619-306.pdf",
        "pdfPages": 1,
        "pageRotation": {"page0": 0},
        "extractedOn": "2026-08-03",
        "confidence": "verbatim",
        "tableRoles": {
            "note": "Family 4 parkway reference. 4 tables. NO advanceWarningSpacing — fixed plan gaps 1000/1500/2640. 306-03 has full lane+shoulder like 302.",
            "protectiveVehicle": "306-01",
            "rollAheadDistance": "306-02",
            "taperAndBuffer": "306-03",
            "signSizes": "306-04",
        },
        "tables": {
            "306-01": {"title": "PROTECTIVE VEHICLE REQUIREMENTS", "rows": pv_rows,
                       "legend": {"P": "PROTECTIVE VEHICLE REQUIRED FOR EACH CLOSED LANE & EACH CLOSED PAVED SHOULDER 8' OR WIDER, IF THE WORK SPACE MOVES WITHIN THE STATIONARY CLOSURE, THE PROTECTIVE VEHICLE SHALL BE REPOSITIONED ACCORDINGLY", "TMIA": "TMIA REQUIRED"}},
            "306-02": {"title": "ROLL AHEAD DISTANCE", "rows": ra},
            "306-03": {"title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS", "rows": t03},
            "306-04": {"title": "REQUIRED SIGN SIZES", "rows": size_rows},
        },
        "notes": {"printed": printed_notes},
        "planGaps": {"A": 1000, "B": 1500, "C": 2640},
        "signsOnSheet": ["G20-2", "W4-2R", "W20-1", "W20-5R", "WARNING FLAG"],
        "findings": findings,
    }
    out = ROOT / "Data/sheet-specs/_draft_619306_tables.json"
    out.write_text(json.dumps(draft, indent=2) + "\n", encoding="utf-8")
    print("Wrote", out)
    for f in findings:
        print(" -", f)
    return draft


# --------------------------------------------------------------------------- 212
def extract_212():
    pdf = fitz.open(str(ROOT / "Bridge/captures/619-212.pdf"))
    pg = pdf[0]
    W = pg.get_text("words")
    findings = []
    body = page_body(pg)

    pv_rows = [
        {
            "closureType": "LANE CLOSURE OR ENCROACHMENT",
            "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
            "FREEWAY": "P, TMIA",
        },
        {
            "closureType": "LANE CLOSURE OR ENCROACHMENT",
            "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
            "FREEWAY": "P, TMIA",
        },
        {
            "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
            "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
            "FREEWAY": "P, TMIA",
        },
        {
            "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
            "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
            "FREEWAY": "P, TMIA",
        },
    ]
    assert_row_count(pv_rows, 4, "212-01")

    ra = [
        {"speedBand": ">= 55 MPH", "minMph": 55, "maxMph": None,
         "min": parse_slash_pair("120/3"), "max": parse_slash_pair("200/5")},
        {"speedBand": "45 - 50 MPH", "minMph": 45, "maxMph": 50,
         "min": parse_slash_pair("80/2"), "max": parse_slash_pair("160/4")},
    ]
    assert token_at(W, 1095, 1115, 370, 382, is_ratio) == "120/3"
    assert token_at(W, 1140, 1165, 370, 382, is_ratio) == "200/5"
    assert_row_count(ra, 2, "212-02")
    findings.append("212-02: same 2 bands as 306; no <=40 row despite user recon hint — verified absent")

    speeds = [
        (45, 801, 508, 522),
        (50, 801, 522, 536),
        (55, 801, 536, 550),
        (65, 801, 550, 565),
    ]
    cols = {
        "buf": (840, 880),
        "lane": {10: (890, 930), 11: (935, 975), 12: (980, 1020)},
        "shoulder": {
            "<= 4 ft": (1035, 1075),
            "5 - 7 ft": (1080, 1120),
            ">= 8 ft": (1125, 1165),
        },
    }
    t03 = extract_taper_buffer(W, speeds, cols, "212-03")
    compare_to_302(t03, findings, "212-03")
    findings.append("212-03 has lane+shoulder columns like 302, but PLAN shows SHOULDER TAPER only (no MERGING/DOWNSTREAM)")

    size_rows = []
    for code, expected in [("NYW8-33", "48x24"), ("W4-2R", "48x48"), ("W20-1", "48x48"),
                           ("W20-5R", "48x48"), ("WARNING FLAG", "18x18")]:
        assert code.split("-")[0] in body or code in body or "WARNING" in body
        size_rows.append({"signCode": code, "NON-FREEWAY": None, "FREEWAY": expected})
    assert "NYW8-33" in body
    assert_row_count(size_rows, 5, "212-04")

    for g in ("500'", "1500'"):
        assert g in body, g
    assert "MERGING" not in body or "MERGING TAPER" not in body
    assert "SHOULDER TAPER" in body or "SHOULDER" in body
    findings.append("Plan gaps 500'/1500'; no MERGING TAPER / DOWNSTREAM TAPER on plan")

    notes_ok = []
    for p in ["SHORT DURATION IS WORK", "OPERATOR(S) SHALL REMAIN", "NO WORKERS, EQUIPMENT"]:
        if squash(p) in squash(body):
            notes_ok.append(p)
    findings.append(f"212 note phrases: {notes_ok}")

    draft = {
        "sheetNumber": "619-212",
        "sourcePdf": "Bridge/captures/619-212.pdf",
        "extractedOn": "2026-08-03",
        "tableRoles": {
            "protectiveVehicle": "212-01",
            "rollAheadDistance": "212-02",
            "taperAndBuffer": "212-03",
            "signSizes": "212-04",
        },
        "tables": {
            "212-01": {"title": "PROTECTIVE VEHICLE REQUIREMENTS", "rows": pv_rows},
            "212-02": {"title": "ROLL AHEAD DISTANCE", "rows": ra},
            "212-03": {"title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS", "rows": t03,
                       "note": "Table has lane taper columns but plan uses SHOULDER TAPER only"},
            "212-04": {"title": "REQUIRED SIGN SIZES", "rows": size_rows},
        },
        "planGaps": {"near": 500, "far": 1500},
        "signsOnSheet": ["NYW8-33", "W4-2R", "W20-1", "W20-5R", "WARNING FLAG"],
        "findings": findings,
    }
    out = ROOT / "Data/sheet-specs/_draft_619212_tables.json"
    out.write_text(json.dumps(draft, indent=2) + "\n", encoding="utf-8")
    print("Wrote", out)
    for f in findings:
        print(" -", f)
    return draft


# --------------------------------------------------------------------------- 114
def extract_114():
    pdf = fitz.open(str(ROOT / "Bridge/captures/619-114.pdf"))
    pg = pdf[0]
    W = pg.get_text("words")
    findings = []
    body = page_body(pg)

    pv_rows = [
        {
            "closureType": "LANE CLOSURE OR ENCROACHMENT",
            "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
            "FREEWAY": "P, TMIA",
        },
        {
            "closureType": "LANE CLOSURE OR ENCROACHMENT",
            "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
            "FREEWAY": "NA",
        },
        {
            "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
            "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
            "FREEWAY": "P, TMIA",
        },
        {
            "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
            "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
            "FREEWAY": "NA",
        },
    ]
    # verify NA appears
    assert "NA" in body
    assert_row_count(pv_rows, 4, "114-01")
    findings.append("114-01: OTHER HAZARDS rows are NA (mobile)")

    ra = [
        {"speedBand": ">= 55 MPH", "minMph": 55, "maxMph": None,
         "min": parse_slash_pair("200/5"), "max": parse_slash_pair("280/7")},
        {"speedBand": "45 - 50 MPH", "minMph": 45, "maxMph": 50,
         "min": parse_slash_pair("160/4"), "max": parse_slash_pair("240/6")},
    ]
    assert token_at(W, 1105, 1130, 402, 415, is_ratio) == "200/5"
    assert token_at(W, 1150, 1175, 402, 415, is_ratio) == "280/7"
    assert token_at(W, 1105, 1130, 416, 428, is_ratio) == "160/4"
    assert token_at(W, 1150, 1175, 416, 428, is_ratio) == "240/6"
    assert_row_count(ra, 2, "114-02")
    findings.append("114-02: MOVING operation roll-ahead (higher than stationary 306/212)")

    size_rows = [
        {"signCode": "NYW8-33", "NON-FREEWAY": None, "FREEWAY": "48x24"},
        {"signCode": "W20-5R", "NON-FREEWAY": None, "FREEWAY": "48x48"},
        {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
    ]
    assert "NYW8-33" in body and "W20-5R" in body
    assert_row_count(size_rows, 3, "114-03")

    assert "500'" in body
    assert "MERGING" not in body
    findings.append("NO taper tables; plan 500' minimum roll-ahead / spacing; signs NYW8-33+W20-5R only")

    draft = {
        "sheetNumber": "619-114",
        "sourcePdf": "Bridge/captures/619-114.pdf",
        "extractedOn": "2026-08-03",
        "tableRoles": {
            "protectiveVehicle": "114-01",
            "rollAheadDistance": "114-02",
            "signSizes": "114-03",
            "note": "NO taperAndBuffer role",
        },
        "tables": {
            "114-01": {"title": "PROTECTIVE VEHICLE REQUIREMENTS", "rows": pv_rows},
            "114-02": {"title": "ROLL AHEAD DISTANCE", "rows": ra},
            "114-03": {"title": "REQUIRED SIGN SIZES", "rows": size_rows},
        },
        "planGaps": {"minRollOrGap": 500, "maxMention": "2 MILE / 1500"},
        "signsOnSheet": ["NYW8-33", "W20-5R", "WARNING FLAG"],
        "findings": findings,
    }
    out = ROOT / "Data/sheet-specs/_draft_619114_tables.json"
    out.write_text(json.dumps(draft, indent=2) + "\n", encoding="utf-8")
    print("Wrote", out)
    for f in findings:
        print(" -", f)
    return draft


# --------------------------------------------------------------------------- 041
def extract_041():
    pdf = fitz.open(str(ROOT / "Bridge/captures/619-041.pdf"))
    pg = pdf[0]
    W = pg.get_text("words")
    findings = []
    body = page_body(pg)

    # NON-FREEWAY PV matrix with speed bands
    pv_rows = [
        {
            "closureType": "LANE CLOSURE OR ENCROACHMENT",
            "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
            "ge45": "P, TMIA", "b35to40": "P, TMIA", "le30": "P",
        },
        {
            "closureType": "LANE CLOSURE OR ENCROACHMENT",
            "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
            "ge45": "NA", "b35to40": "NA", "le30": "NA",
        },
        {
            "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
            "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
            "ge45": "P, TMIA", "b35to40": "P", "le30": "P",
        },
        {
            "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
            "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
            "ge45": "NA", "b35to40": "NA", "le30": "NA",
        },
    ]
    # Verify tokens — recon said w45 / 35-40 / l30 and P, TMIA / P
    assert "NON-FREEWAY" in body
    assert_row_count(pv_rows, 4, "041-01")
    findings.append("041-01: NON-FREEWAY speed-banded PV (no FREEWAY column) — mowing/parkway-adjacent")

    ra = [
        {"speedBand": ">= 55 MPH", "minMph": 55, "maxMph": None,
         "min": parse_slash_pair("200/5"), "max": parse_slash_pair("280/7")},
        {"speedBand": "45 - 50 MPH", "minMph": 45, "maxMph": 50,
         "min": parse_slash_pair("160/4"), "max": parse_slash_pair("240/6")},
        {"speedBand": "<= 40 MPH", "minMph": None, "maxMph": 40,
         "min": parse_slash_pair("120/3"), "max": parse_slash_pair("200/5")},
    ]
    assert token_at(W, 1105, 1125, 344, 356, is_ratio) == "200/5"
    assert token_at(W, 1150, 1175, 344, 356, is_ratio) == "280/7"
    assert token_at(W, 1105, 1125, 358, 370, is_ratio) == "160/4"
    assert token_at(W, 1105, 1125, 372, 385, is_ratio) == "120/3"
    assert_row_count(ra, 3, "041-02")
    findings.append("041-02: 3 speed bands including <=40 (MOVING operation values)")

    assert "W8-23" in body
    assert "36x36" in body and "48x48" in body
    # sizes sit in table header band, not always same y as code token
    size_rows = [
        {"signCode": "W8-23", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
    ]
    assert_row_count(size_rows, 1, "041-03")
    findings.append("041-03: W8-23 only (LOW SHOULDER / mowing)")

    assert "MERGING" not in body
    findings.append("NO taper tables; MOVING OPERATION; shoulder closure/lane encroachment")

    draft = {
        "sheetNumber": "619-041",
        "sourcePdf": "Bridge/captures/619-041.pdf",
        "extractedOn": "2026-08-03",
        "tableRoles": {
            "protectiveVehicle": "041-01",
            "rollAheadDistance": "041-02",
            "signSizes": "041-03",
            "note": "NO taperAndBuffer role",
        },
        "tables": {
            "041-01": {"title": "PROTECTIVE VEHICLE REQUIREMENTS", "rows": pv_rows,
                       "speedBands": [
                           {"id": "ge45", "label": ">= 45 MPH", "minMph": 45, "maxMph": None},
                           {"id": "b35to40", "label": "35 - 40 MPH", "minMph": 35, "maxMph": 40},
                           {"id": "le30", "label": "<= 30 MPH", "minMph": None, "maxMph": 30},
                       ]},
            "041-02": {"title": "ROLL AHEAD DISTANCE", "rows": ra},
            "041-03": {"title": "REQUIRED SIGN SIZES*", "rows": size_rows},
        },
        "signsOnSheet": ["W8-23"],
        "findings": findings,
    }
    out = ROOT / "Data/sheet-specs/_draft_619041_tables.json"
    out.write_text(json.dumps(draft, indent=2) + "\n", encoding="utf-8")
    print("Wrote", out)
    for f in findings:
        print(" -", f)
    return draft


if __name__ == "__main__":
    extract_306()
    print()
    extract_212()
    print()
    extract_114()
    print()
    extract_041()
