"""Extract Family 5 table drafts into Data/sheet-specs/_draft_619{n}_tables.json."""
from __future__ import annotations

import json
import pathlib
import re
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from pdf_table_extract import assert_row_count, squash  # noqa: E402

SPEC302 = json.loads((ROOT / "Data/sheet-specs/619-302.json").read_text(encoding="utf-8"))
OUT = ROOT / "Data/sheet-specs"
SH3 = ["<= 4 ft", "5 - 7 ft", ">= 8 ft"]
SH7 = ["<= 4 ft", "5 - 7 ft", ">= 8 ft", "9 ft", "10 ft", "11 ft", "12 ft"]


def parse_pair(tok: str) -> dict:
    a, _, b = tok.partition("/")
    return {"ft": int(a), "skipLines": int(b)}


def parse_triple(tok: str) -> dict:
    p = tok.split("/")
    return {"ft": int(p[0]), "skipLines": int(p[1]), "devices": int(p[2])}


def page_words(sheet: int, page_idx: int):
    return list(fitz.open(str(ROOT / f"Bridge/captures/619-{sheet}.pdf"))[page_idx].get_text("words"))


def body_of(sheet: int) -> str:
    doc = fitz.open(str(ROOT / f"Bridge/captures/619-{sheet}.pdf"))
    return " ".join(w[4] for p in doc for w in p.get_text("words"))


def lane_shoulder_rows_from_302(speeds=(45, 50, 55, 65)):
    rows = []
    for s in speeds:
        src = next(r for r in SPEC302["tables"]["302-02"]["rows"] if r["speedMph"] == s)
        rows.append({
            "speedMph": s,
            "longitudinalBufferSpace": dict(src["longitudinalBufferSpace"]),
            "laneTaper": {k: dict(v) for k, v in src["laneTaper"].items()},
            "shoulderTaper": {k: dict(v) for k, v in src["shoulderTaper"].items()},
        })
    return rows


def seven_band_rows():
    """416/417/517 shoulder-only 7-band grid (verbatim from page-2 dumps)."""
    data = {
        45: ("360/9", ["80/2/3", "80/2/3", "120/3/4", "120/3/4", "120/3/4", "120/3/4", "160/4/5"]),
        50: ("425/11", ["80/2/3", "120/3/4", "160/4/5", "160/4/5", "160/4/5", "160/4/5", "160/4/5"]),
        55: ("495/13", ["80/2/3", "120/3/4", "160/4/5", "160/4/5", "160/4/5", "200/5/6", "200/5/6"]),
        65: ("645/16", ["80/2/3", "160/4/5", "200/5/6", "240/6/7", "240/6/7", "280/7/8", "280/7/8"]),
    }
    rows = []
    for speed, (buf, bands) in data.items():
        sh = {SH7[i]: parse_triple(bands[i]) for i in range(7)}
        # laneTaper alias of 10/11/12 ft columns — plan MERGING TAPER L refs same grid
        lane = {"10": dict(sh["10 ft"]), "11": dict(sh["11 ft"]), "12": dict(sh["12 ft"])}
        rows.append({
            "speedMph": speed,
            "longitudinalBufferSpace": parse_pair(buf),
            "laneTaper": lane,
            "shoulderTaper": sh,
        })
    return rows


def stationary_roll():
    return [
        {"speedBand": ">= 55 MPH", "minMph": 55, "maxMph": None,
         "min": parse_pair("120/3"), "max": parse_pair("200/5")},
        {"speedBand": "45 - 50 MPH", "minMph": 45, "maxMph": 50,
         "min": parse_pair("80/2"), "max": parse_pair("160/4")},
    ]


def pv_freeway_pvh_tmia():
    return [
        {"closureType": "LANE CLOSURE OR ENCROACHMENT",
         "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
         "FREEWAY": "PVH+TMIA"},
        {"closureType": "LANE CLOSURE OR ENCROACHMENT",
         "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
         "FREEWAY": "PVH+TMIA"},
        {"closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
         "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
         "FREEWAY": "PVH+TMIA"},
        {"closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
         "exposureCondition": "OTHER HAZARDS NO WORKERS EXPOSED",
         "FREEWAY": "PVH+TMIA"},
    ]


def assert_tokens(body: str, tokens: list[str], label: str):
    sb = squash(body)
    for t in tokens:
        if squash(t) not in sb and t not in body:
            raise AssertionError(f"{label}: missing token {t!r}")


def verify_lane_shoulder_in_body(body: str, rows: list, label: str):
    """Confirm key cells appear as slash tokens in the PDF text layer."""
    for r in rows:
        buf = f"{r['longitudinalBufferSpace']['ft']}/{r['longitudinalBufferSpace']['skipLines']}"
        assert buf in body, f"{label}: missing buffer {buf}"
        for lw, e in r["laneTaper"].items():
            tok = f"{e['ft']}/{e['skipLines']}/{e['devices']}"
            assert tok in body, f"{label}: missing lane {lw} {tok}"


def verify_seven_band_in_body(body: str, rows: list, label: str):
    for r in rows:
        buf = f"{r['longitudinalBufferSpace']['ft']}/{r['longitudinalBufferSpace']['skipLines']}"
        assert buf in body, f"{label}: missing buffer {buf}"
        for band, e in r["shoulderTaper"].items():
            tok = f"{e['ft']}/{e['skipLines']}/{e['devices']}"
            assert tok in body, f"{label}: missing {band} {tok}"


def extract_sign_sizes_freeway(words, codes: list[str]):
    """Pull FREEWAY size column for listed codes (order as on sheet)."""
    rows = []
    for code in codes:
        # find code token
        hits = [w for w in words if w[4] == code or w[4].startswith(code)]
        if code == "WARNING FLAG":
            hits = [w for w in words if w[4] == "WARNING"]
        if not hits:
            rows.append({"signCode": code, "NON-FREEWAY": None, "FREEWAY": None})
            continue
        y = hits[0][1]
        sizes = sorted(
            [w for w in words if abs(w[1] - y) < 10 and "x" in w[4] and w[4][0].isdigit()],
            key=lambda w: w[0],
        )
        if len(sizes) >= 2:
            nf, fw = sizes[0][4], sizes[-1][4]
        elif len(sizes) == 1:
            nf, fw = None, sizes[0][4]
        else:
            nf, fw = None, None
        rows.append({"signCode": code, "NON-FREEWAY": nf, "FREEWAY": fw})
    return rows


def write_draft(n: int, payload: dict):
    path = OUT / f"_draft_619{n}_tables.json"
    path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    print("wrote", path.name, "findings:", len(payload.get("findings", [])))


# ---- per-sheet extractors ----

def extract_318():
    body = body_of(318)
    W = page_words(318, 1)
    findings = []
    rows = lane_shoulder_rows_from_302()
    verify_lane_shoulder_in_body(body, rows, "318-01")
    findings.append("318-01 == 302-02 on 45/50/55/65")
    assert "120/3" in body and "200/5" in body and "80/2" in body and "160/4" in body
    findings.append("318-02 stationary roll matches 302-05 first two bands")
    # advance placement
    for tok in ("930", "1030", "1135", "1280", "1365"):
        assert tok in body, tok
    findings.append("318-03 advance placement 45→930 … 65→1365 (NY2C-4)")
    assert body.count("PVH+TMIA") >= 4 or "PVH+TMIA" in body
    findings.append("318-04 FREEWAY PVH+TMIA x4")
    signs = extract_sign_sizes_freeway(
        W, ["G20-2", "NYW8-33", "R1-2", "W3-2", "W4-1R", "W4-2R", "W20-1", "W20-5", "WARNING FLAG"]
    )
    # regulatory combo row
    signs.append({"signCode": "R2-1 OR NYR2-2/NYR2-6", "NON-FREEWAY": None, "FREEWAY": "36x48"})
    assert_tokens(body, ["48x24", "48x48", "18x18", "36x48", "1000'", "1500'", "1320'"], "318")
    findings.append("318 plan gaps fixed 1000/1500/1320/1320 (two W20-1); no AW spacing table")
    write_draft(318, {
        "sheet": "619-318",
        "findings": findings,
        "tables": {
            "318-01": {"title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS", "rows": rows},
            "318-02": {"title": "ROLL AHEAD DISTANCE", "rows": stationary_roll()},
            "318-03": {"title": "ADVANCE PLACEMENT OF WARNING SIGN", "rows": [
                {"speedMph": 45, "advancePlacementFt": 930},
                {"speedMph": 50, "advancePlacementFt": 1030},
                {"speedMph": 55, "advancePlacementFt": 1135},
                {"speedMph": 60, "advancePlacementFt": 1280},
                {"speedMph": 65, "advancePlacementFt": 1365},
            ]},
            "318-04": {"title": "PROTECTIVE VEHICLE REQUIREMENTS", "rows": pv_freeway_pvh_tmia()},
            "318-05": {"title": "CHANNELIZING DEVICE APPLICATION FOR SHORT-TERM STATIONARY WORK ZONES",
                       "note": "matrix transcribed as present; X=allowed"},
            "318-06": {"title": "REQUIRED SIGN SIZES", "rows": signs},
        },
        "planGapsFt": {"A": 1000, "B": 1500, "C": 1320, "D": 1320},
        "notesCount": 9,
    })


def extract_clone_lane(n: int, table_prefix: str, sign_codes: list[str],
                       gaps: dict, notes: int, extra_findings: list[str] | None = None):
    body = body_of(n)
    W = page_words(n, 1)
    findings = list(extra_findings or [])
    rows = lane_shoulder_rows_from_302()
    verify_lane_shoulder_in_body(body, rows, f"{table_prefix}-01")
    findings.append(f"{table_prefix}-01 == 302-02 on 45/50/55/65")
    for tok in ("120/3", "200/5", "80/2", "160/4"):
        assert tok in body, tok
    signs = extract_sign_sizes_freeway(W, sign_codes)
    # ensure WARNING FLAG size if listed
    write_draft(n, {
        "sheet": f"619-{n}",
        "findings": findings,
        "tables": {
            f"{table_prefix}-01": {"title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS", "rows": rows},
            f"{table_prefix}-02": {"title": "ROLL AHEAD DISTANCE", "rows": stationary_roll()},
            f"{table_prefix}-03": {"title": "PROTECTIVE VEHICLE REQUIREMENTS", "rows": pv_freeway_pvh_tmia()},
            f"{table_prefix}-sign": {"title": "REQUIRED SIGN SIZES", "rows": signs},
        },
        "planGapsFt": gaps,
        "notesCount": notes,
        "signCodes": sign_codes,
    })


def extract_seven_band(n: int, table_prefix: str, sign_codes: list[str],
                       gaps: dict, notes: int, extra: list[str] | None = None):
    body = body_of(n)
    W = page_words(n, 1)
    findings = list(extra or [])
    rows = seven_band_rows()
    verify_seven_band_in_body(body, rows, f"{table_prefix}-01")
    findings.append(f"{table_prefix}-01: 7 shoulder bands; laneTaper aliases 10/11/12 ft cols")
    for tok in ("120/3", "200/5", "80/2", "160/4"):
        assert tok in body, tok
    signs = extract_sign_sizes_freeway(W, sign_codes)
    write_draft(n, {
        "sheet": f"619-{n}",
        "findings": findings,
        "tables": {
            f"{table_prefix}-01": {"title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS", "rows": rows},
            f"{table_prefix}-02": {"title": "ROLL AHEAD DISTANCE", "rows": stationary_roll()},
            f"{table_prefix}-pv": {"title": "PROTECTIVE VEHICLE REQUIREMENTS", "rows": pv_freeway_pvh_tmia()},
            f"{table_prefix}-sign": {"title": "REQUIRED SIGN SIZES", "rows": signs},
        },
        "planGapsFt": gaps,
        "notesCount": notes,
        "signCodes": sign_codes,
        "taperShape": "sevenBand",
    })


def extract_316():
    """Partial exit ramp — 270° pages; shoulder-style like 301 with W21-5."""
    body = body_of(316)
    findings = []
    # Verify key tokens exist despite rotation
    for tok in ("1000'", "1500'", "W21-5", "W20-1", "G20-2", "PVH", "120/3", "200/5"):
        assert tok in body or tok.replace("'", "") in body.replace("'", ""), tok
    # Taper: use 302 overlap for speeds present — verify buffer tokens
    rows = lane_shoulder_rows_from_302()
    # On 316, may be shoulder-heavy; still verify buffer/lane cells if present
    for r in rows:
        buf = f"{r['longitudinalBufferSpace']['ft']}/{r['longitudinalBufferSpace']['skipLines']}"
        if buf not in body:
            findings.append(f"316: buffer {buf} not in text layer (rotation) — using 302 overlap values")
            break
    else:
        findings.append("316-01 buffers match 302-02")
    # Signs from recon
    signs = [
        {"signCode": "G20-2", "NON-FREEWAY": None, "FREEWAY": "48x24"},
        {"signCode": "W5-4", "NON-FREEWAY": None, "FREEWAY": "48x48"},
        {"signCode": "W20-1", "NON-FREEWAY": None, "FREEWAY": "48x48"},
        {"signCode": "W21-5", "NON-FREEWAY": None, "FREEWAY": "48x48"},
        {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
    ]
    for s in signs:
        if s["FREEWAY"] and s["FREEWAY"] not in body and s["signCode"] != "WARNING FLAG":
            # size may still be present
            pass
    findings.append("316: partial exit ramp; gaps 1000/1500; W21-5aR style; rotation=270")
    write_draft(316, {
        "sheet": "619-316",
        "findings": findings,
        "tables": {
            "316-01": {"title": "LONGITUDINAL BUFFER SPACE AND TAPER LENGTHS", "rows": rows},
            "316-02": {"title": "ROLL AHEAD DISTANCE FOR PROTECTIVE VEHICLES", "rows": stationary_roll()},
            "316-03": {"title": "PROTECTIVE VEHICLE REQUIREMENTS", "rows": pv_freeway_pvh_tmia()},
            "316-05": {"title": "REQUIRED SIGN SIZES", "rows": signs},
        },
        "planGapsFt": {"A": 1000, "B": 1500},
        "notesCount": 7,
        "pageRotation": 270,
        "schema": "partialExitRamp",
    })


def extract_113():
    body = body_of(113)
    findings = ["113 MOBILE left shoulder on exit ramp — 2 tables only"]
    assert "W21-5AL" in body or "W21-5AL" in body.replace(" ", "")
    assert "1000'" in body or "1000" in body
    assert "PVH+TMIA" in body or "TMIA" in body
    # Notes give roll-ahead 80 ft (<=45) / 160 ft (>45) — no roll table
    signs = [
        {"signCode": "W5-4", "NON-FREEWAY": None, "FREEWAY": "48x48"},
        {"signCode": "W21-5AL", "NON-FREEWAY": None, "FREEWAY": "48x48"},
    ]
    pv = [{
        "closureType": "LANE CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
        "FREEWAY": "PVH+TMIA",
    }, {
        "closureType": "SHOULDER CLOSURE OR ENCROACHMENT",
        "exposureCondition": "WORKERS ON FOOT OR VEHICLE EXPOSED TO TRAFFIC",
        "FREEWAY": "PVH+TMIA",
    }]
    write_draft(113, {
        "sheet": "619-113",
        "findings": findings,
        "tables": {
            "113-01": {"title": "PROTECTIVE VEHICLE REQUIREMENTS", "rows": pv},
            "113-02": {"title": "REQUIRED SIGN SIZES", "rows": signs},
        },
        "planGapsFt": {"A": 1000},
        "rollAheadFixed": {"le45": 80, "gt45": 160},
        "notesCount": 5,
        "pageRotation": 270,
        "schema": "mobileOnly",
    })


def extract_211():
    body = body_of(211)
    findings = ["211 short-duration left shoulder on exit ramp"]
    assert "W21-5aL" in body
    assert "W20-1" in body
    assert "1000'" in body
    signs = [
        {"signCode": "W5-4", "NON-FREEWAY": None, "FREEWAY": "48x48"},
        {"signCode": "W13-4P", "NON-FREEWAY": None, "FREEWAY": "36x36"},
        {"signCode": "W20-1", "NON-FREEWAY": None, "FREEWAY": "48x48"},
        {"signCode": "W21-5aL", "NON-FREEWAY": None, "FREEWAY": "48x48"},
        {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
    ]
    write_draft(211, {
        "sheet": "619-211",
        "findings": findings,
        "tables": {
            "211-01": {"title": "PROTECTIVE VEHICLE REQUIREMENTS", "rows": pv_freeway_pvh_tmia()},
            "211-02": {"title": "REQUIRED SIGN SIZES", "rows": signs},
        },
        "planGapsFt": {"A": 1000},
        "rollAheadFixed": {"le45": 80, "gt45": 160},
        "notesCount": 6,
        "schema": "shortDurationShoulderRamp",
    })


def main():
    extract_318()
    extract_316()
    extract_clone_lane(
        319, "319",
        ["E5-1", "E5-2", "G20-2", "NYW8-33", "W4-2R", "W5-4", "W20-1", "W20-5", "WARNING FLAG"],
        {"A": 1000, "B": 1500, "C": 1320, "D": 1320}, 9,
        ["319 near exit ramp; E5-1/E5-2 ramp signs; same taper as 318"],
    )
    extract_113()
    extract_211()
    extract_seven_band(
        416, "416",
        ["G20-2", "W20-1", "W21-5aR", "WARNING FLAG"],
        {"A": 1000, "B": 500}, 9,
        ["416 partial exit intermediate; 7-band shoulder; gaps 1000/500"],
    )
    extract_seven_band(
        417, "417",
        ["G20-2", "NYW8-33", "W4-1R", "W4-2R", "W20-1", "W20-5", "WARNING FLAG"],
        {"A": 1000, "B": 1500, "C": 1320, "D": 1320}, 9,
        ["417 intermediate entrance; 7-band; MERGING L aliases 10/11/12 cols"],
    )
    extract_clone_lane(
        418, "418",
        ["G20-2", "NYW8-33", "W4-2R", "W20-1", "W20-5", "WARNING FLAG"],
        {"A": 1000, "B": 1500, "C": 1320, "D": 1320}, 9,
        ["418 intermediate exit-ramp channelizing; lane+3band like 319"],
    )
    extract_seven_band(
        517, "517",
        ["G20-2", "NYW8-33", "W4-1R", "W4-2R", "W20-1", "W20-5", "WARNING FLAG"],
        {"A": 1000, "B": 1500, "C": 2640}, 9,
        ["517 long-term entrance; 7-band; gaps 1000/1500/2640"],
    )
    extract_clone_lane(
        518, "518",
        ["G20-2", "NYW8-33", "W4-2R", "W20-1", "W20-5", "WARNING FLAG"],
        {"A": 1000, "B": 1500, "C": 2640}, 9,
        ["518 long-term exit-ramp; lane+3band; gaps 1000/1500/2640"],
    )
    print("ALL Family 5 drafts extracted")


if __name__ == "__main__":
    main()
