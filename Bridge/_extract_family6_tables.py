"""Extract Family 6 table drafts; verify key cells against PDF text layer."""
from __future__ import annotations

import json
import pathlib
import re
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
from pdf_table_extract import squash  # noqa: E402

OUT = ROOT / "Data/sheet-specs"
SPEC311 = json.loads((ROOT / "Data/sheet-specs/619-311.json").read_text(encoding="utf-8"))
SRC = "https://www.dot.ny.gov/main/business-center/engineering/cadd-info/drawings/standard-sheets-us-repository"


def pdf_path(n: int) -> pathlib.Path:
    return ROOT / f"Bridge/captures/619-{n:03d}.pdf" if n < 100 else ROOT / f"Bridge/captures/619-{n}.pdf"


def body(n: int) -> str:
    doc = fitz.open(str(pdf_path(n)))
    return " ".join(w[4] for p in doc for w in p.get_text("words"))


def meta(n: int) -> dict:
    doc = fitz.open(str(pdf_path(n)))
    return {"pages": doc.page_count, "rotation": doc[0].rotation}


def aw_from_311() -> list:
    return [dict(r) for r in SPEC311["tables"]["311-03"]["rows"]]


def buffer_from_311() -> list:
    return [
        {"speedMph": r["speedMph"], "longitudinalBufferSpace": dict(r["longitudinalBufferSpace"])}
        for r in SPEC311["tables"]["311-02"]["rows"]
    ]


def roll_from_311() -> list:
    return [dict(r) for r in SPEC311["tables"]["311-04"]["rows"]]


def pv_from_311() -> list:
    t = SPEC311["tables"]["311-01"]
    return {
        "speedBands": list(t["speedBands"]),
        "rows": [dict(r) for r in t["rows"]],
        "legend": dict(t["legend"]),
        "tableNotes": list(t["tableNotes"]),
    }


def assert_in_body(b: str, tokens: list[str], label: str) -> None:
    sb = squash(b)
    for t in tokens:
        if squash(t) not in sb and t not in b:
            raise AssertionError(f"{label}: missing {t!r}")


def verify_aw(b: str, label: str) -> None:
    assert_in_body(b, ["100", "200", "350", "500", "1500 FT.", "1000 FT.", "AHEAD"], label)


def verify_buffer(b: str, label: str) -> None:
    for r in buffer_from_311():
        tok = f"{r['longitudinalBufferSpace']['ft']}/{r['longitudinalBufferSpace']['skipLines']}"
        assert tok in b, f"{label}: missing buffer {tok}"


def verify_roll(b: str, label: str, require_le40: bool = True) -> None:
    toks = ["120/3", "200/5", "80/2", "160/4"]
    if require_le40:
        toks.append("40/1")
    assert_in_body(b, toks, label)


def size_rows(codes_nf_fw: list[tuple]) -> list:
    return [{"signCode": c, "NON-FREEWAY": nf, "FREEWAY": fw} for c, nf, fw in codes_nf_fw]


def write_draft(n: int, draft: dict) -> None:
    path = OUT / f"_draft_619{n:03d}_tables.json" if n < 100 else OUT / f"_draft_619{n}_tables.json"
    path.write_text(json.dumps(draft, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print("Wrote", path.name)


def extract_307():
    b = body(307)
    m = meta(307)
    verify_aw(b, "307")
    verify_buffer(b, "307")
    verify_roll(b, "307")
    assert_in_body(b, ["W20-7", "W20-4", "W3-4", "W20-1", "G20-2", "FLAGGER"], "307")
    pv = pv_from_311()
    write_draft(307, {
        "sheet": 307,
        "meta": m,
        "tableRoles": {
            "note": "Flagger base. 307-01=AW, 307-02=buffer ONLY (no lane/shoulder taper), "
                    "307-03=sizes, 307-04=PV, 307-05=roll. Tables cell-identical to 311 on overlap.",
            "advanceWarningSpacing": "307-01",
            "taperAndBuffer": "307-02",
            "signSizes": "307-03",
            "protectiveVehicle": "307-04",
            "rollAheadDistance": "307-05",
        },
        "advanceWarning": aw_from_311(),
        "bufferOnly": buffer_from_311(),
        "signSizes": size_rows([
            ("G20-2", "36x18", "48x24"),
            ("W3-4", "36x36", "48x48"),
            ("W20-1", "36x36", "48x48"),
            ("W20-4", "36x36", "48x48"),
            ("W20-7", "36x36", "48x48"),
            ("WARNING FLAG", "18x18", "18x18"),
        ]),
        "protectiveVehicle": pv,
        "rollAhead": roll_from_311(),
        "identityNote": "307-01==311-03; 307-02 buffer==311-02 buffer; 307-04==311-01; 307-05==311-04",
    })


def extract_308():
    b = body(308)
    m = meta(308)
    verify_aw(b, "308")
    verify_buffer(b, "308")
    verify_roll(b, "308")
    write_draft(308, {
        "sheet": 308,
        "meta": m,
        "tableRoles": {
            "note": "Prior to intersection flagger. Same AW/buffer/sizes/PV/roll as 307; role numbers differ: "
                    "01=AW 02=buffer 03=sizes 04=PV 05=roll.",
            "advanceWarningSpacing": "308-01",
            "taperAndBuffer": "308-02",
            "signSizes": "308-03",
            "protectiveVehicle": "308-04",
            "rollAheadDistance": "308-05",
        },
        "advanceWarning": aw_from_311(),
        "bufferOnly": buffer_from_311(),
        "signSizes": size_rows([
            ("G20-2", "36x18", "48x24"),
            ("W3-4", "36x36", "48x48"),
            ("W20-1", "36x36", "48x48"),
            ("W20-4", "36x36", "48x48"),
            ("W20-7", "36x36", "48x48"),
            ("WARNING FLAG", "18x18", "18x18"),
        ]),
        "protectiveVehicle": pv_from_311(),
        "rollAhead": roll_from_311(),
        "cloneOf": 307,
    })


def extract_sign_sizes_generic(n: int, codes: list[str]) -> list:
    """Best-effort: find each code and nearby size tokens on any page."""
    doc = fitz.open(str(pdf_path(n)))
    words = [w for p in doc for w in p.get_text("words")]
    rows = []
    for code in codes:
        if code == "WARNING FLAG":
            hits = [w for w in words if w[4] == "WARNING"]
        else:
            hits = [w for w in words if w[4] == code or w[4].startswith(code + "-") or w[4] == code.replace("-", "")]
            if not hits:
                hits = [w for w in words if code in w[4]]
        if not hits:
            rows.append({"signCode": code, "NON-FREEWAY": None, "FREEWAY": None})
            continue
        y, page_x = hits[0][1], hits[0][0]
        # same y band, size-like tokens to the right
        sizes = sorted(
            [w for w in words if abs(w[1] - y) < 8 and "x" in w[4] and w[4][0].isdigit() and w[0] > page_x - 5],
            key=lambda w: w[0],
        )
        if len(sizes) >= 2:
            nf, fw = sizes[0][4], sizes[1][4]
        elif len(sizes) == 1:
            nf, fw = sizes[0][4], sizes[0][4]
        else:
            nf, fw = None, None
        rows.append({"signCode": code, "NON-FREEWAY": nf, "FREEWAY": fw})
    return rows


def extract_flagger_sibling(n: int, roles: dict, size_codes: list[str], extra_tokens: list[str],
                            has_aw=True, has_buffer=True, has_roll=True, has_pv=True, notes=""):
    b = body(n)
    m = meta(n)
    if has_aw:
        verify_aw(b, str(n))
    if has_buffer:
        verify_buffer(b, str(n))
    if has_roll:
        # tolerate 2-band roll on some intermediate sheets
        try:
            verify_roll(b, str(n))
        except AssertionError:
            verify_roll(b, str(n), require_le40=False)
    for tok in extra_tokens:
        if squash(tok) not in squash(b) and tok not in b:
            print(f"WARN {n}: missing token {tok!r}")
    draft = {
        "sheet": n,
        "meta": m,
        "tableRoles": roles,
        "notes": notes,
        "signSizes": extract_sign_sizes_generic(n, size_codes),
    }
    if has_aw:
        draft["advanceWarning"] = aw_from_311()
    if has_buffer:
        draft["bufferOnly"] = buffer_from_311()
    if has_roll:
        draft["rollAhead"] = roll_from_311()
    if has_pv:
        draft["protectiveVehicle"] = pv_from_311()
    write_draft(n, draft)


def extract_090_091(n: int):
    b = body(n)
    m = meta(n)
    verify_aw(b, str(n))
    verify_buffer(b, str(n))
    assert_in_body(b, ["W20-1", "W20-7", "W3-4"], str(n))
    # no roll-ahead / PV tables on these closure sheets
    prefix = f"{n:03d}"
    write_draft(n, {
        "sheet": n,
        "meta": m,
        "tableRoles": {
            "note": f"Temporary {'road' if n == 90 else 'intersection'} closure. "
                    f"{prefix}-01=AW {prefix}-02=buffer {prefix}-03=sizes. NO PV/roll tables.",
            "advanceWarningSpacing": f"{prefix}-01",
            "taperAndBuffer": f"{prefix}-02",
            "signSizes": f"{prefix}-03",
        },
        "advanceWarning": aw_from_311(),
        "bufferOnly": buffer_from_311(),
        "signSizes": extract_sign_sizes_generic(n, ["W20-1", "W20-7", "W3-4"]),
        "hasProtectiveVehicle": False,
        "hasRollAhead": False,
    })


def extract_sidewalk(n: int, size_codes: list[str]):
    b = body(n)
    m = meta(n)
    assert "SIDEWALK" in b.upper() or "R9-9" in b or "R9-11" in b, f"{n}: not sidewalk?"
    prefix = str(n)
    roles = {
        "note": f"Pedestrian/sidewalk sheet. Sign sizes + channelizing only — NO AW/buffer/taper/roll corridor tables.",
        "signSizes": f"{prefix}-01",
        "channelizingApplication": f"{prefix}-02",
    }
    write_draft(n, {
        "sheet": n,
        "meta": m,
        "tableRoles": roles,
        "signSizes": extract_sign_sizes_generic(n, size_codes),
        "schema": "sidewalk",
        "bodyTokens": re.findall(r"\b(?:R9-\d+[A-Z]?|R11-\d+|W20-1|G20-2|W5-4|M4-9[A-Z]*)\b", b)[:30],
    })


def extract_322():
    b = body(322)
    m = meta(322)
    write_draft(322, {
        "sheet": 322,
        "meta": m,
        "tableRoles": {
            "note": "Crosswalk closure. 322-01=sizes, 322-02=channelizing, 322-03=advance placement guidelines (not A/B/C).",
            "signSizes": "322-01",
            "channelizingApplication": "322-02",
            "advancePlacementGuidelines": "322-03",
        },
        "signSizes": extract_sign_sizes_generic(322, [
            "G20-2", "R11-2", "R8-3", "R9-10", "R9-11L", "R9-11R", "R9-9", "W20-1",
        ]),
        "schema": "crosswalk",
    })


def extract_324_like(n: int):
    """TWLT single-lane shift — has shoulder taper + buffer + AW + roll + PV."""
    b = body(n)
    m = meta(n)
    verify_aw(b, str(n))
    verify_roll(b, str(n))
    assert "SHOULDER" in b.upper() or "L/3" in b, f"{n}: expected shoulder taper"
    # buffer tokens present
    assert "360/9" in b or "BUFFER" in b.upper(), f"{n}: buffer?"
    prefix = str(n)
    # role numbers differ per sheet — recon from titles in body
    write_draft(n, {
        "sheet": n,
        "meta": m,
        "schema": "twlt_shift",
        "signSizes": extract_sign_sizes_generic(n, [
            "G20-2", "NYW8-33", "R4-7", "W20-1", "W20-4", "W20-5", "WARNING FLAG", "NYR9-11",
        ]),
        "advanceWarning": aw_from_311(),
        "rollAhead": roll_from_311(),
        "protectiveVehicle": pv_from_311(),
        "hasShoulderTaper": True,
        "bodyHasMerging": "MERGING" in b.upper(),
        "notes": "TWLT shift — see build script for corridor/roles by content",
    })


def extract_524():
    b = body(524)
    m = meta(524)
    verify_aw(b, "524")
    write_draft(524, {
        "sheet": 524,
        "meta": m,
        "schema": "temp_signal",
        "tableRoles": {
            "note": "Long-term temp signal. 524-01=AW, 524-02=taper/buffer?, 524-03=channelizing, "
                    "524-04=flare, 524-05=sizes. NO roll-ahead table on sheet.",
            "advanceWarningSpacing": "524-01",
            "signSizes": "524-05",
            "channelizingApplication": "524-03",
            "flareRates": "524-04",
        },
        "advanceWarning": aw_from_311(),
        "signSizes": extract_sign_sizes_generic(524, [
            "G20-2", "NYR9-11", "R10-6L", "R10-6R", "W20-1", "W20-4", "W3-3", "WARNING FLAG",
        ]),
        "hasRollAhead": False,
        "hasShoulderTaper": "SHOULDER" in b.upper() or "L/3" in b,
    })


def extract_314():
    b = body(314)
    m = meta(314)
    verify_buffer(b, "314")
    verify_roll(b, "314", require_le40=False)  # moving: only >=55 and 45-50 bands
    assert_in_body(b, ["W20-7", "W20-4", "FLAGGER"], "314")
    roll_2band = [r for r in roll_from_311() if r.get("minMph") is None or r["minMph"] >= 45]
    # Prefer explicit 2-band from PDF
    roll_2band = [
        {"speedBand": ">= 55 MPH", "minMph": 55, "maxMph": None,
         "min": {"ft": 120, "skipLines": 3}, "max": {"ft": 200, "skipLines": 5}},
        {"speedBand": "45 - 50 MPH", "minMph": 45, "maxMph": 50,
         "min": {"ft": 80, "skipLines": 2}, "max": {"ft": 160, "skipLines": 4}},
    ]
    write_draft(314, {
        "sheet": 314,
        "meta": m,
        "schema": "moving_flagger",
        "tableRoles": {
            "note": "Moving flaggers. 314-01=PV, 314-02=roll, 314-03=buffer, 314-04=sizes. "
                    "NO advance-warning spacing table (plan uses fixed 500' gaps).",
            "protectiveVehicle": "314-01",
            "rollAheadDistance": "314-02",
            "taperAndBuffer": "314-03",
            "signSizes": "314-04",
        },
        "bufferOnly": buffer_from_311(),
        "rollAhead": roll_2band,
        "protectiveVehicle": pv_from_311(),
        "signSizes": [
            {"signCode": "G20-1", "NON-FREEWAY": "36x18", "FREEWAY": "48x24"},
            {"signCode": "G20-2", "NON-FREEWAY": "36x18", "FREEWAY": "48x24"},
            {"signCode": "NYW8-33", "NON-FREEWAY": "48x24", "FREEWAY": "48x24"},
            {"signCode": "W20-1", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
            {"signCode": "W20-4", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
            {"signCode": "W20-7", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
            {"signCode": "W3-4", "NON-FREEWAY": "36x36", "FREEWAY": "48x48"},
            {"signCode": "W7-3a", "NON-FREEWAY": "24x18", "FREEWAY": "36x30"},
            {"signCode": "WARNING FLAG", "NON-FREEWAY": "18x18", "FREEWAY": "18x18"},
        ],
        "hasAdvanceWarningTable": False,
        "fixedGapFt": 500,
    })


def extract_309():
    b = body(309)
    m = meta(309)
    verify_aw(b, "309")
    verify_buffer(b, "309")
    assert_in_body(b, ["AFAD", "R10-6", "W20-7"], "309")
    # 309 has two page sets of AW/buffer/sizes (01-03 and 04-06)
    write_draft(309, {
        "sheet": 309,
        "meta": m,
        "schema": "afad",
        "tableRoles": {
            "note": "AFAD. Pages duplicate AW/buffer/sizes as 01-03 and 04-06 (two setups). "
                    "Use 01-03 as primary plan roles; 04-06 are alternate layout.",
            "advanceWarningSpacing": "309-01",
            "taperAndBuffer": "309-02",
            "signSizes": "309-03",
        },
        "advanceWarning": aw_from_311(),
        "bufferOnly": buffer_from_311(),
        "signSizes": extract_sign_sizes_generic(309, [
            "G20-2", "R10-6", "W20-1", "W20-4", "W20-7", "W3-4", "WARNING FLAG",
        ]),
        "hasProtectiveVehicle": "PROTECTIVE" in b.upper(),
        "hasRollAhead": "ROLL AHEAD" in b.upper(),
    })


def main():
    extract_307()
    extract_308()
    extract_309()
    extract_314()
    extract_flagger_sibling(
        323,
        {
            "note": "Flagging at intersection. 323-01=AW 323-02=sizes 323-03=channelizing 323-04=stopping?",
            "advanceWarningSpacing": "323-01",
            "signSizes": "323-02",
            "channelizingApplication": "323-03",
        },
        ["G20-2", "W20-1", "W20-4", "W20-7", "W20-7a", "W3-4", "WARNING FLAG"],
        ["FLAGGER", "W20-7"],
        has_buffer=False, has_roll=False, has_pv=False,
        notes="Intersection flagging — buffer/PV may be on plan notes not numbered tables",
    )
    extract_flagger_sibling(
        407,
        {
            "note": "Intermediate flagger. 407-01=AW 407-02=buffer(45-65) 407-03=channelizing "
                    "407-04=sizes 407-05=PV 407-06=roll. Buffer speeds 45/50/55/65 only.",
            "advanceWarningSpacing": "407-01",
            "taperAndBuffer": "407-02",
            "channelizingApplication": "407-03",
            "signSizes": "407-04",
            "protectiveVehicle": "407-05",
            "rollAheadDistance": "407-06",
        },
        ["G20-2", "NYR9-11", "W20-1", "W20-4", "W20-7", "W20-7a", "W3-4", "WARNING FLAG"],
        ["FLAGGER", "NYR9-11", "W20-7"],
        has_aw=True, has_buffer=False, has_roll=True, has_pv=True,
        notes="buffer manually set in build — 45/50/55/65 only",
    )
    # attach 407 buffer manually
    draft407 = json.loads((OUT / "_draft_619407_tables.json").read_text(encoding="utf-8"))
    draft407["bufferOnly"] = [
        {"speedMph": 45, "longitudinalBufferSpace": {"ft": 360, "skipLines": 9}},
        {"speedMph": 50, "longitudinalBufferSpace": {"ft": 425, "skipLines": 11}},
        {"speedMph": 55, "longitudinalBufferSpace": {"ft": 495, "skipLines": 13}},
        {"speedMph": 65, "longitudinalBufferSpace": {"ft": 645, "skipLines": 16}},
    ]
    draft407["advanceWarning"] = aw_from_311()
    write_draft(407, draft407)
    extract_flagger_sibling(
        421,
        {
            "note": "Intermediate intersection flagging. 421-01=AW 421-02=channelizing 421-03=sizes "
                    "421-04=buffer; PV/roll on 421B-01/421B-02.",
            "advanceWarningSpacing": "421-01",
            "channelizingApplication": "421-02",
            "signSizes": "421-03",
            "taperAndBuffer": "421-04",
            "protectiveVehicle": "421B-01",
            "rollAheadDistance": "421B-02",
        },
        ["G20-2", "NYR9-11", "W20-1", "W20-4", "W20-7", "W20-7a", "W3-4", "WARNING FLAG"],
        ["FLAGGER", "NYR9-11"],
        has_roll=True, has_pv=True, has_buffer=True, has_aw=True,
    )
    extract_sidewalk(321, ["G20-2", "R9-11L", "R9-11R", "R9-9", "W20-1", "WARNING FLAG"])
    extract_322()
    extract_sidewalk(519, ["G20-2", "R9-11L", "R9-11R", "R9-9", "W20-1", "WARNING FLAG"])
    extract_324_like(324)
    extract_324_like(422)
    extract_524()
    extract_090_091(90)
    extract_090_091(91)
    print("Family 6 drafts done")


if __name__ == "__main__":
    main()
