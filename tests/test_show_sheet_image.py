"""The agent must be able to SHOW a 619 sheet, not just describe it.

Regression net for 2026-08-20: asked "can you show me what standard sheet
619-311 looks like", the agent replied "I can't render or display the actual
NYSDOT PDF image in this chat -- I have no image-display tool, only
text-returning lookups." That was false. The panel has displayed images since
2026-08-02 (_log_screenshot / _show_reference_image), and all 68 sheet PDFs are
on disk. The capability existed with nothing exposing it, so the model answered
honestly from what it could see.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

import wztc_ops as ops  # noqa: E402


def _png_header(path: str) -> bytes:
    with open(path, "rb") as f:
        return f.read(8)


def test_sheet_with_prerendered_png():
    """619-311 carries a localRender -- use it rather than re-rendering."""
    r = ops.show_sheet_image("619-311")
    assert r["status"] == "OK" and r["found"] is True
    assert r["source"] == "localRender"
    assert os.path.getsize(r["imagePath"]) > 10_000
    assert "MULTILANE UNDIVIDED" in r["title"].upper()


@pytest.mark.parametrize("sheet", ["619-301", "619-415"])
def test_sheet_without_prerendered_png_renders_on_demand(sheet):
    """Only 4 of 68 specs have a localRender; every one has a localPdf.

    Looking only for a pre-made PNG would work for 4 sheets and fail for 64.
    """
    r = ops.show_sheet_image(sheet)
    assert r["status"] == "OK", r
    assert r["source"] == "localPdf"
    assert _png_header(r["imagePath"]).startswith(b"\x89PNG")
    assert os.path.getsize(r["imagePath"]) > 50_000, "suspiciously small — blank page?"
    assert r["pageCount"] >= 1


def test_every_spec_sheet_can_be_shown():
    """No sheet may be un-showable: spec present => PDF present."""
    import glob
    import json
    import io

    unshowable = []
    for f in glob.glob(str(ROOT / "Data" / "sheet-specs" / "*.json")):
        try:
            d = json.load(io.open(f, encoding="utf-8"))
        except Exception:
            continue
        sh = d.get("sheet") if isinstance(d, dict) else None
        if not isinstance(sh, dict):
            continue
        pdf = sh.get("localPdf")
        if not pdf or not (ROOT / str(pdf)).exists():
            unshowable.append(sh.get("number"))
    assert unshowable == [], f"sheets with a spec but no usable PDF: {unshowable}"


def test_multipage_sheet_page_two():
    r = ops.show_sheet_image("619-415", page=2)
    if r.get("pageCount", 1) < 2:
        pytest.skip("619-415 is single page in this copy")
    assert r["status"] == "OK" and r["page"] == 2


def test_page_out_of_range_is_a_clean_error():
    r = ops.show_sheet_image("619-311", page=99)
    assert r["status"] == "ERROR"
    assert "out of range" in r["note"]
    assert r["pageCount"] >= 1


def test_unknown_sheet_does_not_guess():
    r = ops.show_sheet_image("619-999")
    assert r["status"] == "ERROR" and r["found"] is False
    assert "619-999" in r["note"]


def test_missing_sheet_num():
    assert ops.show_sheet_image("")["status"] == "ERROR"


def test_driver_hook_displays_only_for_this_tool(monkeypatch):
    """The panel hook must fire for show_sheet_image and nothing else."""
    import chat_driver

    shown = []
    monkeypatch.setattr(chat_driver, "_log_screenshot", lambda p: shown.append(str(p)))

    ok = {"status": "OK", "imagePath": r"C:\x\sheet.png"}
    chat_driver._show_sheet_image("show_sheet_image", ok)
    assert shown == [r"C:\x\sheet.png"]

    shown.clear()
    # Wrong tool, failed lookup, and missing path must all be no-ops.
    chat_driver._show_sheet_image("get_sheet_requirements", ok)
    chat_driver._show_sheet_image("show_sheet_image", {"status": "ERROR"})
    chat_driver._show_sheet_image("show_sheet_image", {"status": "OK"})
    chat_driver._show_sheet_image("show_sheet_image", "not a dict")
    assert shown == []


def test_driver_hook_never_fails_the_turn(monkeypatch):
    """A display hiccup must not break an otherwise good lookup."""
    import chat_driver

    def boom(_):
        raise OSError("no such file")

    logged = []
    monkeypatch.setattr(chat_driver, "_log_screenshot", boom)
    monkeypatch.setattr(chat_driver.LOG, "error", lambda m: logged.append(m))
    chat_driver._show_sheet_image("show_sheet_image", {"status": "OK", "imagePath": "x.png"})
    assert logged and "sheet-image display failed" in logged[0]


def test_tool_is_reachable_by_the_agent():
    """A tool the driver cannot call is not wired (2026-08-02 registration gap)."""
    import re
    import io

    src = io.open(ROOT / "mcp-server" / "chat_driver.py", encoding="utf-8").read()
    names: set[str] = set()
    for key in ("_BASE_OP_NAMES", "_WZTC_OP_NAMES"):
        m = re.search(rf"^{key} = \[(.*?)^\]", src, re.S | re.M)
        names |= set(re.findall(r'"([a-z_0-9]+)"', m.group(1)))
    assert "show_sheet_image" in names


def test_search_reference_manual_documents_that_it_shows_an_image():
    """The model only knows what the docstring tells it."""
    import io

    src = io.open(ROOT / "mcp-server" / "server.py", encoding="utf-8").read()
    i = src.index("def search_reference_manual")
    doc = src[i:i + 2200]
    assert "DISPLAYS AN IMAGE" in doc
    assert "show_sheet_image" in doc


# ---------------------------------------------------------------- open_sheet_pdf
# show_sheet_image is for LOOKING (static bitmap in the panel); open_sheet_pdf
# is for WORKING (real viewer: zoom, pan, markup, search, print, save-a-copy).
# Engineer request 2026-08-20.


def test_open_sheet_pdf_shells_the_right_file(monkeypatch):
    opened = []
    monkeypatch.setattr(ops.os, "startfile", lambda p: opened.append(p), raising=False)
    r = ops.open_sheet_pdf("619-311")
    assert r["status"] == "OK" and r["found"] is True
    assert len(opened) == 1
    assert opened[0].lower().endswith("619-311.pdf")
    assert Path(opened[0]).exists()
    assert r["pdfPath"] == opened[0]


def test_open_sheet_pdf_unknown_sheet_opens_nothing(monkeypatch):
    opened = []
    monkeypatch.setattr(ops.os, "startfile", lambda p: opened.append(p), raising=False)
    r = ops.open_sheet_pdf("619-999")
    assert r["status"] == "ERROR" and r["found"] is False
    assert opened == [], "must not shell anything for an unknown sheet"


def test_open_sheet_pdf_requires_sheet_num(monkeypatch):
    opened = []
    monkeypatch.setattr(ops.os, "startfile", lambda p: opened.append(p), raising=False)
    assert ops.open_sheet_pdf("")["status"] == "ERROR"
    assert opened == []


def test_open_sheet_pdf_viewer_failure_still_reports_the_path(monkeypatch):
    """If no viewer is registered, tell them where the file is."""
    def boom(_):
        raise OSError("no application associated")

    monkeypatch.setattr(ops.os, "startfile", boom, raising=False)
    r = ops.open_sheet_pdf("619-311")
    assert r["status"] == "ERROR"
    assert r["pdfPath"].lower().endswith("619-311.pdf")
    assert "manually" in r["note"]


def test_open_sheet_pdf_refuses_paths_outside_the_project(monkeypatch):
    """A bad spec must never shell open something elsewhere on disk."""
    monkeypatch.setattr(
        ops.sheet_spec, "load",
        lambda sn: {"sheet": {"localPdf": r"../../../Windows/System32/calc.exe"}})
    opened = []
    monkeypatch.setattr(ops.os, "startfile", lambda p: opened.append(p), raising=False)
    r = ops.open_sheet_pdf("619-311")
    assert r["status"] == "ERROR"
    assert "outside the project" in r["note"]
    assert opened == []


def test_open_sheet_pdf_is_reachable_by_the_agent():
    import re
    import io

    src = io.open(ROOT / "mcp-server" / "chat_driver.py", encoding="utf-8").read()
    names: set[str] = set()
    for key in ("_BASE_OP_NAMES", "_WZTC_OP_NAMES"):
        m = re.search(rf"^{key} = \[(.*?)^\]", src, re.S | re.M)
        names |= set(re.findall(r'"([a-z_0-9]+)"', m.group(1)))
    assert "open_sheet_pdf" in names
