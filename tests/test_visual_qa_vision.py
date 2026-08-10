"""chat_driver visual-QA vision attach (no MicroStation required)."""
from __future__ import annotations

import base64
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))


def test_qa_capture_rows_and_vision_blocks(tmp_path, monkeypatch):
    import chat_driver as cd

    png = tmp_path / "qa_full.png"
    # Minimal valid-ish PNG bytes not required — we only read_bytes for base64.
    png.write_bytes(b"\x89PNG\r\n\x1a\nfake-qa-frame")

    logged: list[str] = []

    def fake_log_screenshot(path: Path) -> None:
        logged.append(str(path))

    monkeypatch.setattr(cd, "_log_screenshot", fake_log_screenshot)

    result = {
        "status": "OK",
        "visualQaPassed": True,
        "checklist": ["Dims"],
        "captures": [
            {"frame": "full_corridor", "path": str(png)},
            {"frame": "upstream", "path": str(png)},
        ],
    }
    assert len(cd._qa_capture_rows(result)) == 2
    blocks = cd._vision_blocks_for_qa_captures(result)
    assert blocks is not None
    assert blocks[0]["type"] == "text"
    payload = json.loads(blocks[0]["text"])
    assert "captures" not in payload
    assert payload["visualQaPassed"] is True
    images = [b for b in blocks if b.get("type") == "image"]
    assert len(images) == 2
    assert len(logged) == 2
    data = images[0]["source"]["data"]
    assert base64.standard_b64decode(data).startswith(b"\x89PNG")


def test_vision_none_without_captures():
    import chat_driver as cd

    assert cd._vision_blocks_for_qa_captures({"status": "OK"}) is None
    assert cd._qa_capture_rows({"captures": []}) == []
