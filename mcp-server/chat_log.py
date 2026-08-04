"""ChatLog: appends structured lines to chat-log.tsv for WZTCChatPanel.frm
to poll and render. Extracted from chat_driver.py (2026-08-04 split) as a
self-contained I/O concern.
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path


def _flatten(text: str) -> str:
    """Collapse a value to one physical line -- WZTCChatTimer.bas reads
    chat-log.tsv one Line Input# line at a time, so an embedded newline
    would look like a second, malformed log entry."""
    return text.replace("\t", "    ").replace("\r\n", " ").replace("\n", " ")


class ChatLog:
    """timestamp\tTYPE\tkey=val... -- same convention as
    Bridge/wztc-journal.tsv. CRLF is required (confirmed live during M7
    Stage 1 bring-up: VBA's Line Input# reads a bare-LF file as a single
    giant line, silently breaking the panel's line-count-based rendering)."""

    def __init__(self, path: Path, archive_dir: Path | None = None, max_bytes: int = 2_000_000):
        self.path = path
        self.archive_dir = archive_dir if archive_dir is not None else path.parent / "archive"
        self.max_bytes = max_bytes

    def _rotate_if_oversized(self) -> None:
        """Archive (rename, never delete) chat-log.tsv once it passes
        max_bytes, so it doesn't grow forever across sessions -- it was
        hit 54KB after one day with nothing ever trimming it. Safe to do
        at any time (not just process startup): WZTCChatTimer.bas's
        polling loop now detects the file getting smaller than what it's
        already delivered and resyncs from scratch instead of silently
        going stale (see the n < mLastLineCount check added alongside
        this). Best-effort -- a failed rotation just means the next write
        appends to the still-oversized file instead of blocking it."""
        try:
            if self.path.exists() and self.path.stat().st_size >= self.max_bytes:
                self.archive_dir.mkdir(exist_ok=True)
                ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
                self.path.rename(self.archive_dir / f"chat-log-{ts}.tsv")
        except OSError:
            pass

    def _write(self, line_type: str, **fields: str) -> None:
        self._rotate_if_oversized()
        kv = "\t".join(f"{k}={_flatten(str(v))}" for k, v in fields.items())
        line = f"{datetime.now()}\t{line_type}" + (f"\t{kv}" if kv else "")
        with open(self.path, "a", encoding="utf-8", newline="\r\n") as f:
            f.write(line + "\n")

    def thinking(self, text: str) -> None:
        if text.strip():
            self._write("THINKING", text=text)

    def tool_call(self, name: str, tool_input: dict) -> None:
        self._write("TOOL_CALL", name=name, input=json.dumps(tool_input, ensure_ascii=False, default=str))

    def tool_result(self, name: str, status: str, summary: str) -> None:
        self._write("TOOL_RESULT", name=name, status=status, summary=summary)

    def screenshot(self, path: str) -> None:
        self._write("SCREENSHOT", path=path)

    def reference_image(self, path: str, source_name: str, heading: str, page: int) -> None:
        self._write("REFERENCE_IMAGE", path=path, source=source_name, heading=heading, page=page)

    def ask_user_choice(self, question: str, options: list[dict],
                        allow_point_pick: bool, allow_element_pick: bool = False) -> None:
        fields = {
            "question": question,
            "allowPointPick": "Y" if allow_point_pick else "N",
            "allowElementPick": "Y" if allow_element_pick else "N",
        }
        for i, opt in enumerate(options[:4], start=1):
            fields[f"option{i}Label"] = opt.get("label", "")
            fields[f"option{i}Detail"] = opt.get("description", "")
        self._write("ASK_USER_CHOICE", **fields)

    def ask_user(self, question: str) -> None:
        self._write("ASK_USER", question=question)

    def final(self, text: str) -> None:
        self._write("FINAL", text=text)

    def error(self, note: str) -> None:
        self._write("ERROR", note=note)

    def mode_changed(self, mode: str, description: str) -> None:
        self._write("MODE_CHANGED", mode=mode, description=description)
