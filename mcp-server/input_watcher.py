"""InputWatcher: shared cursor into chat-input.tsv. Extracted from
chat_driver.py (2026-08-04 split) as a self-contained I/O concern.
"""
from __future__ import annotations

import time
from pathlib import Path


class InputWatcher:
    """Shared cursor into chat-input.tsv. Both the main loop (waiting for
    the next top-level user message) and ask_user (waiting for a reply
    mid-turn) pull from this same cursor -- they never run concurrently
    (ask_user's wait is nested inside a tool call inside the main loop's
    own turn, and Python here is single-threaded), so there's no risk of
    either one double-consuming a line meant for the other."""

    def __init__(self, path: Path):
        self.path = path
        self._next_idx = 0

    def _read_lines(self) -> list[str]:
        if not self.path.exists():
            return []
        text = self.path.read_text(encoding="utf-8", errors="replace")
        return [ln for ln in text.splitlines() if ln.strip()]

    def skip_existing(self) -> None:
        """Call once at startup so lines from a previous session aren't
        replayed as new input."""
        self._next_idx = len(self._read_lines())

    def wait_for_next(self, poll_s: float = 0.5) -> str:
        while True:
            lines = self._read_lines()
            if len(lines) > self._next_idx:
                line = lines[self._next_idx]
                self._next_idx += 1
                # WZTCChatPanel.btnSend_Click writes "<timestamp>\t<message>".
                parts = line.split("\t", 1)
                return parts[1] if len(parts) > 1 else parts[0]
            time.sleep(poll_s)
