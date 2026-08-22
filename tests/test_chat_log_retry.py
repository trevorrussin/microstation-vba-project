"""chat-log writes must survive VBA holding the file open.

Live 2026-08-20: chat_driver's append hit "[Errno 13] Permission denied" on
Bridge/chat-log.tsv. WZTCChatTimer.bas and WZTCChatPanel.frm both poll that
file from VBA while this process appends to it, and VBA's open can hold a
brief exclusive handle. That one landed on a status line and self-recovered,
but the same race on a FINAL write loses the turn's visible answer -- which in
the panel is indistinguishable from a hang.
"""
from __future__ import annotations

import errno
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

from chat_log import ChatLog  # noqa: E402


def _log(tmp_path: Path) -> ChatLog:
    return ChatLog(tmp_path / "chat-log.tsv", tmp_path / "archive", 2_000_000)


def test_write_succeeds_normally(tmp_path):
    log = _log(tmp_path)
    log.final("hello")
    assert "hello" in (tmp_path / "chat-log.tsv").read_text(encoding="utf-8")


def test_retries_through_a_transient_lock(monkeypatch, tmp_path):
    """The exact failure: first opens denied, a later one succeeds."""
    log = _log(tmp_path)
    real_open = open
    calls = {"n": 0}

    def flaky(path, *a, **kw):
        if str(path).endswith("chat-log.tsv"):
            calls["n"] += 1
            if calls["n"] <= 3:
                raise PermissionError(errno.EACCES, "Permission denied", str(path))
        return real_open(path, *a, **kw)

    monkeypatch.setattr("builtins.open", flaky)
    monkeypatch.setattr(ChatLog, "_WRITE_BACKOFF_S", 0.001)
    log.final("survived the lock")
    assert calls["n"] == 4, "should have retried past the denials"
    assert "survived the lock" in (tmp_path / "chat-log.tsv").read_text(encoding="utf-8")


def test_turn_does_not_die_when_retries_are_exhausted(monkeypatch, tmp_path, capsys):
    """A lost log line must never take the turn down with it."""
    log = _log(tmp_path)

    def always_denied(path, *a, **kw):
        raise PermissionError(errno.EACCES, "Permission denied", str(path))

    monkeypatch.setattr("builtins.open", always_denied)
    monkeypatch.setattr(ChatLog, "_WRITE_BACKOFF_S", 0.001)
    log.final("this line is lost")  # must not raise
    assert "dropped FINAL" in capsys.readouterr().out


def test_non_sharing_errors_surface_immediately(monkeypatch, tmp_path):
    """A bad path or full disk will not fix itself -- fail fast, don't spin."""
    log = _log(tmp_path)
    attempts = {"n": 0}

    def no_space(path, *a, **kw):
        attempts["n"] += 1
        raise OSError(errno.ENOSPC, "No space left on device", str(path))

    monkeypatch.setattr("builtins.open", no_space)
    with pytest.raises(OSError) as ei:
        log.final("x")
    assert ei.value.errno == errno.ENOSPC
    assert attempts["n"] == 1, "must not retry an error that cannot resolve"


@pytest.mark.skipif(sys.platform != "win32", reason="Windows sharing semantics")
def test_real_windows_exclusive_lock_is_survived(monkeypatch, tmp_path):
    """Not a mock: take a real exclusive handle, release it, verify recovery."""
    import msvcrt
    import threading
    import time

    log = _log(tmp_path)
    p = tmp_path / "chat-log.tsv"
    p.write_text("", encoding="utf-8")
    monkeypatch.setattr(ChatLog, "_WRITE_BACKOFF_S", 0.02)

    holder_done = threading.Event()

    def hold():
        with open(p, "a+b") as f:
            try:
                msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
                time.sleep(0.25)
                msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
            except OSError:
                pass
        holder_done.set()

    t = threading.Thread(target=hold)
    t.start()
    time.sleep(0.02)
    log.final("written despite a real lock")   # must not raise
    t.join(timeout=5)
    assert holder_done.is_set()
    assert "written despite a real lock" in p.read_text(encoding="utf-8")
