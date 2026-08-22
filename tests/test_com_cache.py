"""COM type-library cache: keep it off %TEMP%, self-heal it, and never lie
about why an attach failed.

Regression net for 2026-08-20: a temp cleaner removed every .py source under
the MicroStation gen_py entry, leaving __pycache__. The gutted directory still
imported as an empty namespace package, so pywin32 raised
"has no attribute 'CLSIDToClassMap'" -- and ms_connect swallowed it in a bare
except, reporting "no MicroStation instances found running at all" while
MicroStation sat there with the file open.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))

pytest.importorskip("win32com", reason="pywin32 only available on Windows")

import win32com  # noqa: E402

import com_cache  # noqa: E402


def test_default_cache_is_not_in_temp():
    """%TEMP% is disposable by definition -- that is what broke it."""
    d = com_cache.default_cache_dir()
    tmp = os.environ.get("TEMP") or os.environ.get("TMP") or ""
    assert d
    if tmp:
        assert os.path.normcase(tmp) not in os.path.normcase(d)
    # Version-scoped: a 3.12 cache must not be handed to 3.13.
    assert f"{sys.version_info.major}.{sys.version_info.minor}" in d


def test_env_override(monkeypatch, tmp_path):
    monkeypatch.setenv("WZTC_GEN_PY_DIR", str(tmp_path / "custom"))
    assert com_cache.default_cache_dir() == str(tmp_path / "custom")


def test_relocate_sets_both_the_string_and_the_module_path(tmp_path):
    """win32com resolves imports via win32com.gen_py.__path__, not the string.

    Setting only __gen_path__ leaves generation and import pointing at
    different directories.
    """
    target = tmp_path / "relocated"
    got = com_cache.relocate_cache(str(target))
    try:
        assert os.path.normcase(got) == os.path.normcase(str(target))
        assert os.path.isdir(got)
        assert os.path.normcase(win32com.__gen_path__) == os.path.normcase(got)
        gen_py = sys.modules.get("win32com.gen_py")
        assert gen_py is not None
        assert os.path.normcase(gen_py.__path__[0]) == os.path.normcase(got)
    finally:
        com_cache.relocate_cache(com_cache.CACHE_DIR)


def test_relocate_is_idempotent(tmp_path):
    a = com_cache.relocate_cache(str(tmp_path / "x"))
    b = com_cache.relocate_cache(str(tmp_path / "x"))
    try:
        assert a == b
    finally:
        com_cache.relocate_cache(com_cache.CACHE_DIR)


def test_corruption_signature_detection():
    assert com_cache._looks_corrupt(
        AttributeError("module 'win32com.gen_py.X' has no attribute 'CLSIDToClassMap'"))
    assert com_cache._looks_corrupt(
        AttributeError("module 'win32com.gen_py.X' has no attribute 'CLSIDToPackageMap'"))
    # An unrelated AttributeError must NOT trigger a cache rebuild.
    assert not com_cache._looks_corrupt(AttributeError("'NoneType' has no attribute 'foo'"))
    # Nor must a different exception type.
    assert not com_cache._looks_corrupt(RuntimeError("CLSIDToClassMap"))


def test_dispatch_reraises_unrelated_errors_untouched(monkeypatch):
    """A real COM problem must not be disguised as a cache problem."""
    boom = RuntimeError("COM server rejected the call")

    def fake_ensure(_):
        raise boom

    called = []
    monkeypatch.setattr(com_cache.gencache, "EnsureDispatch", fake_ensure)
    monkeypatch.setattr(com_cache, "regenerate_typelib",
                        lambda *a: called.append(a) or True)
    with pytest.raises(RuntimeError) as ei:
        com_cache.dispatch_with_repair(object())
    assert ei.value is boom
    assert called == [], "must not rebuild the cache for an unrelated error"


def test_dispatch_repairs_then_retries_once(monkeypatch):
    attempts = {"n": 0}
    sentinel = object()

    def fake_ensure(_):
        attempts["n"] += 1
        if attempts["n"] == 1:
            raise AttributeError(
                "module 'win32com.gen_py.X' has no attribute 'CLSIDToClassMap'")
        return sentinel

    monkeypatch.setattr(com_cache.gencache, "EnsureDispatch", fake_ensure)
    monkeypatch.setattr(com_cache, "_typelib_spec", lambda d: ("{GUID}", 0, 1, 0))
    monkeypatch.setattr(com_cache, "regenerate_typelib", lambda *a: True)
    assert com_cache.dispatch_with_repair(object()) is sentinel
    assert attempts["n"] == 2, "exactly one retry"


def test_dispatch_raises_original_when_repair_fails(monkeypatch):
    err = AttributeError("module 'win32com.gen_py.X' has no attribute 'CLSIDToClassMap'")
    monkeypatch.setattr(com_cache.gencache, "EnsureDispatch",
                        lambda _: (_ for _ in ()).throw(err))
    monkeypatch.setattr(com_cache, "_typelib_spec", lambda d: ("{GUID}", 0, 1, 0))
    monkeypatch.setattr(com_cache, "regenerate_typelib", lambda *a: False)
    with pytest.raises(AttributeError) as ei:
        com_cache.dispatch_with_repair(object())
    assert ei.value is err


def test_bind_failure_is_not_reported_as_not_running(monkeypatch):
    """The 2026-08-20 misdiagnosis, locked down.

    MicroStation registered in the ROT but unbindable must NOT produce
    "no MicroStation instances found running at all".
    """
    import ms_connect

    def fake_candidates(prog_id=ms_connect.DEFAULT_PROG_ID, bind_errors=None):
        if bind_errors is not None:
            bind_errors.append(
                "AttributeError: module 'win32com.gen_py.X' has no "
                "attribute 'CLSIDToClassMap'")
        return iter(())

    monkeypatch.setattr(ms_connect, "_candidate_apps", fake_candidates)
    with pytest.raises(RuntimeError) as ei:
        ms_connect.get_microstation_app("Test")
    msg = str(ei.value)
    assert "no MicroStation instances found running at all" not in msg
    assert "IS registered" in msg
    assert "CLSIDToClassMap" in msg
    assert "com_cache" in msg


def test_genuinely_absent_still_reports_absent(monkeypatch):
    """The opposite case must keep its original, correct message."""
    import ms_connect

    def fake_candidates(prog_id=ms_connect.DEFAULT_PROG_ID, bind_errors=None):
        return iter(())

    monkeypatch.setattr(ms_connect, "_candidate_apps", fake_candidates)
    with pytest.raises(RuntimeError) as ei:
        ms_connect.get_microstation_app("Test")
    assert "no MicroStation instances found running at all" in str(ei.value)
