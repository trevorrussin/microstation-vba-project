"""Keep the pywin32 COM type-library cache alive and self-healing.

pywin32 generates early-bound wrappers for MicroStation's type library into
`win32com.__gen_path__`, which defaults to `%TEMP%\\gen_py\\<pyver>`. That is a
directory Windows treats as disposable, and on 2026-08-20 a temp cleaner
removed every `.py` source under the MicroStation entry while leaving
`__pycache__` behind. A directory with no `__init__.py` still imports as an
empty namespace package, so pywin32 blew up with:

    module 'win32com.gen_py.CF9F97BF-...' has no attribute 'CLSIDToClassMap'

`ms_connect._candidate_apps` swallowed that in a bare `except Exception:
continue`, so the bridge reported "no MicroStation instances found running at
all" while MicroStation sat right there with the file open. Two failures, one
symptom: a cache that lives somewhere disposable, and an error that pointed at
the wrong thing.

This module fixes both halves:

  1. RELOCATE -- move the cache out of %TEMP% to a stable per-user directory
     nothing sweeps. Must be done before any generation, and must set BOTH
     `win32com.__gen_path__` and `sys.modules["win32com.gen_py"].__path__`
     (win32com/__init__.py builds that synthetic module from __gen_path__ at
     import time, and imports resolve through its __path__, not the string).

  2. SELF-HEAL -- `dispatch_with_repair()` catches the corruption signature,
     regenerates the type library from the live object's own typelib, and
     retries once. A wiped cache becomes a slow first call instead of a
     confusing outage.

Import this BEFORE win32com.client.gencache is used for generation.
Override the location with WZTC_GEN_PY_DIR if needed.
"""
from __future__ import annotations

import importlib
import os
import sys

import pythoncom
import win32com
import win32com.client.gencache as gencache
from win32com.client import makepy

# AttributeError text pywin32 raises when the cached package is present but
# gutted. Both spellings occur -- EnsureDispatch trips CLSIDToClassMap, plain
# Dispatch trips CLSIDToPackageMap.
_CORRUPT_MARKERS = ("CLSIDToClassMap", "CLSIDToPackageMap", "CLSIDToClassMap")


def default_cache_dir() -> str:
    """A stable per-user, per-Python-version cache directory outside %TEMP%."""
    override = os.environ.get("WZTC_GEN_PY_DIR", "").strip()
    if override:
        return override
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "microstation-vba-project", "gen_py",
                        f"{sys.version_info.major}.{sys.version_info.minor}")


def relocate_cache(path: str = "") -> str:
    """Point pywin32's generated-code cache at `path`. Returns the path used.

    Idempotent and safe to call from every entry point.
    """
    target = os.path.abspath(path or default_cache_dir())
    os.makedirs(target, exist_ok=True)
    if os.path.normcase(getattr(win32com, "__gen_path__", "")) == os.path.normcase(target):
        return target
    win32com.__gen_path__ = target
    # win32com/__init__.py synthesises this module with __path__ = [gen_path];
    # imports resolve through it, so updating the string alone is not enough.
    gen_py = sys.modules.get("win32com.gen_py")
    if gen_py is not None:
        gen_py.__path__ = [target]
    importlib.invalidate_caches()
    return target


def _typelib_spec(idisp):
    """(guid, lcid, major, minor) for the type library behind a live object."""
    ti = idisp.GetTypeInfo(0)
    tlb, _index = ti.GetContainingTypeLib()
    attr = tlb.GetLibAttr()
    guid, lcid, major, minor = attr[0], attr[1], attr[3], attr[4]
    return str(guid), int(lcid), int(major), int(minor)


def regenerate_typelib(guid: str, lcid: int, major: int, minor: int) -> bool:
    """Rewrite a type library's cached wrappers. Never deletes anything.

    Uses bForDemand=True so the PACKAGE layout (__init__.py + lazy submodules)
    is written into the existing directory -- a stale empty directory of the
    same name would otherwise shadow a single-file module of that name, since
    Python resolves packages before modules.
    """
    for name in [n for n in list(sys.modules) if guid.strip("{}").lower() in n.lower()]:
        del sys.modules[name]
    importlib.invalidate_caches()
    try:
        tlb = pythoncom.LoadRegTypeLib(guid, major, minor, lcid)
    except Exception:
        return False
    try:
        makepy.GenerateFromTypeLibSpec(tlb, bForDemand=True, bBuildHidden=True)
    except Exception:
        return False
    importlib.invalidate_caches()
    return True


def _looks_corrupt(err: BaseException) -> bool:
    return isinstance(err, AttributeError) and any(
        marker in str(err) for marker in _CORRUPT_MARKERS)


def dispatch_with_repair(idisp):
    """gencache.EnsureDispatch, regenerating the cache once if it is gutted.

    Raises the ORIGINAL error if repair does not help, so a genuinely
    different COM problem is never disguised as a cache problem.
    """
    try:
        return gencache.EnsureDispatch(idisp)
    except Exception as first:
        if not _looks_corrupt(first):
            raise
        try:
            spec = _typelib_spec(idisp)
        except Exception:
            raise first
        if not regenerate_typelib(*spec):
            raise first
        return gencache.EnsureDispatch(idisp)


# Relocating on import is the point -- every consumer goes through ms_connect,
# which imports this first, so the cache is off %TEMP% before anything
# generates into it.
CACHE_DIR = relocate_cache()
