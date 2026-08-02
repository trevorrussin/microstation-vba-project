"""
Deterministic MicroStation COM attachment.

Replaces the ambiguous `win32com.client.GetObject(Class="MicroStationDGN.
Application")` / `gencache.EnsureDispatch("MicroStationDGN.Application")`
pattern used throughout this repo's tooling (bridge_client.py, hot_reload.py,
view_capture.py, scripts/keyin_batch.py, scripts/recipe_batch.py) before
this module existed. Both of those attach to WHICHEVER MicroStation instance
happens to be first in the Running Object Table -- if more than one is
running, which one you get is undefined.

This is not a hypothetical risk: it's the confirmed likely trigger for a
real MicroStation crash on 2026-08-02 -- a test launch + screenshot landed
on an ambiguous instance while two MicroStation processes were running, and
MicroStation went down shortly after (see Claude Code memory
feedback_microstation_crash_two_instances.md for the full incident).

get_microstation_app() enumerates every MicroStationDGN.Application entry
actually registered in the ROT and matches by which one has the target VBA
project loaded (the real invariant every caller in this toolchain depends
on -- WZTCBridge, hot-reload, and view navigation only work in the specific
session that has it loaded; filtering by open *file* isn't reliable, since
a real engineering session could have any file open, not just the
disposable DELETE.dgn this repo's own testing uses). Raises RuntimeError --
never silently guesses -- if zero or more than one instance qualifies,
matching the "immediate stop, don't probe further" rule from that incident.

Only the single-instance path (the overwhelmingly common case) has been
exercised live. The zero-match and multiple-match branches are constructed
from documented ROT/COM semantics but deliberately NOT live-tested against
a real second MicroStation instance -- reproducing that scenario on purpose
is exactly the situation that caused the crash this module exists to guard
against. If either branch's error message ever looks wrong in practice,
that's real signal to revisit this function, not a sign to force a second
instance for a test run.
"""
from __future__ import annotations

import winreg

import pythoncom
import win32com.client
import win32com.client.gencache as gencache

DEFAULT_PROJECT_NAME = "Test"
DEFAULT_PROG_ID = "MicroStationDGN.Application"


def _clsid_for_progid(prog_id: str) -> str:
    """MicroStation registers itself in the Running Object Table under a
    class moniker keyed by CLSID (display name like "!{6BA41DED-...}"), NOT
    a friendly ProgID-based name moniker -- confirmed live 2026-08-02 after
    an earlier version of this function searched for "microstationdgn.
    application" in moniker names and matched nothing at all, despite
    MicroStation actually being registered (its CLSID string was sitting
    right there in the enumeration, unrecognized). Reading the CLSID from
    the registry (HKCR\\<prog_id>\\CLSID) is what GetObject(Class=prog_id)
    does internally too -- this just exposes it so multiple ROT entries for
    the same CLSID can be enumerated instead of GetObject's implicit
    pick-one behavior."""
    key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{prog_id}\\CLSID")
    try:
        clsid_str, _ = winreg.QueryValueEx(key, "")
        return clsid_str
    finally:
        winreg.CloseKey(key)


def _candidate_apps(prog_id: str = DEFAULT_PROG_ID):
    """Yields (open_file_description, app) for every distinct instance of
    prog_id currently registered in the Running Object Table. A moniker
    that fails to bind (a stale/dying process still lingering in the ROT)
    is skipped rather than raised -- that's a normal, harmless occurrence,
    not something callers need to handle."""
    target = _clsid_for_progid(prog_id).lower()

    rot = pythoncom.GetRunningObjectTable()
    bind_ctx = pythoncom.CreateBindCtx(0)
    enum_moniker = rot.EnumRunning()
    while True:
        monikers = enum_moniker.Next(1)
        if not monikers:
            return
        moniker = monikers[0]
        try:
            display_name = moniker.GetDisplayName(bind_ctx, None)
        except Exception:
            continue
        if target not in display_name.lower():
            continue
        try:
            unknown = rot.GetObject(moniker)
            # rot.GetObject returns a plain PyIUnknown -- QueryInterface to
            # IDispatch first (confirmed live: passing PyIUnknown straight
            # to Dispatch()/EnsureDispatch() fails with "no attribute
            # GetTypeInfo"), then gencache.EnsureDispatch (not plain
            # Dispatch) so early-bound type info is available -- callers
            # like view_capture.navigate_view() pass Point3d structs, which
            # fail under plain dynamic dispatch (see Claude Code memory
            # feedback_vba_compile_error_recovery.md's "com_record" note).
            idisp = unknown.QueryInterface(pythoncom.IID_IDispatch)
            app = gencache.EnsureDispatch(idisp)
        except Exception:
            continue

        try:
            open_file = app.ActiveDesignFile.FullName
        except Exception:
            open_file = "(could not read ActiveDesignFile)"
        yield open_file, app


def get_microstation_app(project_name: str = DEFAULT_PROJECT_NAME):
    """Returns the MicroStation Application COM object for the ONE running
    instance that has project_name's VBA project loaded (checked via
    app.VBE.VBProjects(project_name), the same existence check hot_reload.py
    already used before this module existed). Raises RuntimeError if zero or
    more than one instance qualifies, listing every open file found across
    all running instances so the caller can tell the user exactly what's
    running and what to close."""
    matches = []
    diagnostics = []
    for open_file, app in _candidate_apps():
        try:
            app.VBE.VBProjects(project_name)
            matches.append(app)
            diagnostics.append(f"{open_file} -- HAS {project_name!r} loaded")
        except Exception:
            diagnostics.append(f"{open_file} -- does not have {project_name!r} loaded")

    if len(matches) == 1:
        return matches[0]

    found = "\n  ".join(diagnostics) if diagnostics else "(no MicroStation instances found running at all)"
    if len(matches) == 0:
        raise RuntimeError(
            f"No running MicroStation instance has the {project_name!r} VBA project "
            f"loaded. Instances found:\n  {found}"
        )
    raise RuntimeError(
        f"{len(matches)} running MicroStation instances all have {project_name!r} "
        f"loaded -- can't tell which one to use. Close the extra instance(s) first. "
        f"Instances found:\n  {found}"
    )
