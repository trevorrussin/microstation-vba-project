"""
Hot-reload VBA source files' code directly into the live MicroStation VBA
IDE, in place -- no VBComponents.Remove/Import, no manual delete+reimport
in the IDE (the File Sync Protocol in CLAUDE.md).

Replaces an existing component's CodeModule text (VBIDE.CodeModule.
DeleteLines + AddFromString) via Application.VBE, confirmed as a supported
MicroStation VBA automation path per Bentley KB0026620. Editing in place
instead of Remove+Import sidesteps two known issues: VBComponents.Remove
not always taking effect until the calling procedure finishes (moot here
since this project's own code is never removed), and Import needing an
`Attribute VB_Name` line to resolve the component name (this repo omits
that attribute everywhere per CLAUDE.md, so Import would be unreliable
without extra parsing -- targeting an existing component by name sidesteps
that entirely).

Only touches an existing component's code. Does NOT:
  - create a new component (nothing to target yet for a brand-new file)
  - touch a UserForm's control layout (that's the Designer part of the
    VBComponent, separate from CodeModule -- untouched here)
Both of those still need the manual File Sync Protocol.

This repo's .bas/.frm/.cls files are plain code with no VERSION/Begin-End
designer header and no Attribute lines (confirmed by inspection), so the
whole file's text IS the component's code -- no parsing needed, unlike a
file freshly exported from the IDE.

Usage:
    python hot_reload.py Modules/SharedState.bas Modules/WZTCBridge.bas
    python hot_reload.py --project Test UserForms/WZTCChatPanel.frm

Requires MicroStation open with the target VBA project already loaded,
and (per Office's equivalent setting -- unconfirmed whether MicroStation
gates this the same way) "Trust access to the VBA project object model"
allowed. If that's the blocker, the COM error names it explicitly.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pythoncom
import win32com.client

from bridge_client import PROJECT_NAME


def reload_file(project, file_path: Path) -> None:
    component_name = file_path.stem
    try:
        component = project.VBComponents(component_name)
    except Exception as exc:
        raise RuntimeError(
            f"no component named {component_name!r} in VBA project "
            f"{project.Name!r} -- if this is a new file, it needs a manual "
            f"Import first (File Sync Protocol, CLAUDE.md); hot-reload only "
            f"updates existing components. ({exc})"
        ) from exc

    code_mod = component.CodeModule
    # rstrip: AddFromString counts a trailing "\n" in the string as one more
    # (empty) line than the file actually has -- confirmed live, a plain
    # read_text() round trip left CodeModule.CountOfLines one higher than
    # the file's line count.
    content = file_path.read_text(encoding="utf-8").rstrip("\r\n")

    if code_mod.CountOfLines > 0:
        code_mod.DeleteLines(1, code_mod.CountOfLines)
    code_mod.AddFromString(content)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("files", nargs="+", help="Path(s) to .bas/.cls/.frm source files to push live")
    parser.add_argument("--project", default=PROJECT_NAME, help=f"VBA project name (default: {PROJECT_NAME!r})")
    args = parser.parse_args()

    paths = [Path(f) for f in args.files]
    missing = [p for p in paths if not p.is_file()]
    if missing:
        for p in missing:
            print(f"ERROR: file not found: {p}", file=sys.stderr)
        return 1

    pythoncom.CoInitialize()
    try:
        try:
            app = win32com.client.GetObject(Class="MicroStationDGN.Application")
        except Exception as exc:
            print(
                f"ERROR: could not attach to a running MicroStation session: {exc}\n"
                "Is MicroStation open?",
                file=sys.stderr,
            )
            return 1

        try:
            vbe = app.VBE
        except Exception as exc:
            print(
                f"ERROR: Application.VBE not accessible: {exc}\n"
                "This is the same gate as Office's 'Trust access to the VBA "
                "project object model' -- check MicroStation's VBA Manager / "
                "VBA IDE Tools menu for an equivalent setting.",
                file=sys.stderr,
            )
            return 1

        try:
            project = vbe.VBProjects(args.project)
        except Exception as exc:
            print(
                f"ERROR: VBA project {args.project!r} is not loaded in the "
                f"running MicroStation session: {exc}",
                file=sys.stderr,
            )
            return 1

        exit_code = 0
        for path in paths:
            try:
                reload_file(project, path)
                print(f"OK: reloaded {path.stem} from {path}")
            except Exception as exc:
                print(f"ERROR: {path.stem}: {exc}", file=sys.stderr)
                exit_code = 1
        return exit_code
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
