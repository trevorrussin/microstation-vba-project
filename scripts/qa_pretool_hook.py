"""PreToolUse hook (Bash and PowerShell): refuse a Python script that changes MicroStation through COM when no QA task is open.

The bridge guard (mcp-server/task_guard.py) only sees edits that go through bridge_client. A Python script that attaches
with ms_connect / win32com and calls SendKeyin, HideLevel, DisplayFlag ... edits the drawing without passing it (the
2026-10-24 circles run did exactly that, ~15 times). This hook reads the command and any .py file it names; if the code
attaches to MicroStation AND contains a mutating call, the same rules as the bridge apply: an open cad.drafting task,
research read, failed-attempt budget not spent.

It is a static check on text, so it can be evaded (code built at run time, a module that imports the mutating code).
It exists to stop the accidental case, which is the case that happened. Read-only COM scripts pass.
Maintenance scripts (hot reload, VBA recovery, QA gate, playbook runner) are exempt; scratch scripts that are not drawing
work put QA_ALLOW_NO_TASK=1 in the command.

Claude Code (.claude/settings.json "PreToolUse", matcher Bash|PowerShell): stdin JSON, exit 2 + stderr = refused.
"""
from __future__ import annotations

import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "mcp-server"))

# Real usage only (an import, a call, an assignment): a regex or a comment that merely names these words is not an attach or an edit.
ATTACH = re.compile(r"(?:from|import)\s+ms_connect|get_microstation_app\s*\(|win32com\.client|GetActiveObject\s*\(|"
                    r"Dispatch\s*\(\s*[\"']MicroStationDGN", re.I)
MUTATE = re.compile(
    r"\.SendKeyin\s*\(|\.SendCommand\s*\(|\.CadInputQueue\b|\.HideLevel\s*\(|\.ShowLevel\s*\(|\.IsDisplayed\s*=[^=]|\.IsFrozen\s*=[^=]|"
    r"\.IsLocked\s*=[^=]|\.DisplayFlag\w*\s*=[^=]|\.Display\s*=[^=]|\.Displays?\w*\s*=[^=]|\.Delete\w*\s*\(|\.AddElement\s*\(|"
    r"\.Rewrite\s*\(|\.Move\w*\s*\(|\.SaveSettings\s*\(|\.Save\s*\(|\.SaveAs\s*\(|\.ActiveLevel\s*=[^=]|\.ActiveColor\s*=[^=]", re.I)
EXEMPT = {"hot_reload.py", "vba_compile_recovery.py", "diag_vba_ping.py", "qa_gate.py", "playbook_runner.py",
          "view_capture.py", "restart_chat_driver.py", "qa_pretool_hook.py", "qa_stop_hook.py"}
MAX_READ = 400_000


_PYRUN = re.compile(r"\b(?:python[\d.]*(?:\.exe)?|py)\b[ \t]+([^;&|\r\n]*)", re.I)


def _executed(command: str, cwd: str) -> tuple[list[str], bool]:
    """-> (script files a python invocation in the command runs, whether the command itself is run as code: -c or stdin).
    A command that only names a .py file (grep, cat, git diff) runs nothing."""
    out, inline = [], False
    for m in _PYRUN.finditer(command):
        args = m.group(1).strip()
        if args.startswith("-c") or args == "-" or args.startswith("- ") or args.startswith("-<") or args.startswith("-\t"):
            inline = True
            continue
        if args.startswith("-m"):
            continue
        for tok in re.findall(r"[\w./\\:~-]+\.py\b", args):
            for base in (cwd, ROOT, os.path.join(ROOT, "scripts")):
                p = tok if os.path.isabs(tok) else os.path.join(base, tok)
                if os.path.isfile(p):
                    out.append(p)
                    break
            break  # the first .py is the script; the rest are its arguments
    return out, inline


def assess(command: str, cwd: str = "") -> str:
    """-> "" when the command may run, else the refusal message."""
    if not re.search(r"\bpython[\d.]*\b|\bpy\b|\.py\b", command, re.I) or "QA_ALLOW_NO_TASK=1" in command or "QA_ALLOW_NO_TASK = '1'" in command:
        return ""
    if re.search(r"(?:^|[;&|]\s*)(?:python[\d.]*\s+-m\s+)?pytest\b|(?:^|[;&|]\s*)python[\d.]*\s+-m\s+unittest\b", command):  # test runs scan fixtures that hold COM samples as strings
        return ""
    files, inline = _executed(command, cwd or os.getcwd())
    texts, names = ([command] if inline else []), []
    for p in files:
        if "/tests/" in p.replace("\\", "/"):
            continue
        names.append(os.path.basename(p))
        try:
            texts.append(open(p, encoding="utf-8", errors="replace").read(MAX_READ))
        except OSError:
            pass
    if names and all(n in EXEMPT for n in names):
        return ""
    blob = "\n".join(texts)
    if not (ATTACH.search(blob) and MUTATE.search(blob)):
        return ""
    import task_guard
    try:
        task_guard._cache["at"] = 0.0  # fresh read: a task opened a moment ago must count
        task_guard.require_open_task(["COM_SCRIPT_EDIT"])
    except task_guard.QAGateRequired as ex:
        what = ", ".join(sorted(set(m.group(0).strip() for m in MUTATE.finditer(blob)))[:5])
        return (f"This script changes MicroStation through COM ({what}) and the QA gate refuses it: {ex} "
                "Display and level changes are drawing state too. Read the playbook for the topic first "
                "(references-levels-models-measuring S3 for hidden or unwanted linework), and hand GUI steps (who: engineer) to the engineer.")
    return ""


def main() -> int:
    try:
        data = json.load(sys.stdin)
        cmd = (data.get("tool_input") or {}).get("command", "")
        msg = assess(cmd, data.get("cwd") or os.getcwd())
    except Exception:
        return 0  # a broken hook must never trap the session
    if msg:
        print(msg, file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
