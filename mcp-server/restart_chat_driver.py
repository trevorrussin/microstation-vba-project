"""
Safely restart the chat_driver.py process so Python-side code changes
(wztc_ops.py, server.py's tool wrappers, chat_driver.py itself) take
effect -- mirrors hot_reload.py's role for VBA changes, for the one kind
of update hot-reload explicitly cannot cover (per CLAUDE.md's File Sync
Protocol: "mcp-server/*.py changes are not picked up by hot-reload --
restart the Python process instead").

Safety rule this script exists to enforce: NEVER kill chat_driver.py
while a turn (or a pending ask_user_choice point-pick) looks in
progress -- that would strand whatever the engineer or the agent was
mid-way through. Restarting is only safe once the main loop has
returned to its blocking wait_for_next() call, which the log makes
observable: chat-log.tsv's LAST line is always either FINAL (a turn
completed normally) or ERROR (a turn failed and was caught) right
before the loop goes back to waiting -- anything else (THINKING,
TOOL_CALL, TOOL_RESULT, ASK_USER_CHOICE, MODE_CHANGED as the latest
line) means something is still active. This script refuses to proceed
unless the last line is FINAL/ERROR, or the log is empty/missing
(nothing has ever run).

Usage:
    python restart_chat_driver.py
    python restart_chat_driver.py --force   # skip the idle check (only
                                             # if you've confirmed by
                                             # other means it's safe --
                                             # e.g. the process already
                                             # isn't running at all)
"""
from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path

BRIDGE_DIR = Path(r"c:\repos\microstation-vba-project\Bridge")
CHAT_LOG_FILE = BRIDGE_DIR / "chat-log.tsv"
MCP_DIR = Path(r"c:\repos\microstation-vba-project\mcp-server")

# Matches the interpreter chat_driver.py has actually been launched with
# every time observed live this session (distinct from server.py's,
# which resolves via plain "python" instead) -- hardcoded rather than
# re-derived from the running process, so a restart works even when
# nothing is currently running.
PYTHON_EXE = r"C:\Users\RussinT\AppData\Local\Programs\Python\Python312\python.exe"

IDLE_LOG_TYPES = {"FINAL", "ERROR"}


def _run_ps(command: str, timeout: int = 15) -> str:
    result = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
        capture_output=True, text=True, timeout=timeout,
    )
    return result.stdout.strip()


def last_log_type() -> str | None:
    """Returns the TYPE field of chat-log.tsv's last non-blank line, or
    None if the file is missing/empty (nothing has ever run -- treated
    as idle, not as a reason to refuse)."""
    if not CHAT_LOG_FILE.exists():
        return None
    with open(CHAT_LOG_FILE, "r", encoding="utf-8", errors="replace") as f:
        lines = [ln for ln in f if ln.strip()]
    if not lines:
        return None
    parts = lines[-1].split("\t")
    return parts[1] if len(parts) > 1 else None


def find_chat_driver_pids() -> list[int]:
    """Returns EVERY matching PID, not just one -- a naive single-value
    parse (e.g. str.isdigit() on the raw output) silently returns
    nothing at all when two-or-more processes match, which is exactly
    the case this function most needs to catch reliably (confirmed
    live 2026-08-02: that exact bug made this script blind to an
    already-running process and launch a duplicate instead of stopping
    it -- see the restart flow below, which now refuses rather than
    picking one arbitrarily when more than one PID comes back).

    The match pattern requires 'chat_driver.py' to be preceded by
    start-of-string, whitespace, or a path separator -- a plain
    substring match also caught THIS SCRIPT's own process (python
    restart_chat_driver.py), since 'restart_chat_driver.py' contains
    'chat_driver.py' as a literal substring. That self-match produced a
    repeatable phantom "duplicate" every single time this script ran
    (confirmed live 2026-08-02 -- three false alarms in a row before the
    pattern was traced to this). \b alone does not fix it: '_' is a
    word character, so there is no regex word boundary between
    'restart_' and 'chat_driver.py' either."""
    out = _run_ps(
        "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | "
        "Where-Object { $_.CommandLine -match '(^|[\\\\/ ])chat_driver\\.py' } | "
        "Select-Object -ExpandProperty ProcessId"
    )
    return [int(line) for line in out.splitlines() if line.strip().isdigit()]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--force", action="store_true", help="skip the idle check")
    args = parser.parse_args()

    if not args.force:
        last_type = last_log_type()
        if last_type is not None and last_type not in IDLE_LOG_TYPES:
            print(
                f"REFUSED: chat-log.tsv's last entry is {last_type!r}, not FINAL/ERROR -- "
                "a turn or a pending ask_user_choice point-pick looks in progress. "
                "Not restarting (pass --force to override once you've confirmed by other "
                "means it's actually safe).",
                file=sys.stderr,
            )
            return 1
        print(f"Idle check passed (last log entry: {last_type!r}).")

    old_pids = find_chat_driver_pids()
    if len(old_pids) > 1:
        print(
            f"REFUSED: {len(old_pids)} chat_driver.py processes already running (PIDs {old_pids}) -- "
            "an unexpected duplicate is a stop-and-check situation, not something to resolve by "
            "picking one automatically. Stop the extras by hand, confirm exactly one (or zero) "
            "remain, then re-run this script.",
            file=sys.stderr,
        )
        return 1
    if old_pids:
        old_pid = old_pids[0]
        print(f"Stopping chat_driver.py (PID {old_pid})...")
        _run_ps(f"Stop-Process -Id {old_pid} -Force -ErrorAction SilentlyContinue")
        time.sleep(1)
        if old_pid in find_chat_driver_pids():
            print(f"ERROR: PID {old_pid} did not stop.", file=sys.stderr)
            return 1
    else:
        print("No running chat_driver.py found -- starting fresh.")

    print("Starting new chat_driver.py in a new visible console window...")
    # CREATE_NEW_CONSOLE (not CREATE_NO_WINDOW): the engineer's normal
    # workflow is a terminal window they can see output in and Ctrl+C by
    # hand -- an earlier version of this script launched fully hidden
    # (no window at all), which left a genuinely healthy process running
    # with no visible way for the engineer to check on or interrupt it,
    # and looked exactly like "nothing is running" from their side
    # (confirmed live 2026-08-02). stdout is NOT redirected to a file
    # here for the same reason -- it needs to land in that visible
    # window, not be captured away from it.
    subprocess.Popen(
        [PYTHON_EXE, "chat_driver.py"],
        cwd=str(MCP_DIR),
        creationflags=subprocess.CREATE_NEW_CONSOLE,
    )

    time.sleep(3)
    new_pids = find_chat_driver_pids()
    if len(new_pids) == 1:
        print(f"OK: chat_driver.py running (PID {new_pids[0]}) in its own console window.")
        return 0
    if len(new_pids) > 1:
        print(f"ERROR: {len(new_pids)} chat_driver.py processes running after restart (PIDs {new_pids}) "
              "-- the old one may not have actually stopped. Check manually.", file=sys.stderr)
        return 1
    print("WARNING: could not confirm the new process started -- check the new console window directly.",
          file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
