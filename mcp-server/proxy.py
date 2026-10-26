"""
Stable stdio proxy in front of server.py (wztc-designer MCP server).

Claude Code owns the MCP process and cannot be told to restart it, so tool
registration changes in server.py / code_reload.py needed an engineer-run
/mcp reconnect. This proxy is the process Claude Code owns. It runs
server.py as a child and forwards every JSON-RPC line both ways. When a
watched file changes it syntax-checks it, restarts the child, replays the
MCP handshake to the new child, and sends notifications/tools/list_changed
so Claude Code re-reads the tool list without a reconnect.

Watched files: server.py, code_reload.py (reload_code covers the rest).
Add more with WZTC_PROXY_WATCH="a.py,b.py". Syntax error -> keep the old
child, report on stderr, retry only after the file changes again.
Manual restart: touch mcp-server/.proxy-restart.
Requests in flight when the child restarts get a JSON-RPC error reply.
Change to proxy.py itself needs one /mcp reconnect.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time

HERE = os.path.dirname(os.path.abspath(__file__))
SERVER = os.path.join(HERE, "server.py")
FLAG = os.path.join(HERE, ".proxy-restart")
WATCH = [s.strip() for s in os.environ.get("WZTC_PROXY_WATCH", "server.py,code_reload.py").split(",") if s.strip()]

out_lock = threading.Lock()      # serialises writes to Claude Code
child_lock = threading.RLock()   # guards child + handshake state
child: subprocess.Popen | None = None
init_req: str | None = None
init_id = None
initialized_note: str | None = None
pending: set = set()             # ids of client requests awaiting a child reply
swallow_id = None                # id of the replayed initialize whose reply is dropped
swallow_evt = threading.Event()
retired: set = set()             # children the proxy stopped on purpose
last_start = 0.0
quick_deaths = 0                 # consecutive children that died within 5 s of starting


def log(msg: str) -> None:
    sys.stderr.write(f"[wztc-proxy] {msg}\n")
    sys.stderr.flush()


def to_client(line: str) -> None:
    with out_lock:
        sys.stdout.write(line if line.endswith("\n") else line + "\n")
        sys.stdout.flush()


def to_child(line: str) -> None:
    with child_lock:
        try:
            child.stdin.write(line if line.endswith("\n") else line + "\n")
            child.stdin.flush()
        except Exception as e:  # child died; pump_child restarts it
            log(f"write to child failed: {e}")
            try:
                msg = json.loads(line)
            except Exception:
                return
            rid = json.dumps(msg.get("id"))
            if msg.get("method") and rid in pending:
                pending.discard(rid)
                to_client(json.dumps({"jsonrpc": "2.0", "id": msg["id"],
                                      "error": {"code": -32603, "message": "server not running; retry the call"}}))


def start_child() -> None:
    global child, last_start
    last_start = time.time()
    child = subprocess.Popen(
        [sys.executable, SERVER], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=None, text=True, encoding="utf-8", bufsize=1, cwd=HERE,
    )
    threading.Thread(target=pump_child, args=(child,), daemon=True).start()


def pump_child(proc: subprocess.Popen) -> None:
    global swallow_id
    for line in proc.stdout:
        if proc is not child:
            return
        try:
            msg = json.loads(line)
        except Exception:
            to_client(line)
            continue
        mid = msg.get("id")
        if "method" not in msg and mid is not None:
            if swallow_id is not None and mid == swallow_id:
                swallow_id = None
                swallow_evt.set()
                continue
            with child_lock:
                pending.discard(json.dumps(mid))
            if mid == init_id and "result" in msg:
                caps = msg["result"].setdefault("capabilities", {})
                caps.setdefault("tools", {})["listChanged"] = True
                line = json.dumps(msg)
        to_client(line)
    if proc is child and proc not in retired:
        global quick_deaths
        quick_deaths = quick_deaths + 1 if time.time() - last_start < 5 else 0
        log("child exited on its own")
        if quick_deaths >= 3:
            log("child died 3 times within 5 s of starting; not restarting until a watched file changes")
            return
        threading.Thread(target=_safe_restart, args=("child exited",), daemon=True).start()


def _safe_restart(reason: str) -> None:
    try:
        restart(reason)
    except Exception as e:
        log(f"restart failed: {e}")


def restart(reason: str) -> None:
    global swallow_id, quick_deaths
    with child_lock:
        log(f"restarting child: {reason}")
        if reason != "child exited":
            quick_deaths = 0
        old = child
        if old is not None:
            retired.add(old)
        for rid in list(pending):
            to_client(json.dumps({"jsonrpc": "2.0", "id": json.loads(rid),
                                  "error": {"code": -32603, "message": "server restarting; retry the call"}}))
        pending.clear()
        if old is not None:
            try:
                old.stdin.close()
                old.terminate()
                old.wait(timeout=5)
            except Exception:
                old.kill()
        start_child()
        if init_req is not None:
            swallow_evt.clear()
            swallow_id = init_id
            to_child(init_req)
            if not swallow_evt.wait(timeout=30):
                log("new child did not answer initialize within 30 s")
                return
            if initialized_note:
                to_child(initialized_note)
            to_client(json.dumps({"jsonrpc": "2.0", "method": "notifications/tools/list_changed"}))
        log("restart complete")


def mtimes() -> dict:
    d = {}
    for n in WATCH:
        try:
            d[n] = os.stat(os.path.join(HERE, n)).st_mtime
        except OSError:
            pass
    return d


def watcher() -> None:
    seen = mtimes()
    while True:
        time.sleep(1.0)
        reason = None
        if os.path.exists(FLAG):
            try:
                os.remove(FLAG)
            except OSError:
                pass
            reason = "manual flag"
        cur = mtimes()
        changed = [n for n in cur if cur[n] != seen.get(n)]
        if changed:
            seen = cur
            bad = None
            for n in changed:
                try:
                    with open(os.path.join(HERE, n), "rb") as f:
                        compile(f.read(), n, "exec")
                except Exception as e:
                    bad = f"{n}: {e}"
                    break
            if bad:
                log(f"syntax error, keeping old child: {bad}")
                changed = []
            elif reason is None:
                reason = "changed: " + ", ".join(changed)
        if reason:
            try:
                restart(reason)
            except Exception as e:
                log(f"restart failed: {e}")


def main() -> None:
    global init_req, init_id, initialized_note
    sys.stdin.reconfigure(encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8", newline="\n")
    start_child()
    threading.Thread(target=watcher, daemon=True).start()
    for line in sys.stdin:
        try:
            msg = json.loads(line)
        except Exception:
            to_child(line)
            continue
        method = msg.get("method")
        mid = msg.get("id")
        with child_lock:
            if method == "initialize":
                init_req, init_id = line, mid
            elif method == "notifications/initialized":
                initialized_note = line
            if method and mid is not None:
                pending.add(json.dumps(mid))
            to_child(line)
    with child_lock:
        try:
            child.stdin.close()
            child.terminate()
        except Exception:
            pass


if __name__ == "__main__":
    main()
