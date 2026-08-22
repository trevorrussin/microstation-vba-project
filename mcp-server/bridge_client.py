"""
Low-level transport to WZTCBridge.bas.

Writes Bridge/request.tsv, triggers WZTCBridge.RunRequest via a COM-sent
keyin, reads Bridge/response.tsv, and resolves any resultFile= pointer into
parsed rows. This is the exact pipeline proved end-to-end in scratch scripts
during M1/M2 (com_trigger_test.py, com_query_test.py) — same file paths,
same keyin form, same project name — just packaged for reuse instead of
being copy-pasted per test.

Confirmed synchronous against this install: app.CadInputQueue.SendKeyin(...)
does not return until WZTCBridge.RunRequest has finished and response.tsv is
already updated (~0.4s round trip for a single op). The retry loop below is
kept as a safety net rather than a requirement — if a future batch or a
slower op ever breaks that synchronicity, this fails loudly instead of
silently.
"""
from __future__ import annotations

import itertools
import threading
import time
from pathlib import Path
from typing import Any

import pythoncom

import ms_connect

# Per-thread cached MicroStation COM handle.
#
# Measured live 2026-08-20 on this install, per bridge round trip:
#     ms_connect.get_microstation_app()  ~680 ms   <-- was paid EVERY call
#       of which app.VBE.VBProjects()    ~500 ms
#     app.CadInputQueue.SendKeyin()      ~100 ms
# So ~87% of a small op's wall time was re-proving which MicroStation to talk
# to, not doing the work. A 309-element road ran ~1248 s largely on this.
#
# The module docstring's note that caching "raised CoInitialize has not been
# called" was right about the cause but not the fix: COM apartments are
# per-thread, so ONE shared handle breaks when the MCP SDK dispatches the
# next call on a different worker thread. Keying the cache on the thread
# fixes that without marshalling — each thread initializes its own apartment
# once and reuses its own handle. Correspondingly we must NOT CoUninitialize
# after each call any more; that would tear down the apartment the cached
# handle lives in.
#
# Staleness (MicroStation closed/restarted) is handled by letting the keyin
# fail and re-attaching once, rather than probing before every call — even
# the cheapest liveness probe measured 65-75 ms, which would give back a
# tenth of the win on every single op.
_tls = threading.local()


def _attach_app():
    """Fresh attach + apartment init for the CURRENT thread."""
    pythoncom.CoInitialize()
    app = ms_connect.get_microstation_app(PROJECT_NAME)
    _tls.app = app
    return app


def _thread_app():
    app = getattr(_tls, "app", None)
    return app if app is not None else _attach_app()


def _send_keyin(keyin: str) -> None:
    """SendKeyin on the cached handle, re-attaching once if it went stale."""
    try:
        _thread_app().CadInputQueue.SendKeyin(keyin)
    except Exception:
        # Stale handle (MicroStation restarted, apartment torn down, RPC
        # server gone). Drop it and try exactly once more with a fresh
        # attach; a second failure is a real error and propagates.
        _tls.app = None
        _attach_app().CadInputQueue.SendKeyin(keyin)

BRIDGE_DIR = Path(r"c:\repos\microstation-vba-project\Bridge")
REQUEST_FILE = BRIDGE_DIR / "request.tsv"
RESPONSE_FILE = BRIDGE_DIR / "response.tsv"
RUN_SUB = "WZTCBridge.RunRequest"

# M7 — chat_driver.py's own file pair + VBA entry point (WZTCBridge.
# RunChatToolRequest), so it doesn't race the module-level `bridge`
# singleton below on the same request.tsv/response.tsv.
CHAT_TOOL_REQUEST_FILE = BRIDGE_DIR / "chat-tool-request.tsv"
CHAT_TOOL_RESPONSE_FILE = BRIDGE_DIR / "chat-tool-response.tsv"
CHAT_TOOL_RUN_SUB = "WZTCBridge.RunChatToolRequest"

# Confirmed working VBA project name against this MicroStation install
# (see scratchpad com_trigger_test.py). If VBA project files are ever
# renamed/reorganized, this is the one place to update.
PROJECT_NAME = "Test"

_req_counter = itertools.count(1)


class BridgeError(RuntimeError):
    """Raised for transport-level failures (no response, malformed line,
    COM attach failure) — distinct from an ERROR status the bridge itself
    returned for a well-formed request, which callers see as a normal
    {"status": "ERROR", "note": ...} dict instead of an exception."""


def _next_req_id() -> str:
    return f"P{next(_req_counter)}"


def _response_has_req_ids(text: str, req_ids: list[str]) -> bool:
    """Match whole reqId tokens. 'P1' must not match leftover 'P12490'."""
    heads = {ln.split("\t", 1)[0] for ln in text.splitlines() if ln.strip()}
    return all(rid in heads for rid in req_ids)


def _parse_response_line(line: str) -> dict[str, Any]:
    parts = line.split("\t")
    if len(parts) < 2:
        raise BridgeError(f"malformed response line: {line!r}")
    fields: dict[str, Any] = {"reqId": parts[0], "status": parts[1]}
    for kv in parts[2:]:
        if "=" in kv:
            k, v = kv.split("=", 1)
            fields[k] = v
    return fields


def _read_result_rows(result_file: str) -> list[dict[str, str]]:
    text = Path(result_file).read_text(encoding="utf-8", errors="replace")
    lines = [l for l in text.splitlines() if l.strip()]
    if not lines:
        return []
    header = lines[0].split("\t")
    rows = []
    for line in lines[1:]:
        # Bounded split: a data row can legitimately have MORE tabs than the
        # header declares (GET_JOURNAL's single "line" column holds a raw
        # multi-field journal entry; LIST_DEFERRED_HANDOFFS's "detail"
        # column holds several key=val pairs). An unbounded split() zipped
        # by index silently drops everything past the header's column
        # count instead of keeping it in the last column — confirmed live,
        # this is why get_journal() returned bare timestamps with the rest
        # of each line truncated away.
        cols = line.split("\t", len(header) - 1)
        rows.append({header[i]: (cols[i] if i < len(cols) else "") for i in range(len(header))})
    return rows


class Bridge:
    def __init__(self, request_file: Path = REQUEST_FILE, response_file: Path = RESPONSE_FILE,
                 run_sub: str = RUN_SUB):
        """Defaults are the module-level paths/Sub every caller has always
        used — the module-level `bridge` singleton below is unchanged by
        this constructor existing. A second instance with its own file pair
        + Sub (e.g. M7's chat_driver.py holding its own COM/keyin connection
        alongside this server's, via WZTCBridge.RunChatToolRequest) avoids
        two processes racing on the same request.tsv/response.tsv, since
        each process's reqId counter independently starts at P1 (see
        _next_req_id) — without separate files, two concurrent processes
        could both write reqId P1 to the same request.tsv. run_sub must
        name a VBA Sub that reads request_file/writes response_file using
        the exact same protocol as WZTCBridge.RunRequest (currently just
        RunRequest itself, or RunChatToolRequest for the chat-tool pair) —
        this class has no way to verify that at runtime, it just sends
        whatever Sub name it's given."""
        self.request_file = request_file
        self.response_file = response_file
        self.run_sub = run_sub

    def call(self, op_type: str, **params: Any) -> dict[str, Any]:
        """Send a single op, return its parsed response as a dict with at
        least {reqId, status, note?}. When the op returned a resultFile
        (multi-row query), it's resolved automatically into a 'rows' key —
        callers never need to open that file themselves."""
        return self.call_batch([(op_type, params)])[0]

    def call_batch(self, ops: list[tuple[str, dict[str, Any]]]) -> list[dict[str, Any]]:
        """Send several ops in one request/response round trip, in order.
        Each still gets its own reqId and journal entry on the VBA side —
        batching only saves keyin round trips, it does not change semantics."""
        req_ids = []
        lines = []
        for op_type, params in ops:
            req_id = _next_req_id()
            req_ids.append(req_id)
            kv = "\t".join(f"{k}={v}" for k, v in params.items() if v is not None and v != "")
            line = f"{req_id}\t{op_type}" + (f"\t{kv}" if kv else "")
            lines.append(line)

        # Text-mode write: Python translates \n -> \r\n on Windows, which is
        # required (VBA's Line Input # reads a bare-LF file as one giant line).
        self.request_file.write_text("\n".join(lines) + "\n", encoding="utf-8")

        # COM objects are apartment-threaded, so the handle is cached PER
        # THREAD (see _tls above) instead of re-attached every call. The
        # deterministic attach itself (ms_connect, 2026-08-02) is unchanged —
        # it still refuses to guess between multiple MicroStation instances;
        # it just isn't re-run for every op now.
        _send_keyin(f"VBA RUN [{PROJECT_NAME}]{self.run_sub}")

        resp_text = self._read_response_with_retry(req_ids)
        resp_lines = {l.split("\t", 1)[0]: l for l in resp_text.splitlines() if l.strip()}

        results = []
        for req_id in req_ids:
            line = resp_lines.get(req_id)
            if line is None:
                raise BridgeError(
                    f"no response for reqId {req_id} — is MicroStation open with "
                    f"the {PROJECT_NAME!r} VBA project loaded, and WZTCBridge imported?"
                )
            parsed = _parse_response_line(line)
            if "resultFile" in parsed:
                parsed["rows"] = _read_result_rows(parsed["resultFile"])
            results.append(parsed)
        return results

    def _read_response_with_retry(self, req_ids: list[str], timeout_s: float = 90.0) -> str:
        deadline = time.time() + timeout_s
        text = self.response_file.read_text(encoding="utf-8", errors="replace")
        if _response_has_req_ids(text, req_ids):
            return text
        # Adaptive backoff instead of a flat 0.1 s: SendKeyin is synchronous on
        # this install, so the response is usually already there and a fixed
        # 100 ms sleep was pure latency on every op that did miss the first
        # read. Start fine, widen for genuinely slow ops (big batches, hatch)
        # so a long wait doesn't spin the file read.
        delay = 0.005
        while time.time() < deadline:
            time.sleep(delay)
            text = self.response_file.read_text(encoding="utf-8", errors="replace")
            if _response_has_req_ids(text, req_ids):
                return text
            if delay < 0.1:
                delay = min(delay * 2, 0.1)
        raise BridgeError(
            f"response.tsv never contained all reqIds {req_ids} within {timeout_s}s "
            "of sending the keyin — MicroStation may be busy, closed, or the "
            "keyin trigger may not be synchronous under this condition (see "
            "Layer 5 risks in the plan)."
        )


# Module-level singleton. Holds no COM state itself (see call_batch) — it's
# just a convenient shared instance, not a cached connection.
bridge = Bridge()

# M7 — same, for chat_driver.py. Separate file pair + VBA Sub so it can run
# alongside `bridge` (used by this stdio MCP server) without file collisions.
chat_bridge = Bridge(CHAT_TOOL_REQUEST_FILE, CHAT_TOOL_RESPONSE_FILE, CHAT_TOOL_RUN_SUB)
