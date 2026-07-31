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
import time
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

BRIDGE_DIR = Path(r"c:\repos\microstation-vba-project\Bridge")
REQUEST_FILE = BRIDGE_DIR / "request.tsv"
RESPONSE_FILE = BRIDGE_DIR / "response.tsv"

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
        REQUEST_FILE.write_text("\n".join(lines) + "\n", encoding="utf-8")

        # COM objects are apartment-threaded: a handle obtained on one thread
        # cannot be used from another without marshaling. The MCP SDK runs
        # each synchronous tool call via a worker-thread dispatch, which is
        # not guaranteed to be the same OS thread call to call (confirmed by
        # testing: caching one COM handle at server startup and reusing it
        # here raised "CoInitialize has not been called" on the first real
        # tool call). So this initializes COM and re-attaches fresh on
        # whichever thread is actually calling, every time, rather than
        # caching anything across calls.
        pythoncom.CoInitialize()
        try:
            app = win32com.client.GetObject(Class="MicroStationDGN.Application")
            keyin = f"VBA RUN [{PROJECT_NAME}]WZTCBridge.RunRequest"
            app.CadInputQueue.SendKeyin(keyin)
        finally:
            pythoncom.CoUninitialize()

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

    def _read_response_with_retry(self, req_ids: list[str], timeout_s: float = 5.0) -> str:
        deadline = time.time() + timeout_s
        text = RESPONSE_FILE.read_text(encoding="utf-8", errors="replace")
        if all(rid in text for rid in req_ids):
            return text
        while time.time() < deadline:
            time.sleep(0.1)
            text = RESPONSE_FILE.read_text(encoding="utf-8", errors="replace")
            if all(rid in text for rid in req_ids):
                return text
        raise BridgeError(
            f"response.tsv never contained all reqIds {req_ids} within {timeout_s}s "
            "of sending the keyin — MicroStation may be busy, closed, or the "
            "keyin trigger may not be synchronous under this condition (see "
            "Layer 5 risks in the plan)."
        )


# Module-level singleton. Holds no COM state itself (see call_batch) — it's
# just a convenient shared instance, not a cached connection.
bridge = Bridge()
