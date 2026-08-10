"""Continue the interrupted 619-311 live test after history repair."""
from __future__ import annotations

import json
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "mcp-server"))
import chat_history  # noqa: E402

BRIDGE = Path(__file__).resolve().parent.parent / "Bridge"
HIST = BRIDGE / "chat-history.json"
CHAT_INPUT = BRIDGE / "chat-input.tsv"
CHAT_LOG = BRIDGE / "chat-log.tsv"

UP = [1020224.49, 218516.28, 0.0]
DN = [1020079.64, 218512.19, 0.0]


def append_input(message: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
    flat = message.replace("\t", " ").replace("\r\n", " ").replace("\n", " ")
    with open(CHAT_INPUT, "a", encoding="utf-8", newline="") as f:
        f.write(f"{ts}\t{flat}\n")
    print(f"[INPUT] {flat[:240]}", flush=True)


def main() -> int:
    bak = BRIDGE / (
        "chat-history.pre-continue."
        + datetime.now().strftime("%Y%m%d-%H%M%S")
        + ".bak.json"
    )
    if HIST.exists():
        shutil.copy2(HIST, bak)
        msgs = json.loads(HIST.read_text(encoding="utf-8"))
        chat_history._repair_tool_pairing(msgs)
        # Keep locked-plan context if present; else clear.
        # Safer for this continue: clear and nudge with full instruction.
        HIST.write_text("[]\n", encoding="utf-8")
        print("cleared history after repair", flush=True)

    start = sum(
        1 for ln in CHAT_LOG.read_text(encoding="utf-8", errors="replace").splitlines()
        if ln.strip()
    )
    append_input(
        "Continue 619-311 sheet build. Designer inputs already known: "
        "speed=45, Non-Freeway, lane=12, shoulder=8 ft, area_type=RURAL. "
        "If no order table this session, build_wztc_order_table with those. "
        "Then IMMEDIATELY call run_sheet_build("
        f"upstream_edge={UP}, downstream_edge={DN}). "
        "Do NOT ask for edge picks. Do NOT fish Default linework. "
        "After run_sheet_build finishes, FINAL with brief status."
    )

    seen = 0
    t0 = time.time()
    saw_rsb = False
    while time.time() - t0 < 600:
        lines = [
            ln for ln in CHAT_LOG.read_text(encoding="utf-8", errors="replace").splitlines()
            if ln.strip()
        ][start:]
        while seen < len(lines):
            line = lines[seen]
            seen += 1
            parts = line.split("\t")
            typ = parts[1] if len(parts) > 1 else "?"
            print(
                f"[{typ}] " + "\t".join(parts[1:3])[:220].encode("ascii", "replace").decode(),
                flush=True,
            )
            if typ == "TOOL_CALL" and "run_sheet_build" in line:
                saw_rsb = True
            if typ == "ERROR":
                print("[FAIL] ERROR", flush=True)
                return 2
            if typ in ("ASK_USER", "ASK_USER_CHOICE"):
                # Refuse re-asking edges — push executor again
                time.sleep(0.5)
                start = sum(
                    1 for ln in CHAT_LOG.read_text(encoding="utf-8", errors="replace").splitlines()
                    if ln.strip()
                )
                seen = 0
                append_input(
                    f"Do not ask again. Call run_sheet_build("
                    f"upstream_edge={UP}, downstream_edge={DN}) now."
                )
                break
            if typ == "FINAL":
                print(f"[DONE] run_sheet_build={saw_rsb}", flush=True)
                return 0 if saw_rsb else 1
        else:
            time.sleep(0.5)
    print("[FAIL] timeout", flush=True)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
