"""One-shot live 619-311 regression: queue build, auto-answer known test
inputs/edges from the last successful afternoon run, print log until FINAL.
"""
from __future__ import annotations

import re
import time
from datetime import datetime
from pathlib import Path

BRIDGE = Path(__file__).resolve().parent.parent / "Bridge"
CHAT_INPUT = BRIDGE / "chat-input.tsv"
CHAT_LOG = BRIDGE / "chat-log.tsv"

# Last successful 619-311 live picks (2026-08-04 afternoon)
UPSTREAM = [1020224.49, 218516.28, 0.00]
DOWNSTREAM = [1020079.64, 218512.19, 0.00]

START_MSG = (
    "Build NYSDOT standard sheet 619-311 fresh. Enter wztc mode if needed. "
    "Deterministic path only: get_sheet_requirements, ask designer inputs once, "
    "build_wztc_order_table, then work-area edges, then "
    "run_sheet_build(upstream_edge, downstream_edge). "
    "Do NOT fish Default linework, do NOT freestyle define_alignment_segment, "
    "do NOT re-ask locked inputs. Prefer get_plan_status. "
    "If prior plan junk is in the way, clear_plan_elements(keep_alignments=False) first. "
    "NYW8-33 stays a handoff (vehicle-mounted)."
)


def append_input(message: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")
    flat = message.replace("\t", " ").replace("\r\n", " ").replace("\n", " ")
    with open(CHAT_INPUT, "a", encoding="utf-8", newline="") as f:
        f.write(f"{ts}\t{flat}\n")
    print(f"[INPUT] {flat[:240]}", flush=True)


def log_len() -> int:
    if not CHAT_LOG.exists():
        return 0
    return sum(1 for ln in CHAT_LOG.read_text(encoding="utf-8", errors="replace").splitlines() if ln.strip())


def reply_for_ask(line: str) -> str:
    blob = line.lower()
    # Both edges in one prompt — give both, labeled (do not send a single point).
    if re.search(r"\btwo\b|\bboth\b", blob) and re.search(r"edge|upstream|downstream", blob):
        return (
            f"upstream_edge=[{UPSTREAM[0]}, {UPSTREAM[1]}, 0]; "
            f"downstream_edge=[{DOWNSTREAM[0]}, {DOWNSTREAM[1]}, 0]. "
            "Call run_sheet_build with those exact lists. Proceed."
        )
    if re.search(r"upstream", blob) and not re.search(r"downstream", blob):
        return f"({UPSTREAM[0]}, {UPSTREAM[1]}, 0.00)"
    if re.search(r"downstream", blob):
        return f"({DOWNSTREAM[0]}, {DOWNSTREAM[1]}, 0.00)"
    if re.search(r"direction|bearing|east|west|north|south|words", blob):
        return (
            f"Travel is roughly westbound along the corridor. "
            f"upstream_edge=[{UPSTREAM[0]}, {UPSTREAM[1]}, 0] "
            f"downstream_edge=[{DOWNSTREAM[0]}, {DOWNSTREAM[1]}, 0]. "
            "Call run_sheet_build now. Do not ask again."
        )
    if re.search(r"45 mph|12ft|8ft|rural|option1", blob) or (
        re.search(r"speed|lane|shoulder|area", blob)
    ):
        return "45 mph / 12ft lane / 8ft shoulder / Rural"
    if re.search(r"order table|looks right|confirm|proceed|ok to draw|corridor", blob):
        return (
            "Order table OK. Call run_sheet_build("
            f"upstream_edge=[{UPSTREAM[0]}, {UPSTREAM[1]}, 0], "
            f"downstream_edge=[{DOWNSTREAM[0]}, {DOWNSTREAM[1]}, 0]). Proceed."
        )
    if re.search(r"arrow.?panel|trailer|vehicle", blob):
        return "trailer"
    return (
        "45 mph, Non-Freeway, 12 ft lane, 8 ft shoulder, area_type=RURAL, sheet 619-311. "
        f"Call run_sheet_build(upstream_edge=[{UPSTREAM[0]}, {UPSTREAM[1]}, 0], "
        f"downstream_edge=[{DOWNSTREAM[0]}, {DOWNSTREAM[1]}, 0]). Proceed."
    )


def main() -> int:
    start = log_len()
    append_input(START_MSG)
    seen = 0
    continues = 0
    t0 = time.time()
    timeout = 900
    saw_run_sheet = False
    saw_build = False
    saw_visual = False

    while time.time() - t0 < timeout:
        lines = [
            ln for ln in CHAT_LOG.read_text(encoding="utf-8", errors="replace").splitlines()
            if ln.strip()
        ][start:]
        while seen < len(lines):
            line = lines[seen]
            seen += 1
            parts = line.split("\t")
            typ = parts[1] if len(parts) > 1 else "?"
            summary = "\t".join(parts[1:3])[:220]
            print(f"[{typ}] {summary.encode('ascii', 'replace').decode('ascii')}", flush=True)

            low = line.lower()
            if typ == "TOOL_CALL" and "run_sheet_build" in low:
                saw_run_sheet = True
            if typ == "TOOL_CALL" and "build_wztc_order_table" in low:
                saw_build = True
            if typ == "TOOL_CALL" and "run_visual_qa" in low:
                saw_visual = True

            if typ == "ERROR":
                print("[FAIL] ERROR — intervening may be needed", flush=True)
                return 2

            if typ in ("ASK_USER", "ASK_USER_CHOICE"):
                if continues >= 12:
                    print("[FAIL] too many asks", flush=True)
                    return 3
                continues += 1
                reply = reply_for_ask(line)
                print(f"[AUTO-REPLY #{continues}] {reply[:160]}", flush=True)
                # Wait for ask tool to block on input
                time.sleep(0.8)
                start = log_len()
                seen = 0
                append_input(reply)
                break

            if typ == "FINAL":
                print(
                    f"[STATS] build_table={saw_build} run_sheet_build={saw_run_sheet} "
                    f"visual_qa={saw_visual} auto_replies={continues}",
                    flush=True,
                )
                text = line.lower()
                # If FINAL is just asking for inputs / edges, continue
                if re.search(r"posted speed|lane width|what.*(speed|edge)|need.*(input|edge)|identical|downstream", text) and continues < 12:
                    continues += 1
                    reply = (
                        f"Call run_sheet_build now with "
                        f"upstream_edge=[{UPSTREAM[0]}, {UPSTREAM[1]}, 0], "
                        f"downstream_edge=[{DOWNSTREAM[0]}, {DOWNSTREAM[1]}, 0]. "
                        "Do not ask for edges again. Then FINAL after it completes."
                    )
                    print(f"[AUTO-REPLY to FINAL ask #{continues}]", flush=True)
                    start = log_len()
                    seen = 0
                    append_input(reply)
                    break
                if saw_run_sheet or (saw_build and saw_visual):
                    print("[DONE] deterministic path exercised", flush=True)
                    return 0
                if continues < 8 and not saw_run_sheet:
                    continues += 1
                    nudge = (
                        "Continue 619-311. Locked path: if order table built, call "
                        f"run_sheet_build(upstream_edge={UPSTREAM}, "
                        f"downstream_edge={DOWNSTREAM}). Do not fish. Then FINAL."
                    )
                    print(f"[NUDGE #{continues}]", flush=True)
                    start = log_len()
                    seen = 0
                    append_input(nudge)
                    break
                print("[DONE] FINAL without run_sheet_build — check transcript", flush=True)
                return 0
        else:
            time.sleep(0.5)

    print("[FAIL] timeout", flush=True)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
