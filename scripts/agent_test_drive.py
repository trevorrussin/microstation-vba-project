"""Drive one chat_driver turn (or a short multi-turn) via chat-input.tsv
and print new chat-log lines until FINAL/ERROR or timeout. Auto-replies
to ASK_USER / ASK_USER_CHOICE with a canned continue so Cursor can run
the live agent test without the panel UI.

For Designer-style questions (speed / lane / shoulder / sheet), answers
come from --designer-* flags or the built-in Non-Freeway right-lane
closure defaults — not the generic continue message.
"""
from __future__ import annotations

import argparse
import re
import time
from datetime import datetime
from pathlib import Path

BRIDGE = Path(r"c:\repos\microstation-vba-project\Bridge")
CHAT_INPUT = BRIDGE / "chat-input.tsv"
CHAT_LOG = BRIDGE / "chat-log.tsv"

DEFAULT_CONTINUE = (
    "Looks good — proceed with the next step of the full standard-sheet plan. "
    "Keep going until: order table built from the full sheet sign list, "
    "place_order_table_stations done, EVERY isSign place_sign done at outward "
    "perp tips, and sheet elements addressed (place_element_run / place_cell / "
    "handoff). Do not stop after one sign or one tick."
)

# Defaults matching a typical Non-Freeway short-term right lane closure test
DEFAULT_SPEED = "45 mph"
DEFAULT_ROAD_TYPE = "Non-Freeway"
DEFAULT_LANE = "12 ft"
DEFAULT_SHOULDER = "8 ft"
DEFAULT_SHEET = "619-311"
DEFAULT_ALIGN_Y = "217040"


def append_input(message: str) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    flat = message.replace("\t", " ").replace("\r\n", " ").replace("\n", " ")
    with open(CHAT_INPUT, "a", encoding="utf-8", newline="") as f:
        f.write(f"{ts}\t{flat}\n")
    print(f"[INPUT] {flat[:200]}{'...' if len(flat) > 200 else ''}", flush=True)


def log_line_count() -> int:
    if not CHAT_LOG.exists():
        return 0
    return sum(1 for ln in CHAT_LOG.read_text(encoding="utf-8", errors="replace").splitlines() if ln.strip())


def read_log_from(start_idx: int) -> list[str]:
    if not CHAT_LOG.exists():
        return []
    lines = [ln for ln in CHAT_LOG.read_text(encoding="utf-8", errors="replace").splitlines() if ln.strip()]
    return lines[start_idx:]


def parse_type(line: str) -> str:
    parts = line.split("\t")
    return parts[1] if len(parts) > 1 else ""


def ask_blob(line: str) -> str:
    """Flatten ASK_* fields for keyword matching."""
    return line.lower()


def reply_for_ask(line: str, args: argparse.Namespace) -> str:
    """Map Designer-style questions to concrete answers."""
    blob = ask_blob(line)
    y = getattr(args, "align_y", DEFAULT_ALIGN_Y)

    # Sheets without advanceWarningSpacing must not invent URBAN/RURAL.
    if re.search(r"area[_\s-]*type|urban|rural", blob) and not re.search(
        r"\b(speed|lane\s*width|shoulder)\b", blob
    ):
        return (
            f"Do NOT invent URBAN/RURAL. Sheet {args.sheet} may have no "
            f"advanceWarningSpacing role — call build_wztc_order_table WITHOUT "
            f"area_type (leave it empty). Road type = {args.road_type}; "
            f"speed = {args.speed}; lane = {args.lane}; shoulder = {args.shoulder}. "
            f"Build on Y={y} only. Proceed."
        )

    if re.search(r"too short|4595|extend|corridor.*(length|ft)|how should i resolve", blob):
        y = getattr(args, "align_y", DEFAULT_ALIGN_Y)
        return (
            f"EXTEND the corridor. Define Upstream at Y={y} from X=1018000 to "
            f"X=1024000 (6000 ft). Downstream as needed. Approved — proceed "
            "stations, every isSign, dims/labels/WS/chan/PV/AP, view_drawing / run_visual_qa_captures."
        )

    # Multi-field asks: answer everything at once
    wants_speed = bool(re.search(r"speed|mph|posted", blob))
    wants_road = bool(re.search(r"road\s*type|freeway|non-freeway", blob))
    wants_lane = bool(re.search(r"lane\s*width", blob))
    wants_shoulder = bool(re.search(r"shoulder\s*width|shoulder\s*band", blob))
    wants_sheet = bool(re.search(r"619-|sheet|short\s*term|short\s*duration|duration", blob))

    parts: list[str] = []
    if wants_speed:
        parts.append(f"posted speed = {args.speed}")
    if wants_road:
        parts.append(f"road type = {args.road_type}")
    if wants_lane:
        parts.append(f"lane width = {args.lane}")
    if wants_shoulder:
        parts.append(f"shoulder width = {args.shoulder}")
    if wants_sheet:
        parts.append(f"use sheet {args.sheet} (Short Term). Keep this sheet number.")

    if parts:
        return (
            "; ".join(parts)
            + f". Corridor Y={y} (X~1019735 to 1022735). If the tool rejects "
            "missing area_type only when that sheet has advanceWarningSpacing; "
            "otherwise omit area_type. Proceed with get_sheet_requirements + "
            "build_wztc_order_table using the FULL sheet sign list."
        )

    # Alignment / workspace / proceed
    if re.search(r"work\s*area|how long|length should|work span", blob):
        return (
            f"Use a default work-area length of 200 ft downstream of path "
            f"start (station 0) at Y={y} unless the sheet prints a fixed "
            f"length. Call place_order_table_workspace / channelizing / "
            f"PV with Buffer Space fallback — do not block on this. Proceed."
        )

    if re.search(r"align|which line|click|point|workspace|boundary|reference", blob):
        return (
            f"Use the ~3000 ft E-W alignment at Y={y} "
            f"(X=1019735 to X=1022735). Upstream = that line west→east. "
            "For workspace, a simple rectangle hugging that corridor is fine. Proceed."
        )

    if re.search(r"order table|looks good|confirm|ok to|proceed|approve", blob):
        return args.continue_msg

    return args.continue_msg


def drive(message: str, args: argparse.Namespace) -> int:
    start = log_line_count()
    append_input(message)
    continues = 0
    pending_ask_line: str | None = None
    saw_order_table_stations = False
    saw_build_order_table = False
    saw_get_sheet = False
    saw_place_sign = False
    place_sign_count = 0
    saw_ask_inputs = False
    t0 = time.time()
    seen = 0

    while time.time() - t0 < args.timeout:
        new_lines = read_log_from(start)
        while seen < len(new_lines):
            line = new_lines[seen]
            seen += 1
            typ = parse_type(line)
            parts = line.split("\t")
            summary = "\t".join(parts[1:3])[:220] if len(parts) > 1 else line[:220]
            safe = summary.encode("ascii", "replace").decode("ascii")
            print(f"[{typ}] {safe}", flush=True)

            low = line.lower()
            if typ == "TOOL_CALL" and "place_order_table_stations" in low:
                saw_order_table_stations = True
            if typ == "TOOL_CALL" and "build_wztc_order_table" in low:
                saw_build_order_table = True
            if typ == "TOOL_CALL" and "get_sheet_requirements" in low:
                saw_get_sheet = True
            if typ == "TOOL_CALL" and "name=place_sign" in low.replace(" ", ""):
                saw_place_sign = True
                place_sign_count += 1
            if typ in ("ASK_USER", "ASK_USER_CHOICE"):
                saw_ask_inputs = True
                pending_ask_line = line
            if typ == "ERROR":
                print("[FAIL] ERROR", flush=True)
                return 2
            if typ == "FINAL":
                print(
                    f"[STATS] get_sheet={saw_get_sheet} build_table={saw_build_order_table} "
                    f"stations={saw_order_table_stations} place_sign_n={place_sign_count} "
                    f"asked={saw_ask_inputs}",
                    flush=True,
                )
                # Agent sometimes asks Designer inputs in FINAL text instead of
                # ask_user — treat that as an ask and supply values.
                if re.search(
                    r"posted speed|lane width|shoulder width|619-|sheet/duration|what.?s the speed",
                    line,
                    re.I,
                ) and not (
                    saw_get_sheet or saw_build_order_table
                ):
                    if continues < args.max_continues:
                        continues += 1
                        pending_ask_line = None
                        y = getattr(args, "align_y", DEFAULT_ALIGN_Y)
                        reply = (
                            f"posted speed = {args.speed}; road type = {args.road_type}; "
                            f"lane width = {args.lane}; shoulder width = {args.shoulder}; "
                            f"use sheet {args.sheet} (Short Term). Alignment = the new "
                            f"~3000 ft E-W line around Y={y} (X~1019735 to 1022735). "
                            f"Omit area_type unless the sheet has advanceWarningSpacing. "
                            f"Proceed with get_sheet_requirements + full order-table plan."
                        )
                        print(
                            f"[AUTO-REPLY #{continues} to FINAL Designer-input ask]",
                            flush=True,
                        )
                        start = log_line_count()
                        seen = 0
                        append_input(reply)
                        break

                # Full success: sheet + table + stations + multiple signs
                if (
                    saw_get_sheet
                    and saw_build_order_table
                    and saw_order_table_stations
                    and place_sign_count >= 3
                ):
                    print("[DONE] full-sheet path exercised (sheet+table+stations+>=3 signs)", flush=True)
                    return 0
                # Partial success: order-table path at least
                if saw_order_table_stations and (saw_build_order_table or saw_place_sign):
                    if continues < args.max_continues:
                        continues += 1
                        pending_ask_line = None
                        print(
                            f"[AUTO-CONTINUE #{continues} after FINAL — need more signs/elements]",
                            flush=True,
                        )
                        start = log_line_count()
                        seen = 0
                        append_input(args.continue_msg)
                        break
                    print("[DONE] order-table path exercised (partial signs)", flush=True)
                    return 0
                if continues < args.max_continues:
                    continues += 1
                    pending_ask_line = None
                    print(f"[AUTO-CONTINUE #{continues} after FINAL checkpoint]", flush=True)
                    start = log_line_count()
                    seen = 0
                    append_input(args.continue_msg)
                    break
                print("[DONE] FINAL (no more continues)", flush=True)
                return 0

        if pending_ask_line is not None:
            ask_line = pending_ask_line
            pending_ask_line = None
            continues += 1
            if continues > args.max_continues:
                print("[FAIL] too many continues", flush=True)
                return 3
            reply = reply_for_ask(ask_line, args)
            print(f"[AUTO-REPLY #{continues} to ASK_*]", flush=True)
            append_input(reply)

        time.sleep(0.5)

    print(
        f"[FAIL] timeout  get_sheet={saw_get_sheet} build={saw_build_order_table} "
        f"stations={saw_order_table_stations} signs={place_sign_count} asked={saw_ask_inputs}",
        flush=True,
    )
    return 1


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--message", required=True)
    ap.add_argument("--continue-msg", default=DEFAULT_CONTINUE)
    ap.add_argument("--timeout", type=float, default=900)
    ap.add_argument("--max-continues", type=int, default=8,
                    help="Auto-continue/ASK reply budget (default 8; was 16 — "
                         "lower to control live-agent API cost)")
    ap.add_argument("--fresh-history", action="store_true",
                    help="Clear Bridge/chat-history.json before posting "
                         "(avoids re-billing a huge poisoned history)")
    ap.add_argument("--speed", default=DEFAULT_SPEED)
    ap.add_argument("--road-type", default=DEFAULT_ROAD_TYPE)
    ap.add_argument("--lane", default=DEFAULT_LANE)
    ap.add_argument("--shoulder", default=DEFAULT_SHOULDER)
    ap.add_argument("--sheet", default=DEFAULT_SHEET)
    ap.add_argument("--align-y", default=DEFAULT_ALIGN_Y,
                    help="Corridor Y for auto-replies (311≈217040, 301 test≈216840)")
    args = ap.parse_args()
    if args.fresh_history:
        hist = BRIDGE / "chat-history.json"
        backup = BRIDGE / "chat-history.prev.json"
        if hist.exists():
            hist.replace(backup)
            print(f"[fresh-history] moved prior history to {backup.name}", flush=True)
        hist.write_text("[]", encoding="utf-8")
        print("[fresh-history] chat-history.json reset to []", flush=True)
    return drive(args.message, args)


if __name__ == "__main__":
    raise SystemExit(main())
