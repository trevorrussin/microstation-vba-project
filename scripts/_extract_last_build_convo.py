"""One-shot: format last afternoon 619-311 build into readable markdown."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BRIDGE = ROOT / "Bridge"


def trunc(s: str, n: int = 800) -> str:
    s = s or ""
    return s if len(s) <= n else s[:n] + f" …[{len(s) - n} more chars]"


def parse_input_ts(ts_s: str) -> datetime:
    if ts_s.startswith("2026-") or ts_s.startswith("2025-"):
        return datetime.fromisoformat(ts_s)
    return datetime.strptime(ts_s, "%m/%d/%Y %I:%M:%S %p")


def main() -> None:
    inputs: list[tuple[datetime, str]] = []
    for line in (BRIDGE / "chat-input.tsv").read_text(encoding="utf-8", errors="replace").splitlines():
        if "\t" not in line:
            continue
        ts_s, text = line.split("\t", 1)
        try:
            ts = parse_input_ts(ts_s)
        except ValueError:
            continue
        if ts.date().isoformat() == "2026-08-04" and ts.hour >= 16:
            inputs.append((ts, text.strip()))

    rows: list[tuple[datetime, str, dict]] = []
    extract = BRIDGE / "last-build-session-extract.tsv"
    if not extract.exists():
        # rebuild slice
        log = (BRIDGE / "chat-log.tsv").read_text(encoding="utf-8", errors="replace").splitlines()
        start = next(
            (i for i, line in enumerate(log)
             if line.startswith("2026-08-04 16:52")
             or line.startswith("2026-08-04 16:53")
             or line.startswith("2026-08-04 16:54")
             or line.startswith("2026-08-04 16:55")),
            None,
        )
        if start is None:
            raise SystemExit("could not find afternoon session in chat-log.tsv")
        extract.write_text("\n".join(log[start:]) + "\n", encoding="utf-8")

    for line in extract.read_text(encoding="utf-8", errors="replace").splitlines():
        parts = line.split("\t")
        if len(parts) < 2:
            continue
        ts = datetime.fromisoformat(parts[0])
        kind = parts[1]
        fields: dict[str, str] = {}
        for p in parts[2:]:
            if "=" in p:
                k, v = p.split("=", 1)
                fields[k] = v
        rows.append((ts, kind, fields))

    events: list[tuple[str, datetime, dict]] = [("USER", ts, {"text": t}) for ts, t in inputs]
    events += [(kind, ts, fields) for ts, kind, fields in rows]
    events.sort(key=lambda x: x[1])

    out: list[str] = [
        "# Last build conversation — 619-311 (2026-08-04 afternoon)",
        "",
        "Source: `Bridge/chat-input.tsv` + `Bridge/chat-log.tsv` "
        "(~16:52 through FINAL at 17:09).",
        "Includes engineer prompts, agent THINKING, tool calls/results, "
        "asks, screenshots, and FINALs.",
        "",
    ]

    for kind, ts, fields in events:
        stamp = ts.strftime("%H:%M:%S")
        if kind == "USER":
            out += [f"## [{stamp}] ENGINEER", "", fields["text"], ""]
        elif kind == "THINKING":
            out += [f"### [{stamp}] THINKING", "", trunc(fields.get("text", ""), 1500), ""]
        elif kind == "TOOL_CALL":
            name = fields.get("name", "?")
            out.append(f"- **[{stamp}] TOOL_CALL** `{name}`")
            out.append(f"  - input: `{trunc(fields.get('input', ''), 600)}`")
        elif kind == "TOOL_RESULT":
            name = fields.get("name", "?")
            status = fields.get("status", "?")
            out.append(
                f"- **[{stamp}] TOOL_RESULT** `{name}` → **{status}** — "
                f"{trunc(fields.get('summary', ''), 500)}"
            )
        elif kind == "ASK_USER_CHOICE":
            out += [f"### [{stamp}] ASK_USER_CHOICE", "", "```"]
            for k, v in fields.items():
                out.append(f"{k}={trunc(v, 400)}")
            out += ["```", ""]
        elif kind == "ASK_USER":
            out += [f"### [{stamp}] ASK_USER", "", trunc(fields.get("question", ""), 1000), ""]
        elif kind == "SCREENSHOT":
            out.append(f"- **[{stamp}] SCREENSHOT** `{fields.get('path', '')}`")
        elif kind == "FINAL":
            out += [f"## [{stamp}] FINAL", "", fields.get("text", ""), ""]
        elif kind == "ERROR":
            out += [f"## [{stamp}] ERROR", "", trunc(fields.get("note", ""), 1200), ""]
        elif kind == "MODE_CHANGED":
            out.append(
                f"- **[{stamp}] MODE** → {fields.get('mode')} "
                f"({fields.get('description', '')})"
            )

    path = BRIDGE / "last-build-conversation.md"
    path.write_text("\n".join(out), encoding="utf-8")
    print(f"wrote {path} ({path.stat().st_size} bytes, {len(events)} events)")


if __name__ == "__main__":
    main()
