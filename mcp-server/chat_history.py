"""Conversation history persistence, trimming, and prompt-cache-marker
management for chat_driver.py. Extracted (2026-08-04 split) as a
self-contained concern: everything here operates on `messages: list[dict]`
and a couple of size caps, independent of the tool-calling loop itself.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

# Hard cap on persisted conversation length. Live 2026-08-03: chat-history
# hit ~616 messages / 1.1MB; every new turn re-sent that context and the
# agent also refused to retry a fixed tool based on stale errors in that
# history. Keep a bounded recent window. Override with WZTC_MAX_HISTORY_MESSAGES.
MAX_HISTORY_MESSAGES = int(os.environ.get("WZTC_MAX_HISTORY_MESSAGES", "40"))
# Secondary char budget on serialized content (rough); trim oldest until under.
MAX_HISTORY_CHARS = int(os.environ.get("WZTC_MAX_HISTORY_CHARS", "350000"))
# Trimming from the FRONT of `messages` changes the prompt-cache prefix for
# every later request (the API caches an exact byte-prefix match), forcing a
# full-price cache rewrite the next turn -- confirmed live 2026-08-03:
# cacheWrite jumped to 238,912 tokens ($1.44) immediately after a trim.
# Trimming down to the cap on every turn that exceeds it means that rewrite
# happens on almost every turn once history fills up. Trimming further below
# the cap (hysteresis) leaves headroom so the rewrite is rare instead of
# continuous, at the cost of carrying a somewhat shorter live window.
HISTORY_TRIM_TARGET_MESSAGES = int(os.environ.get("WZTC_HISTORY_TRIM_TARGET_MESSAGES", "24"))
HISTORY_TRIM_TARGET_CHARS = int(os.environ.get("WZTC_HISTORY_TRIM_TARGET_CHARS", "220000"))

# After a turn finishes, image tool_results (view_drawing) stay in the
# in-memory messages list and get re-sent on every later round-trip until
# Anthropic's clear_tool_uses edit ages them out -- which, measured live
# 2026-08-02, did NOT keep three ~300KB base64 screenshots out of
# chat-history.json. A cache-miss turn then billed ~243k input tokens
# (~$0.73) twice in a few seconds. Strip images (and truncate other giant
# text tool_results) ourselves once the turn that needed them is over.
_MAX_TOOL_RESULT_CHARS = 12_000
_IMAGE_STUB = (
    "[screenshot omitted from history to control cost — "
    "call view_drawing again if you still need to see the view]"
)

# +1 for the marker this turn is about to add, +1 for the system block's own
# marker = 4 total per request, exactly the API's hard per-request cap.
MAX_KEPT_MESSAGE_CACHE_MARKERS = 2


def _to_jsonable(obj):
    """Serialize a message for chat-history.json in a form that's valid to
    send straight back to the API on reload. Plain obj.model_dump(mode="json")
    is not enough: parsed response blocks (e.g. text blocks from tool_runner's
    use of messages.parse()) carry SDK-internal fields like parsed_output
    that the request schema rejects outright -- confirmed live, a reloaded
    history crashed the next turn with "content.0.text.parsed_output: Extra
    inputs are not permitted". The SDK's own outbound request transform
    (anthropic/_utils/_transform.py) strips exactly these via each model's
    __api_exclude__ attribute before sending; mirroring the same
    exclude_unset/exclude combination here keeps a reload byte-faithful to
    what the API actually accepts as input."""
    if hasattr(obj, "model_dump"):
        return obj.model_dump(
            mode="json", exclude_unset=True, by_alias=True, exclude=getattr(obj, "__api_exclude__", None)
        )
    if isinstance(obj, list):
        return [_to_jsonable(x) for x in obj]
    if isinstance(obj, dict):
        return {k: _to_jsonable(v) for k, v in obj.items()}
    return obj


def _content_char_len(content) -> int:
    if isinstance(content, str):
        return len(content)
    if not isinstance(content, list):
        return 0
    n = 0
    for block in content:
        if isinstance(block, dict):
            n += len(str(block.get("text", "") or ""))
            n += len(str(block.get("thinking", "") or ""))
            inner = block.get("content")
            if inner is not None:
                n += _content_char_len(inner)
        else:
            n += 64
    return n


def _strip_old_thinking(messages: list) -> None:
    """Remove thinking blocks from all but the newest assistant message.
    Prior-turn thinking is not required for correctness (Anthropic allows
    removing it) and was a quiet cost driver on long sessions. Drop the
    blocks entirely rather than stubbing — stubbed thinking without a
    signature can confuse the API."""
    last_asst = -1
    for i, msg in enumerate(messages):
        if isinstance(msg, dict) and msg.get("role") == "assistant":
            last_asst = i
    for i, msg in enumerate(messages):
        if i == last_asst:
            continue
        if not isinstance(msg, dict) or msg.get("role") != "assistant":
            continue
        content = msg.get("content")
        if not isinstance(content, list):
            continue
        msg["content"] = [
            block for block in content
            if not (isinstance(block, dict) and block.get("type") == "thinking")
        ]


def _block_get(block, key: str, default=None):
    """Read a field from a content block whether it is a plain dict (loaded
    history) or an SDK model object (fresh tool_runner append). Repair used
    to only see dicts — unanswered tool_use SDK blocks survived save_history
    and 400'd the next turn (live 2026-08-05)."""
    if isinstance(block, dict):
        return block.get(key, default)
    return getattr(block, key, default)


def _block_type(block) -> str | None:
    t = _block_get(block, "type")
    return str(t) if t is not None else None


def _msg_tool_use_ids(msg: dict) -> set[str]:
    content = msg.get("content")
    if not isinstance(content, list):
        return set()
    ids: set[str] = set()
    for block in content:
        if _block_type(block) == "tool_use":
            tid = _block_get(block, "id")
            if tid:
                ids.add(str(tid))
    return ids


def _msg_tool_result_ids(msg: dict) -> set[str]:
    content = msg.get("content")
    if not isinstance(content, list):
        return set()
    ids: set[str] = set()
    for block in content:
        if _block_type(block) == "tool_result":
            tid = _block_get(block, "tool_use_id")
            if tid:
                ids.add(str(tid))
    return ids


def _is_tool_result_only_user(msg: dict) -> bool:
    """True when a user message is only tool_result blocks (no text).
    A front-trim that lands here leaves an orphan tool_use_id with no
    preceding tool_use — API 400 on the next turn (live 2026-08-04)."""
    if not isinstance(msg, dict) or msg.get("role") != "user":
        return False
    content = msg.get("content")
    if isinstance(content, str):
        return False
    if not isinstance(content, list) or not content:
        return False
    saw_result = False
    for block in content:
        t = _block_type(block)
        if t is None and not isinstance(block, dict):
            # Unknown SDK object — not a pure tool_result message.
            return False
        if t == "tool_result":
            saw_result = True
        elif t == "text" and str(_block_get(block, "text") or "").strip():
            return False
        else:
            return False
    return saw_result


def _safe_trim_start_index(messages: list, want_keep: int) -> int:
    """Index to slice from so messages[start:] keeps ~want_keep msgs and
    does not begin mid tool-call chain. Walks earlier if needed to find a
    user text message (or any non-tool_result-only user)."""
    if want_keep <= 0 or want_keep >= len(messages):
        return 0
    start = len(messages) - want_keep
    while start > 0 and _is_tool_result_only_user(messages[start]):
        start -= 1
    # Prefer starting on a real user turn (text), not an assistant mid-chain.
    while start > 0:
        m = messages[start]
        if isinstance(m, dict) and m.get("role") == "user" and not _is_tool_result_only_user(m):
            break
        if isinstance(m, dict) and m.get("role") == "user":
            break
        start -= 1
    return start


def _repair_tool_pairing(messages: list) -> None:
    """Drop orphan tool_result user messages / broken tool_use chains so
    reloaded history is valid for the Messages API.

    Live 2026-08-04: front-trim left messages[0] as a lone tool_result
    (toolu_0142…); the next user turn failed with
    'unexpected tool_use_id found in tool_result blocks'."""
    # 1) Drop leading tool_result-only user messages (no prior tool_use).
    while messages and _is_tool_result_only_user(messages[0]):
        messages.pop(0)
        print("[history] dropped leading orphan tool_result user message", flush=True)

    # 2) Walk pairs: each user tool_result must reference ids from the
    #    immediately preceding assistant tool_use. Drop the user message
    #    (and optionally a dangling assistant tool_use with no result) when
    #    the pairing is broken.
    i = 0
    while i < len(messages):
        msg = messages[i]
        if not isinstance(msg, dict):
            i += 1
            continue
        if msg.get("role") == "user" and _msg_tool_result_ids(msg):
            prev = messages[i - 1] if i > 0 else None
            prev_ids = _msg_tool_use_ids(prev) if isinstance(prev, dict) and prev.get("role") == "assistant" else set()
            need = _msg_tool_result_ids(msg)
            if not need.issubset(prev_ids):
                print(
                    f"[history] dropped unpaired tool_result user message "
                    f"(need={sorted(need - prev_ids)[:3]})",
                    flush=True,
                )
                messages.pop(i)
                continue
        i += 1

    # 2b) Drop assistant messages that issued tool_use without an immediate
    #     following user tool_result (live 2026-08-05: FINAL turn saved
    #     assistant(tool_use) then assistant(text) with no tool_result in
    #     between → next turn 400 'tool_use ids without tool_result').
    i = 0
    while i < len(messages):
        msg = messages[i]
        if not isinstance(msg, dict) or msg.get("role") != "assistant":
            i += 1
            continue
        uses = _msg_tool_use_ids(msg)
        if not uses:
            i += 1
            continue
        nxt = messages[i + 1] if i + 1 < len(messages) else None
        ok = (
            isinstance(nxt, dict)
            and nxt.get("role") == "user"
            and uses.issubset(_msg_tool_result_ids(nxt))
        )
        if ok:
            i += 1
            continue
        print(
            f"[history] dropped assistant with unanswered tool_use "
            f"(ids={sorted(uses)[:3]})",
            flush=True,
        )
        messages.pop(i)

    # 3) Drop a trailing assistant that still has unanswered tool_use
    #    (interrupted turn) — next user text would also 400.
    while messages:
        last = messages[-1]
        if not isinstance(last, dict) or last.get("role") != "assistant":
            break
        uses = _msg_tool_use_ids(last)
        if not uses:
            break
        print("[history] dropped trailing assistant with unanswered tool_use",
              flush=True)
        messages.pop()

    # 4) Messages API requires the first message to be role=user. After a
    #    front-trim that landed mid tool-chain we often start on assistant.
    #    Drop leading assistant + matching tool_result pairs until a real
    #    user text message remains; if none exists, clear (fresh session is
    #    safer than a synthetic stub that confuses the model).
    while messages:
        first = messages[0]
        if isinstance(first, dict) and first.get("role") == "user" and not _is_tool_result_only_user(first):
            break
        if isinstance(first, dict) and first.get("role") == "assistant":
            messages.pop(0)
            if messages and _is_tool_result_only_user(messages[0]):
                messages.pop(0)
            print("[history] dropped leading assistant(+tool_result) so history "
                  "starts on a user message", flush=True)
            continue
        if _is_tool_result_only_user(first):
            messages.pop(0)
            print("[history] dropped leading orphan tool_result user message", flush=True)
            continue
        break
    has_user_text = any(
        isinstance(m, dict) and m.get("role") == "user" and not _is_tool_result_only_user(m)
        for m in messages
    )
    if messages and not has_user_text:
        print("[history] no user-text message left after repair — clearing history",
              flush=True)
        messages.clear()


def harness_history_issues(messages: list) -> list[str]:
    """Detect remaining Messages-API breakers after _repair_tool_pairing.

    Empty list = safe to send. Non-empty = harness P0 — clear rather than
    nudging the model through a 400 loop.
    """
    issues: list[str] = []
    if not messages:
        return issues
    first = messages[0]
    if not (isinstance(first, dict) and first.get("role") == "user"
            and not _is_tool_result_only_user(first)):
        issues.append("history does not start with a real user text message")
    for i, msg in enumerate(messages):
        if not isinstance(msg, dict):
            issues.append(f"messages[{i}] is not a dict")
            continue
        if msg.get("role") == "assistant":
            uses = _msg_tool_use_ids(msg)
            if uses:
                nxt = messages[i + 1] if i + 1 < len(messages) else None
                if not (
                    isinstance(nxt, dict)
                    and nxt.get("role") == "user"
                    and uses.issubset(_msg_tool_result_ids(nxt))
                ):
                    issues.append(
                        f"messages[{i}] assistant has unanswered tool_use "
                        f"{sorted(uses)[:3]}"
                    )
        if msg.get("role") == "user" and _msg_tool_result_ids(msg):
            prev = messages[i - 1] if i > 0 else None
            prev_ids = (
                _msg_tool_use_ids(prev)
                if isinstance(prev, dict) and prev.get("role") == "assistant"
                else set()
            )
            need = _msg_tool_result_ids(msg)
            if not need.issubset(prev_ids):
                issues.append(
                    f"messages[{i}] tool_result without matching prior tool_use "
                    f"{sorted(need - prev_ids)[:3]}"
                )
    return issues


def harness_preflight_or_clear(messages: list) -> list[str]:
    """Repair, re-check, CLEAR history if still broken (HARNESS_P0).

    Returns issue strings that triggered a clear (empty if healthy).
    """
    _repair_tool_pairing(messages)
    issues = harness_history_issues(messages)
    if issues:
        messages.clear()
        print(
            f"[history] HARNESS_P0 cleared broken history: {issues[:4]}",
            flush=True,
        )
    return issues


def _trim_history_window(messages: list) -> list:
    """Keep only the newest MAX_HISTORY_MESSAGES, then trim oldest further
    if the remaining text still exceeds MAX_HISTORY_CHARS. Always leaves at
    least 2 messages when possible (last user+assistant exchange).

    Trims down to HISTORY_TRIM_TARGET_MESSAGES/_CHARS -- below the cap --
    rather than to the cap itself. See the cache-prefix comment on those
    constants: trimming to the exact cap busts the prompt cache on nearly
    every turn once history fills up; trimming further below leaves several
    turns of headroom before the next (expensive) rewrite is needed.

    Cuts only at safe boundaries (not mid tool_use/tool_result chain) and
    repairs any residual orphans afterward."""
    if not messages:
        return messages
    if len(messages) > MAX_HISTORY_MESSAGES:
        target = min(HISTORY_TRIM_TARGET_MESSAGES, MAX_HISTORY_MESSAGES)
        start = _safe_trim_start_index(messages, target)
        dropped = start
        messages[:] = messages[start:]
        print(f"[history] trimmed {dropped} older messages "
              f"(cap={MAX_HISTORY_MESSAGES}, target={target}, "
              f"kept={len(messages)})", flush=True)
    total = sum(_content_char_len(m.get("content")) for m in messages
                if isinstance(m, dict))
    if total > MAX_HISTORY_CHARS:
        target_chars = min(HISTORY_TRIM_TARGET_CHARS, MAX_HISTORY_CHARS)
        while len(messages) > 2:
            total = sum(_content_char_len(m.get("content")) for m in messages
                        if isinstance(m, dict))
            if total <= target_chars:
                break
            # Prefer dropping a whole safe prefix unit, not a lone tool_result.
            if len(messages) > 2 and _is_tool_result_only_user(messages[1]):
                # Drop assistant+results together when possible.
                messages.pop(0)
                if messages and _is_tool_result_only_user(messages[0]):
                    messages.pop(0)
            else:
                messages.pop(0)
            print(f"[history] trimmed oldest message (chars target={target_chars})",
                  flush=True)
    _repair_tool_pairing(messages)
    return messages


def _shrink_tool_result_content(inner):
    if isinstance(inner, list):
        out = []
        for part in inner:
            if isinstance(part, dict) and part.get("type") == "image":
                out.append({"type": "text", "text": _IMAGE_STUB})
            elif isinstance(part, dict) and part.get("type") == "text":
                text = part.get("text", "")
                if len(text) > _MAX_TOOL_RESULT_CHARS:
                    text = text[:_MAX_TOOL_RESULT_CHARS] + "\n[truncated — re-query with a tighter scope if needed]"
                out.append({**part, "text": text})
            else:
                out.append(part)
        return out
    if isinstance(inner, str) and len(inner) > _MAX_TOOL_RESULT_CHARS:
        return inner[:_MAX_TOOL_RESULT_CHARS] + "\n[truncated — re-query with a tighter scope if needed]"
    return inner


def _strip_bulky_history(messages: list) -> None:
    """Mutate messages in place: drop base64 image payloads from prior
    tool_results, truncate oversized text tool_results, stub old thinking,
    and enforce the message/char window. Safe on a mix of plain dicts
    (loaded history / prior turns) and SDK objects (this turn's freshly
    appended content) -- non-dict blocks are left alone for image/text
    shrink; thinking/window trim only touches dict messages."""
    for msg in messages:
        if not isinstance(msg, dict):
            continue
        content = msg.get("content")
        if not isinstance(content, list):
            continue
        for i, block in enumerate(content):
            if not isinstance(block, dict):
                continue
            if block.get("type") != "tool_result":
                continue
            inner = block.get("content")
            content[i] = {**block, "content": _shrink_tool_result_content(inner)}
    _strip_old_thinking(messages)
    _trim_history_window(messages)


def load_history(history_file: Path) -> list[dict]:
    if not history_file.exists():
        return []
    raw = history_file.read_text(encoding="utf-8-sig")
    messages = json.loads(raw)
    before = len(messages)
    before_fingerprint = (
        messages[0].get("role") if messages and isinstance(messages[0], dict) else None,
        _msg_tool_result_ids(messages[0]) if messages and isinstance(messages[0], dict) else set(),
    )
    _strip_bulky_history(messages)
    after_fingerprint = (
        messages[0].get("role") if messages and isinstance(messages[0], dict) else None,
        _msg_tool_result_ids(messages[0]) if messages and isinstance(messages[0], dict) else set(),
    )
    if len(messages) < before or before_fingerprint != after_fingerprint:
        # Persist trim/repair so the next restart doesn't re-load orphans.
        save_history(history_file, messages)
    return messages


def save_history(history_file: Path, messages: list[dict]) -> None:
    _strip_bulky_history(messages)
    serializable = [{"role": m["role"], "content": _to_jsonable(m["content"])} for m in messages]
    history_file.write_text(json.dumps(serializable, indent=2, ensure_ascii=False), encoding="utf-8")


def trim_cache_control(messages: list[dict]) -> None:
    """Keeps only the newest MAX_KEPT_MESSAGE_CACHE_MARKERS pre-existing
    cache_control markers in messages, stripping older ones, before the
    caller adds one more for this turn's user message.

    An earlier version of this function stripped ALL old markers down to
    zero, keeping only the single newest one. That avoided the 4-marker
    cap error ("A maximum of 4 blocks with cache_control may be provided")
    but broke caching outright on this long-running conversation --
    confirmed live, the next turn's input_tokens jumped to ~223,000
    (essentially uncached) instead of reading from cache, at ~$1.15/turn.
    Why: the API only finds a cache hit by walking back at most 20 content
    blocks from a breakpoint (prompt-caching.md "20-block lookback
    window"). With only one marker at the very end of an hours-long,
    many-turn conversation, the nearest actual cached prefix was hundreds
    of blocks further back -- outside that window -- so everything in
    between got billed at full price. The original (buggy) code
    accidentally got this part right: leaving a marker on every turn kept
    consecutive breakpoints close together, always within reach of the
    walk-back. Keeping a couple of recent markers instead of stripping to
    one preserves that locality while still staying under the 4-marker
    cap. Not a perfect guarantee against the 20-block gap on a single
    unusually tool-call-heavy turn (this tracks marker count, not actual
    block count) -- but a large improvement over both the original bug
    and the first fix's regression.

    Mutates messages in place. Safe to call on a mix of plain dicts (loaded
    from chat-history.json, or a prior turn's tool_result) and raw SDK
    content-block objects (this turn's freshly-appended assistant content,
    not yet round-tripped through save_history) -- only dict-shaped blocks
    can carry a cache_control key this code itself set, so non-dict blocks
    are left untouched rather than guessed at."""
    markers = []
    for msg in messages:
        content = msg.get("content")
        if not isinstance(content, list):
            continue
        for block in content:
            if isinstance(block, dict) and "cache_control" in block:
                markers.append(block)

    if len(markers) > MAX_KEPT_MESSAGE_CACHE_MARKERS:
        for block in markers[:-MAX_KEPT_MESSAGE_CACHE_MARKERS]:
            block.pop("cache_control", None)
