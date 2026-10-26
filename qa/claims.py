"""Evidence labels for statements an agent makes about a task.

OBSERVED    directly measured or read back (needs the evidence: tool, read-back, screenshot path)
INFERRED    follows from observations but was not tested (needs the evidence it follows from)
HYPOTHESIS  plausible, untested
VERIFIED    confirmed by an independent test: the evidence must name a gate check that PASSED this task
UNKNOWN     not enough evidence
SUPPORTED   evidence exists but the conclusion needs inference (needs the evidence it rests on)
CONTRADICTED available evidence conflicts with the claim (needs the conflicting evidence)

`close` refuses while a hypothesis, unknown or contradicted claim is open, and lists inferred and supported claims as
NOT verified. A claim is upgraded with the same id (`qa_gate claim --id N --level verified --evidence "COM-01"`) or withdrawn.
"""
from __future__ import annotations

import re

LEVELS = ("observed", "inferred", "supported", "hypothesis", "verified", "unknown", "contradicted")
NEEDS_EVIDENCE = ("observed", "inferred", "supported", "verified", "contradicted")
OPEN_LEVELS = ("hypothesis", "unknown", "contradicted")
_CHECK_ID = re.compile(r"\b[A-Z]{2,5}(?:-[A-Z]{2,5})?-\d{2}\b")


def passed_checks(latest: dict | None) -> set[str]:
    return {r["id"] for r in (latest or {}).get("results", []) if r.get("status") == "PASS"}


def validate(level: str, text: str, evidence: str, passed: set[str]) -> str:
    """-> "" when the claim may be recorded, else the reason it may not."""
    if level not in LEVELS:
        return f"level must be one of {', '.join(LEVELS)}"
    if len((text or "").strip()) < 10:
        return "state the claim in a sentence (>= 10 characters)"
    ev = (evidence or "").strip()
    if level in NEEDS_EVIDENCE and len(ev) < 5:
        return f"a {level} claim needs --evidence (the read-back, tool result, screenshot or check that supports it)"
    if level == "verified" and not (set(_CHECK_ID.findall(ev)) & passed):
        return ("a verified claim must cite a gate check that PASSED on this task (e.g. COM-01, REQ-GEOM-01). "
                f"Passed now: {', '.join(sorted(passed)) or 'none'}. Otherwise record it as inferred.")
    return ""


def add(task: dict, level: str, text: str, evidence: str, passed: set[str], claim_id: int | None = None) -> dict:
    claims = task.setdefault("claims", [])
    if claim_id is not None:
        c = next((x for x in claims if x["id"] == claim_id), None)
        if c is None:
            raise ValueError(f"no claim {claim_id}")
        why = validate(level, text or c["text"], evidence, passed)
        if why:
            raise ValueError(why)
        c.update(level=level, evidence=evidence, text=(text or c["text"]).strip())
        return c
    why = validate(level, text, evidence, passed)
    if why:
        raise ValueError(why)
    c = {"id": max([x["id"] for x in claims] or [0]) + 1, "level": level, "text": text.strip(), "evidence": evidence.strip()}
    claims.append(c)
    return c


def withdraw(task: dict, claim_id: int) -> None:
    before = len(task.get("claims", []))
    task["claims"] = [c for c in task.get("claims", []) if c["id"] != claim_id]
    if len(task["claims"]) == before:
        raise ValueError(f"no claim {claim_id}")


def open_claims(task: dict) -> list[dict]:
    return [c for c in task.get("claims", []) if c["level"] in OPEN_LEVELS]


def inferred_claims(task: dict) -> list[dict]:
    return [c for c in task.get("claims", []) if c["level"] in ("inferred", "supported")]
