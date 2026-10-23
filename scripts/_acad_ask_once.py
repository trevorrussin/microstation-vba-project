"""One AutoCAD/textbox ask, then exit — agent must read the answer and decide.

Writes scripts/_acad_ask_last.json for the Cursor agent to reason over.
Does NOT auto-chain into point picks or the next question.
"""
from __future__ import annotations

import json
import os
import sys
from datetime import datetime, timezone

sys.path.insert(0, r"c:\repos\microstation-vba-project\scripts")
from acad_ask import ask_choice

OUT = r"c:\repos\microstation-vba-project\scripts\_acad_ask_last.json"


def main():
    title = os.environ.get(
        "ACAD_ASK_TITLE",
        "Flare / CAD ask",
    )
    message = os.environ.get(
        "ACAD_ASK_MESSAGE",
        "NW flares stay untouched.\n"
        "Tell me what you want next in your own words.\n"
        "I (the agent) will read this and either act or ask a follow-up — "
        "this window will NOT auto-jump into point picks.",
    )
    options_raw = os.environ.get(
        "ACAD_ASK_OPTIONS",
        "All non-NW pockets|"
        "MidN only|"
        "MidS only|"
        "East only|"
        "MidN NORTH (35.01) only|"
        "Stop / nothing yet",
    )
    options = [o.strip() for o in options_raw.split("|") if o.strip()]

    answer = ask_choice(
        title=title,
        message=message,
        options=options,
        allow_free_text=True,
    )
    payload = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "title": title,
        "message": message,
        "options": options,
        "answer": answer,
    }
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)
    print("ANSWER_JSON", OUT)
    print("ANSWER:", answer)


if __name__ == "__main__":
    main()
