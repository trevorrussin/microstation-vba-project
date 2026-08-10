import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "mcp-server"))
import chat_history  # noqa: E402

bridge = Path(__file__).resolve().parent.parent / "Bridge"
hist = bridge / "chat-history.json"
bak = bridge / (
    "chat-history.pre-tooluse-repair."
    + datetime.now().strftime("%Y%m%d-%H%M%S")
    + ".bak.json"
)
shutil.copy2(hist, bak)
print("backed up", bak)
msgs = json.loads(hist.read_text(encoding="utf-8"))
print("before", len(msgs), [m.get("role") for m in msgs])
chat_history._repair_tool_pairing(msgs)
print("after repair", len(msgs), [m.get("role") for m in msgs])
# Fresh sheet test: clear so prior FINAL does not confuse the new build.
hist.write_text("[]\n", encoding="utf-8")
print("cleared history for fresh 619-311 run")
