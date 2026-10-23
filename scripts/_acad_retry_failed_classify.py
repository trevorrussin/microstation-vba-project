"""Retry failed Location SGD sheets one-by-one with hardened COM."""
from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(r"c:\repos\microstation-vba-project")
LIST = ROOT / "outputs" / "curb-batch-retry-list.txt"
SCRIPT = ROOT / "scripts" / "_acad_batch_classify_locations.py"
LOG = ROOT / "outputs" / "curb-batch-retry-console.txt"


def main():
    names = [ln.strip() for ln in LIST.read_text(encoding="utf-8").splitlines() if ln.strip()]
    # only need unique token after Location_
    tokens = []
    for n in names:
        # HWP24QC1_Location_10_SGD.dwg -> 10_SGD or Location_10
        if "Location_" in n:
            tokens.append(n.split("Location_", 1)[1].replace(".dwg", ""))
        else:
            tokens.append(n)
    print(f"retrying {len(tokens)} sheets", flush=True)
    with LOG.open("w", encoding="utf-8") as out:
        out.write(f"retry {len(tokens)}\n")
    for tok in tokens:
        cmd = [sys.executable, str(SCRIPT), "--only", tok, "--force"]
        print(f"\n>>> RETRY {tok}", flush=True)
        with LOG.open("a", encoding="utf-8") as out:
            out.write(f"\n>>> RETRY {tok}\n")
            p = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True)
            out.write(p.stdout or "")
            out.write(p.stderr or "")
            print(p.stdout or "")
            if p.stderr:
                print(p.stderr)
            out.write(f"exit={p.returncode}\n")
        time.sleep(4.0)
    print("RETRY BATCH DONE", flush=True)


if __name__ == "__main__":
    main()
