"""Resume paused curb-batch retries. Respects outputs/curb-batch-paused.txt."""
from __future__ import annotations

import os
import subprocess
import sys
import time

ROOT = r"c:\repos\microstation-vba-project"
OUT = os.path.join(ROOT, "outputs")
PAUSE = os.path.join(OUT, "curb-batch-paused.txt")
LIST = os.path.join(OUT, "curb-batch-retry-list.txt")
LIVE = os.path.join(OUT, "curb-batch-retry-live.txt")
SCRIPT = os.path.join(ROOT, "scripts", "_acad_batch_classify_locations.py")


def paused() -> bool:
    if not os.path.isfile(PAUSE):
        return False
    try:
        return open(PAUSE, encoding="utf-8").read().strip().lower().startswith("paused")
    except Exception:
        return False


def main() -> int:
    sheets = [
        ln.strip()
        for ln in open(LIST, encoding="utf-8")
        if ln.strip() and not ln.strip().startswith("#")
    ]
    # Clear pause so we actually run
    open(PAUSE, "w", encoding="utf-8").write("running\n")
    with open(LIVE, "w", encoding="utf-8") as live:
        live.write(f"retrying {len(sheets)} sheets\n")
        live.flush()
        print(f"retrying {len(sheets)} sheets", flush=True)
        fails = []
        for name in sheets:
            if paused():
                live.write("PAUSED by engineer\n")
                print("PAUSED", flush=True)
                return 2
            tag = name.replace("HWP24QC1_Location_", "").replace(".dwg", "")
            live.write(f"\n>>> RETRY {tag}\n")
            live.flush()
            print(f"\n>>> RETRY {tag}", flush=True)
            # --only matches basename substring; force reclassify
            only = tag  # e.g. 10_SGD or 11-12_SGD
            cmd = [
                sys.executable,
                "-u",
                SCRIPT,
                "--only",
                only,
                "--force",
            ]
            p = subprocess.run(
                cmd,
                cwd=ROOT,
                stdout=live,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            live.write(f"exit={p.returncode}\n")
            live.flush()
            if p.returncode != 0:
                fails.append(name)
                live.write(
                    "BLOCKED — Failure Recovery Loop; stopping retry batch\n"
                )
                live.flush()
                open(PAUSE, "w", encoding="utf-8").write(
                    f"blocked\nfile={name}\n"
                    "ACTION: diagnose before next sheet\n"
                )
                print(f"BLOCKED on {name} — stopping", flush=True)
                return 1
            time.sleep(2.0)
        live.write(f"\n===== RETRY DONE fails={len(fails)} =====\n")
        for f in fails:
            live.write(f"  FAIL {f}\n")
        live.flush()
    open(PAUSE, "w", encoding="utf-8").write("idle\n")
    print(f"RETRY DONE fails={len(fails)}", flush=True)
    return 1 if fails else 0


if __name__ == "__main__":
    raise SystemExit(main())
