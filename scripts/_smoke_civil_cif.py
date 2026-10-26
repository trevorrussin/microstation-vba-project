"""Live smoke of Civil CIF reads after Bridge Import + new impl."""
from __future__ import annotations

import json
import pythoncom

pythoncom.CoInitialize()
from bridge_client import Bridge  # noqa: E402

b = Bridge()


def call(cmd: str, **args):
    return b.call("GET_CIF", cmd=cmd, json=json.dumps(args, separators=(",", ":")))


def main() -> int:
    ping = call("ping")
    print("ping", {k: ping.get(k) for k in ("status", "impl", "connection", "geometricModel")})
    if ping.get("status") != "OK":
        print("FAIL ping", ping)
        return 1

    for cmd, args in (
        ("list_alignments", {}),
        ("list_corridors", {}),
        ("list_terrains", {}),
        ("get_active_terrain", {}),
        ("list_profiles", {}),
        ("get_corridor_properties", {"idx": 0}),
    ):
        r = call(cmd, **args)
        status = r.get("status")
        n = len(r.get("rows") or [])
        note = r.get("note") or ""
        print(f"{cmd}: status={status} rows={n} note={note[:80]}")
        if status != "OK" and cmd != "get_corridor_properties":
            # empty corridors may still be OK with 0 data rows
            print("  raw keys", list(r.keys())[:12])
        if cmd == "list_alignments" and n:
            print("  first", r["rows"][0])
        if cmd == "get_active_terrain":
            print("  rows", r.get("rows"))
        if cmd == "get_corridor_properties":
            print("  sample", (r.get("rows") or [])[:3], "err", note)

    # new mutate cmds present? unknown cmd vs missing corridor
    r = call("get_corridor_properties", corridorIdx=0)
    print("get_corridor_properties corridorIdx=0:", r.get("status"), r.get("note"), "rows", len(r.get("rows") or []))

    # Bridge level display
    r2 = b.call("GET_LEVEL_DISPLAY", level="Default", viewIndex=1)
    print("GET_LEVEL_DISPLAY:", r2.get("status"), r2.get("note") or list(r2.keys())[:8])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
