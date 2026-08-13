"""Diagnose VBA RUN after compile fix."""
from __future__ import annotations

import time
from pathlib import Path

import pythoncom

import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "mcp-server"))
import ms_connect  # noqa: E402

BRIDGE = Path(r"c:\repos\microstation-vba-project\Bridge")


def main() -> None:
    pythoncom.CoInitialize()
    app = ms_connect.get_microstation_app("Test")
    vbe = app.VBE
    proj = vbe.VBProjects.Item("Test")
    cm = proj.VBComponents.Item("WZTCBridge").CodeModule
    probe = "Public Sub CursorBridgePing()"
    found = any(probe in cm.Lines(ln, 1) for ln in range(1, cm.CountOfLines + 1))
    if not found:
        cm.InsertLines(
            cm.CountOfLines + 1,
            "\n".join([
                "Public Sub CursorBridgePing()",
                '    Open "c:\\repos\\microstation-vba-project\\Bridge\\_ping.txt" For Output As #1',
                '    Print #1, "ping-ok"',
                "    Close #1",
                "End Sub",
                "",
            ]),
        )
        print("inserted CursorBridgePing")
    else:
        print("CursorBridgePing already present")

    ping = BRIDGE / "_ping.txt"
    if ping.exists():
        ping.unlink()

    app.CadInputQueue.SendKeyin("VBA RUN [Test]WZTCBridge.CursorBridgePing")
    time.sleep(1.5)
    print("ping exists", ping.exists(), ping.read_text() if ping.exists() else "")

    # Also try RunChatToolRequest
    req = BRIDGE / "chat-tool-request.tsv"
    resp = BRIDGE / "chat-tool-response.tsv"
    req.write_text("P55\tPLACE_LINE\tx1=1\ty1=1\tx2=2\ty2=2\treason=diag\n", encoding="utf-8")
    app.CadInputQueue.SendKeyin("VBA RUN [Test]WZTCBridge.RunChatToolRequest")
    time.sleep(1.5)
    print("chat-resp", resp.read_text(encoding="utf-8", errors="replace")[:200])
    pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
