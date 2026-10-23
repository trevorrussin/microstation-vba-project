"""Replace text on existing BTS page-3 callouts from green EDIT FreeTexts only."""
from __future__ import annotations

import sys
import time

import pythoncom

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import ms_connect

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
mr = app.ActiveModelReference
q = app.CadInputQueue


def get_tn(host_id: int):
    e = mr.GetElementByID64(host_id)
    if e.IsTextNodeElement:
        return e
    if int(e.Type) == 2:
        ee = e.AsComplexElement.GetSubElements()
        while ee.MoveNext():
            if ee.Current.IsTextNodeElement:
                return ee.Current
    return None


def read_lines(tn_el):
    tn = tn_el.AsTextNodeElement
    return [tn.TextLine(i) for i in range(1, int(tn.TextLinesCount) + 1)]


def set_lines(tn_el, lines):
    tn = tn_el.AsTextNodeElement
    n = int(tn.TextLinesCount)
    for i, line in enumerate(lines, start=1):
        if i <= n:
            tn.SetTextLine(i, line)
        else:
            tn.AddTextLine(line)
    for i in range(len(lines) + 1, n + 1):
        tn.SetTextLine(i, " ")
    tn_el.Rewrite()


# Final text for each existing host (EDIT chips only — no NEW callouts).
# Sources:
#   [0]/[58] -> 6132 (+ same pay items already on 5474 S-1)
#   [44] -> 6120
#   [45]/[47]/[49] -> 6126 / 6129 / 6280
#   [51]/[52] -> 3770  (818122011 (4); bare (1) on 680.8223)
#   [43]/[38] -> 6214  (818122011 (5); (2) already on 83720011)

FINAL = {
    6132: [
        "VEHICULAR SIGNAL HEADS AND APS PUSH BUTTON",
        "(ITEM NOS. 680.818122011 (1), 680.81822011 (1), 680.8223 (1), 680.8151 (1))",
    ],
    5474: [
        "S-1 POLE WITH VEHICULAR HEAD ",
        "(ITEM NOS. 680.5001, 680.20720011 (1), 680.8223 (1),",
        "680.818122011 (1), 680.81822011 (1))",
        "(SEE NOTE 3)",
    ],
    6120: [
        "STREETLIGHT W/PEDESTRIAN TRAFFIC SIGNALS AND APS PUSH BUTTON",
        "(ITEM NOS. 680.8141 (2), 680.815002 (2), 680.8151 (1), 680.83080011 (1))",
    ],
    6126: [
        "APS W/POST",
        "(ITEM NOS. 680.5001, 680.8152 (1))",
    ],
    6129: [
        "APS W/POST",
        "(ITEM NOS. 680.5001, 680.8152 (1))",
    ],
    6280: [
        "APS W/POST",
        "(ITEM NOS. 680.5001, 680.8152 (1))",
    ],
    3770: [
        "25' TRAFFIC SIGNAL MAST ARM WITH VEHICULAR",
        "AND PEDESTRIAN TRAFFIC SIGNALS",
        "AND APS PUSH BUTTON",
        "(ITEM NOS. 680.5001, 680.8141 (2), 680.815002 (2),",
        "680.8151 (1), 680.818122011 (4), 680.81822011 (4),",
        "680.82050011, 680.8223 (1), 680.82310011, 680.83720011)",
        "(SEE NOTE 3)",
    ],
    6214: [
        "DUAL 25' TRAFFIC SIGNAL MAST ARM WITH VEHICULAR HEADS",
        "(ITEM NOS. 680.5001, 680.818122011 (5), 680.81822011 (5),",
        "680.82050011, 680.8223, 680.82310011, 680.83720011(2))",
        "(SEE NOTE 3)",
    ],
}

print("=== BEFORE ===")
for hid in FINAL:
    tn = get_tn(hid)
    print(f"\n--- {hid} ---")
    for L in read_lines(tn):
        print(" ", L)

print("\n=== APPLY ===")
for hid, lines in FINAL.items():
    tn = get_tn(hid)
    before = read_lines(tn)
    set_lines(tn, lines)
    after = read_lines(get_tn(hid))
    changed = before != after
    print(f"host {hid}: {'UPDATED' if changed else 'already matched'}")
    for L in after:
        print(" ", L)

q.SendKeyin("SAVE DESIGN")
time.sleep(0.4)
print("\nSAVED")
