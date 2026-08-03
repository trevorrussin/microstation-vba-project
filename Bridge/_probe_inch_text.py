"""Probe TEXTEDITOR INSERT_TEXT inch-mark strategies."""
from __future__ import annotations

import time

import pythoncom
import win32com.client
from win32com.client import Dispatch

pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("MicroStationDGN.Application")
mr = app.ActiveModelReference
q = chr(34)
x, y = 1022500.0, 217100.0


def read_texts_near(cx, cy, tol=40):
    oScan = Dispatch("MicroStationDGN.ElementScanCriteria")
    oScan.ExcludeNonGraphical()
    ee = mr.Scan(oScan)
    out = []
    while ee.MoveNext():
        el = ee.Current
        try:
            is_text = el.IsTextElement or el.IsTextNodeElement
        except Exception:
            continue
        if not is_text:
            continue
        rng = el.Range
        tx = (float(rng.High.X) + float(rng.Low.X)) / 2
        ty = (float(rng.High.Y) + float(rng.Low.Y)) / 2
        if abs(tx - cx) > tol or abs(ty - cy) > tol:
            continue
        content = "?"
        try:
            if el.IsTextElement:
                content = el.AsTextElement.Text
            else:
                # text node — concatenate lines if possible
                tn = el.AsTextNodeElement
                parts = []
                for i in range(1, tn.TextLinesCount + 1):
                    parts.append(tn.TextLine(i))
                content = " | ".join(parts)
        except Exception as e:
            content = f"err:{e}"
        out.append((content, el.Color, tx, ty))
    return out


def place_doubled(px, py):
    size = "36" + q + " x 36" + q
    esc = size.replace(q, q + q)
    keyin = "TEXTEDITOR PLAYCOMMAND INSERT_TEXT " + q + esc + q
    app.CadInputQueue.SendCommand("TEXTEDITOR PLACE")
    app.CadInputQueue.SendKeyin("TEXTEDITOR PLAYCOMMAND INSERT_TEXT " + q + "PROBE-A" + q)
    app.CadInputQueue.SendCommand(
        "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP "
        "SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    )
    app.CadInputQueue.SendKeyin(keyin)
    app.CadInputQueue.SendDataPoint(app.Point3dFromXYZ(px, py, 0), 1)
    app.CadInputQueue.SendReset()
    print("A keyin", repr(keyin))


def place_piecewise(px, py):
    inch_keyin = "TEXTEDITOR PLAYCOMMAND INSERT_TEXT " + q + q + q  # INSERT_TEXT """
    app.CadInputQueue.SendCommand("TEXTEDITOR PLACE")
    app.CadInputQueue.SendKeyin("TEXTEDITOR PLAYCOMMAND INSERT_TEXT " + q + "PROBE-B" + q)
    app.CadInputQueue.SendCommand(
        "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP "
        "SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    )
    app.CadInputQueue.SendKeyin("TEXTEDITOR PLAYCOMMAND INSERT_TEXT " + q + "36" + q)
    print("B inch", repr(inch_keyin))
    app.CadInputQueue.SendKeyin(inch_keyin)
    app.CadInputQueue.SendKeyin("TEXTEDITOR PLAYCOMMAND INSERT_TEXT " + q + " x 36" + q)
    app.CadInputQueue.SendKeyin(inch_keyin)
    app.CadInputQueue.SendDataPoint(app.Point3dFromXYZ(px, py, 0), 1)
    app.CadInputQueue.SendReset()


def place_prime(px, py):
    prime = "\u2033"  # double prime
    size = "48" + prime + " x 48" + prime
    keyin = "TEXTEDITOR PLAYCOMMAND INSERT_TEXT " + q + size + q
    app.CadInputQueue.SendCommand("TEXTEDITOR PLACE")
    app.CadInputQueue.SendKeyin("TEXTEDITOR PLAYCOMMAND INSERT_TEXT " + q + "PROBE-C" + q)
    app.CadInputQueue.SendCommand(
        "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP "
        "SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    )
    app.CadInputQueue.SendKeyin(keyin)
    app.CadInputQueue.SendDataPoint(app.Point3dFromXYZ(px, py, 0), 1)
    app.CadInputQueue.SendReset()
    print("C keyin", repr(keyin))


def place_color0(px, py):
    app.CadInputQueue.SendKeyin("ACTIVE COLOR 0")
    app.CadInputQueue.SendKeyin("ACTIVE LEVEL Default")
    app.CadInputQueue.SendCommand("TEXTEDITOR PLACE")
    app.CadInputQueue.SendKeyin("TEXTEDITOR PLAYCOMMAND INSERT_TEXT " + q + "PROBE-D-WHITE" + q)
    app.CadInputQueue.SendCommand(
        "TEXTEDITOR PLAYCOMMAND KEY_DOWN KEY_CODE 0x06 CONTROL_KEY_STATE UP "
        "SHIFT_KEY_STATE UP ALT_KEY_STATE UP"
    )
    prime = "\u2033"
    app.CadInputQueue.SendKeyin(
        "TEXTEDITOR PLAYCOMMAND INSERT_TEXT " + q + "48" + prime + " x 48" + prime + q
    )
    app.CadInputQueue.SendDataPoint(app.Point3dFromXYZ(px, py, 0), 1)
    app.CadInputQueue.SendReset()


app.CadInputQueue.SendKeyin("ACTIVE COLOR 0")
place_doubled(x, y)
time.sleep(0.4)
place_piecewise(x, y - 30)
time.sleep(0.4)
place_prime(x, y - 60)
time.sleep(0.4)
place_color0(x, y - 90)
time.sleep(0.5)

print("\nResults:")
for dy in (0, -30, -60, -90):
    hits = read_texts_near(x, y + dy, tol=25)
    for h in hits:
        print(repr(h[0]), "color", h[1], "at", h[2], h[3])
