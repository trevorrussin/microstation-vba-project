"""Compare stock vs new callout anatomy."""
from __future__ import annotations

import sys

sys.path.insert(0, r"c:\repos\microstation-vba-project\mcp-server")
import pythoncom
import ms_connect
from win32com.client import Dispatch

pythoncom.CoInitialize()
_, app = next(ms_connect._candidate_apps())
mr = app.ActiveModelReference


def dump_gg(label: str, gg: int) -> None:
    print(f"=== {label} GG {gg} ===")
    o = Dispatch("MicroStationDGN.ElementScanCriteria")
    o.ExcludeNonGraphical()
    ee = mr.Scan(o)
    while ee.MoveNext():
        e = ee.Current
        try:
            if int(e.GraphicGroup) != gg:
                continue
            eid = int(e.ID64)
            t = int(e.Type)
            print(
                f" id={eid} type={t} lvl={e.Level.Name if e.Level else None} "
                f"color={e.Color} wt={e.LineWeight} hidden={e.IsHidden}"
            )
            if e.IsTextNodeElement:
                tn = e.AsTextNodeElement
                print(
                    f"  TN origin=({tn.Origin.X:.3f},{tn.Origin.Y:.3f}) "
                    f"just={tn.Justification} lines={tn.TextLinesCount} "
                    f"spacing={tn.LineSpacing}"
                )
                for i in range(1, tn.TextLinesCount + 1):
                    print(f"   L{i}: {tn.TextLine(i)!r}")
                for attr in (
                    "Height",
                    "Width",
                    "Font",
                    "TextStyleName",
                    "NodeNumber",
                    "MaxCharactersPerLine",
                ):
                    try:
                        print(f"  {attr}={getattr(tn, attr)}")
                    except Exception:
                        pass
                # try GetTextPart / TextStyle on component
                try:
                    # enumerate text parts
                    for meth in dir(tn):
                        if "Text" in meth or "Font" in meth or "Height" in meth:
                            pass
                except Exception:
                    pass
            if t == 3:
                le = e.AsLineElement
                s, en = le.StartPoint, le.EndPoint
                ln = ((en.X - s.X) ** 2 + (en.Y - s.Y) ** 2) ** 0.5
                print(
                    f"  LINE ({s.X:.3f},{s.Y:.3f})->({en.X:.3f},{en.Y:.3f}) len={ln:.3f}"
                )
            if t == 33:
                de = e.AsDimensionElement
                print(
                    f"  DIM type={de.DimensionType} primary={de.PrimaryText!r} "
                    f"refs={de.ReferencePointsCount}"
                )
                try:
                    print(f"  style={de.DimensionStyle.Name}")
                except Exception as ex:
                    print(f"  style err {ex}")
                for i in range(1, de.ReferencePointsCount + 1):
                    p = de.Point(i)
                    print(f"   pt{i}=({p.X:.3f},{p.Y:.3f})")
                r = e.Range
                print(
                    f"  range=({r.Low.X:.2f},{r.Low.Y:.2f})-"
                    f"({r.High.X:.2f},{r.High.Y:.2f})"
                )
        except Exception as ex:
            print(" err", ex)


for label, gg in [
    ("STOCK vehicular", 6119),
    ("STOCK APS post", 6117),
    ("STOCK streetlight", 6115),
    ("NEW pull", 7001),
    ("NEW conduit mid", 7002),
    ("NEW cable23", 7012),
]:
    dump_gg(label, gg)
