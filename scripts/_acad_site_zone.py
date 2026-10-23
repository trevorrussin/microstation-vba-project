"""Site-zone filter for HWP24QC1 Location SGD hosts.

SGD hosts were seeded from one template and carry foreign content:
  - a PROP-D ped-ramp kit (*U226..*U231, *U273, *U274) + PRO-S polylines
    E52F / FC58, copied and moved per sheet, often hundreds of ft off-site
  - survey drop/TC blocks from other locations (~1012000,207000 / ~1002000,210000)

Authority for "this sheet's site" = the sheet's survey xref (T4318* / Survey*).
A ramp / FL is on-site only if the survey has dense content around it.
Loc20 evidence (2026-10-01): real site cells 300-830 survey ents per 100 ft cell;
template kit 560 ft away = 0.
"""
from __future__ import annotations

import time

from pywintypes import com_error

CELL_FT = 100.0
MIN_SURVEY_ENTS = 100  # summed over the 3x3 cell neighborhood (~150 ft)
BUSY = -2147418111


def _r(fn, tries=100, delay=0.4):
    last = None
    for _ in range(tries):
        try:
            return fn()
        except com_error as e:
            last = e
            if e.args and e.args[0] == BUSY:
                time.sleep(delay)
                continue
            raise
    raise last  # type: ignore[misc]


def is_survey_xref(name: str, path: str) -> bool:
    blob = f"{name}|{path}".upper()
    return "T4318" in blob or name.upper().startswith("SURVEY")


def survey_cells(doc, log=print) -> dict | None:
    """Return {(ix,iy): count} of survey-xref entity centers, or None if no survey."""
    cells: dict[tuple[int, int], int] = {}
    found = False
    for j in range(_r(lambda: doc.Blocks.Count)):
        try:
            b = _r(lambda j=j: doc.Blocks.Item(j))
            if not _r(lambda: b.IsXRef):
                continue
            name = _r(lambda: b.Name)
            path = _r(lambda: b.Path or "")
        except Exception:
            continue
        if not is_survey_xref(name, path):
            continue
        n = _r(lambda: b.Count)
        if n <= 0:
            log(f"    site-zone: survey xref {name} empty/unloaded")
            continue
        found = True
        for i in range(n):
            try:
                mn, mx = _r(lambda i=i: b.Item(i).GetBoundingBox())
                cx = (mn[0] + mx[0]) / 2.0
                cy = (mn[1] + mx[1]) / 2.0
            except Exception:
                continue
            k = (int(cx // CELL_FT), int(cy // CELL_FT))
            cells[k] = cells.get(k, 0) + 1
        log(f"    site-zone: survey xref {name} ents={n}")
    return cells if found else None


def density(cells: dict, x: float, y: float) -> int:
    ix, iy = int(x // CELL_FT), int(y // CELL_FT)
    return sum(cells.get((ix + dx, iy + dy), 0) for dx in (-1, 0, 1) for dy in (-1, 0, 1))


def on_site(cells: dict, x: float, y: float) -> bool:
    return density(cells, x, y) >= MIN_SURVEY_ENTS


def corner_labels(doc, log=print) -> list[tuple[str, float, float]]:
    """(corner_id, x, y) from 'CORNER ID ... 10xxxxx' labels in the plan Location xref.

    Point = MLeader arrowhead (first leader vertex) when present, else text bbox center.
    Used to assign islands to Corner IDs by proximity instead of sort order.
    """
    import re

    from win32com.client.dynamic import Dispatch as DD

    pat = re.compile(r"CORNER\s*ID[^0-9]*(?:\(\s*\d+\s*\))?[^0-9]*(\d{7})", re.I)
    out: list[tuple[str, float, float]] = []
    for j in range(_r(lambda: doc.Blocks.Count)):
        try:
            b = _r(lambda j=j: doc.Blocks.Item(j))
            if not _r(lambda: b.IsXRef):
                continue
            name = _r(lambda: b.Name)
            path = _r(lambda: b.Path or "")
        except Exception:
            continue
        if is_survey_xref(name, path):
            continue
        n = _r(lambda: b.Count)
        for i in range(n):
            try:
                e = DD(_r(lambda i=i: b.Item(i)))
                on = e.ObjectName
                if on not in ("AcDbMLeader", "AcDbMText", "AcDbText"):
                    continue
                m = pat.search(e.TextString or "")
                if not m:
                    continue
                x = y = None
                if on == "AcDbMLeader":
                    try:
                        v = e.GetLeaderLineVertices(0)
                        x, y = float(v[0]), float(v[1])
                    except Exception:
                        pass
                if x is None:
                    mn, mx = e.GetBoundingBox()
                    x, y = (mn[0] + mx[0]) / 2.0, (mn[1] + mx[1]) / 2.0
                out.append((m.group(1), x, y))
            except Exception:
                continue
    log(f"    corner labels: {[(c, round(x), round(y)) for c, x, y in out]}")
    return out
