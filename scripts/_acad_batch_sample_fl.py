"""Batch-sample many AutoCAD curves in chunked LISP SendCommand calls."""
from __future__ import annotations

import os
import time
import pythoncom
from win32com.client import Dispatch

OUT = r"c:\repos\microstation-vba-project\scripts\_fl_batch_pts.txt"
OUT_LISP = OUT.replace("\\", "/")
MAX_LISP_CHARS = 1800


def _lisp_chunk(handles, step: float, file_mode: str, write_done: bool):
    parts = [
        "(progn (vl-load-com)",
        f'(setq f (open "{OUT_LISP}" "{file_mode}"))',
        f"(setq step {step:.4f})",
    ]
    for h in handles:
        parts.append(
            f'(progn (setq e (handent "{h}")) '
            f"(if e (progn "
            f"(setq len (vlax-curve-getDistAtParam e (vlax-curve-getEndParam e))) "
            f"(setq d 0.0) "
            f"(while (<= d (+ len 1e-6)) "
            f"(setq p (vlax-curve-getPointAtDist e (min d len))) "
            f'(if p (write-line (strcat "{h}," (rtos (car p) 2 6) "," '
            f'(rtos (cadr p) 2 6)) f)) '
            f"(setq d (+ d step))) "
            f"))) "
        )
    if write_done:
        parts.append('(write-line "DONE" f) (close f) (princ))')
    else:
        parts.append("(close f) (princ))")
    return " ".join(parts) + " "


def _split_handles(handles):
    chunks = []
    cur = []
    for h in handles:
        trial = cur + [h]
        if cur and len(_lisp_chunk(trial, 0.4, "w", True)) > MAX_LISP_CHARS:
            chunks.append(cur)
            cur = [h]
        else:
            cur = trial
    if cur:
        chunks.append(cur)
    return chunks


def sample_many(doc, handles, step: float = 0.35, timeout: float = 20.0):
    """Return {handle: [(x,y), ...]} using chunked LISP batch writes."""
    handles = [h for h in handles if h]
    if not handles:
        return {}
    if os.path.exists(OUT):
        try:
            os.remove(OUT)
        except Exception:
            pass

    chunks = _split_handles(handles)
    per_chunk = max(5.0, timeout / max(len(chunks), 1))
    for i, chunk in enumerate(chunks):
        mode = "w" if i == 0 else "a"
        done = i == len(chunks) - 1
        lisp = _lisp_chunk(chunk, step, mode, done)
        doc.SendCommand(lisp)
        t0 = time.time()
        while time.time() - t0 < per_chunk:
            if os.path.exists(OUT):
                if not done:
                    break
                with open(OUT, encoding="utf-8", errors="ignore") as f:
                    tail = f.read()
                if "DONE" in tail.splitlines()[-5:]:
                    break
            time.sleep(0.2)

    if not os.path.exists(OUT):
        raise FileNotFoundError(f"batch sample produced no file: {OUT}")

    result = {h: [] for h in handles}
    with open(OUT, encoding="utf-8", errors="ignore") as f:
        for line in f:
            if line.strip() == "DONE":
                continue
            a = line.strip().split(",")
            if len(a) >= 3:
                h, x, y = a[0], float(a[1]), float(a[2])
                if h in result:
                    pts = result[h]
                    if not pts or abs(pts[-1][0] - x) > 1e-4 or abs(pts[-1][1] - y) > 1e-4:
                        pts.append((x, y))
    return result


if __name__ == "__main__":
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    doc = acad.ActiveDocument
    hs = ["13D80A", "13D2EA", "13D425", "13A7EA"]
    d = sample_many(doc, hs, step=0.5)
    for h, pts in d.items():
        print(h, len(pts))
