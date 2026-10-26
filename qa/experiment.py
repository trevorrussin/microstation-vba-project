"""Reversible A/B experiment for display and state questions: change ONE thing, look, and always put it back.

"Is this graphic controlled by setting X?" is answered by an intervention, not by correlation:

    1. snapshot the state            (caller-supplied: dict of whatever the caller can read back)
    2. capture the view              (caller-supplied: returns an image path)
    3. apply the one change          (caller-supplied)
    4. capture the view, snapshot again
    5. restore                       (always runs, also when step 3 or 4 raised)
    6. snapshot again and require it equals step 1

Verdicts:
    SUPPORTED     the target region changed and nothing else did (protected regions unchanged, no unexpected state keys changed)
    REFUTED       the target region did not change: the setting is not what draws it
    SIDE_EFFECT   something outside the target changed (a protected region of the picture, or a state key not named in
                  `expect_changed`): the change hid or altered more than intended, so the picture proves nothing about the target
    INCONCLUSIVE  the two captures cannot be compared (missing, different size)
    RESTORE_FAILED  the state after restore differs from the state before: stop and tell the engineer exactly what differs

Only the caller knows how to read and set the state, so nothing here talks to MicroStation. There is no per-view Level Display
read/write op in the bridge yet (LIST_LEVELS reports the global IsDisplayed flag only): until one exists the snapshot is whatever
the caller can read (describe_drawing_state, LIST_LEVELS, element fingerprints), and view toggles stay engineer GUI steps.
"""
from __future__ import annotations

from typing import Callable


def image_diff(before: str, after: str, region: tuple | None = None, threshold: int = 24) -> dict:
    """Fraction of pixels whose channel difference exceeds `threshold`, over `region` (x0, y0, x1, y1 in pixels) or the whole image."""
    import numpy as np
    from PIL import Image
    a, b = Image.open(before).convert("RGB"), Image.open(after).convert("RGB")
    if a.size != b.size:
        return {"comparable": False, "reason": f"image sizes differ {a.size} vs {b.size}"}
    x0, y0, x1, y1 = region or (0, 0, a.size[0], a.size[1])
    x0, y0, x1, y1 = max(0, int(x0)), max(0, int(y0)), min(a.size[0], int(x1)), min(a.size[1], int(y1))
    if x1 <= x0 or y1 <= y0:
        return {"comparable": False, "reason": "region is empty"}
    pa = np.asarray(a, dtype=np.int16)[y0:y1, x0:x1]
    pb = np.asarray(b, dtype=np.int16)[y0:y1, x0:x1]
    changed = np.abs(pa - pb).max(axis=2) > threshold
    n = int(changed.sum())
    box = None
    if n:
        ys, xs = np.nonzero(changed)
        box = (int(xs.min()) + x0, int(ys.min()) + y0, int(xs.max()) + x0, int(ys.max()) + y0)
    return {"comparable": True, "changed_px": n, "changed_frac": n / changed.size, "bbox": box}


def _state_diff(a: dict, b: dict) -> list[str]:
    return sorted(k for k in set(a) | set(b) if a.get(k) != b.get(k))


def run_ab(hypothesis: str, snapshot: Callable[[], dict], apply: Callable[[], None], restore: Callable[[], None],
           capture: Callable[[], str], region: tuple | None = None, protect: list[tuple] | None = None,
           expect_changed: set[str] | None = None, min_changed: float = 0.001, max_protect_changed: float = 0.001) -> dict:
    """Run the experiment. `region` is where the graphic should disappear (None = whole picture); `protect` are regions that
    must NOT change (linework that has to stay visible); `expect_changed` are the state keys the apply step is allowed to change."""
    out = {"hypothesis": hypothesis, "verdict": "INCONCLUSIVE", "notes": []}
    s0 = snapshot()
    img0 = capture()
    s1, img1, err = None, None, None
    try:
        apply()
        img1 = capture()
        s1 = snapshot()
    except Exception as ex:  # noqa: BLE001 - the restore below must still run
        err = ex
    finally:
        restore_err = None
        try:
            restore()
        except Exception as ex:  # noqa: BLE001
            restore_err = ex
    s2 = snapshot()
    left = _state_diff(s0, s2)
    if restore_err or left:
        out.update(verdict="RESTORE_FAILED", restore_diff=left)
        out["notes"].append(f"restore {'raised ' + repr(restore_err) if restore_err else 'left these state keys different'}: {left or ''}")
        return out
    if err is not None:
        out["notes"].append(f"apply/capture raised {err!r}; state restored")
        return out
    extra = [k for k in _state_diff(s0, s1) if k not in (expect_changed or set())]
    d = image_diff(img0, img1, region)
    out["target"] = d
    if not d.get("comparable"):
        out["notes"].append(d.get("reason", ""))
        return out
    prot = [image_diff(img0, img1, r) for r in (protect or [])]
    out["protected"] = prot
    if extra:
        out.update(verdict="SIDE_EFFECT", unexpected_state=extra)
        out["notes"].append(f"state keys changed that the experiment did not name: {extra}")
    elif any(p.get("comparable") and p["changed_frac"] > max_protect_changed for p in prot):
        out.update(verdict="SIDE_EFFECT")
        out["notes"].append("a protected region changed: the change hid more than the target")
    elif d["changed_frac"] >= min_changed:
        out["verdict"] = "SUPPORTED"
    else:
        out["verdict"] = "REFUTED"
    return out


def as_tried(result: dict) -> dict:
    """Row for experience memory (`tried`): the intervention and what it showed."""
    return {"what": f"A/B: {result['hypothesis']}", "result": result["verdict"] + ("; " + "; ".join(result["notes"]) if result["notes"] else "")}
