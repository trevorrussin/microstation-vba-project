"""Harden batch COM + never reuse FL across islands."""
from __future__ import annotations

from pathlib import Path

P = Path(r"c:\repos\microstation-vba-project\scripts\_acad_batch_classify_locations.py")
text = P.read_text(encoding="utf-8")

# Insert com_retry helper after log()
HELPER = '''
RPC_E_CALL_REJECTED = -2147418111


def com_retry(fn, *args, tries: int = 12, label: str = "COM", **kwargs):
    """Retry on AutoCAD 'call rejected by callee' / busy."""
    last = None
    for attempt in range(tries):
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass
        try:
            return fn(*args, **kwargs)
        except Exception as ex:
            last = ex
            msg = str(ex)
            busy = (
                getattr(ex, "args", None)
                and len(ex.args) > 0
                and ex.args[0] == RPC_E_CALL_REJECTED
            ) or "rejected" in msg.lower() or "busy" in msg.lower() or "<unknown>" in msg
            if not busy and attempt > 2:
                raise
            log(f"  {label} retry {attempt+1}/{tries}: {ex}")
            time.sleep(1.5 + attempt * 0.75)
    raise RuntimeError(f"{label} failed after retries: {last}")


def close_all_except_keep(acad, keep_substrings=("01-02", "RECOVERED")) -> None:
    """Ensure only Loc01 (or keep list) stays open — leftover docs cause COM rejects."""
    for _ in range(8):
        try:
            n = acad.Documents.Count
        except Exception:
            time.sleep(2.0)
            continue
        closed_any = False
        for i in range(n - 1, -1, -1):
            try:
                d = acad.Documents.Item(i)
                name = d.Name
            except Exception:
                continue
            if any(k.upper() in name.upper() for k in keep_substrings):
                continue
            try:
                d.Close(False)
                log(f"  preclose leftover {name}")
                closed_any = True
            except Exception as ex:
                log(f"  preclose FAIL {name}: {ex}")
            time.sleep(0.5)
        if not closed_any:
            break
        time.sleep(1.0)

'''

if "def com_retry(" not in text:
    text = text.replace(
        "def log(msg: str) -> None:\n    print(msg, flush=True)\n",
        "def log(msg: str) -> None:\n    print(msg, flush=True)\n" + HELPER,
    )

# Fix pick_inner signature to accept exclude
old_sig = "def pick_inner_fl(fls: list[dict], island: list[dict]) -> tuple[dict | None, list[str]]:"
new_sig = "def pick_inner_fl(fls: list[dict], island: list[dict], exclude: set[str] | None = None) -> tuple[dict | None, list[str]]:"
if old_sig in text:
    text = text.replace(old_sig, new_sig, 1)
    # add exclude filter after near = []
    needle = '''    near = []
    for f in fls:
        d = math.hypot(f["cx"] - icx, f["cy"] - icy)
'''
    repl = '''    exclude = exclude or set()
    near = []
    for f in fls:
        if f["handle"] in exclude:
            continue
        d = math.hypot(f["cx"] - icx, f["cy"] - icy)
'''
    if needle in text:
        text = text.replace(needle, repl, 1)

# Call site: pass used_inners + used_outers as exclude
text = text.replace(
    "        inner, outers = pick_inner_fl(fls, island)",
    "        inner, outers = pick_inner_fl(fls, island, exclude=set(used_inners) | set(used_outers))",
)

# After ok, mark outers used too so next island can't pick them
text = text.replace(
    """        if ok:
            claimed.update(used)
            used_inners.append(h)
            # Do NOT hide sources/ramps — engineer needs color-coded geometry visible
            if HIDE_SOURCES:
                v7.hide_entity(doc, h, f"island{ii+1} classified")
""",
    """        if ok:
            claimed.update(used)
            used_inners.append(h)
            used_outers.extend(outers)
            # Do NOT hide sources/ramps — engineer needs color-coded geometry visible
            if HIDE_SOURCES:
                v7.hide_entity(doc, h, f"island{ii+1} classified")
""",
)

# Harden open_doc: preclose leftovers, activate, longer waits
OLD_OPEN = '''def open_doc(acad, path: str):
    key = os.path.basename(path).lower()
    last = None
    for attempt in range(10):
        try:
            n = acad.Documents.Count
            for i in range(n):
                d = acad.Documents.Item(i)
                if d.Name.lower() == key:
                    d.Activate()
                    return d, "already_open"
            d = acad.Documents.Open(path, False)
            return d, "opened"
        except Exception as ex:
            last = ex
            log(f"  Open/enum retry {attempt+1}/10: {ex}")
            time.sleep(2.0 + attempt * 0.5)
    raise RuntimeError(f"Open failed after retries: {last}")
'''

NEW_OPEN = '''def open_doc(acad, path: str):
    key = os.path.basename(path).lower()
    close_all_except_keep(acad)
    last = None
    for attempt in range(12):
        try:
            pythoncom.PumpWaitingMessages()
            n = acad.Documents.Count
            for i in range(n):
                d = acad.Documents.Item(i)
                if d.Name.lower() == key:
                    d.Activate()
                    time.sleep(0.5)
                    return d, "already_open"
            d = acad.Documents.Open(path, False)
            time.sleep(1.0)
            try:
                d.Activate()
            except Exception:
                pass
            return d, "opened"
        except Exception as ex:
            last = ex
            log(f"  Open/enum retry {attempt+1}/12: {ex}")
            time.sleep(2.5 + attempt * 0.8)
            close_all_except_keep(acad)
    raise RuntimeError(f"Open failed after retries: {last}")
'''

if OLD_OPEN in text:
    text = text.replace(OLD_OPEN, NEW_OPEN, 1)
else:
    print("WARN: open_doc block not exact match")

# Longer close settle
text = text.replace(
    """def close_doc(doc) -> None:
    try:
        doc.Close(False)  # False = do not save again (already tried Save)
        log("  closed")
    except Exception as ex:
        log(f"  close FAIL: {ex}")
    time.sleep(1.0)
""",
    """def close_doc(doc) -> None:
    try:
        doc.Close(False)  # False = do not save again (already tried Save)
        log("  closed")
    except Exception as ex:
        log(f"  close FAIL: {ex}")
    time.sleep(2.5)
    try:
        pythoncom.PumpWaitingMessages()
    except Exception:
        pass
""",
)

# Longer between sheets
text = text.replace(
    "        time.sleep(1.5)  # let AutoCAD settle between sheets",
    "        time.sleep(3.0)  # let AutoCAD settle between sheets",
)

# Wrap ensure_layers and classify scan in com_retry via classify_doc start
# Soften Esc send - already try/except

P.write_text(text, encoding="utf-8")
print("hardened", P)
