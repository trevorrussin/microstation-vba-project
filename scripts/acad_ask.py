"""General-purpose AutoCAD ask (WZTC-style).

HARD RULE: every prompt shows a real textbox the engineer can type into.
Suggested buttons only fill the box — they never replace free text.
No getkword-only / alert-only asks without a typeable field.

API:
  ask_message  — ack; optional notes in textbox (blank OK)
  ask_text     — free string (required unless allow_empty)
  ask_choice   — suggestions + free textbox (always typeable)
  ask_points   — textbox gate first (skip/cancel/notes), then AutoCAD picks
"""
from __future__ import annotations

import os
import time
import pythoncom
from win32com.client import Dispatch

SCRIPTS = r"c:\repos\microstation-vba-project\scripts"
JOB_LSP = os.path.join(SCRIPTS, "_acad_ask_job.lsp")
RESULT = os.path.join(SCRIPTS, "_acad_ask_result.txt")
RESULT_LISP = RESULT.replace("\\", "/")


def _escape_lisp_str(s: str) -> str:
    return (
        s.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r\n", "\\n")
        .replace("\n", "\\n")
    )


def _acad():
    pythoncom.CoInitialize()
    acad = Dispatch("AutoCAD.Application")
    acad.Visible = True
    try:
        import win32con
        import win32gui

        hwnd = int(acad.HWND)
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    doc = None
    last = None
    for _ in range(60):
        try:
            doc = acad.ActiveDocument
            _ = doc.Name
            break
        except Exception as ex:
            last = ex
            time.sleep(0.5)
    if doc is None:
        raise RuntimeError(
            "AutoCAD is busy (modal prompt still open). "
            "Press Esc in AutoCAD until idle, then retry."
        ) from last
    return acad, doc


def _send(doc, cmd: str, tries: int = 20):
    last = None
    for _ in range(tries):
        try:
            doc.SendCommand(cmd)
            return
        except Exception as ex:
            last = ex
            time.sleep(0.5)
    raise RuntimeError(
        "AutoCAD rejected SendCommand — press Esc until idle, then retry."
    ) from last


def _wait_result(timeout: float = 600.0) -> str:
    t0 = time.time()
    while time.time() - t0 < timeout:
        if os.path.exists(RESULT) and os.path.getsize(RESULT) > 2:
            time.sleep(0.35)
            with open(RESULT, encoding="utf-8") as f:
                txt = f.read().strip()
            if txt.startswith("CANCEL"):
                raise RuntimeError("user cancelled AutoCAD ask")
            if (
                txt.startswith("OK")
                or txt.startswith("POINT")
                or txt.startswith("CHOICE")
                or txt.startswith("TEXT")
            ):
                return txt
        time.sleep(0.4)
    raise TimeoutError("AutoCAD ask timed out — no result file")


def _run_job_lsp(body: str, timeout: float = 600.0) -> str:
    if os.path.exists(RESULT):
        try:
            os.remove(RESULT)
        except Exception:
            pass
    with open(JOB_LSP, "w", encoding="utf-8") as f:
        f.write(body)
    path = JOB_LSP.replace("\\", "/")
    _, doc = _acad()
    _send(doc, f'(load "{path}") ')
    return _wait_result(timeout)


def _textbox_dialog(
    *,
    title: str,
    message: str,
    options: list[str] | None = None,
    prompt: str = "Type here (always available)",
    allow_empty: bool = False,
    ok_label: str = "OK",
) -> str:
    """Topmost window that ALWAYS includes a typeable textbox."""
    import tkinter as tk
    from tkinter import ttk

    options = options or []
    result: list[str | None] = []

    root = tk.Tk()
    root.title(title)
    root.attributes("-topmost", True)
    root.resizable(True, True)
    root.geometry("600x560+180+60")

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frm, text=title, font=("Segoe UI", 12, "bold")).pack(anchor="w")

    msg = tk.Text(frm, height=9, wrap=tk.WORD, font=("Segoe UI", 10))
    msg.insert("1.0", message)
    msg.insert(
        tk.END,
        "\n\n---\nYou can always type freely in the box below "
        "(suggestions only fill the box; they do not lock you in).",
    )
    if options:
        msg.insert(tk.END, "\n\nSuggestions:\n")
        for i, opt in enumerate(options, 1):
            msg.insert(tk.END, f"  {i}. {opt}\n")
    msg.configure(state=tk.DISABLED)
    msg.pack(fill=tk.BOTH, expand=True, pady=(8, 8))

    ttk.Label(frm, text=prompt, font=("Segoe UI", 10, "bold")).pack(anchor="w")
    entry = ttk.Entry(frm, font=("Segoe UI", 11))
    entry.pack(fill=tk.X, pady=(4, 8))
    entry.focus_set()

    if options:
        btn_row = ttk.Frame(frm)
        btn_row.pack(fill=tk.X, pady=(0, 8))

        def fill(opt: str):
            entry.delete(0, tk.END)
            entry.insert(0, opt)
            entry.focus_set()

        for i, opt in enumerate(options, 1):
            ttk.Button(
                btn_row,
                text=f"{i}. {opt}",
                command=lambda o=opt: fill(o),
            ).pack(fill=tk.X, pady=1)

    def submit(_event=None):
        val = entry.get().strip()
        if not val and not allow_empty:
            return
        result.append(val)
        root.destroy()

    def cancel():
        result.append(None)
        root.destroy()

    actions = ttk.Frame(frm)
    actions.pack(fill=tk.X)
    ttk.Button(actions, text=ok_label, command=submit).pack(side=tk.RIGHT, padx=4)
    ttk.Button(actions, text="Cancel", command=cancel).pack(side=tk.RIGHT)
    root.bind("<Return>", submit)
    root.bind("<Escape>", lambda e: cancel())

    try:
        root.lift()
        root.focus_force()
    except Exception:
        pass
    root.mainloop()

    if not result or result[0] is None:
        raise RuntimeError("user cancelled AutoCAD ask")
    return result[0]


def ask_message(*, title: str, message: str) -> str:
    """Ack dialog — always has a textbox. Blank notes + OK is fine.

    Returns typed notes (may be empty string).
    """
    return _textbox_dialog(
        title=title,
        message=message,
        options=["OK — continue", "Stop / cancel"],
        prompt="Type notes here, or leave blank and OK to continue",
        allow_empty=True,
        ok_label="OK",
    )


def ask_text(
    *,
    title: str,
    message: str,
    prompt: str = "Type your answer here",
    allow_empty: bool = False,
) -> str:
    """Free-typed string via on-screen textbox (always present)."""
    return _textbox_dialog(
        title=title,
        message=message,
        prompt=prompt,
        allow_empty=allow_empty,
    )


def ask_choice(
    *,
    title: str,
    message: str,
    options: list[str] | None = None,
    allow_free_text: bool = True,
) -> str:
    """Suggestions + mandatory free textbox. Typing always wins.

    allow_free_text is kept for callers but textbox is always shown.
    """
    del allow_free_text  # always on
    options = options or []
    raw = _textbox_dialog(
        title=title,
        message=message,
        options=options,
        prompt="Type your answer here (number, suggestion, or anything else)",
        allow_empty=False,
    )
    if raw.isdigit() and options:
        idx = int(raw) - 1
        if 0 <= idx < len(options):
            return options[idx]
    low = raw.lower()
    for opt in options:
        if low == opt.lower() or opt.lower().startswith(low) or low in opt.lower():
            return opt
    return raw


def ask_points(
    *,
    title: str,
    message: str,
    prompts: list[str],
    zoom: tuple[float, float, float] | None = None,
    banner: bool = True,
    rubber_band_from_first: bool = True,
    timeout: float = 600.0,
) -> list[tuple[float, float]]:
    """Textbox gate first (always), then N getpoint picks in AutoCAD.

    No mid-pick alert dialogs (those had no textbox). Esc in AutoCAD cancels a pick.
    """
    if not prompts:
        raise ValueError("prompts required")

    gate = _textbox_dialog(
        title=title,
        message=(
            message
            + "\n\nNext: click "
            + str(len(prompts))
            + " points in AutoCAD ("
            + " → ".join(prompts)
            + ").\n"
            "Type skip/cancel to abort, leave blank (or type continue) to start picking."
        ),
        options=["Continue — start picking", "Skip", "Cancel"],
        prompt="Type here (continue / skip / cancel / any notes)",
        allow_empty=True,
        ok_label="OK",
    )
    glow = gate.lower().strip()
    if glow in ("skip", "cancel", "stop", "abort") or glow.startswith("cancel"):
        raise RuntimeError(f"user cancelled point pick: {gate or 'cancel'}")

    # Build LISP getpoint sequence — NO (alert) mid-steps (no textbox there)
    parts = ["(progn (vl-load-com)"]
    if zoom is not None:
        cx, cy, h = zoom
        parts.append(f'(command "_.ZOOM" "_C" "{cx},{cy}" "{h}")')
    if banner:
        ban = _escape_lisp_str(title + "\\P" + message.replace("\n", "\\P"))
        parts.append(
            "(setq _ask_banner (entmakex (list "
            "'(0 . \"MTEXT\") '(100 . \"AcDbEntity\") '(8 . \"C-ANNO-NPLT\") '(62 . 2) "
            "'(100 . \"AcDbMText\") "
            f"'(10 {(zoom[0] - 12) if zoom else 0:.3f} {(zoom[1] + 12) if zoom else 0:.3f} 0.0) "
            "'(40 . 1.0) '(41 . 28.0) "
            f'(cons 1 "{ban}"))))'
        )
    else:
        parts.append("(setq _ask_banner nil)")

    parts.append("(setq _ask_pts nil)")
    parts.append("(setq _ask_prev nil)")
    for i, prompt in enumerate(prompts):
        p = _escape_lisp_str(prompt)
        if i == 0 or not rubber_band_from_first:
            parts.append(f'(setq _p (getpoint "\\n{p}: "))')
        else:
            parts.append(f'(setq _p (getpoint _ask_prev "\\n{p}: "))')
        parts.append(
            "(if (null _p) (progn "
            "(if _ask_banner (entdel _ask_banner)) "
            f'(setq f (open "{RESULT_LISP}" "w")) (write-line "CANCEL" f) (close f) (exit)))'
        )
        parts.append("(setq _ask_pts (cons _p _ask_pts))")
        if i == 0:
            parts.append("(setq _ask_prev _p)")

    parts.append("(if _ask_banner (entdel _ask_banner))")
    parts.append("(setq _ask_pts (reverse _ask_pts))")
    parts.append(f'(setq f (open "{RESULT_LISP}" "w"))')
    parts.append('(write-line "OK" f)')
    parts.append(
        "(foreach p _ask_pts "
        '(write-line (strcat "POINT," (rtos (car p) 2 6) "," (rtos (cadr p) 2 6)) f))'
    )
    parts.append("(close f) (princ \"\\nASK_OK\") (princ))")
    txt = _run_job_lsp("\n".join(parts) + "\n", timeout=timeout)
    pts = []
    for line in txt.splitlines():
        if line.startswith("POINT,"):
            _, x, y = line.split(",", 2)
            pts.append((float(x), float(y)))
    if len(pts) != len(prompts):
        raise RuntimeError(f"expected {len(prompts)} points, got {len(pts)}: {txt}")
    return pts


if __name__ == "__main__":
    notes = ask_message(
        title="acad_ask smoke",
        message="Textbox should be visible. Leave blank and OK, or type a note.",
    )
    print("notes:", repr(notes))
    print("smoke OK")
