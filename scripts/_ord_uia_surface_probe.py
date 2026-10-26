"""Read-only UIA surface map of live OpenRoads Designer.

Enumerate / Exists / Name / ClassName / ControlType only.
No Invoke Close, no SetValue, no CAD keyins that modify the model.
Does NOT open Key-in via F9 — only probes if already present.
"""
from __future__ import annotations

import json
import sys
import time
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

import uiautomation as auto

OUT_DIR = Path(__file__).resolve().parents[1] / "outputs" / "uia"
JSON_PATH = OUT_DIR / "ord_uia_surface.json"
MD_PATH = OUT_DIR / "ord_uia_surface.md"

INTERESTING_NAME_FRAGMENTS = (
    "Level",
    "Key-in",
    "Key in",
    "Explorer",
    "Display",
    "Properties",
    "Element Selection",
    "Barmenu",
    "Dialog",
)

NAMED_CONTROL_TYPES = {
    "ButtonControl",
    "EditControl",
    "ListControl",
    "ListItemControl",
    "ComboBoxControl",
    "TreeItemControl",
    "TreeControl",
    "CheckBoxControl",
    "TabItemControl",
    "MenuItemControl",
    "HyperlinkControl",
    "TextControl",
    "DocumentControl",
}

USEFUL_HINTS = (
    "level",
    "key-in",
    "key in",
    "explorer",
    "display",
    "properties",
    "element selection",
    "search",
    "filter",
    "browse",
    "command",
)


def _safe_name(c) -> str:
    try:
        return (c.Name or "").strip()
    except Exception:
        return ""


def _safe_class(c) -> str:
    try:
        return (c.ClassName or "").strip()
    except Exception:
        return ""


def _safe_type(c) -> str:
    try:
        return (c.ControlTypeName or "").strip()
    except Exception:
        return ""


def _rect_tuple(c):
    try:
        r = c.BoundingRectangle
        if r is None:
            return None
        return {
            "left": int(r.left),
            "top": int(r.top),
            "right": int(r.right),
            "bottom": int(r.bottom),
            "width": int(r.width()),
            "height": int(r.height()),
        }
    except Exception:
        return None


def find_ord_main():
    """Find OpenRoads / MicroStation RadRibbonWindow main window."""
    candidates = []
    root = auto.GetRootControl()
    for w in root.GetChildren():
        name = _safe_name(w)
        cls = _safe_class(w)
        if not name and not cls:
            continue
        score = 0
        low = name.lower()
        if "openroads" in low or "microstation" in low:
            score += 10
        if "fea_rwy" in low or ".dgn" in low or "x727" in low:
            score += 5
        if cls == "RadRibbonWindow":
            score += 8
        if "bentley" in low:
            score += 2
        if score:
            candidates.append((score, w, name, cls))
    candidates.sort(key=lambda x: -x[0])
    if not candidates:
        return None
    return candidates[0][1]


def list_toplevel_related(main):
    """Top-level windows that look like ORD dialogs / children of the session."""
    main_name = _safe_name(main)
    main_pid = None
    try:
        main_pid = main.ProcessId
    except Exception:
        pass

    dialogs = []
    seen = set()
    root = auto.GetRootControl()
    for w in root.GetChildren():
        try:
            hwnd = w.NativeWindowHandle
        except Exception:
            hwnd = 0
        if hwnd and hwnd in seen:
            continue
        if hwnd:
            seen.add(hwnd)

        name = _safe_name(w)
        cls = _safe_class(w)
        ctype = _safe_type(w)
        if w == main or (hwnd and hwnd == getattr(main, "NativeWindowHandle", None)):
            continue

        related = False
        try:
            if main_pid and w.ProcessId == main_pid:
                related = True
        except Exception:
            pass

        # Also catch MStnChild / known Bentley dialogs even if ProcessId flaky
        low_cls = cls.lower()
        low_name = name.lower()
        if any(
            x in low_cls
            for x in ("mstnchild", "bentley", "dgn", "toolsettings", "dockable")
        ):
            related = True
        if any(
            frag.lower() in low_name
            for frag in (
                "Key-in",
                "Level Display",
                "Element Selection",
                "Properties",
                "Explorer",
                "Barmenu",
                "Tool Settings",
                "Message Center",
                "AccuDraw",
                "View Groups",
                "References",
                "Models",
                "Drawing",
            )
        ):
            related = True

        if not related:
            continue
        if not name and not cls:
            continue

        dialogs.append(
            {
                "control": w,
                "name": name,
                "class": cls,
                "control_type": ctype,
                "rect": _rect_tuple(w),
                "hwnd": int(hwnd) if hwnd else None,
            }
        )

    # Also walk immediate children of main that are windows/panes
    try:
        for ch in main.GetChildren():
            name = _safe_name(ch)
            cls = _safe_class(ch)
            ctype = _safe_type(ch)
            if ctype not in ("WindowControl", "PaneControl", "CustomControl"):
                continue
            try:
                hwnd = ch.NativeWindowHandle
            except Exception:
                hwnd = 0
            if hwnd and hwnd in seen:
                continue
            if hwnd:
                seen.add(hwnd)
            dialogs.append(
                {
                    "control": ch,
                    "name": name,
                    "class": cls,
                    "control_type": ctype,
                    "rect": _rect_tuple(ch),
                    "hwnd": int(hwnd) if hwnd else None,
                    "parent": "main_child",
                }
            )
    except Exception as e:
        dialogs.append(
            {
                "control": None,
                "name": f"(main children walk error: {e})",
                "class": "",
                "control_type": "",
                "rect": None,
                "hwnd": None,
            }
        )

    return dialogs


def walk_named_controls(root_ctrl, max_depth=8, max_nodes=2500, cap_named=40):
    """DFS walk collecting named interesting controls + node count."""
    named = []
    type_counts = Counter()
    empty_panes = []
    node_count = 0
    errors = 0

    def walk(ctrl, depth):
        nonlocal node_count, errors
        if depth > max_depth or node_count >= max_nodes or len(named) >= cap_named * 3:
            return
        node_count += 1
        try:
            name = _safe_name(ctrl)
            cls = _safe_class(ctrl)
            ctype = _safe_type(ctrl)
            type_counts[ctype or "?"] += 1

            if ctype in NAMED_CONTROL_TYPES and name:
                if len(named) < cap_named:
                    named.append(
                        {
                            "name": name[:120],
                            "control_type": ctype,
                            "class": cls[:80] if cls else "",
                            "depth": depth,
                        }
                    )
            elif name and any(f.lower() in name.lower() for f in INTERESTING_NAME_FRAGMENTS):
                if len(named) < cap_named:
                    named.append(
                        {
                            "name": name[:120],
                            "control_type": ctype,
                            "class": cls[:80] if cls else "",
                            "depth": depth,
                            "interesting": True,
                        }
                    )

            # Empty MStnChild panes
            if "MStnChild" in cls or (ctype == "PaneControl" and not name):
                try:
                    kids = ctrl.GetChildren()
                    kid_n = len(kids) if kids else 0
                except Exception:
                    kid_n = -1
                if kid_n == 0 and len(empty_panes) < 20:
                    empty_panes.append(
                        {
                            "name": name[:80] or "(unnamed)",
                            "class": cls[:80],
                            "control_type": ctype,
                            "depth": depth,
                        }
                    )
        except Exception:
            errors += 1
            return

        try:
            children = ctrl.GetChildren()
        except Exception:
            errors += 1
            return
        for ch in children or []:
            walk(ch, depth + 1)

    if root_ctrl is not None:
        walk(root_ctrl, 0)

    return {
        "node_count": node_count,
        "named_controls": named[:cap_named],
        "type_counts": dict(type_counts.most_common(30)),
        "empty_panes": empty_panes,
        "errors": errors,
    }


def verdict_for(name: str, cls: str, walk_result: dict) -> str:
    named = walk_result.get("named_controls") or []
    node_count = walk_result.get("node_count") or 0
    empty_panes = walk_result.get("empty_panes") or []

    if node_count <= 1 and not named:
        return "empty"
    if "MStnChild" in (cls or "") and not named and node_count < 5:
        return "empty"

    useful_hits = 0
    for nc in named:
        nlow = (nc.get("name") or "").lower()
        if any(h in nlow for h in USEFUL_HINTS):
            useful_hits += 1
        if nc.get("control_type") in (
            "EditControl",
            "ListControl",
            "TreeControl",
            "TreeItemControl",
            "ComboBoxControl",
            "CheckBoxControl",
        ):
            useful_hits += 1

    name_low = (name or "").lower()
    if any(
        x in name_low
        for x in (
            "key-in",
            "level display",
            "explorer",
            "properties",
            "element selection",
            "references",
            "models",
        )
    ):
        if named or node_count > 10:
            return "useful"

    if useful_hits >= 2 or (useful_hits >= 1 and node_count > 15):
        return "useful"
    if named and useful_hits == 0 and node_count > 5:
        # Has named chrome but no civil/dialog content
        return "chrome_only"
    if not named and node_count > 1:
        return "chrome_only"
    if empty_panes and not named:
        return "empty"
    return "chrome_only"


def find_ribbon(main):
    """Locate ribbon control under main window."""
    hits = []

    def walk(ctrl, depth, path):
        if depth > 6 or len(hits) >= 5:
            return
        name = _safe_name(ctrl)
        cls = _safe_class(ctrl)
        ctype = _safe_type(ctrl)
        low = (name + " " + cls).lower()
        if (
            "ribbon" in low
            or ctype == "TabControl"
            or "RadRibbon" in cls
            or name in ("Home", "View", "Drawing", "Annotate")
        ):
            hits.append({"control": ctrl, "name": name, "class": cls, "type": ctype, "path": path})
        try:
            children = ctrl.GetChildren()
        except Exception:
            return
        for i, ch in enumerate(children or []):
            walk(ch, depth + 1, path + [i])

    walk(main, 0, [])
    return hits


def walk_ribbon_named_tools(main, search_frags=None):
    """Walk ribbon area; collect named tools; search Level/Key-in/Explorer/Display."""
    if search_frags is None:
        search_frags = ("Level", "Key-in", "Key in", "Explorer", "Display")

    ribbon_named = []
    search_hits = []
    tab_names = []

    # Prefer ControlType Tab / Button / MenuItem under main with shallow-deep walk
    # focused on ribbon-like names
    visited = [0]
    MAX = 4000

    def walk(ctrl, depth, under_ribbon):
        if depth > 12 or visited[0] >= MAX:
            return
        visited[0] += 1
        name = _safe_name(ctrl)
        cls = _safe_class(ctrl)
        ctype = _safe_type(ctrl)
        low_cls = cls.lower()
        low_name = name.lower()

        is_ribbonish = under_ribbon or (
            "ribbon" in low_cls
            or "ribbon" in low_name
            or ctype == "TabControl"
            or name in ("Ribbon", "Home", "View", "Drawing", "Annotate", "OpenRoads Modeling")
        )

        if ctype == "TabItemControl" and name:
            tab_names.append(name)

        if is_ribbonish and name and ctype in (
            "ButtonControl",
            "MenuItemControl",
            "SplitButtonControl",
            "CheckBoxControl",
            "TabItemControl",
            "HyperlinkControl",
            "ListItemControl",
            "ToolBarControl",
        ):
            if len(ribbon_named) < 120:
                ribbon_named.append(
                    {
                        "name": name[:120],
                        "control_type": ctype,
                        "class": cls[:80],
                        "depth": depth,
                        "tab_hint": None,
                    }
                )
            for frag in search_frags:
                if frag.lower() in low_name:
                    search_hits.append(
                        {
                            "name": name[:120],
                            "control_type": ctype,
                            "class": cls[:80],
                            "matched": frag,
                            "depth": depth,
                        }
                    )
                    break

        # Also catch interesting names anywhere under ribbon depth
        if is_ribbonish and name:
            for frag in search_frags:
                if frag.lower() in low_name and not any(
                    h.get("name") == name[:120] and h.get("matched") == frag for h in search_hits
                ):
                    search_hits.append(
                        {
                            "name": name[:120],
                            "control_type": ctype,
                            "class": cls[:80],
                            "matched": frag,
                            "depth": depth,
                        }
                    )

        try:
            children = ctrl.GetChildren()
        except Exception:
            return
        for ch in children or []:
            walk(ch, depth + 1, is_ribbonish)

    walk(main, 0, False)

    # Filter to Home/View interest if tabs found
    home_view_tools = [
        t
        for t in ribbon_named
        if any(
            x in (t.get("name") or "").lower()
            for x in ("home", "view", "level", "key", "explorer", "display", "attributes")
        )
    ]

    return {
        "tabs_seen": sorted(set(tab_names))[:40],
        "named_tools": ribbon_named,
        "home_view_related": home_view_tools[:60],
        "search_hits": search_hits[:80],
        "nodes_visited": visited[0],
    }


def probe_keyin_if_present(dialogs_meta):
    """If Key-in already exists, probe Edit controls only — no F9 open, no keyins."""
    keyin = None
    for d in dialogs_meta:
        n = (d.get("name") or "").lower()
        if n == "key-in" or "key-in" in n:
            keyin = d
            break
    if not keyin or keyin.get("control") is None:
        return {
            "present": False,
            "opened_via_f9": False,
            "note": "Key-in window not found among top-level dialogs; F9 open skipped (read-only policy).",
            "edits": [],
        }

    ctrl = keyin["control"]
    edits = []
    named = []

    def walk(c, depth):
        if depth > 8 or len(edits) >= 30:
            return
        name = _safe_name(c)
        ctype = _safe_type(c)
        cls = _safe_class(c)
        if ctype == "EditControl":
            edits.append(
                {
                    "name": name[:120],
                    "class": cls[:80],
                    "depth": depth,
                    "rect": _rect_tuple(c),
                    # Do NOT read/set Value — some civil edits may side-effect; Name only.
                }
            )
        if name and ctype in NAMED_CONTROL_TYPES and len(named) < 40:
            named.append({"name": name[:120], "control_type": ctype, "class": cls[:80]})
        try:
            for ch in c.GetChildren() or []:
                walk(ch, depth + 1)
        except Exception:
            return

    walk(ctrl, 0)
    return {
        "present": True,
        "opened_via_f9": False,
        "name": keyin.get("name"),
        "class": keyin.get("class"),
        "edits": edits,
        "named_controls_sample": named,
        "note": "Probed existing Key-in only; no keyins sent.",
    }


def summarize(main_info, dialogs, ribbon, keyin):
    useful = [d for d in dialogs if d.get("verdict") == "useful"]
    chrome = [d for d in dialogs if d.get("verdict") == "chrome_only"]
    empty = [d for d in dialogs if d.get("verdict") == "empty"]

    top_useful = []
    for d in useful:
        top_useful.append(
            {
                "name": d.get("name"),
                "class": d.get("class"),
                "node_count": d.get("node_count"),
                "named_sample": [c.get("name") for c in (d.get("named_controls") or [])[:8]],
            }
        )
    top_useful = top_useful[:15]

    top_chrome_empty = []
    for d in chrome + empty:
        top_chrome_empty.append(
            {
                "name": d.get("name") or "(unnamed)",
                "class": d.get("class"),
                "verdict": d.get("verdict"),
                "node_count": d.get("node_count"),
            }
        )
    top_chrome_empty = top_chrome_empty[:15]

    recommendations = []
    if keyin.get("present") and keyin.get("edits"):
        recommendations.append(
            {
                "target": "Key-in (existing window)",
                "approach": "UIA Edit Name/Exists + SendKeys only for non-destructive status keyins; prefer bridge COM for model ops",
                "why": f"Already present with {len(keyin.get('edits') or [])} EditControl(s)",
            }
        )
    else:
        recommendations.append(
            {
                "target": "Key-in",
                "approach": "Open via F9 only when idle, then UIA probe — or use CadInputQueue/SendKeyin via VBA bridge",
                "why": "Key-in not currently visible in UIA top-level set",
            }
        )

    levelish = [d for d in useful if "level" in (d.get("name") or "").lower()]
    if levelish:
        recommendations.append(
            {
                "target": "Level Display dialog",
                "approach": "UIA (Lists/Tree/CheckBoxes) if named controls exist; else computer-use for view toggles",
                "why": f"Found useful dialog(s): {[d.get('name') for d in levelish]}",
            }
        )
    else:
        recommendations.append(
            {
                "target": "Level Display",
                "approach": "computer-use or ribbon click — UIA may be chrome-only/empty for MStnChild panes",
                "why": "No useful Level Display surface in this probe",
            }
        )

    recommendations.append(
        {
            "target": "Ribbon tools (Level / Explorer / Key-in)",
            "approach": "UIA Invoke only if ButtonControl has stable Name; else computer-use click by screenshot",
            "why": f"Ribbon search hits: {len(ribbon.get('search_hits') or [])}; named tools: {len(ribbon.get('named_tools') or [])}",
        }
    )
    recommendations.append(
        {
            "target": "Model geometry / civil fields",
            "approach": "bridge COM / CIF add-in — NOT UIA SetValue",
            "why": "Civil fields and model edits are unsafe/unreliable via UIA",
        }
    )

    return {
        "dialog_counts": {
            "useful": len(useful),
            "chrome_only": len(chrome),
            "empty": len(empty),
            "total": len(dialogs),
        },
        "top_useful": top_useful,
        "top_chrome_or_empty": top_chrome_empty,
        "keyin": {
            "present": keyin.get("present"),
            "edit_count": len(keyin.get("edits") or []),
        },
        "ribbon": {
            "tabs_seen": ribbon.get("tabs_seen"),
            "named_tool_count": len(ribbon.get("named_tools") or []),
            "search_hit_count": len(ribbon.get("search_hits") or []),
        },
        "recommendations": recommendations,
        "probed_at": datetime.now(timezone.utc).isoformat(),
        "main_window_name": main_info.get("name"),
    }


def write_markdown(report: dict) -> str:
    lines = []
    lines.append("# OpenRoads Designer UIA surface map")
    lines.append("")
    lines.append(f"Probed at: `{report.get('summary', {}).get('probed_at', '')}`")
    lines.append("")
    main = report.get("main_window") or {}
    lines.append("## Main window")
    lines.append(f"- **Name:** {main.get('name')}")
    lines.append(f"- **Class:** {main.get('class')}")
    lines.append(f"- **ControlType:** {main.get('control_type')}")
    lines.append(f"- **ProcessId:** {main.get('process_id')}")
    lines.append("")

    summary = report.get("summary") or {}
    dc = summary.get("dialog_counts") or {}
    lines.append("## Dialog verdicts")
    lines.append(
        f"- useful: **{dc.get('useful', 0)}** · chrome_only: **{dc.get('chrome_only', 0)}** · empty: **{dc.get('empty', 0)}** · total: **{dc.get('total', 0)}**"
    )
    lines.append("")

    lines.append("## Top useful UIA targets")
    for i, u in enumerate((summary.get("top_useful") or [])[:10], 1):
        sample = ", ".join((u.get("named_sample") or [])[:5]) or "(no named sample)"
        lines.append(
            f"{i}. **{u.get('name') or '(unnamed)'}** (`{u.get('class')}`, nodes={u.get('node_count')}) — {sample}"
        )
    lines.append("")

    lines.append("## Top chrome-only / empty areas")
    for i, c in enumerate((summary.get("top_chrome_or_empty") or [])[:10], 1):
        lines.append(
            f"{i}. **{c.get('name')}** (`{c.get('class')}`) — {c.get('verdict')}, nodes={c.get('node_count')}"
        )
    lines.append("")

    ribbon = report.get("ribbon_named_tools") or {}
    if isinstance(ribbon, dict):
        lines.append("## Ribbon")
        lines.append(f"- Tabs seen: {', '.join(ribbon.get('tabs_seen') or []) or '(none)'}")
        lines.append(f"- Named tools collected: {len(ribbon.get('named_tools') or [])}")
        lines.append(f"- Search hits (Level/Key-in/Explorer/Display): {len(ribbon.get('search_hits') or [])}")
        for h in (ribbon.get("search_hits") or [])[:20]:
            lines.append(
                f"  - [{h.get('matched')}] {h.get('name')} ({h.get('control_type')})"
            )
        lines.append("")

    keyin = report.get("keyin") or {}
    lines.append("## Key-in")
    lines.append(f"- Present: {keyin.get('present')}")
    lines.append(f"- Opened via F9: {keyin.get('opened_via_f9')}")
    lines.append(f"- Note: {keyin.get('note')}")
    lines.append(f"- Edit controls: {len(keyin.get('edits') or [])}")
    lines.append("")

    lines.append("## Recommended next automation targets")
    for r in summary.get("recommendations") or []:
        lines.append(f"- **{r.get('target')}** — approach: {r.get('approach')}")
        lines.append(f"  - why: {r.get('why')}")
    lines.append("")

    lines.append("## All dialogs (brief)")
    for d in report.get("dialogs") or []:
        lines.append(
            f"- `{d.get('verdict')}` **{d.get('name') or '(unnamed)'}** / `{d.get('class')}` — nodes={d.get('node_count')}, named={len(d.get('named_controls') or [])}"
        )
    lines.append("")
    return "\n".join(lines)


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    t0 = time.time()

    main_win = find_ord_main()
    if not main_win:
        report = {
            "main_window": None,
            "dialogs": [],
            "ribbon_named_tools": [],
            "keyin": {"present": False, "note": "ORD main window not found"},
            "summary": {"error": "OpenRoadsDesigner RadRibbonWindow not found"},
        }
        JSON_PATH.write_text(json.dumps(report, indent=2), encoding="utf-8")
        MD_PATH.write_text("# ORD UIA — main window not found\n", encoding="utf-8")
        print("FAIL: no main window", file=sys.stderr)
        return 1

    main_info = {
        "name": _safe_name(main_win),
        "class": _safe_class(main_win),
        "control_type": _safe_type(main_win),
        "process_id": getattr(main_win, "ProcessId", None),
        "hwnd": int(main_win.NativeWindowHandle) if main_win.NativeWindowHandle else None,
        "rect": _rect_tuple(main_win),
        "child_windows_brief": [],
    }

    # List immediate children of main (Name/Class)
    try:
        for ch in main_win.GetChildren() or []:
            main_info["child_windows_brief"].append(
                {
                    "name": _safe_name(ch)[:100],
                    "class": _safe_class(ch)[:80],
                    "control_type": _safe_type(ch),
                }
            )
    except Exception as e:
        main_info["child_windows_error"] = str(e)

    raw_dialogs = list_toplevel_related(main_win)
    dialogs = []
    for rd in raw_dialogs:
        ctrl = rd.get("control")
        walk = walk_named_controls(ctrl) if ctrl is not None else {
            "node_count": 0,
            "named_controls": [],
            "type_counts": {},
            "empty_panes": [],
            "errors": 0,
        }
        entry = {
            "name": rd.get("name"),
            "class": rd.get("class"),
            "control_type": rd.get("control_type"),
            "hwnd": rd.get("hwnd"),
            "rect": rd.get("rect"),
            "parent": rd.get("parent"),
            "node_count": walk["node_count"],
            "named_controls": walk["named_controls"],
            "type_counts": walk["type_counts"],
            "empty_panes": walk["empty_panes"],
            "verdict": verdict_for(rd.get("name") or "", rd.get("class") or "", walk),
        }
        dialogs.append(entry)

    # Keep control refs for keyin probe before stripping
    keyin = probe_keyin_if_present(raw_dialogs)

    ribbon = walk_ribbon_named_tools(main_win)

    summary = summarize(main_info, dialogs, ribbon, keyin)
    summary["elapsed_sec"] = round(time.time() - t0, 2)

    report = {
        "main_window": main_info,
        "dialogs": dialogs,
        "ribbon_named_tools": ribbon,
        "keyin": keyin,
        "summary": summary,
    }

    JSON_PATH.write_text(json.dumps(report, indent=2, default=str), encoding="utf-8")
    MD_PATH.write_text(write_markdown(report), encoding="utf-8")

    print(f"Wrote {JSON_PATH}")
    print(f"Wrote {MD_PATH}")
    print(
        f"dialogs={len(dialogs)} useful={summary['dialog_counts']['useful']} "
        f"chrome={summary['dialog_counts']['chrome_only']} empty={summary['dialog_counts']['empty']} "
        f"elapsed={summary['elapsed_sec']}s"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
