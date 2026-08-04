# NYSDOT WZTC Designer — Instructions for Claude Code

This file is read automatically at the start of every Claude Code session.
It contains project rules, patterns, and constraints that apply to all work in this repo.

---

## Project Summary

MicroStation 2023 CONNECT VBA tool for NYSDOT Workzone Traffic Control plan design.
Language: VBA. Files: `.bas` (modules), `.frm` (user forms), `.cls` (class modules).
There is no standalone compiler — code is run inside the MicroStation VBA IDE (Alt+F11).

**8-step sequential workflow:**
WZTCDesigner → DrawWorkSpace → AlignDraw → PlacePerp → frmSignSubColors → PlaceSign → PlaceElements → PlaceCells

State persists across form unloads via 40+ public variables in `SharedState.bas`.
All forms are modeless (`vbModeless`).

---

## Rules — Always Follow

### Before Writing Any CadInputQueue Code
Read the relevant Legacy file first. These contain the exact working patterns for
MicroStation VBA command sequences that are known to work:
- `Legacy Files/LegacyPrototype.bas` — sign face / PLACE CELL ICON / ATTACH LIBRARY pattern
- `Legacy Files/LegacyElements.bas` — HATCH ICON / PLACE SHAPE CONSTRAINED pattern

Do not guess CadInputQueue sequences. Match the Legacy pattern.

### MicroStation VBA API Constraints
- `ElementScanCriteria` valid methods: `ExcludeNonGraphical`, `ExcludeAllTypes`, `IncludeType`
  — `IncludeOnlyType` does NOT exist and will cause a "variable not defined" compile error
- `SendCommand` is for tool activation: `PLACE LINE CONSTRAINED`, `HATCH ICON`, `PLACE CELL ICON`
- `SendKeyin` is for settings: `ACTIVE LEVEL`, `ACTIVE COLOR`, `ACTIVE WEIGHT`, `ATTACH LIBRARY`, `TEXTEDITOR PLAYCOMMAND`
- `ae.CenterPoint` works in MicroStation 2023 CONNECT; `TryComputeCenter()` in PerpPlacement.bas is the fallback
- Do NOT add `Attribute VB_Name` to any module — not needed in MicroStation VBA

### Element Level / Color / Weight — Do Not Mix Up
| Element type | Level | Color | Weight |
|---|---|---|---|
| Alignment lines, perp tick lines | Default | 0 (white) | 0 |
| Dimension elements | Default | 2 (yellow) | 0 |
| WZTC drawing elements (Work Space, Channelizing Devices, Barriers, etc.) | Per DrawElements.bas GetElementLevel() | 6 | 2 |

When fixing a bug or adding a feature, only change the element properties of the element being worked on.
Do NOT change the level/color/weight of other element types.

### Form Behavior
- All forms are modeless — the GetInput loop works while the form is visible
- Only hide the form for cell placement (PlaceSelectedCell) and callout placement (TEXTEDITOR PLACENOTE)
- Standard drawing commands (PLACE LINE CONSTRAINED, HATCH ICON) do NOT require hiding the form

### Code Style
- Make minimal changes — do not refactor code surrounding a bug fix
- Do not add docstrings, comments, or type annotations to code that was not changed
- Do not add error handling for scenarios that cannot happen in the workflow
- Status messages go in `lblStatus.Caption` on the active form

---

## File Sync Protocol

**After Claude edits an existing `.bas`/`.frm`/`.cls` module on disk (standard path):**
Run `python mcp-server/hot_reload.py <path(s)>` to push the change into the open MicroStation VBA
project in place — do not ask the user to manually delete+re-import for routine code updates.
Never hot-reload while VBA is `[running]`/`[break]` (a bridge op, keyin batch, or GetInput loop
still executing) — this can wedge the project or crash MicroStation. If a call may still be in
flight, stop and ask the user to Reset (or Ctrl+Break) in the VBA IDE first.

Manual delete+Import File is still required, not hot-reload, when:
- the module/form does not exist in the IDE yet (brand-new file)
- adding or changing **UserForm Designer controls** (layout is not part of the CodeModule)
- hot-reload fails after a clean Reset (fall back to manual import for the affected files)

`mcp-server/*.py` changes (chat_driver.py, wztc_ops.py, etc.) are not picked up by hot-reload.
Restart it yourself — run `python mcp-server/restart_chat_driver.py` (it checks the process is
idle before stopping it, then relaunches in a new visible console window). Do not just tell the
user to restart it manually; this is the same "agent restarts, doesn't ask" rule as VBA hot-reload
above.

**After user edits in the MicroStation VBA IDE:**
The disk file is out of date. User must:
1. Right-click module → Export File
2. Tell Claude what changed before asking for further edits

If the user describes a manual change they made in the IDE, treat it as authoritative — do not
revert or overwrite it. Ask for clarification if the change is unclear.

---

## Cross-Tool Work Log (Cursor + Claude Code)

This project is sometimes worked on from Cursor as well as Claude Code. Neither tool can see
the other's session history or persistent memory, so `dev-notes/agent-log.md` is a manual
bridge: Cursor is instructed (via `.cursor/rules/agent-log.mdc`) to append an entry there after
finishing non-trivial work; Claude Code should do the same.

**Check `dev-notes/agent-log.md` whenever the user mentions work happened in Cursor** (or another
tool) since your last session, or at the start of a session if it's been a while. Fold anything
load-bearing into your own persistent memory the same way you would if the user had described it
directly in chat — don't just leave it sitting in the log unprocessed.

---

## Debugging Tips for This Project

When the user reports a bug, ask for:
1. The exact error message (if any) from the VBA IDE
2. The output from the Immediate window (Debug.Print lines)
3. Which step in the 8-step workflow triggered the issue
4. Which file/sub the issue is in, if known

Do not guess the root cause without this information for CadInputQueue timing issues —
these are notoriously hard to diagnose from symptoms alone.

---

## Future Development (Planned in ROADMAP.md)
- v1.2.0: Cost estimate tool
- v1.3.0: Detour plan designer
- v1.4.0: Workzone summary tables
- v1.5.0: Workzone cross sections

For any new workflow feature (new form, new module), use plan mode before writing code.