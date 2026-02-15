# Environment Requirements

## MicroStation

### Required Version
- **Bentley MicroStation V8i** (SelectSeries 3 or later recommended)
- **File format:** DGN V8 (`.dgn`)
- **VBA runtime:** The built-in MicroStation VBA engine (MSO) — no external VBA installation needed

### MicroStation Configuration
| Setting | Required Value | Why |
|---------|---------------|-----|
| Master Units | Feet | All spacing calculations are in feet |
| VBA Macros | Enabled | Tool runs as a VBA macro |
| Workspaces | NYSDOT standard workspace active | Required for ProjectWise path mapping |

### How to Check Your MicroStation Version
1. In MicroStation: **Help → About MicroStation**
2. The version number appears in the dialog (e.g., "V8i SS4 08.11.09.xxx")

### Enabling VBA Macros
1. **Utilities → Macro → Project Manager**
2. If prompted about macro security, set to allow macros from trusted locations

---

## ProjectWise / WorkSpace

### Required Cell Libraries
The tool uses two hardcoded cell library paths. These must exist and be accessible:

| Library | Path | Contents |
|---------|------|---------|
| NYSDOT Sign Faces | `c:\pwworking\usny\d0119093\ny_plan_nmutcd_signface.cel` | MUTCD sign face cells (R, W, G series) |
| NYSDOT WZTC Symbols | `c:\pwworking\usny\d0119091\ny_plan_wztc.cel` | TWZAP_P, TWZSGN_P, TWZFLG_P, etc. |

These paths are provided by the **NYSDOT ProjectWise WorkSpace**. If your workstation has ProjectWise WorkSpace mounted, these paths will exist automatically.

### Checking Library Availability
In MicroStation, open the **Cell Library dialog** (Element → Cells) and try:
- **File → Attach** → navigate to `c:\pwworking\usny\d0119091\ny_plan_wztc.cel`

If the path doesn't exist, contact your ProjectWise administrator to confirm the correct local workspace path, then update the constant `WZTC_CELL_LIB` in `ModuleWZTCCells.bas` and the library paths in `Module3.bas` and `ModTest.bas`.

### ProjectWise Client
- **Bentley ProjectWise Explorer** (any version compatible with your organization's PWE server)
- WorkSpace configuration must map `d0119091` and `d0119093` datasets to `c:\pwworking\usny\`

---

## Development Environment (for editing and extending the tool)

### Recommended Code Editor
- **Cursor** (AI-assisted code editor, fork of VS Code) — recommended for this project
  - Download: [cursor.com](https://cursor.com)
  - Supports `.bas`, `.frm`, `.cls`, `.md` file editing with syntax highlighting

### Alternative Editors
- **VS Code** with VBA syntax extension (e.g., `VBA` by *abiggerhammer*)
- **Notepad++** with VBA language mode

### Version Control
- **Git** — the project is a git repository
  - Recommended client: **GitHub Desktop** or command-line git
  - All `.bas`, `.frm`, `.cls`, and `.md` files are tracked

### AI Coding Assistant
- **Claude Code** (Anthropic CLI) — used to develop and refactor this project
  - See [CLAUDE_CODE_GUIDE.md](CLAUDE_CODE_GUIDE.md) for setup and token optimization tips

---

## VBA IDE (Built into MicroStation)

The VBA IDE is accessed by pressing **Alt + F11** inside MicroStation. It provides:
- Project Explorer (file tree)
- Code editor with basic IntelliSense for MicroStation object model
- Immediate Window for running commands
- Breakpoints and debugging

### Exporting Files from VBA IDE to Disk
When you modify code in the VBA IDE, export the updated file back to disk so git can track it:

1. Right-click the module/form in Project Explorer
2. **Export File...**
3. Save to the project folder (overwrite existing file)
4. Commit the change in git

### Importing Files into VBA IDE
When pulling updates from git, re-import changed files:

1. Right-click the project in Project Explorer
2. **Import File...**
3. Select the updated `.bas` / `.frm` / `.cls` file
4. Delete the old version of the module if it already exists (VBA will not auto-replace)

---

## System Requirements

| Component | Minimum | Recommended |
|-----------|---------|-------------|
| OS | Windows 7 (64-bit) | Windows 10 / Windows 11 (64-bit) |
| RAM | 4 GB | 8 GB+ |
| MicroStation | V8i SS2 | V8i SS4 or later |
| ProjectWise Client | Any PWE version | Latest compatible with org server |
| Git | 2.x | Latest |
| Cursor / VS Code | Any recent version | Latest |

---

## Known Incompatibilities

- **MicroStation CONNECT Edition** — This tool uses the V8i VBA object model (`CadInputQueue`, `ElementEnumerator`, `ArcElement.Origin`, etc.). CONNECT Edition uses a different API (MicroStation MVBA / .NET). Significant porting work would be required.
- **32-bit MicroStation** — All modern NYSDOT deployments are 64-bit. If running a 32-bit version, DLong arithmetic in `ElIDAsDouble()` may behave differently.
- **Non-NYSDOT cell libraries** — Sign cell names and library paths are specific to NYSDOT ProjectWise standards. Other state DOT deployments would need different library paths and potentially different cell naming conventions.
