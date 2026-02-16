# Contributing to NYSDOT Workzone Traffic Control Designer

Thank you for your interest in contributing. This guide covers the development workflow, coding conventions, and pull request process.

---

## Development Workflow

### Branch Strategy

| Branch | Purpose |
|--------|---------|
| `main` | Stable, release-ready code. All merges require a pull request. |
| `feature/<name>` | New features (e.g., `feature/wztc-estimate-tool`) |
| `fix/<name>` | Bug fixes (e.g., `fix/spacing-table-lookup`) |
| `refactor/<name>` | Code restructuring with no behavior change |

1. Create a feature or fix branch from `main`
2. Make your changes in small, focused commits
3. Open a pull request back to `main` when ready

### Getting Started

1. Clone the repository
2. Open MicroStation V8i and press **Alt + F11** to open the VBA Editor
3. Import all `.bas`, `.cls`, and `.frm` files from the project (see [INSTALLATION_GUIDE.md](INSTALLATION_GUIDE.md))
4. Run the debug tests in the `Debug/` folder to verify everything loads correctly

---

## Naming Conventions

### Files

| Type | Convention | Example |
|------|-----------|---------|
| Standard modules | PascalCase, descriptive verb+noun | `DrawSign.bas`, `CellPlacer.bas` |
| Class modules | PascalCase, noun-based | `PlaceButtons.cls`, `SignNumBox.cls` |
| UserForms | PascalCase, short action name | `PlacePerp.frm`, `PlaceSign.frm` |
| Documentation | UPPER_SNAKE or PascalCase `.md` | `INSTALLATION_GUIDE.md`, `CONTRIBUTING.md` |

### VBA Code

| Element | Convention | Example |
|---------|-----------|---------|
| Public Sub/Function | PascalCase | `StartSignPlacement`, `GetPointAndTangent` |
| Private Sub/Function | PascalCase | `RefreshDisplay`, `PopulateCategories` |
| Public variables | camelCase with `wztc` prefix | `wztcSignCount`, `wztcBufferSpace` |
| Private variables | camelCase | `rowCount`, `lastPoint` |
| Constants | UPPER_SNAKE_CASE | `TABLE_START_TOP`, `INITIAL_ROWS` |
| UDTs (User Defined Types) | PascalCase | `PathSeg`, `signData` |
| Form controls | Hungarian-style prefix | `lblStatus`, `cboCategory`, `btnSubmit` |

### Control Prefixes

| Prefix | Control Type |
|--------|-------------|
| `lbl` | Label |
| `cbo` / `cmb` | ComboBox |
| `txt` | TextBox |
| `btn` / `cmd` | CommandButton |
| `opt` | OptionButton |
| `frame` | Frame |
| `lst` | ListBox |

---

## Formatting Rules

- **`Option Explicit`** at the top of every module, class, and form
- **Section headers** using `' ============` block comments to separate logical sections
- **Indentation:** 4 spaces (VBA IDE default)
- **Line continuation:** Use ` _` (space + underscore) for lines exceeding ~100 characters
- **Blank lines:** One blank line between procedures, two blank lines between major sections

### Example

```vba
Option Explicit

' ============================================================
' MODULE PURPOSE DESCRIPTION
' ------------------------------------------------------------
' Additional context about what this module does.
' ============================================================

Private Const MAX_ITEMS As Integer = 50

' ============================================================
' PUBLIC ENTRY POINT
' ============================================================
Public Sub DoSomething()
    Dim i As Integer
    For i = 1 To MAX_ITEMS
        ' ... implementation ...
    Next i
End Sub
```

---

## Comment and Documentation Style

- Every module starts with a block comment describing its purpose
- Every `Public Sub` or `Public Function` has a one-line comment above it explaining what it does
- Inline comments are used sparingly, only where the logic is not self-evident
- Use `' TODO:` for planned improvements and `' NOTE:` for important caveats

### Form Modules

Each `.frm` file includes a header listing the controls that must be added manually in the VBA IDE form designer:

```vba
' Controls to add manually in the VBA IDE form designer:
'   lblStatus  - Label          (status / error messages)
'   btnSubmit  - CommandButton  "Submit"
```

---

## Testing Requirements

Before submitting a pull request:

1. **Run `DebugDesignerLoad`** — verifies all WZTCDesigner controls exist
2. **Run `DebugWorkflowSequence`** — verifies all 7 forms load without errors
3. **Run `DebugTest`** — verifies basic control existence on WZTCDesigner
4. **Manual test:** Walk through the full 6-step workflow at least once in MicroStation:
   - Configure a workzone in WZTCDesigner
   - Draw an alignment (at least 2 segments)
   - Place perpendicular lines
   - Draw at least one sign
   - Draw at least one WZTC element
   - Place at least one cell symbol

---

## Pull Request Requirements

### Before Opening a PR

- [ ] All debug tests pass (no FAIL results in Immediate Window)
- [ ] Manual workflow test completed successfully
- [ ] No `Debug.Print` statements left in production code (move to Debug/ if needed)
- [ ] All new public variables added to `SharedState.bas` (not scattered in other modules)
- [ ] `Option Explicit` present in every new file
- [ ] New forms include the control header comment block

### PR Description Format

```
## Summary
Brief description of what this PR does and why.

## Changes
- List of specific changes made

## Testing
- How the changes were tested
- Any edge cases considered

## Screenshots
(If UI changes are involved, include before/after screenshots)
```

### Review Checklist

Reviewers should verify:
- Code follows the naming conventions above
- No hardcoded values that should be constants
- `ControlExists()` guard used before accessing dynamic controls
- Error handlers present in any Sub that interacts with MicroStation API
- State that must persist across form loads is stored in `SharedState.bas`

---

## Questions?

Open an issue on the repository if you have questions about the codebase or contribution process.