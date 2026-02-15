# Using Cursor and Claude Code for VBA Development

A practical guide to using AI-assisted coding tools to develop, debug, and extend this MicroStation VBA project efficiently.

---

## Tools Overview

| Tool | What It Is | Best For |
|------|-----------|---------|
| **Cursor** | AI-assisted code editor (VS Code fork) | Writing and editing `.bas`/`.frm`/`.cls` files, asking questions about code |
| **Claude Code** | Anthropic's AI CLI (command-line tool) | Multi-file edits, refactoring, generating documentation, git operations |
| **Claude.ai** | Web-based chat interface | Architecture questions, design discussions, long explanations |

For this project, **Cursor + Claude Code together** is the recommended workflow. Use Cursor's editor for browsing code and making targeted edits; use Claude Code for larger tasks that touch multiple files.

---

## Setting Up Cursor

1. Download and install from [cursor.com](https://cursor.com)
2. Open the project folder: **File → Open Folder** → select `microstation-vba-project`
3. Cursor indexes all files automatically

### Recommended Cursor Settings for VBA
- Install the **VBA** extension (search "VBA" in Extensions panel) for syntax highlighting
- Set Tab size to 4 spaces (VBA convention)
- In Cursor settings, add `.bas`, `.frm`, `.cls` as known file types if they don't open correctly

### Using Cursor AI Chat
Press **Ctrl + L** to open the inline chat panel. You can:
- Ask questions about the code: *"What does BuildAlignmentPath do?"*
- Request edits: *"Add error handling to this function"*
- Reference specific files: drag a file into the chat or use `@filename`

---

## Setting Up Claude Code (CLI)

Claude Code is a command-line tool that runs Claude as an interactive coding agent inside your terminal. It can read, write, and search files in your project, run git commands, and handle multi-step tasks.

### Installation
```bash
npm install -g @anthropic/claude-code
```

Requires Node.js 18+. Check: `node --version`

### Starting Claude Code
```bash
cd "c:\Users\RussinT\OneDrive - AECOM\Desktop\microstation-vba-project"
claude
```

The `>` prompt means Claude Code is ready. It has access to all files in the current directory.

### Basic Commands Inside Claude Code
| Command | What It Does |
|---------|-------------|
| `/help` | Show all available commands |
| `/compact` | Summarize conversation history to free up context (use when session gets long) |
| `/clear` | Start a fresh conversation (loses all context) |
| `/cost` | Show how many tokens you have used this session |
| `Ctrl+C` | Cancel a running operation |
| `Ctrl+D` or `/exit` | Exit Claude Code |

---

## Workflow: Making Changes with Claude Code

### Typical Session Pattern

1. Open terminal in the project folder
2. Start: `claude`
3. Describe your task clearly, including:
   - Which form or module is affected
   - What the current behavior is
   - What you want it to do instead
4. Review the proposed changes before approving
5. After changes are made, export updated files from the MicroStation VBA IDE to disk, then commit to git

### Example Prompts That Work Well

**Bug fix:**
> "In PlacePerp, when the user clicks Place Line for a sign item, the perpendicular line is drawing at the wrong location on arc segments. The issue is in PerpPlacement.bas. Read the arc handling code and fix it."

**New feature:**
> "Add a label next to each Place Line button in PlacePerp that shows the cumulative distance from the start of the alignment to that item."

**Documentation:**
> "Read PerpPlacement.bas and write inline comments explaining the math for the arc tangent calculation."

**Refactoring:**
> "The module DrawSign.bas has a misleading name. Identify all places that reference it and tell me what I'd need to change to rename it to ModuleSignDrawing."

---

## Understanding Token Usage and Rate Limits

### What Are Tokens?
Tokens are the units Claude processes. One token ≈ 4 characters of text. Sending a large file to Claude uses many input tokens; receiving a long response uses output tokens. Both affect your usage and rate limits.

### Rate Limits by Plan
| Plan | Approximate Daily Token Budget | Notes |
|------|-------------------------------|-------|
| Anthropic API (pay-as-you-go) | Unlimited (billed per token) | Recommended for heavy development work |
| Claude Pro (web) | ~45 messages / 5 hours | Resets every 5 hours |
| Claude Code (CLI) | Uses API credits | Same rate limits as API tier |

Check your current usage at: [console.anthropic.com](https://console.anthropic.com)

### Viewing Token Usage in Claude Code
- During a session: `/cost` shows tokens used so far this session
- In the status bar (if enabled): token count is shown continuously
- Budget warnings appear automatically when approaching context limits

### Context Window
Claude has a maximum context window (currently 200,000 tokens for Sonnet). As a session grows long, older messages are compressed. You can trigger this manually with `/compact` to summarize history and free up space without losing the key facts.

---

## Optimizing Token Usage

### Strategy 1: Be Specific About Which Files to Read
**Wasteful:**
> "Look at all the VBA files and fix the arc bug."

**Efficient:**
> "Read PerpPlacement.bas lines 150-220 (the BuildAlignmentPath arc handling block) and fix the arc center derivation."

Directing Claude to specific files and line ranges dramatically reduces input tokens.

### Strategy 2: Use `/compact` Proactively
When you've finished a major task and are starting something new, run `/compact` before continuing. This compresses the prior context into a summary (preserving key facts) and frees up room for new work.

Good time to compact:
- After completing a bug fix and before starting a new feature
- When the session has gone through 10+ message exchanges
- Before a task that will require reading many large files

### Strategy 3: Ask Claude to Read Targeted Sections
Instead of having Claude read an entire 1500-line form file, tell it which section matters:

> "Read only the `btnSubmit_Click` sub in WZTCDesigner.frm (around line 1526) and the `RestoreState` sub near the end of the file."

### Strategy 4: One Task Per Session for Complex Work
For completely unrelated tasks, starting a new session (`/clear` or just restarting `claude`) avoids context from Task A polluting Task B and keeps context usage lower.

### Strategy 5: Use Background Agents for Parallel Reads
Claude Code can launch sub-agents to read multiple files in parallel. If you're doing a refactor that touches many files, say:
> "I need to understand how the five-module workflow connects — read AlignmentTool.bas, PerpPlacement.bas, SignPlacer.bas, DrawElements.bas, and CellPlacer.bas in parallel and summarize their interfaces."

### Strategy 6: Keep Context Files Updated
The `MEMORY.md` file in `.claude/projects/` is loaded automatically into every Claude Code session. Keep it accurate — it saves re-explaining the project architecture in every session. See [memory notes below](#memory-and-context-files).

---

## Memory and Context Files

Claude Code maintains a persistent memory directory for this project at:
```
C:\Users\RussinT\.claude\projects\c--Users-RussinT-OneDrive---AECOM-Desktop-microstation-vba-project\memory\
```

### Files in This Directory
| File | What to Put There |
|------|------------------|
| `MEMORY.md` | Project overview, module roles, key patterns — loaded automatically into every session |
| `debugging.md` | Common errors, their causes, and fixes |
| `patterns.md` | Repeated code patterns, naming conventions |

### How to Update Memory
Tell Claude Code: *"Remember that [fact]."* Claude will update the appropriate memory file.

Or edit the files directly in Cursor.

### MEMORY.md Best Practices
- Keep it under 200 lines (lines beyond 200 are truncated)
- Use tables for module roles — they compress well
- Record stable facts only (architecture, patterns, preferences)
- Remove entries that are no longer accurate

---

## Common Pitfalls and How to Avoid Them

### Pitfall: Claude Stops Mid-Task
Sometimes Claude will say "I'll wait for the background agents to finish" and pause. If this happens, just tell it: *"Keep going, don't wait."* or *"Continue the task."*

### Pitfall: Edit Tool Fails — "File Has Not Been Read"
Claude's Edit tool requires a recent Read of the file in the current session. If you get this error, tell Claude: *"Read [filename] first, then make the edit."*

### Pitfall: Large File Reads Eating Context
Some `.frm` files are 1500+ lines. Reading an entire form file uses ~3,000–5,000 tokens. Instead, use targeted reads:
> "Read WZTCDesigner.frm lines 1580-1640 (the btnSubmit_Click handler)."

### Pitfall: Out-of-Sync VBA IDE
After Claude edits a `.frm` or `.bas` file on disk, the MicroStation VBA IDE still has the old version loaded. You must **re-import** the file in the VBA IDE (delete the old module, then File → Import File) to pick up the changes.

### Pitfall: Forgetting to Export After IDE Changes
If you modify code directly in the MicroStation VBA IDE (not via Cursor/Claude Code), you must **Export File** before Claude can see your changes. Otherwise Claude will be reading the old disk version.

---

## Recommended Development Cycle

```
1. Edit code (via Claude Code or Cursor)
         ↓
2. In MicroStation VBA IDE:
   - Delete old module
   - File → Import File (new version)
         ↓
3. Test in MicroStation (Alt+F11 → run macro)
         ↓
4. If bug found → back to step 1
         ↓
5. Export working file from VBA IDE to disk
         ↓
6. git add [file] && git commit -m "description"
```

### Using Claude Code for Git
Claude Code can run git commands directly:
> "Commit the changes to PerpPlacement.bas with a message describing the arc fix."

> "Show me what changed in WZTCDesigner.frm since the last commit."

---

## Tips for Writing Good Prompts for This Project

1. **Reference the architecture** — mention which step of the workflow you're in (alignment placement, sign drawing, etc.)
2. **Describe what the user does** — Claude understands user actions (click, right-click, draw arc) and can reason about the MicroStation input model
3. **Cite module names** — use exact module names (e.g., `PerpPlacement.bas`, not "the alignment module")
4. **Specify MicroStation constraints** — mention when something must work as a modeless form, or when you need `GetInput` loop behavior
5. **Say "do not delete"** when you only want additions — Claude defaults to minimal changes but it's good to be explicit about preserved functionality

---

## Tracking Rate Limits

### Anthropic Console
- Visit [console.anthropic.com](https://console.anthropic.com) → **Usage** to see daily/monthly token consumption
- Set up usage alerts to notify when approaching spend limits

### Claude Code Token Display
The Claude Code status line shows token usage in real time. Enable it with:
```
/statusline
```

### If You Hit Rate Limits
- Wait for the limit window to reset (usually 1 hour for Tier 1, or next calendar day)
- Switch to Claude.ai web interface for simple questions that don't require file access
- Use `/compact` to reduce context before continuing heavy file-reading tasks
- Break large tasks into smaller sessions across multiple days

---

## Quick Reference Card

| Task | How |
|------|-----|
| Start Claude Code | `claude` in project terminal |
| Show token usage | `/cost` |
| Compress history | `/compact` |
| Exit | `/exit` or Ctrl+D |
| Open Cursor chat | Ctrl+L in Cursor |
| Import file to VBA IDE | Right-click project → Import File |
| Export file from VBA IDE | Right-click module → Export File |
| Check API usage | console.anthropic.com |
