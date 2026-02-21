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
> "I need to understand how the workflow connects — read AlignmentTool.bas, PerpPlacement.bas, DrawSign.bas, DrawElements.bas, and CellPlacer.bas in parallel and summarize their interfaces."

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

## CLAUDE.md — Project-Level Persistent Instructions

### What It Is

`CLAUDE.md` is a special file placed at the **root of your project folder**. Claude Code reads it automatically at the start of **every session**, before any conversation begins — including brand new sessions where Claude has no prior context.

Think of it as the "always read this first" file. Anything written there is in scope for every conversation without you having to repeat it.

```
c:\repos\microstation-vba-project\
├── CLAUDE.md          ← read automatically at session start
├── Modules\
├── UserForms\
└── ...
```

### How It Differs from MEMORY.md

| | CLAUDE.md | MEMORY.md |
|---|---|---|
| **Location** | Project root (checked into git) | Claude's local memory directory (not in git) |
| **When read** | Start of every session, automatically | Start of every session, automatically |
| **Best for** | Rules, constraints, behavioral instructions | Facts, module tables, architecture summaries |
| **Who edits it** | You (in Cursor or any editor) | Claude (when you say "remember that…") |
| **Shared with team** | Yes — it's in the repo | No — it's local to your machine |
| **Line limit** | No hard limit | 200 lines (content after line 200 is truncated) |

**Use CLAUDE.md for:**
- Things Claude should always do or never do in this project
- API constraints specific to MicroStation VBA (e.g., IncludeOnlyType doesn't exist)
- Which Legacy Files to read before implementing certain patterns
- Element level/color/weight rules that apply project-wide
- File sync protocol reminders

**Use MEMORY.md for:**
- Module roles table
- Key SharedState variable names
- Architecture summary
- Known bugs and their fixes

### What's in This Project's CLAUDE.md

The `CLAUDE.md` at the root of this project contains:

1. **Project summary** — what the tool does and the 6-step workflow
2. **"Read Legacy Files first" rule** — instructs Claude to check LegacyPrototype.bas and LegacyElements.bas before writing any CadInputQueue command sequence
3. **MicroStation VBA API constraints** — documents which ElementScanCriteria methods exist, when to use SendCommand vs SendKeyin, and the CenterPoint fallback chain
4. **Element properties table** — which level/color/weight applies to each element type so Claude never mixes them up
5. **File sync protocol** — reminds Claude about the Import/Export step between disk edits and the VBA IDE
6. **Debugging guidance** — what information to ask for when a CadInputQueue bug is reported

### How to Update CLAUDE.md

Edit it directly in Cursor or any text editor. It's a plain markdown file.

Good reasons to add something to CLAUDE.md:
- You've told Claude the same rule three times across different sessions
- Claude keeps making the same type of mistake (wrong API method, wrong level, etc.)
- You've established a new project-wide pattern that should apply everywhere

Example additions you might make over time:
```markdown
## New Rule Added After v1.2.0
When writing cost estimate code, always read DesignerRef.bas first — it contains
all the spacing constants and item codes used in cost calculations.
```

### CLAUDE.md vs CLAUDE_CODE_GUIDE.md

This guide (`CLAUDE_CODE_GUIDE.md`) is **documentation for you** — explaining how to use Claude Code effectively.

`CLAUDE.md` is **instructions for Claude** — it speaks directly to the AI in the imperative ("do this", "never do that").

Write CLAUDE.md as if you're writing instructions for a new employee who reads it before every shift.

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
6. **See Appendix A** For specific examples of how Claude Code reasons.

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


# Appendix A: Claude Reasoning — Extended Excerpts & Commentary


# Claude Reasoning — Extended Excerpts & Commentary (How Claude Thinks)

> **Purpose:** This appendix shows *how Claude reasons* — not just the technical conclusions it reached, but the **mental moves**, **self‑corrections**, **re‑evaluation**, and **decision‑making patterns** Claude displays. The excerpts below are from real conversations, with commentary explaining *how Claude organizes information, shifts hypotheses, and decides what to do next*.

---

## 1) Recognizing a Pattern → Switching to a New Hypothesis Mid‑Thought

**Excerpt**

> *“The `TEXTEDITOR PLACE` command might still be active when the function returns … queued commands are asynchronous … actually wait — if `StartDefaultCommand` fires too early, that could leave MicroStation expecting input even though VBA has moved on.”*

### What this shows about Claude’s reasoning
- Claude begins with **Hypothesis A** (“TEXTEDITOR PLACE still active”).
- Then something in the logic triggers a *second* insight, and it shifts to **Hypothesis B** (“StartDefaultCommand happening too early”).
- Notice the phrase: **“actually wait —”** → Claude is *overriding its own initial reasoning*. 
- This is characteristic of Claude’s chain‑of‑thought: when new implications appear, it **re-evaluates the situation** instead of sticking to its original guess.

### How it organizes thoughts
- Claude keeps multiple hypotheses active at once and rotates between them.
- It continuously checks: *“Does this explanation fit all observed symptoms?”*
- When it realizes a better match exists, it pivots.

This is Claude practising **iterative hypothesis refinement** — one of its strongest reasoning behaviors.

---

## 2) Shifting From Local Logic → Recognizing a Global Pattern

**Excerpt**

> *“frmAlignmentPlacement stays visible and doesn’t call `GetInput` inside the click handler … Let’s replicate that.”*

### What this reveals
Claude spots a **pattern that already works** somewhere else in your codebase and decides:
- “Instead of debugging from scratch…”
- “…use a **known-good example** as the model.”

This is Claude engaging in:
- **Pattern matching**
- **Analogy-based reasoning**
- **Local-to-global generalization**

It compares two workflows, identifies the difference (“this one hides the form, the other doesn’t”), and infers a solution.

### This is Claude’s reasoning move: *“Find the stable system and copy its rules.”*

---

## 3) Noticing a Conflict → Declaring a Root Cause

**Excerpt**

> *“Locking arises from asynchronous CadInputQueue state + re‑entrancy … the two sequences interleave and conflict.”*

### What this means in terms of reasoning
Claude identifies **two independent forces** causing the same issue and realizes they interact:
- Force 1: Asynchronous queue behavior
- Force 2: UI re-entrancy and event overlap

Then it merges them into a single conceptual diagnosis:
- **“Interleaving”**

This shows:
- **Causal chaining** (“A leads to B leads to C”)
- **Synthesis** (“These two problems are actually one larger problem”)
- **Conflict detection** (“Two things running at the wrong time”) 

Claude often combines multiple signals into a unified model.

---

## 4) Realizing a Hidden Rule → Rewriting the Mental Model

**Excerpt**

> *“The cell name is the sign number.”*

This appears obvious *after Claude says it*, but prior to that, the system didn’t have the rule explicitly.

### What Claude is doing
- Detecting a **data invariant** (“text and cell names must match because they represent the same selected item”).
- Updating its internal ruleset: *“The system’s behavior is governed by this relationship.”*

This is **constraint discovery** — Claude finds hidden rules that unify several observed behaviors.

---

## 5) Local Reasoning → Jump to a Better Global Strategy

**Excerpt**

> *“A cleaner fix would be to use the MicroStation VBA API directly … but the more likely minimal fix is to keep the form modeless and disable buttons.”*

### What Claude is doing
Claude is weighing **two different solution classes**:
1. **Ideal long-term fix** (direct API calls), and
2. **Low-risk short-term fix** (modifying UI/tool sequencing).

This demonstrates:
- **Cost–benefit reasoning** (risk vs reward)
- **Prioritization** (“better to fix the timing issue first, then refactor later”)
- **Software engineering judgment**, not just syntax

Claude often lays out multiple solution layers, then selects the one matching your constraints (low-risk, fast, minimal code changes).

---

## 6) Re-evaluating Earlier Thoughts When New Evidence Appears

**Excerpt**

> *“More likely, the real problem is that clicking Next Sign triggers GetInput which blocks while waiting for user input … the queued flow isn’t completing cleanly, so the screen locks up.”*

### What this shows
Claude starts with several guesses, but once it connects the symptoms:
- Blocking behavior from `GetInput`
- Queue processing not flushed
- User clicking before tool ends

…it **re-prioritizes its hypotheses**, promoting this one to the new “most likely cause.”

This is:
- **Bayesian-like reasoning** (reweighting hypotheses)
- **Diagnostic narrowing** (eliminate less consistent theories)
- **Confidence adjustment** (“more likely the problem is…”)

---

## 7) “Thinking Out Loud” — Internal Monologue Moments

Claude sometimes reveals parts of its self-dialogue:

**Examples from your excerpts:**
- *“actually wait—”* → catching a contradiction
- *“more likely the real problem is…”* → reprioritizing hypotheses
- *“let’s replicate that”* → pattern selection
- *“a cleaner fix would be…”* → optimization planning
- *“but the minimal fix is…”* → constraint‑aware decision

These moments expose **its internal decision tree**, even though Claude typically summarizes rather than dumping raw chain‑of‑thought.

---

## 8) Macro-Level Reasoning Structure Claude Uses

Across all excerpts, Claude repeatedly falls into the same reasoning architecture:

### 1. **Symptom analysis**
“What exactly is going wrong?”

### 2. **Hypothesis generation**
“Possibilities include A, B, C…”

### 3. **Pattern matching**
“Where have I seen similar logic in this project?”

### 4. **Conflict detection**
“These two actions overlap incorrectly.”

### 5. **Constraint discovery**
“Oh — cell name = sign number, that’s a rule.”

### 6. **Refinement**
“Actually, this explanation fits better…”

### 7. **Decision & action plan**
“Minimal fix: change UI flow. Long-term fix: API implementation.”

This is how Claude organizes and evolves its reasoning mid-dialogue.

---

## 9) Why This Commentary Matters

These examples teach you **how to structure your prompts** to get Claude’s best work:
- Give symptoms → Claude generates hypotheses
- Describe patterns → Claude checks analogies
- Ask for minimal fixes → Claude avoids over-refactoring
- Ask “is there a better way?” → Claude explores alternate strategies

Understanding Claude’s reasoning moves lets you *steer* it and recognize when it is:
- Revising a hypothesis
- Spotting a pattern
- Detecting a conflict
- Discovering a rule
- Scaling up to a better solution

Use this document as a guide for reading Claude’s thought process and leveraging its strengths in your MicroStation VBA workflows.

---

**Maintainer Note:** These excerpts and interpretations come from your real working notes and conversations. They are intended to help future maintainers understand *how* Claude reasons so they can replicate successful interactions.