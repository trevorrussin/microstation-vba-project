---
name: fresh-context-review
description: >-
  Independent review of a change you did not write. Use after large CAD or
  code edits, or when the engineer asks for a second-pass review. Compare
  against requirements only; report defects; do not rewrite in the same turn
  unless asked.
---

# Fresh-context review

You did **not** author the change under review (or treat prior work as someone else’s). Do not defend it.

## Process

1. Restate the engineer’s **requirements** and **done means** in bullets.  
2. Inspect the artifact (diff, DWG inventory, script output, sheet QA) — not the author’s narrative.  
3. Check for:
   - Missing or inverted requirements  
   - Silent deletes / wrong file / wrong product (MicroStation vs AutoCAD)  
   - Outer vs inner mixups; authority-order violations  
   - Verify steps that were skipped or that cannot catch the failure mode  
   - Tests/scripts that print OK without proving the claim  
4. Report **concrete** defects (path, handle, LF, test name). No vague “looks fine.”  
5. Do **not** rewrite unless the engineer asks. Prefer a defect list + recommended next plan.

## Loc01 curb review extras

Read `.cursor/skills/loc01-curb-classify/SKILL.md`. Confirm:

- Depressed only flare-to-flare with splits  
- Straight includes long colinear runs  
- Corner only short-chord arcs outside depressed  
- Outers intact; sources not Deleted  
- `scripts/_acad_verify_loc01_curb.py` was run  
