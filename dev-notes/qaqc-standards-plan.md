# QA/QC + Standards Retrieval Plan (all MicroStation / CAD / engineering tasks)

Status: APPROVED + BUILT 2026-10-23 (engineer: no risk tiers — all tasks equal; PDFs in Standards/). Built: standards library + registry, structure-aware hybrid index (BM25 + vectors + rerank + cross-refs + OCR), qa_gate.py + criteria/failure-mode catalogs + checkers + Stop hooks (Claude Code + Cursor), chat-agent FINAL gate, engcalc/ drainage basics, fixtures (tests/qa). NOT yet live-tested: new VBA CAD read-back ops (MicroStation closed).
Scope: every drafting and engineering task in this repo (WZTC sheet builds, plan
markups, curb classification, quantities, future drainage / roadway / detour work),
not one project or discipline.

---

## 1. Why the current process fails on real design work

Evidence from this repo, not theory. Each row is a failure that a rule told the
agent to avoid, and the agent still shipped or nearly shipped.

| # | Real incident | Root cause class |
|---|---|---|
| 1 | Curb verify gate passed sheets on line *count*, not feet (early-bound `e.Length` silently failed → 0.01 ft/line). Loc19 falsely BLOCKED; every LF number was fiction. | Verifier had its own silent bug; never tested against a known-bad case |
| 2 | Flare read returned 0 flares on every sheet (early-bound `Coordinate`) → Depressed measured from DWS pad only, ~2–3 ft short per ramp. Earlier "PASS" sheets were wrong. | Silent failure inside the generator; gate only checked `Depressed > 0` |
| 3 | Loc10_SGD held Loc03-04's drawing. Gate said PASS (steel present). | Gate checks *presence*, not *identity* (does this file belong to this sheet?) |
| 4 | Template ramp kit copied into every SGD; batch classified the kit as islands and `E52F`/`FC58` as INNER on 5+ sheets. | No check that inputs are in-scope (site / authority) |
| 5 | Corner IDs assigned by sort order — 3/5 wrong on Loc03-04, wrote wrong rows toward the estimate. | Mapping by assumption, not by evidence (labels) |
| 6 | 2026-08-10: screenshot-based "confirmed" claims contradicted what the engineer saw. | Visual-only evidence accepted as verification |
| 7 | Loc 11 → 12 → 13 same failure, batch kept going (fake progress). | Gate was advisory; agent could continue |

Pattern: **the rules are good, the enforcement is not.** Rules live in
`.cursor/rules/*.mdc` / `CLAUDE.md` / skills and depend on the agent remembering
them. Verifiers share code paths (and bugs) with the generator. No gate proves it
can catch a bad result. Standards are pulled ad hoc (3 PDFs indexed) instead of
resolved per task.

## 2. Target architecture

```
 Task intake ─► Standards resolution ─► Criteria matrix ─► Execute ─► Independent verify ─► Gate ─► Evidence record
 (task card)    (registry + index)      (req → check)     (agent)    (deterministic +       PASS/FAIL/  (qa-record.json,
                                                                      AI review + human)    UNCERTAIN    regression set)
```

Principle: **the thing that produced the result never decides it is correct.**
Only a checker exit code / record status clears the gate. UNCERTAIN is FAIL.

## 3. Standards library (what to pull, how to pick the right ones)

### 3.1 Sources — have / add

| Source | Have? | Use |
|---|---|---|
| NYSDOT Configuration 2024 ZIP (`Standards/NYSDOT_Configuration2024_01.zip`) | yes | Machine-readable truth: level/feature dgnlibs, line styles, cells, sheet seeds, fonts, print styles, `Superelevation/HDM_2019_NHS.xml` + `NonNHS.xml`, `Sight Visibility ... .xml`, `NY_PayItemLookup.xlsx`, `Sign Lookup.xlsx` |
| NYC DDC Infrastructure General Requirements, Apr 2024 (`Standards/INFRA-GR-...pdf`, 122 pp) | yes | Governs DDC projects (HWP24QC1). §2.4 lists governing criteria (NYC DOT Std Details, DDC Design Guidelines & Directives, Street Design Manual, DEP Sewer / Water Main Standards, NYC Stormwater Manual); §3 plotting scales (1"=30' plan, 1"=2'/1"=5' profile), CAD format; §4 task-by-task deliverables |
| MUTCD Part 6, NYS MUTCD Supplement, NYSDOT Std Sheets Book 3 (`Project Documentation/`) | yes, indexed **(exists)** in `Data/manual-index.sqlite` (394 FTS chunks) | WZTC |
| NYSDOT HDM chapters (index: dot.ny.gov/.../dqab/hdm) — at minimum Ch 2, 5, 8, 10, 16, 18, 20, 21 | **add** | Design criteria. Ch 20 §20.5.4 = CAD QC checklist; Ch 8 rev 104 = drainage (effective design approval ≥ 2026-08-01) |
| NYSDOT Standard Specifications + Standard Sheets (other books) | **add** | Details / materials |
| NYSDOT Engineering Instructions / Bulletins (EIIS) | **add** | Supersede manual text between revisions — must be checked per HDM's own instruction |
| PDM, 3R Standards, CPDM, Utility Accommodation, Land Survey Std, ROW Mapping Manual, Pay Item Catalog | add as tasks need them | Scope / survey / estimates |
| NYC DOT Std Highway Specs + Std Details, Street Design Manual, DEP Sewer Design Standards / Stormwater Manual, DDC Spec Bulletins | **add** (DDC GR §2.4 names them) | City projects |
| Project-specific: Specific Requirements, design criteria memos, approved revisions | per project | Highest authority after law |

### 3.2 Standards registry (new: `Standards/registry.yaml`)

One row per document — the system must be able to say *which revision it used*.

```yaml
- id: nysdot-hdm-08
  title: HDM Chapter 8 Highway Drainage
  agency: NYSDOT
  discipline: [drainage]
  revision: "104"
  issued: 2026-04-24
  effective: "design approval >= 2026-08-01 or letting >= 2027-01-01"
  supersedes: nysdot-hdm-08-r103
  source_url: https://www.dot.ny.gov/divisions/engineering/design/dqab/hdm/hdm-repository/chapt_08.pdf
  local_path: Standards/nysdot/hdm/chapt_08.pdf   # gitignored binary
  sha256: ...
  jurisdictions: [state, federal-aid]
  chunking: section        # section | page | table
```

`scripts/standards_status.py` re-checks `source_url` revision/date monthly and
flags changes (no silent use of a superseded chapter).

### 3.3 Jurisdiction resolution — pick the right standards per task

A **task card** (new, see §4) states: project, owner agency, funding, discipline,
deliverable, design stage. Rules (data, not prompt text) in
`Standards/jurisdiction-rules.yaml`, e.g.:

- Owner = NYC DDC → DDC GR + its §2.4 list; add NYSDOT HDM only if State/Federal
  funded (DDC GR §4.1.3 says so explicitly).
- Owner = NYSDOT → HDM + Std Sheets + Std Specs + EIs newer than the HDM revision.
- Any MicroStation deliverable → NYSDOT config ZIP symbology + HDM Ch 20 (or the
  agency CAD standard the contract names).
- Conflict between two sources → UNCERTAIN, ask the engineer; never auto-pick.

### 3.4 Retrieval (build on `ingest_manuals.py` **(exists)**)

1. Generalize the ingester to be registry-driven (any PDF, per-doc chunk mode,
   store agency / doc id / revision / section / page as columns).
2. Hybrid lookup: exact section lookup (`8.7.4.4`), FTS5 keyword (exists), metadata
   filters (agency, discipline, effective). Vector search is optional later — exact
   numbers and section ids matter more than semantic similarity here.
3. New MCP + chat-driver tool `get_task_standards(task_card)` → returns the
   governing document list **and** a starter criteria list with citations. Must be
   added to `chat_driver._OP_NAMES` (known registration gap).
4. Every cited requirement carries `{doc_id, revision, section, page}`; the gate
   re-opens the page text and confirms the quoted value is present (evidence check,
   not just "a citation exists").

## 4. Task cards + criteria matrices

**Task card** (`tasks/<id>/task.yaml`, written by the agent at intake, confirmed by
the engineer for non-trivial work):

```yaml
id: hwp24qc1-curb-loc20
type: cad.classify.curb          # → selects criteria set
project: HWP24QC1
owner: NYC-DDC
inputs: [Screenshots/HWP24QC1_Location_20_SGD.dwg]
authority: [ddc-gr-2024, loc01-recovered-reference, engineer-chat-2026-10-01]
deliverable: steel curb layers + LF by Corner ID
risk_tier: 2                     # 1 cosmetic · 2 quantity/plan · 3 safety/structural/hydraulic
```

**Criteria matrix** per task type (`qa/criteria/<type>.yaml`): each row =
requirement → source → check → severity → who can clear it.

```yaml
- id: CURB-DEP-01
  req: Depressed spans flare-base to flare-base on INNER curb
  source: engineer rule 2026-09-29 (skill loc01-curb-classify)
  check: script:qa/checks/curb_depressed_span.py   # deterministic
  severity: 2
- id: CAD-SYM-01
  req: Elements By-Level; level exists in NYSDOT level library; proposed weight 2
  source: nysdot-hdm-20 §20.7.1
  check: script:qa/checks/cad_symbology.py
  severity: 2
```

Starter sets to write first: `cad.symbology` (all MicroStation work),
`cad.classify.curb`, `wztc.sheet_build` (wrap existing scorecard/rules gate),
`plan.markup` (wrap existing text-overlap gate), `estimate.quantity`.

## 5. The gate (make it automatic, not a rule)

### 5.1 One runner: `scripts/qa_gate.py`

- Input: task card. Loads the criteria set, runs every check, writes
  `outputs/qa/<task-id>/qa-record.json`:
  `{criterion, status: PASS|FAIL|UNCERTAIN, observed, expected, evidence:[file, handle/ID, page], checker_version}`.
- Overall = PASS only if every severity ≥ 2 row is PASS. UNCERTAIN → BLOCKED.
- Exit code 0 / 1 / 2 (PASS / FAIL / UNCERTAIN). Updates the existing
  `outputs/engineering-harness-manifest.json` pattern **(exists)** instead of a new format.

### 5.2 Enforcement points (the part that is missing today)

| Where | Enforcement |
|---|---|
| Claude Code | `Stop` hook in `.claude/settings.json`: if any `tasks/*/task.yaml` is `in_progress` and its latest qa-record ≠ PASS, block stop with the failing rows. Agent cannot end the turn claiming done. |
| Cursor | Same script from `.cursor/hooks.json` `stop`. |
| Batch scripts | Every batch calls `qa_gate.py` per item; save/advance only on exit 0 (generalize what `_acad_batch_classify_locations.py` does now). |
| In-MicroStation agent | `chat_driver` refuses FINAL unless the task's qa-record is PASS — generalize existing `geometry_qa_passed` / `visual_qa_passed` gates **(exists)** in `plan_workflow.py`. |
| Estimate writes | `_acad_live_curb_qty.py` (and any future writer) refuses to write rows whose source task is not PASS. |

### 5.3 Independence rules for checkers (from incidents 1–3)

1. **Read back the saved file**, reopened read-only — never trust in-memory state or the generator's log.
2. A checker **must not import the generator's measurement code**. Recompute (e.g. LF from line endpoints, flare stations from raw solids).
3. **Every checker has known-bad fixtures it must FAIL** (`qa/fixtures/<check>/bad_*.json` or tiny DWG/DGN). `pytest qa/` runs them; a checker that passes a bad fixture is itself broken. This catches the 0.01-ft Length bug and the 0-flare bug on day one.
4. **Identity checks** before content checks: file ↔ sheet (survey xref name, location numbers, extents), source handles not template signatures.
5. **Late-bind COM everywhere** (`win32com.client.dynamic.Dispatch`) in checkers; a COM exception is UNCERTAIN, never skipped.
6. Visual captures are supporting evidence only; they never clear a severity ≥ 2 row.

## 6. Turning standards into deterministic checks (highest leverage)

### 6.1 CAD conformance checker (`qa/checks/cad_symbology.py`) — applies to every MicroStation task

Source: HDM Ch 20 §20.5.4.2 + §20.7.1 + §20.9, config ZIP dgnlibs.

- Export the level library once (COM: attach `NY_Features_Levels_ElemTemp_Settings_Annotations.dgnlib`, dump level name/color/weight/style → `Standards/derived/nysdot-levels.tsv`; extends existing `Data/level-categories.tsv`).
- Per changed element: level exists in library; color/weight/style ByLevel (or matches level def); existing = weight 0, proposed = weight 2 (§20.7.1); Z elevation = 0 in 2D / not wild in 3D (§20.5.4.2.b); cell and line-style names exist in NY libraries (§20.7.5); text height matches scale per §20.9.6; working units = survey ft / in (§20.9.1); element extents inside the expected State Plane zone box (§20.5.4.2.e).
- Scope rule (CLAUDE.md): only elements the task touched — diff element IDs before/after.

### 6.2 Design-criteria checks from XML/XLSX in the config ZIP

- Superelevation rate vs radius / design speed: `HDM_2019_NHS.xml`, `HDM_2019_NonNHS.xml`.
- Sight distance: `Sight Visibility Tables and Equations Imperial.xml`.
- Pay items / units: `NY_PayItemLookup.xlsx` (every estimate row's item number + unit valid).
- Sign codes and sizes: `Sign Lookup.xlsx` vs `SignLibrary.bas` (known half-coverage gap).

### 6.3 Engineering calculations (drainage as the pattern)

LLM role: identify governing criteria + extract inputs with citations, list missing inputs. **Never the calculator.**
Python modules `engcalc/hydrology.py`, `engcalc/hydraulics.py` (rational method, Manning, inlet interception, HGL) with unit tracking (`pint`) and hand-checked unit tests against HDM Ch 8 / FHWA HEC examples. Output = compliance matrix: requirement · source · computed · limit · PASS/FAIL/UNCERTAIN. Risk tier 3 ⇒ engineer sign-off row is mandatory; the gate cannot PASS it alone.

## 7. Golden sets, fault injection, metrics

- **Golden set** per task type in `qa/golden/`: engineer-approved outputs with expected values. Curb: Loc01 RECOVERED (OneDrive), Loc20, Loc03-04 (S/D/C LF per corner). WZTC: approved 619-xxx builds.
- **Fault injection** (must be caught): drop flares; inject template kit; swap a file's content (Loc10 case); offset geometry 0.5 ft; wrong level; superseded standard revision.
- **Metamorphic tests**: translate all coordinates, reorder entities, rename handles → classification and LF unchanged.
- **Metrics** in `outputs/qa/metrics.json` per run: false-accept on fixtures (target 0), defect recall, escalation rate, golden-set regression. Every engineer correction becomes a new fixture (closes the loop the failure protocol describes in words today).

## 8. Rollout — each phase has its own verify

| Phase | Build | Done when (verified by) |
|---|---|---|
| **0 — Gate skeleton (first)** | `qa_gate.py`, qa-record schema, task card schema, Stop hook (Claude Code + Cursor). Wrap the two gates we already have: curb verify, sheet scorecard. | Hook blocks a turn with a FAIL record (demo); curb verify fixtures: bad-Length and 0-flare fixtures FAIL, Loc20 PASS |
| **1 — Checker hardening** | Independence rules §5.3 applied to existing verify scripts; identity check (file ↔ sheet); fixtures for incidents 1–5 | `pytest qa/` green; Loc10 now FAILs identity |
| **2 — Standards registry + retrieval** | `registry.yaml`, jurisdiction rules, registry-driven ingester, `get_task_standards` tool (MCP + chat_driver) | For 3 sample task cards (DDC ped ramp, NYSDOT WZTC sheet, NYSDOT drainage) the tool returns the expected doc set; every citation re-verifies against page text |
| **3 — CAD conformance** | Level library export, `cad_symbology.py`, wired into every MicroStation task type | Fixtures: wrong level / non-ByLevel / Z≠0 / bad cell name all FAIL; a known-clean DGN PASSes |
| **4 — Golden + metamorphic + metrics** | `qa/golden`, fault injection suite, metrics file | False-accept = 0 on fixtures across 2 consecutive runs |
| **5 — Engineering calc pattern** | `engcalc/` drainage module + compliance-matrix output + tier-3 sign-off row | Unit tests reproduce HDM/HEC worked examples within tolerance |

## 9. Decisions needed from the engineer

1. Approve the Stop-hook enforcement (it will block me/Cursor from ending a turn while a task's gate is not PASS). Allow an explicit `qa waive <criterion> --reason` path for you only?
2. Which standards to download first beyond what's in `Standards/` (suggest: HDM Ch 2, 8, 16, 18, 20, 21; NYC DOT Std Details; DEP Sewer Design Standards).
3. Where large PDFs live (keep gitignored under `Standards/` with sha256 in the registry, or a shared drive path).
4. Risk tier defaults per task type (proposal: curb/quantities = 2, WZTC sheet = 2, drainage/structural = 3).
5. Golden-set owners: you approve each golden case; the agent never promotes its own output to golden.

## 10. What this plan does not claim

- No standard beyond the three already-indexed PDFs, the config ZIP, and DDC GR has been downloaded or verified yet; URLs above come from dot.ny.gov pages read 2026-10-23.
- AI review is a supplement for semantic checks (omissions, conflicting notes), never the gate of record for severity ≥ 2.
