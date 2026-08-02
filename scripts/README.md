# Key-in batch probe (agentic growth path)

MicroStation has thousands of key-ins. We grow `Data/command-registry.tsv` in
**batches**, not one conversational turn at a time.

## Constraint (read this)

| Work | Parallel? |
|---|---|
| Research / harvest candidate lists into TSV | **Yes** — many agents |
| Live COM `SendKeyin` against MicroStation | **No** — one process, serial |

One MicroStation session cannot safely accept parallel key-in bursts (tool
state races). Harvest widely; probe once.

**Hang policy:** each live `SendKeyin` runs in a child process with a **3 second**
hard timeout. If it doesn't return, the child is killed, the key-in is recorded
as `HANG` → `unsafe-blocked`, and the batch continues. `tool` / `datapoint` /
`dialog` / `file` kinds are never executed (activate-and-wait or UI).

## Files

| File | Role |
|---|---|
| `Data/keyin-candidates.tsv` | Queue of candidates to probe (agents append here) |
| `scripts/keyin_batch.py` | Probe + promote |
| `Bridge/keyin-probe-batch.json` | Latest probe results |
| `Data/command-registry.tsv` | Authoritative gated catalog |

## Commands

MicroStation must be open with **`DELETE.dgn`** active (scratch file only —
never a real project sheet). The probe script refuses to run otherwise.

```bash
# Probe then promote
python scripts/keyin_batch.py run

# Or step by step
python scripts/keyin_batch.py probe
python scripts/keyin_batch.py promote --dry-run
python scripts/keyin_batch.py promote
```

## Candidate TSV format

```
opName  keyin  kind  requiredParams  source  notes
```

`kind` must be one of:

- `settings` / `view` / `lock` — OK probe → `verified-headless-safe`
- `tool` / `datapoint` / `dialog` / `file` — always `unsafe-blocked` (never verified)

### Crash-risk skip (never live-probe)

Fast COM OK ≠ safe. These previously destabilized MicroStation after a wave5
batch and are skipped by `keyin_batch.py` (promoted as `unsafe-blocked`):

- `MDL LOAD` / `MDL SILENTLOAD` / `MDL UNLOAD` / `MDL LIST`
- `EXPAND KEYIN`
- `DESIGN HISTORY *` (including `SET DESIGN HISTORY`)
- `NAMED VIEW ATTACH ALL` / `NAMED VIEW DETACH ALL`
- `PRINT LEVELS` / `PLOT LEVELS` (SendKeyin can hang indefinitely)
- `FACET/HIDDEN/VISIBLE EDGES` (hung wave7)
- `PUBLISHDGN PUBLISH *` (hung / file-writing)
- `PRINTORGANIZER *` (slow/hang)
- `IPLOT *`
- Non-toggle `PRINT VIEW/SCALE/AREA/…` (hung after PRINT COPIES)

The probe also heartbeats COM after each keyin and aborts naming the last
keyin if MicroStation dies.

Lines starting with `#` are ignored. Keep **CRLF** if VBA will ever read this
file; the probe script itself accepts either.

## Wave size (fewer check-ins)

Aim for **~250–400 candidates per probe** (4 harvest agents × 70–90, or one
seed TSV + agent merges). Harvest is parallel; probe is one serial COM pass
(~seconds if all fast). Bigger than ~500 is fine for settings/view/lock, but
hang/crash risk grows — keep the crash-risk skip list current and leave
MicroStation on `DELETE.dgn` only.

You only need to check in when: MicroStation dies, the probe hangs, or you want
a new harvest category. Otherwise one `run` after harvest is enough.

## How parallel agents should help

1. **Harvester agents** (parallel): each owns a category (REFERENCE, LEVEL,
   LOCK SNAP, VIEW, SET, …). They read Bentley Help / DOT manuals and
   **append** rows to `Data/keyin-candidates.tsv` (or write
   `Data/keyin-candidates-<category>.tsv` for a human/agent to concat).
2. **One probe run** (serial): `python scripts/keyin_batch.py run`
3. **Reviewer agent** (optional): read `Bridge/keyin-probe-batch.json`, flag
   suspicious OK results (unexpectedly slow, dubious spelling) for demotion.

Do **not** launch multiple probe processes against the same MicroStation.

## What this does not automate

- Full Help Key-in Index dump (no COM export found on this install)
- `UNDO` / `COMPRESS` / file open — leave `needs-testing` or omit from batches

---

# Drawing recipe probe (geometry bar)

Settings probes only prove a key-in returns. Drawing tools need a full
`COMMAND` + `DATAPOINT` + `RESET` recipe and proof that **geometry appeared**.

| File | Role |
|---|---|
| `Data/recipe-candidates.tsv` | Recipe queue with sample params |
| `scripts/recipe_batch.py` | Probe + promote |
| `Bridge/recipe-probe-batch.json` | Latest recipe probe results |

```bash
# MicroStation must have DELETE.dgn active
python scripts/recipe_batch.py run
```

**Pass bar:** graphical element count increases by `expectedNewElements` (default 1),
COM stays alive, recipe completes within 3s/step. Fast COM OK without a new
element is `NO_ELEMENT` → stays `needs-testing`.

Candidate columns:

```
opName  recipeLines  requiredParams  sampleParams  expectedNewElements  source  notes
```

`sampleParams` example: `x1=100|y1=100|x2=110|y2=100` — substituted into `{x1}` etc.
in `recipeLines` for the live probe. The registry row keeps placeholders.

Do **not** invent CadInputQueue sequences — copy `Legacy Files/` or `WZTCExec.bas`.
Circle/block/arc/smartline stay blocked until a proven in-repo sequence exists.
Hatch is seeded but may need shape association / pattern settings before it
promotes (see latest `recipe-probe-batch.json`).

## Sheet registry extraction helper

`python scripts/extract_sheet_signs.py [619-NNN ...]` dumps Book 3 owned pages +
candidate signs/tables for extending `Data/sheet-registry.tsv`. Review against
the PDF; do not paste blindly.
