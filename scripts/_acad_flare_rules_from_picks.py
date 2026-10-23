"""Flare triangle rules learned from engineer AutoCAD picks (Loc01, 2026-09-29).

Use this on the NEXT drawing so flares can be placed without re-prompting
when the same geometry pattern exists (outer topo + tan PRO-S + DWS/ramp).

## What the engineer drew (P1 / P2 / P3 → AcDbSolid on PROP-D-ROAD-RAMPS)

Not the old NW script recipe (short base on outer + tip inland by fixed depth).
Actual clicks form a **pocket triangle** that fills the gap between:

1. **Outer / green topo** (or a vertex on it)
2. **Tan PRO-S-CURB-LINE** (diagonal curb) — at least one corner ON or along it
   when that line exists and opens the pocket
3. **Third corner** — either another topo/curb vertex OR empty space that closes
   the blue-sketch pocket (OSNAP may need to be off for empty space)

### Observed Loc01 solids (engineer picks)

| Island end | Handle | Notes |
|---|---|---|
| MidN north (35.01) | 13F329 | P1 on PRO tip ~ (1028495, 187488); opens along tan |
| MidN SE | 13F31B | Picks on PRO verts [1]/[0] region |
| MidS north | 13F321 | Small pocket; 3 distinct pts (~2 ft edges) |
| MidS south | 13F323 | Small pocket near diagonal |
| East north | 13F32B | Larger ~6 ft edges |
| East south | 13F327 | Larger ~5–6 ft edges |
| NW reference | 13E9C9/CA | Still gold for NW; ~2 verts near outer, tip inland |

### Automation recipe for next DWG

1. Find each ped-ramp DWS / trail hatch + OUTER curb FL + offset PRO-S tan.
2. At each ramp end, build candidate triangle:
   - A = nearest outer/topo point at DWS end
   - B = point along tan PRO-S past any slope-arrow gap (ray / closest on PRO-S)
   - C = third pocket corner (outer along face OR empty space if no vertex)
3. Reject if triangle area < ~0.5 ft² or two corners coincide (OSNAP duplicate).
4. `AddSolid(A,B,C,C)` on PROP-D-ROAD-RAMPS color 2.
5. If PRO-S does not open a pocket (coincident with outer), fall back to NW-style
   short base on outer + tip toward island ~4–6 ft — but prefer engineer pocket
   when tan is offset.

### Curb layer authority (same session)

- **Corner** = inner curb segments under/within ~0.75 ft of a yellow flare triangle
- **Depressed** = inner curb between the two flares of a ramp (flare-to-flare)
- **Straight** = remaining inner curb beyond flares
- Never modify OUTER C-TOPO-FEAT / outer FL handles

Script: `_acad_reclassify_by_flares.py`
Ask UI: `acad_ask.py` (textbox always required)
"""
print(__doc__)
