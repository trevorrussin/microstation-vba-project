# Strengthen MCP / registry execution (ORD + MicroStation)

Goal: agents default to fast MCP/registry paths with verify-after-write; honest
engineer handoff for dialogs; UIA/vision only as fallback.

Updated: 2026-10-24 (all three research tracks merged).

## Method priority (authority)

See `.cursor/rules/ord-execution-priority.mdc`  
(MCP tool → registry + readback → COM script → UIA chrome → vision → engineer).

## Implement next (ranked)

Cross-cut from registry / MCP-ops / Civil research. Ship in this order unless engineer reprioritizes.

| Rank | Item | Track | Effort | Notes |
|-----:|------|-------|--------|-------|
| 1 | **Registry armed check + disarm + `armed`/`commandName` in result** | registry | M+S | **SHIPPED 2026-10-24** — `WZTCCommandRegistry.ExecuteRecipe` |
| 2 | **`set_level_display(level, on, view)` + readback** | MCP | M | **SHIPPED** — `wztc_ops` / `server.py` / chat_driver |
| 3 | **Enrich level display readback (per-view + library flags)** | MCP | M | **SHIPPED** — `GET_LEVEL_DISPLAY` / `get_level_display` (list stays fast/global) |
| 4 | **`hide_levels_recipe` (BYVIEW → OFF → verify)** | MCP | S | **SHIPPED** — escalate needsEngineer on fail |
| 5 | **Named CIF read wrappers** (`list_alignments`, `list_profiles`, …) | MCP/CIF | S | **SHIPPED 2026-10-24** — wrappers + `Data/cif-ops.yaml` + chat_driver |
| 6 | **`activate_model(name)` + first-class `list_models`** | MCP | S | Drawing vs Default; never switch DGN file |
| 7 | **Expose `get_element_cad_props` on MCP** | MCP | S | Already in wztc_ops; not on server.py |
| 8 | **`set_annotation_scale` + readback** | MCP | S | Thin registry wrap + describe_drawing_state |
| 9 | **`set_reference_display` + readback** | MCP | M | Ladder R4; start with ALL |
| 10 | **CIF `add_template_drop`** | Civil | L | **SHIPPED C# + MCP** (scratch/`production_ok: false` until live promote) |
| 11 | **CIF `set_active_terrain`** | Civil | M | **SDK blocked** — `ActiveSurface` get-only; wrapper returns `needsEngineer` + `get_active_terrain` |
| 12 | Registry promotion requires CLEAN audit; settings effect readback | registry | M | Hardens #1 long-term |

## Do not build

- Create Drawing / Civil Profile|XS Named Boundary / Open Profile Model (DocCenter / pick)
- View Options–Civil via UIA or key-in guessing
- Level Display dialog content via UIA (chrome-only)
- ProjectWise seed/check-in
- Key-in Browser list scrape (use SendKeyin)
- Template Drop / New Corridor status-bar prompts until CIF covers them

## Research detail

### Registry / false-OK

2252 rows (1888 verified, 291 unsafe-blocked). 2026-08-02 audit blocked 67 ARMED keyins. Live `ExecuteRecipe` still = “no COM throw.” Strengthen: post-CommandName check, disarm, return fields, promotion CLEAN gate, settings readback, pytest.

### MCP ops map

~150 MCP tools — drafting place/edit thick; **display/view-state + ORD discoverability** thin. Top gaps align with implement ranks 2–9. Dialog walls stay handoff + verify-after-engineer.

### Civil / CIF wall

Key-ins cannot create Civil. Path: CIF → VBA GET_CIF/CIF_EXEC → MCP (named wrappers + `Data/cif-ops.yaml`). Profile/XS Create Drawing permanent engineer. `add_template_drop` / `set_active_profile` / corridor props shipped; `set_active_terrain` engineer-only until a setter appears. Next: point controls, set_feature_definition, promote template-drop after scratch prove-out.

## Done nearby

- `mcp-server/ord_uia.py` — focus / open-keyin / dismiss-vba / probe
- `Data/ord-ui-catalog.yaml` — UIA surface + dead ends
- `.cursor/rules/ord-execution-priority.mdc` — method ladder
