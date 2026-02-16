# Roadmap

Forward planning for the NYSDOT Workzone Traffic Control Designer project.

---

## Current Release: v1.1.0

The Workzone Traffic Control Designer is feature-complete for its core workflow:
- Workzone configuration with MUTCD NY spacing standards
- Alignment drawing with line/arc chain support
- Automated perpendicular line placement along alignment
- Sign placement with post, face, and text label
- WZTC construction element drawing
- Cell library symbol placement

---

## Planned: v1.2.0 — WZTC Cost Estimate Tool

**Status:** Planning

A new tool that automatically generates a Workzone Traffic Control cost estimate using the values already configured in WZTCDesigner.frm.

### Planned Features
- Reads the spacing and clearances table values (downstream taper, roll ahead, buffer space, etc.) from the current WZTC design
- Reads the sign selection table (sign numbers, quantities, sizes) from the current design
- Calculates unit quantities for each WZTC item based on the configured workzone parameters
- Generates a formatted cost estimate table with item descriptions, quantities, units, and unit prices
- Exports the estimate to a format suitable for inclusion in project documents

### Data Sources
- Spacing values from `SharedState.bas` (wztcDownstreamTaper, wztcBufferSpace, etc.)
- Sign data from `SharedState.bas` (wztcSignNumbers, wztcSignSpacings, wztcSignSizes, wztcSignSides)
- Unit prices from a configurable price table (to be defined)

---

## Planned: v1.3.0 — Detour Plan Designer

**Status:** Concept

Three new forms that allow users to develop detour plans automatically, following a guided workflow similar to WZTCDesigner.frm.

### Planned Features
- **Detour Route Designer form** — Configure detour route parameters including road names, turn-by-turn directions, detour length, and affected intersections
- **Detour Sign Placement form** — Automatically determine required detour signs (M4-8, M4-9, M4-10, etc.) based on the detour route configuration and place them at appropriate locations
- **Detour Plan Sheet form** — Generate a formatted detour plan sheet with route map, sign table, and notes for inclusion in traffic control plan sets

### Architecture
- Will follow the same sequential form workflow pattern as the WZTC Designer
- Will share the existing sign library infrastructure for detour sign data
- Detour state will be stored in `SharedState.bas` alongside existing WZTC state

---

## Planned: v1.4.0 — Workzone Summary Tables

**Status:** Concept

A new form that automatically generates Workzone Traffic Control summary tables from the current design data.

### Planned Features
- Reads all placed WZTC items from the current design session
- Generates a summary table listing all signs, elements, and cells with quantities
- Formats the table for direct placement in MicroStation or export to external documents
- Supports multiple workzone configurations within a single project

---

## Planned: v1.5.0 — Workzone Traffic Control Cross Sections

**Status:** Concept

A new form that automatically generates Workzone Traffic Control cross section views.

### Planned Features
- Generates typical cross section views showing lane configuration, shoulder widths, barrier placement, and sign locations
- Uses the lane width, shoulder width, and road type values from WZTCDesigner.frm
- Shows channelizing device spacing and taper geometry in cross section view
- Supports multiple cross section types (tangent section, taper section, work area section)

---

## Future Considerations

- **Save/Load project files** — Persist WZTC configuration to a file so multi-session projects don't require re-entry
- **Undo support** — Track placed elements for selective removal
- **MicroStation CONNECT Edition support** — Port from V8i VBA to CONNECT SDK or .NET add-in
- **Custom sign library editor** — GUI for adding, editing, and removing signs from the library without code changes