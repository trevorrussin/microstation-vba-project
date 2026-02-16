# Changelog

All notable changes to the NYSDOT Workzone Traffic Control Designer are documented in this file.

This project follows [Semantic Versioning](https://semver.org/): `MAJOR.MINOR.PATCH`
- **MAJOR** — breaking changes to workflow or data model
- **MINOR** — new features, backward-compatible
- **PATCH** — bug fixes and documentation updates

---

## [v1.1.0] - 2025

### Added
- **Full 6-step WZTC workflow:** WZTCDesigner > AlignDraw > PlacePerp > PlaceSign > PlaceElements > PlaceCells
- **WZTCDesigner form** with dropdowns for workzone category, NYSDOT 619 sheet number, road speed, road type, lane width, and shoulder width
- **Spacing and clearances table** auto-populated from MUTCD NY standards based on selected road speed, road type, and category
- **Dynamic sign table** with auto-fill from sign library when user enters a sign number
- **WZTC Order panel** with drag-to-reorder, showing the full placement sequence (spacing labels + signs + Work Area)
- **NYSDOT 619 Sheet Viewer** (SheetViewer.frm) with embedded WebBrowser for viewing standard sheets as PDF reference
- **Alignment drawing tool** supporting line and arc segments with endpoint memory for continuous chain building
- **Path geometry engine** (PerpPlacement.bas) with arc-length interpolation, tangent computation, and perpendicular line placement along arbitrary line/arc chains
- **Sign drawing** with click-to-place post, sign face cell, and text label at perpendicular line locations; supports One Side and Both Sides placement with connecting arc
- **WZTC element drawing** for 5 element types: Work Space (with auto-hatch), Channelizing Devices, Removal Striping, Temporary Barrier, and Temp Barrier w/ Warning Lights
- **Cell library placement** for 16 WZTC symbols from ny_plan_wztc.cel with interactive click-to-place
- **Sign library** with 10 default MUTCD signs, lookup by sign number, and custom sign support
- **WithEvents class modules** (PlaceButtons.cls, SignNumBox.cls) for dynamic control event handling
- **SharedState.bas** centralized state persistence across all form steps
- **Back / Return to Designer navigation** on every workflow form
- **SelectAndShow** method on SheetViewer for compact side-by-side reference from WZTCDesigner
- **Debug test suite** (DebugTest.bas, DebugDesignerLoad.bas, DebugWorkflowSequence.bas)
- **Project documentation:** README with architecture diagram, INSTALLATION_GUIDE with pixel-level control placement, ENVIRONMENT.md, CONTRIBUTING.md, ROADMAP.md, CHANGELOG.md
- **MIT License**

### Legacy
- Legacy prototype files (LegacyPrototype.bas, LegacySignPlace.bas, LegacyCells.bas, LegacyElements.bas, LegacyDesigner.frm, LegacyAlign.frm) moved to `Legacy Files/` folder for reference

---

## [v0.1.0] - 2025

### Added
- Initial prototype with hardcoded sign placement coordinates (LegacyPrototype.bas)
- Basic form for workzone configuration (LegacyDesigner.frm)
- Manual cell placement macro (LegacyCells.bas)
- Manual element drawing macro (LegacyElements.bas)

### Notes
- This was the proof-of-concept phase with hardcoded geometry
- All v0.1.0 code has been superseded by the v1.1.0 modular architecture