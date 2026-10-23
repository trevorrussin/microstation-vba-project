---
title: New MicroStation / ORD file on ProjectWise - seed, shells, coordinate system, references, naming, check-in
when: Creating a new DGN (shell, feature, sheet, alignment file), choosing 2D vs 3D, setting the coordinate system, attaching survey/topo, naming files and checking in.
keywords: new file, new document, projectwise, seed, shell, coordinate system, state plane, ny83-lif, 2263, long island, 2d, 3d, file name, naming convention, check in, hcl, alg, cph, fea
---

# New file on ProjectWise (NYSDOT / AECOM)

```yaml playbook
playbook:
  topic: new-file-setup-and-coordinates
  title: New MicroStation / ORD file on ProjectWise - seed, shells, coordinate system, references, naming, check-in
  when: Creating a new DGN, choosing 2D vs 3D, coordinate system, attaching survey/topo, naming, check-in
  keywords: [new file, new document, projectwise, seed, shell, coordinate system, state plane, ny83-lif, 2263, long island, 2d, 3d, file name, naming convention, check in, hcl, alg, cph, fea]
  sources:
    - "Creating a new Microstation file on ProjectWise-20240418_092215-Meeting Recording.mp4 @ 00:19"
    - "MicrostationTipsAndTricks.pptx slide 4"
    - "MicrostationTipsAndTricks.pptx slide 22"
    - "ProjectWise Overview, ORD H&V Alignments.pptx slide 7"
  preflight:
    - {id: PF1, check: gcs, name: "NY83-LIF", advisory: true, why: "NYC work (Bronx, Manhattan, Staten Island, Brooklyn, Queens) is in the NAD83 NY Long Island zone, US Foot (EPSG 2263)"}
    - {id: PF2, check: state, key: is3D, value: "False", advisory: true, why: "everything except 3D modeling is developed in a 2D file"}
  steps:
    - id: S1
      title: Create the document from a seed (ProjectWise Advanced wizard)
      who: engineer
      manual_reason: "ProjectWise Explorer is an external application; the agent must not drive it"
      how: |
        ProjectWise Explorer: right-click in the folder > New Document > New Advanced Wizard. Pick the folder, then Select
        Template = the NYS seed (NYSDOT standards, 2D or 3D). The seed decides the application; fill in document properties only
        when the exact file is known. NYS seeds create two model views; a file whose model is named "Imperial" was NOT made from the NYS
        seed (first red flag that a file was set up wrong).
      cite: "Creating a new Microstation file on ProjectWise-20240418_092215-Meeting Recording.mp4 @ 00:19"
    - id: S2
      title: Make project shell files once, then copy them
      who: engineer
      manual_reason: "ProjectWise copy / rename"
      how: |
        At project start create project-specific "shell" files (2D and 3D) from the seed, prepare them (coordinate system, survey/aerial
        references), check them in, and COPY the shell for every later file instead of starting over. Copy in ProjectWise (Copy To /
        Paste, then Rename), not by hand-editing paths. Shell names: `<PIN>_shell.dgn`, `<PIN>_shell3D.dgn`.
      cite: "Creating a new Microstation file on ProjectWise-20240418_092215-Meeting Recording.mp4 @ 00:29"
    - id: S3
      title: 2D vs 3D
      who: agent
      how: |
        Develop everything except 3D modeling in a 2D file. A 3D file created by accident is exported to 2D (File > Export > 2D). The agent
        works in 2D models only (Z = 0 everywhere; the QA gate checks Z). Do not mix versions: OpenRoads 2021 R2 and 2022 R3 files are not
        compatible - keep everyone on the same release.
      cite: "MicrostationTipsAndTricks.pptx slide 4"
      test: {action: state, key: is3D, value: "False", advisory: true}
    - id: S4
      title: Coordinate system for NYC work is NY83-LIF
      who: agent
      how: |
        NYSDOT SPCS zones: West, Central, East, Long Island. Long Island also covers Staten Island, the Bronx and the tip of Manhattan:
        pick `NY83-LIF` (NAD83 New York State Plane, Long Island, US Foot, code 2263) for the Bronx projects, and verify against a map before
        trusting a zone name. Wrong zone -> elements land in the wrong place. A wrong system can be corrected later (Geographic Coordinate
        System dialog offers to reproject references/linework).
      cite: "Creating a new Microstation file on ProjectWise-20240418_092215-Meeting Recording.mp4 @ 10:43"
      test: {action: gcs, name: "NY83-LIF", advisory: true}
    - id: S5
      title: Set the coordinate system and attach survey/topo
      who: engineer
      manual_reason: "Geographic Coordinate System and Attach Reference dialogs"
      how: |
        Utilities/Geographic > Select Geographic Coordinate System (Library) > NY83-LIF. Attach the existing survey/topo and the bridge survey
        (ground surface and bridge surface are separate files) with Orientation Coincident - World. Reference attachments that are
        Georeferenced can be set Geographic - Reprojected. If a file looks odd after attaching, turn OFF annotation scale and set the model
        scale (Model > Properties; 480 = 1"=40') back to 1, then Ctrl+F to save settings.
      cite: "Creating a new Microstation file on ProjectWise-20240418_092215-Meeting Recording.mp4 @ 11:23"
    - id: S6
      title: Feature file names - HCL and ALG
      who: agent
      how: |
        NYSDOT PDM Appendix 14 names files `<PIN|D-number>_<cat>_<sub>_<description>.dgn` (category fea = feature, map = mapping, cph = contract
        plans highway, cbp = bridge plans). Practice for linework vs OpenRoads: plain horizontal control line work in `<PIN>_fea_hcl.dgn`,
        OpenRoads alignments in `<PIN>_fea_alg.dgn` (HCL = horizontal control line, ALG = alignment). Sheet files `<PIN>_cph_pro-1.dgn`.
      cite: "Creating a new Microstation file on ProjectWise-20240418_092215-Meeting Recording.mp4 @ 16:24"
      test: {action: name, value: "x73177_fea_alg.dgn", pattern: "^[A-Za-z0-9.]+_fea_(hcl|alg)\\.dgn$"}
    - id: S7
      title: Master reference file for plan sets
      who: engineer
      manual_reason: "Reference attachment settings on the sheet file"
      how: |
        General plans use a master reference file with everything attached; the plan file attaches only the master with Live Nesting,
        Nesting Depth 1, Display Overrides never. Annotation scale differences between file and reference give wrong line styles, text and cell
        sizes - turn annotation scale lock handling on the attachment.
      cite: "MicrostationTipsAndTricks.pptx slide 16"
    - id: S8
      title: Check in
      who: engineer
      manual_reason: "ProjectWise check-in"
      how: |
        Always save the file, save settings (Ctrl+F), then Check In (or Update Server Copy). Nothing may be left checked out overnight. If you forgot
        to check in do NOT re-open the file from ProjectWise: check in the local copy first or you lose work.
      cite: "ProjectWise Overview, ORD H&V Alignments.pptx slide 7"
```
