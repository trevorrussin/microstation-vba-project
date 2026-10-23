---
title: Horizontal and vertical alignments - 2D HCL first, 3D ALG, stationing, best fit, annotation
when: Creating or checking an alignment (horizontal control line, ALG), choosing where to place it, setting start station / station equations, stationing direction, arcs and tangents, annotating stationing, opening a profile model for a vertical.
keywords: alignment, alignments, hcl, alg, horizontal, vertical, station, start station, station equation, stationing, complex by element, feature definition, ac_p, best fit, baseline tie, civil analysis, profile model, arc between points, tangent
---

# Alignments (ORD, NYSDOT / AECOM practice)

```yaml playbook
playbook:
  topic: alignments-horizontal-vertical
  title: Horizontal and vertical alignments - 2D HCL first, 3D ALG, stationing, best fit, annotation
  when: Creating or checking an alignment, start station / station equations, stationing direction, arcs, annotation, profile model
  keywords: [alignment, alignments, hcl, alg, horizontal, vertical, station, start station, station equation, stationing, complex by element, feature definition, ac_p, best fit, baseline tie, civil analysis, profile model, arc between points, tangent]
  sources:
    - "Assign _5 - Bruckner Blvd Alignments-20240626_100533-Meeting Recording.mp4 @ 01:53"
    - "Lunch and Learn - Corridors, Parametric Constraints & Point Controls-20230808_120243-Meeting Recording.mp4 @ 07:50"
    - "000_ORD_Tutorial.pdf page 120"
    - "ProjectWise Overview, ORD H&V Alignments.pptx slide 9"
    - "BRP - Vertical Profile.mp4 @ 00:29"
    - "Importing an .alg file into ORD-20230628_103215-Meeting Recording.mp4 @ 00:17"
    - "ORD Horizontal Alignments - Station Equations-20240315_151601-Meeting Recording.mp4 @ 00:22"
    - "BRP - Horizontal Alignment.mp4 @ 00:46"
    - "LAB - Vertical Alignments-20230710_133442-Meeting Recording.mp4 @ 00:35"
    - "Horizontal & Vertical Alignments and Building a Corridor-20230609_124202-Meeting Recording.mp4 @ 00:46"
  preflight:
    - {id: PF1, check: level, name: AC_P, why: "feature definition AC_P = Alignment, Control, Proposed"}
    - {id: PF2, check: level, name: ACSC_P_T, why: "stationing text (regular/cardinal) levels used by Annotate Element"}
    - {id: PF3, check: level, name: ACSCT_P, why: "stationing cardinal ticks"}
    - {id: PF4, check: gcs, name: "NY83-LIF", advisory: true, why: "Region 11 / NYC projects are in the Long Island zone"}
    - {id: PF5, check: state, key: annotationScaleFactor, value: "480", advisory: true, why: "annotate stationing at the model's design scale (480 = 1in=40ft) or the text is invisible/huge"}
  steps:
    - id: S1
      title: Two files - 2D HCL for plans, 3D ALG for modeling
      who: engineer
      manual_reason: "File creation and ProjectWise organization"
      how: |
        Create the horizontal control line (HCL) in a 2D file (open with MicroStation CONNECT) and the alignment (ALG) in a 3D file (open with OpenRoads). The ALG is
        what corridors and models are built from; keep it in its own file under ORD so nobody opens it to add text and corrupts the whole model. Plan sheets
        reference the 2D HCL (the plan presentation), under Model. Secondary alignments get their own files (`HCL_B`, `ALG_B`...). One alignment per ALG file
        is safest: if the file goes down you lose one alignment, not all.
      cite: "Assign _5 - Bruckner Blvd Alignments-20240626_100533-Meeting Recording.mp4 @ 01:53"
    - id: S2
      title: Design the alignment in 2D first, then trace it into 3D
      who: engineer
      manual_reason: "Civil geometry tools (Line, Arc, Complex By Element) create civil features, which the agent's plain-graphics tools cannot"
      how: |
        In a 3D file the cursor snaps to 3D elements: a 50 ft tangent snapped two feet up in elevation is no longer 50 ft. So draw lines/arcs in the 2D file,
        reference the 2D file into the 3D file and trace it (Line / Arc Between Points). The Tangent snap only exists for plain 2D elements, which is another
        reason to design in 2D; when you trace arcs use start, end, pass-through. Flatten the view/plane if a measured length looks too long.
      pitfall: "A measured 45 ft where 13 ft was expected = snapping to 3D."
      cite: "Assign _5 - Bruckner Blvd Alignments-20240626_100533-Meeting Recording.mp4 @ 04:04"
    - id: S3
      title: Where to put it - not on a vertical face, centered, stretched to the survey limits
      who: agent
      how: |
        Do not draw an alignment on a curb face or abutment: the profile will jog wherever the line is in front of or on top of the curb. Put it in the middle of the road.
        Best fit between two edges: draw a best-fit straight line along each edge, take the midpoint of each, snap perpendicular to the other line, connect
        the midpoints, extend to the limits, then measure the perpendicular distance to each edge at both ends (recorded: 16.01/16.04 both ends; offset 13 ft to the
        proposed curb gave 12.9x / 13.02). Always extend to the survey limits: you can never take length away later but you may have to add some.
      cite: "Assign _5 - Bruckner Blvd Alignments-20240626_100533-Meeting Recording.mp4 @ 16:26"
      test: {action: measure, distance: 13}
    - id: S4
      title: Name, feature definition AC_P, direction, and Complex By Element
      who: engineer
      manual_reason: "Properties dialog and Complex By Element tool prompts"
      how: |
        Set the Feature Definition to AC_P (proposed alignment) BEFORE drawing (Geometry > General Tools > Standards > Feature Definition Toolbar, Use Active Feature
        Definition on) or fix it in Properties. Give the alignment a real name (the road) so you know which one when several are referenced. Draw in the direction of
        stationing: upstate CAD manual = south to north and west to east so plans read like a book; downstate may differ, follow the project. Check true north is up with
        View > Rotate > Top. Join pieces with Geometry > Horizontal > Complex Geometry > Complex By Element (Method Automatic, Maximum Gap 0.033, Name); click the first
        element NEAR ITS START so the direction arrow points the right way. Chain Commands joins consecutive tangents; Simple Arc between two tangents needs a radius
        and Trim/Extend Both.
      cite: "000_ORD_Tutorial.pdf page 129"
    - id: S5
      title: Station format and never start at zero
      who: agent
      how: |
        Stations are written `10+00.00` (ss+ss.ss; in ORD Format Options `ss+ss.ss`). Never start an alignment at 0+00 or in the negatives: start at 10+00 (or 1000)
        so you can add length at the beginning later. The first station of a new alignment is set with Geometry > Horizontal > Modify > Start Station (Start Distance 0,
        Station 10+00); the tool picks 10+00 by default. Keep the station string format when writing station labels or tables.
      cite: "Assign _5 - Bruckner Blvd Alignments-20240626_100533-Meeting Recording.mp4 @ 37:35"
      test: {action: name, value: "10+00.00", pattern: "^[0-9]+\\+[0-9]{2}(\\.[0-9]+)?$"}
    - id: S6
      title: Hold a station at a point, then add a station equation (prefix) from the START
      who: engineer
      manual_reason: "Start Station / Station Equation tool prompts are interactive civil tools"
      how: |
        To make a point on the alignment a given station (for example where a skew meets the alignment must be 95+50): Geometry > Modify > Start Station, pick the alignment, click that
        point as the start distance and type 95+50; ORD works out the rest (the alternative is subtracting the distance along the element from 95+50 yourself). Verify with Civil
        Analysis > Analyze Point at several places along the alignment. A prefix is a Station Equation: Modify > Station Equation, back station = the FIRST station of the alignment
        (click its start), ahead station = the same station with the prefix typed (NB 77+43.18). If you equate at 95+50 the prefix only applies FROM 95+50 on, not to the whole alignment.
        Remove a bad equation in Explorer > OpenRoads Model > the alignment, then redo it from the start. This tool also fixes an alignment that started at 0: back station 0, ahead
        station 10+00.
      pitfall: "A prefix that appears only after a certain station means the back station was not the start of the alignment."
      cite: "ORD Horizontal Alignments - Station Equations-20240315_151601-Meeting Recording.mp4 @ 04:03"
      test: {action: calc, expr: "9550 - 1234.56", expect: 8315.44}
    - id: S7
      title: Offsets from a known edge keep the intent
      who: agent
      how: |
        When the engineer fixes the alignment offset from a feature (13 ft from the proposed curb), place the parallel at exactly that distance and verify it by measuring
        perpendicular to both. ORD offsets are negative LEFT, positive RIGHT; copy_parallel is + LEFT.
      cite: "Assign _5 - Bruckner Blvd Alignments-20240626_100533-Meeting Recording.mp4 @ 21:47"
      test: {op: copy_parallel, args: {element_id: $line, distance: 13}}
    - id: S8
      title: Tie the alignment to something a contractor can find
      who: engineer
      manual_reason: "Needs the survey baseline and baseline ties (project data the agent does not have)"
      how: |
        An alignment must be tied to a physical point in space: use the survey baseline ties, give the field a station/offset dimension to pin the alignment to. If the surveyor
        provided no baseline, ask the survey sub (alignments are not normally delivered). Do not invent the tie.
      cite: "Assign _5 - Bruckner Blvd Alignments-20240626_100533-Meeting Recording.mp4 @ 07:43"
    - id: S9
      title: Annotate stationing at the design scale
      who: engineer
      manual_reason: "Drawing Production > Element Annotation tool"
      how: |
        Check the model scale first: Model Properties design scale 480 (1in=40ft) and Line Style Scale = Global line style scale, otherwise the stationing text does not
        show. Drawing Production > Annotations > Element Annotation > Annotate Element, annotation group Stationing, Locate Elements, Reset to complete. Annotation is dynamic
        and follows the alignment (ORD training uses Annotation Scale Lock 1in=100ft for its example).
      cite: "Assign _5 - Bruckner Blvd Alignments-20240626_100533-Meeting Recording.mp4 @ 38:58"
    - id: S10
      title: Vertical - open the profile model with a terrain active
      who: engineer
      manual_reason: "Open Profile Model needs a view click; terrain activation is a pick"
      how: |
        Geometry > Vertical > Open Profile Model, pick the alignment, click another view. Set the existing terrain ACTIVE first (only one terrain is active at a time),
        then draw the vertical with the vertical Line/Curve tools (best-fit, no tracing). Join with Complex Geometry > Vertical, name it, give it the proposed feature
        definition. A profile can be set active like a terrain, so several profile alternatives can share one horizontal alignment.
      cite: "Lunch and Learn - Corridors, Parametric Constraints & Point Controls-20230808_120243-Meeting Recording.mp4 @ 12:48"
    - id: S11
      title: Delete 3D elements with the Delete tool, not the keyboard Delete key
      who: engineer
      manual_reason: "Deleting civil/3D elements is the engineer's call and must use the tool that removes associated data"
      how: "Use the Delete Element tool (X) on civil/3D elements so every associated item goes with it; the keyboard Delete key can leave orphans."
      cite: "Assign _5 - Bruckner Blvd Alignments-20240626_100533-Meeting Recording.mp4 @ 43:27"
    - id: S12
      title: Vertical - tangents first, then parabolic curves between them
      who: engineer
      manual_reason: "Vertical Lines / Curves tools in the profile view are interactive civil geometry"
      how: |
        Existing-ground terrain ACTIVE, Geometry > Vertical > Open Profile Model (pick the horizontal alignment, click a view). Draw ALL tangent lines first from the start of the project limits,
        then add the vertical curves: Curves > Profile Curve Between Elements > Parabola, pick tangent 1 then tangent 2, key in the curve length (e.g. 100 ft), Trim/Extend Both, Feature
        Definition AC_P. Join with Complex Geometry > Profile Complex By Elements (Method Automatic) so the profile is one element. The status bar (lower left) prompts each step.
      pitfall: "The recording is a tool demonstration, NOT a design: vertical curve lengths, K values and grades must meet the governing standard (state/federal/local). Search the standards library before choosing them."
      cite: "BRP - Vertical Profile.mp4 @ 02:43"
    - id: S13
      title: Edit a vertical alignment later in the table editor
      who: engineer
      manual_reason: "Vertical table editor dialog"
      how: |
        Table Editor > pick the profile: every tangent with its slope and every vertical curve length is listed; type a new curve length (200 ft) and Apply, or raise/lower a PVI elevation
        in the table. This is how a profile is updated later without redrawing.
      cite: "BRP - Vertical Profile.mp4 @ 04:30"
    - id: S14
      title: Do not state grades, K values or curve lengths from memory
      who: agent
      how: |
        Vertical design values come from the governing standard (NYSDOT HDM Chapter 2 and its design-criteria tables), cited by document, section and page via search_standards /
        get_standard_section, and the engineer's profile; the agent never invents a grade or curve length. The QA gate's citation check verifies a quoted requirement exists on the cited page.
      cite: "BRP - Vertical Profile.mp4 @ 01:21"
      test: {action: standard, doc: "nysdot/hdm/chapt_02.pdf", section: "2.6.9"}
    - id: S15
      title: Older project - import the alignments from an .alg file
      who: engineer
      manual_reason: "Import Geometry dialog / ProjectWise file picker"
      how: |
        Open the 3D alignment (ALG) file in OpenRoads, Geometry > Import/Export > Import Geometry, pick the .alg file from ProjectWise (the picker lists project documents; choose the
        alignment .alg, e.g. a copy of the old ORD alignment file). The alignments arrive as civil geometry in the active file; then check names, feature definitions (AC_P) and stationing.
      cite: "Importing an .alg file into ORD-20230628_103215-Meeting Recording.mp4 @ 00:17"
    - id: S16
      title: Read stationing and properties of an existing complex alignment
      who: engineer
      manual_reason: "Properties / Explorer panes are interactive"
      how: |
        Select the complex element: Properties show Feature Name, Feature Definition AC_P, ByLevel color/weight, last modified, locked state. The OpenRoads Model > Alignments list in Explorer shows every
        alignment (AC_P1, AC_P2, Sedgwick, Ramp_CurbEast ...). Modify > Start Station shows Ahead Station / Back Station / Distance (the recording used an ahead station of RCW 95+00 -
        a prefix/name before the station is allowed).
      cite: "Horizontal & Vertical Alignments and Building a Corridor-20230609_124202-Meeting Recording.mp4 @ 00:46"
    - id: S17
      title: Horizontal in the order the BRP recording uses - terrain active, draw plain, trace, complex, station, prefix
      who: engineer
      manual_reason: "Civil geometry tools are interactive"
      how: |
        1) Existing-ground terrain ACTIVE (Terrain > Edit > Active; the status bar says "active terrain model set"). 2) Draw the layout with plain lines and arcs first. 3) Trace it with the OpenRoads
        geometry tools: the first point shows a prompt where you set the Feature Definition (AC_P for NYSDOT settings); Arc Between Points with start, end, pass-through. 4) The pieces are separate features until you
        Complex By Element; the arrow on the first element is the alignment direction (click it backwards and the profile comes out reversed). 5) Start Station (never the default 0; the recording started at 433 for that
        project; start distance 0). 6) Add the prefix with a Station Equation from the first station (ahead station SB 433+00). 7) Verify with Civil Analysis going up and down the alignment.
      cite: "BRP - Horizontal Alignment.mp4 @ 00:46"
    - id: S18
      title: Vertical alternatives - several profiles per horizontal, one active
      who: engineer
      manual_reason: "Profile alternatives and Set Active Profile are interactive"
      how: |
        Design options (a shared-use path following the rail profile on the bridge, a 125 ft vertical curve beginning 14+77.97, grades 4.50% and 0.66%, an approximate location, a 25 mph design speed note) come from plan callouts and the
        governing criteria: read them from the plan PDF, do not guess. Build each option as its own alignment/profile (VOID_S Opt 1, Opt 2 ...), draw with Profile Line Between Points / Parabola, and make ONE profile active
        (Set Active Profile); the corridor uses the active profile. A corridor file separate from the ALG file references the alignment; Analyze Point reads station and elevation back for checking against the callouts.
      cite: "LAB - Vertical Alignments-20230710_133442-Meeting Recording.mp4 @ 08:07"
```


