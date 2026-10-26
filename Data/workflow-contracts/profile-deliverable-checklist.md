# Profile sheet deliverable checklist - DRAFT for engineer correction

Status: DRAFT (2026-10-23). Not yet read by the QA gate. Engineer: strike, add or reword items; the gate should check THIS list, not a list the agent writes for itself.
Sources: NYSDOT PDF "Creating Profiles and Sheets from Alignment 7" p.1-2; playbooks profile-cutting-workflow (N1-N7), profile-sheet-annotation, alignments-horizontal-vertical, corridors-and-templates;
HDM 21.3.5.2 (p.23, preliminary plans for design approval - may not apply unchanged to contract plans), HDM Table 21-4 (p.19), HDM 5.7.4 (p.69-70).

Verified-by: A = agent can measure it; E = engineer confirms; ? = no way to check yet.

## A. Inputs (engineer supplies, never invented)
- [ ] A1 Alignment name, units (ft or m), PVI table (station, elevation, curve length) - E
- [ ] A2 Vertical design values meet the governing criteria (grades, curve lengths, K, sight distance), cited from HDM Ch.2/5 - E (agent cites; no check yet) ?
- [ ] A3 Alignment starts at 10+00.00 or later, never 0+00 or negative (BRP horizontal @03:58; L&L @11:39) - A (preflight blocks a start below 10+00.00)

## B. The civil profile
- [ ] B1 Profile exists on the named alignment, one single element (BRP vertical @04:03) - A (profile_verify: single_profile + continuity)
- [ ] B2 PVI stations/elevations match the table to 0.01 ft; grades to 0.01%; curve lengths as typed (HDM Table 21-4) - A
- [ ] B3 Each curve is parabolic with the typed length; first/last curve length 0; each curve fits its tangents - A
- [ ] B4 Feature definition AC_P (PDF p.1 step 2) - A when agent-built (set + read back); engineer-built: not readable (feature reads empty on existing profiles)
- [ ] B5 Profile is set ACTIVE (PDF p.1 step 3) - A (profile_verify: active_profile reads Alignment.ActiveProfile)
- [ ] B6 Alignment's original profiles unchanged (count and ranges) - A (baseline saved by profile_build)

## C. Annotation (HDM 21.3.5.2 list; keep only what the sheet needs)
- [ ] C1 Percent grades, location and length of vertical curves - E/A
- [ ] C2 Stopping and headlight sight distances - E
- [ ] C3 Banking diagram with horizontal control points and cross slope (superelevation) - E
- [ ] C4 Intersections and ramp take-offs - E
- [ ] C5 Existing and proposed drainage and utility crossings - E
- [ ] C6 Bridges/culverts, with measured vertical clearance; minimum from the Bridge Manual, not HDM 2.6.9 - E
- [ ] C7 Element Annotation with Profile Annotation group (PDF p.1 step 4) - E
- [ ] C8 Existing ground dashed ("( Dashed )"); bridge deck on level BDE; title in ny_Prop Title; scale bar 480_Vert - A

## D. Named boundary, drawing, sheet (PDF p.2)
- [ ] D1 Named Boundary, Civil Profile, Profile Only seed, 1"=40', VE 5, height 50, clearances 1, datum spacing 5/100, Use Active Vertical - A for VE, spacing, Use Active Vertical (profile_sheet_verify); E for seed, height, clearances
- [ ] D2 Boundary start/stop stations cover the PVI range - A (profile_sheet_verify)
- [ ] D3 Profile Drawing model and Sheet model exist, named after profile_name (PIN_cph_pro-N), no apostrophes - A (profile_sheet_verify)
- [ ] D4 Drawing attached to the sheet at 1/detail-scale - A (profile_sheet_verify)
- [ ] D5 NYSDOT border present, title block updated - E
- [ ] D6 Sheet DGNs checked in to ProjectWise - E

## E. Cleanliness
- [ ] E1 No test or scratch objects left in the production file (diff against the start snapshot) - A
- [ ] E2 No elements added outside the intended models; hidden-level/reference changes listed - A
