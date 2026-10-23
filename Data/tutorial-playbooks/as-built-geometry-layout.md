---
title: Laying out as-built geometry - bearings, stations, 2D first, then the 3D alignment
when: Re-creating an existing alignment/geometry from as-built (record) plans: bearings and azimuths, stations of centerlines of bearings, curves with a radius, converting old state-plane coordinates, deciding how closely to hold the as-built.
keywords: as-built, as built, record plan, bearing, azimuth, direction, centerline of bearings, abutment, station difference, radius, coordinates, raster, rasterize, hand drawn, working line
---

# Laying out as-built geometry

Purpose in the recording: give the structural designers an alignment to work from. The goal is a line in space that matches the as-builts, not
every physical feature (signals, utilities, and so on are out of scope).

```yaml playbook
playbook:
  topic: as-built-geometry-layout
  title: Laying out as-built geometry - bearings, stations, 2D first, then the 3D alignment
  when: Re-creating existing geometry from record plans, bearings/azimuths, stations of bearings, radius curves, old coordinates
  keywords: [as-built, as built, record plan, bearing, azimuth, direction, centerline of bearings, abutment, station difference, radius, coordinates, raster, rasterize, hand drawn, working line]
  sources:
    - "Laying out As Built Geometry-20240422_104931-Meeting Recording.mp4 @ 00:05"
    - "Laying out As Built Geometry-20240422_104931-Meeting Recording.mp4 @ 09:13"
  preflight:
    - {id: PF1, check: gcs, name: "NY83-LIF", advisory: true, why: "the working file's coordinate system (Long Island zone for the Bronx)"}
    - {id: PF2, check: level, name: AC_P, why: "alignment, control, proposed"}
  steps:
    - id: S1
      title: Rasterize the as-built plan view and use it as the underlay
      who: engineer
      manual_reason: "Raster attachment (Raster Manager) is a GUI action"
      how: "Convert the as-built plan view to a raster and attach it in the working file; zoom into the stated coordinates and read them carefully (a 2 vs an 8 matters)."
      cite: "Laying out As Built Geometry-20240422_104931-Meeting Recording.mp4 @ 00:05"
    - id: S2
      title: Old coordinates do not convert - hold directions and distances instead
      who: agent
      how: |
        Coordinates on old plans (for example an easting starting 22,000) are usually in a plane that does not convert straightforwardly to NY Long Island State Plane (surveyors
        confirm this). Do not invent a conversion: ask the engineer. Throw the coordinates away and lock the geometry with the angles, azimuths and stationed distances on the plan.
        Directions in degrees are very sensitive: one second of arc off can skew the whole alignment, so expect the result NOT to match the as-builts perfectly, and say so.
      pitfall: "State the achieved tolerance. Never claim the as-built was 'matched' without measuring. Scale of the problem: a 1-degree direction error is 17.45 ft off at 1000 ft."
      cite: "Laying out As Built Geometry-20240422_104931-Meeting Recording.mp4 @ 01:06"
      test: {action: calc, expr: "1000 * 0.01745329", expect: 17.4533, tol: 0.001}
    - id: S3
      title: Quadrant bearings to drawn directions
      who: agent
      how: |
        As-builts give quadrant bearings (N29°54'00"E, S 57 08 20 W). Convert to an azimuth (clockwise from north), then to the drawing angle (counterclockwise from +X = 90 - azimuth);
        ORD key-ins accept N01:33:00:W. The agent's conversion is verified against the geometry it draws: N29°54'00"E has azimuth 29.9 degrees.
      cite: "Laying out As Built Geometry-20240422_104931-Meeting Recording.mp4 @ 09:13"
      test: {action: bearing, value: "N29°54'00\"E", expect_azimuth: 29.9}
    - id: S4
      title: Work in 2D - a snap to a survey point puts a Z on your line
      who: agent
      how: |
        Once a survey is attached the file behaves as 3D: if you snap to a survey point the line gets a Z, and its length becomes the 3D length (a 12 ft lane drawn in 3D is 12 ft in space,
        not 12 ft in plan). Keep the layout in a 2D file/model, check Z in Properties and set it to 0 (the length changes back to the plan length), turn off the survey/raster levels while drawing
        (an ACTIVE level cannot be turned off - switch the active level first). The QA gate's Z check flags a line with Z.
      cite: "Laying out As Built Geometry-20240422_104931-Meeting Recording.mp4 @ 07:00"
      test: {action: z_zero}
    - id: S5
      title: Fill in the puzzle - stations give distances
      who: agent
      how: |
        Draw every line the plan gives with its dimension (working line A, centerline of the east abutment bearings, centerline of bearings at the other end ...). The difference between the stations
        of two centerlines of bearings is the distance between them (for example 971.42 - 928.59 = 42.83 ft); check each piece against its neighbours, the way a puzzle checks out.
      cite: "Laying out As Built Geometry-20240422_104931-Meeting Recording.mp4 @ 09:52"
      test: {action: calc, expr: "971.42 - 928.59", expect: 42.83}
    - id: S6
      title: Curves - arc between points with the stated radius
      who: engineer
      manual_reason: "Arc Between Points is a civil tool; the stated radius and the station difference check are made in the tool prompts"
      how: |
        Where the as-built gives a radius, use Arc Between Points (start, end, pass-through) with that radius; no geometry (civil) element is needed in the layout file. Check the arc length
        against the difference in stations.
      cite: "Laying out As Built Geometry-20240422_104931-Meeting Recording.mp4 @ 11:16"
    - id: S7
      title: Trace the finished layout into the 3D ALG file
      who: engineer
      manual_reason: "Civil geometry tracing in the 3D ALG file"
      how: |
        When the horizontal matches, reference the layout file into the 3D ALG file and trace it there so you never snap to survey elevations; then look at the profile and draw the vertical.
      cite: "Laying out As Built Geometry-20240422_104931-Meeting Recording.mp4 @ 11:56"
    - id: S8
      title: Decide how closely to hold the as-built with the engineer
      who: engineer
      manual_reason: "An engineering judgement: hold the as-built for the structure, or use today's survey geometry"
      how: |
        Engineers differ: some hold the as-built as closely as possible to help the structural design, others use the survey geometry. Ask which, state the choice in the task, and report the
        difference between the as-built and the result.
      cite: "Laying out As Built Geometry-20240422_104931-Meeting Recording.mp4 @ 02:31"
```
