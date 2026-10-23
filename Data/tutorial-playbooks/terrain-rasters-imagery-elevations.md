---
title: Terrain, aerial imagery (rasters) and surface elevations - import, display, attach, analyze
when: Importing or displaying an existing/proposed terrain (contours, triangles, interval), getting an elevation at a point (Civil Analysis), downloading and attaching NYSDOT ortho imagery or other rasters, remapping rasters on ProjectWise, or diagnosing a raster that will not load.
keywords: terrain, dtm, contour, contours, triangles, surface, elevation, civil analysis, analyze point, raster, ortho, orthoimagery, aerial, jp2, kmz, google earth, clearinghouse, proposed terrain, existing terrain, active terrain
---

# Terrain, rasters and surface elevations

```yaml playbook
playbook:
  topic: terrain-rasters-imagery-elevations
  title: Terrain, aerial imagery (rasters) and surface elevations
  when: Terrain import/display, elevations at a point, ortho imagery and raster attachment
  keywords: [terrain, dtm, contour, contours, triangles, surface, elevation, civil analysis, analyze point, raster, ortho, orthoimagery, aerial, jp2, kmz, google earth, clearinghouse, proposed terrain, existing terrain, active terrain]
  sources:
    - "Modifying Terrain Display Settings-20240423_103751-Meeting Recording.mp4 @ 04:57"
    - "Obtaining surface elevations from an ORD Terrain-20240515_123126-Meeting Recording.mp4 @ 02:58"
    - "Creating an Orthoimagery ORD File-20240426_110749-Meeting Recording.mp4 @ 00:03"
    - "Bringing Rasters into ORD-20240318_100529-Meeting Recording.mp4 @ 01:13"
    - "Civil Analysis in ORD-20240702_101546-Meeting Recording.mp4 @ 01:28"
  preflight:
    - {id: PF1, check: level, name: LELTR_CL, why: "Landscape, Elevation, Terrain, Contour Label"}
    - {id: PF2, check: level, name: LELTR_P, why: "Landscape, Elevation, Terrain, Proposed"}
    - {id: PF3, check: gcs, name: "NY83-LIF", advisory: true, why: "rasters carry geographic location: set the coordinate system BEFORE attaching imagery"}
  steps:
    - id: S1
      title: Terrain files - reference them in, copy them into the project civil folder
      who: engineer
      manual_reason: "ProjectWise copy and Reference attach"
      how: |
        Terrain may be its own file or inside the survey. Copy the terrain files from the pursuit folder into the project's civil OpenRoads terrain folder and attach from there (replace them when
        the official survey arrives). Attach with Orientation Coincident - World and check the attachment scale is 1.0 (a terrain attached at a strange scale shows the wrong size).
      cite: "Modifying Terrain Display Settings-20240423_103751-Meeting Recording.mp4 @ 01:13"
    - id: S2
      title: Display - contours vs triangles vs boundary, interval and labels
      who: engineer
      manual_reason: "Reference Properties / terrain display overrides are GUI settings"
      how: |
        Right-click the terrain > Properties: grey settings are the file's own. Turn Override on to pick another display template: Existing Boundaries, Contours, Contours and Triangles (proposed has its own set).
        Contour interval: Calculated Features > Contours (override on): minor 1 ft, major 5 ft by default; set a finer interval (0.2 ft) to find holes and glitches in the survey, then turn the override off to
        return to the simple display. Contour TEXT can be invisible only because the model/reference scale is 1:1: set the scale correctly rather than hunting the level.
      cite: "Modifying Terrain Display Settings-20240423_103751-Meeting Recording.mp4 @ 04:57"
    - id: S3
      title: Elevation at a point - Civil Analysis > Analyze Point
      who: engineer
      manual_reason: "Civil Analysis is an interactive civil tool"
      how: |
        Attach the terrain(s) first. OpenRoads Modeling > Home > Civil Analysis: "Select element to analyze point": click ANY contour of the surface you want (it only selects the surface), then hover the point; the
        status shows the elevation (and station/offset for an alignment). Turn the survey files OFF first: they hold 3D elements and the cursor snaps to their Z, giving survey-versus-proposed differences. Pick the
        right surface: do not use the mainline surface for a ramp or the reverse; "point is external to the terrain" means the wrong surface.
      pitfall: "A reading that disagrees with the surface by a fraction of a foot is usually a snap to a survey Z."
      cite: "Obtaining surface elevations from an ORD Terrain-20240515_123126-Meeting Recording.mp4 @ 02:58"
    - id: S4
      title: Terrain levels are usable and ByLevel
      who: agent
      how: "Contour labels (LELTR_CL) and proposed terrain (LELTR_P) come from the seed; do not recolor them. The agent never edits terrain models."
      cite: "Modifying Terrain Display Settings-20240423_103751-Meeting Recording.mp4 @ 06:17"
      test: {action: levels, levels: [LELTR_CL, LELTR_P]}
    - id: S5
      title: Download NYSDOT ortho imagery
      who: engineer
      manual_reason: "External website download (NYSDOT ortho clearinghouse) and file extraction"
      how: |
        Search "NYSDOT ortho imagery" and open the clearinghouse map, zoom to the project area, expand the sidebar, Ortho Results > Current View, Show Tile List, click each tile to download (zip; extract with 7-Zip).
        The file you need is the .jp2. Everything is delivered in US survey feet. Save locally.
      cite: "Creating an Orthoimagery ORD File-20240426_110749-Meeting Recording.mp4 @ 00:03"
    - id: S6
      title: Build the aerial file LOCALLY, coordinate system first
      who: engineer
      manual_reason: "Raster attachment and local file creation"
      how: |
        Create the aerial DGN locally, not on ProjectWise: MicroStation will not bring rasters in as US survey feet on ProjectWise (it reads them as meters), so they attach wrong. Set the geographic coordinate
        system first (NY83-LIF for NYC projects). Attach Tools > Raster / References, select all the .jp2 files and Open; the ACTIVE level at attach time is the level the rasters come in on, so set it first.
        Rasters hold pixels AND their geographic position, so they are huge: attach all, then clip away what you do not need. Loading takes minutes; the application can seem docked or frozen - do not click.
        Accept the Raster Attachment Options prompt without changes. The message "The file type is not supported" means the raster format/driver is not available in that session.
      cite: "Creating an Orthoimagery ORD File-20240426_110749-Meeting Recording.mp4 @ 02:04"
    - id: S7
      title: Verify placement with a KMZ in Google Earth
      who: engineer
      manual_reason: "Export and Google Earth are external steps"
      how: |
        Draw one test line on its own level, delete (or turn off) the aerials and Save As a test copy so the export does not try to export rasters, File > Export > Visualization File Types > Google Earth KMZ;
        Google Earth opens at the line. If it lands in the right place the aerials are in the right place.
      cite: "Creating an Orthoimagery ORD File-20240426_110749-Meeting Recording.mp4 @ 04:35"
    - id: S8
      title: Move to ProjectWise - re-map the raster links
      who: engineer
      manual_reason: "ProjectWise upload and raster re-mapping"
      how: |
        Moving the file to ProjectWise breaks the links to the locally stored aerials: re-map each raster to its ProjectWise location (Raster Manager). The first open of an aerial file on ProjectWise can take up
        to 15 minutes and looks empty until the links resolve. Have a teammate open the file to prove every aerial shows (red = not found); a raster mapped to your local working folder will not show for them.
        File name clashes in ProjectWise: add a suffix.
      cite: "Creating an Orthoimagery ORD File-20240426_110749-Meeting Recording.mp4 @ 06:30"
    - id: S9
      title: Station and offset of any snapped point
      who: engineer
      manual_reason: "Civil Analysis tool"
      how: |
        Civil Analysis > Analyze Point on an alignment gives station and offset (for example 310+04.53, 53 ft left of a barrier face); the feature you snap to must be a 3D or OpenRoads feature. Hovering a line shows
        which alignment it is. Useful to find the station of a barrier or ramp edge.
      cite: "Civil Analysis in ORD-20240702_101546-Meeting Recording.mp4 @ 01:28"
    - id: S10
      title: Convert State Plane to lat/long to check a location
      who: engineer
      manual_reason: "External tools (Earth Point, Google Earth)"
      how: |
        To check where a State Plane coordinate is, convert it (the recording used Earth Point with zone 3104 Long Island, US Survey Feet, e.g. 1025523.019 E 239406.947 N) and view it in Google Earth. NYSDOT
        Survey Manual Chapter 6 (Coordinate Systems and Datums) defines the zones; Long Island zone 3104 covers the Bronx.
      cite: "Bringing Rasters into ORD-20240318_100529-Meeting Recording.mp4 @ 02:05"
```
