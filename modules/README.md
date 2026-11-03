# VBA source modules

Edit these UTF-8 sources instead of a monolithic `generator.bas`. Build from the
repository root with `cscript //nologo build.vbs`; see the root [README](../README.md)
for build and test instructions.

`project.json` is the ordered source manifest. The first three modules contain
constants, platform declarations, and shared state. Procedure modules follow.
Every BAS file must be listed exactly once, with a matching `VB_Name` header and
`Option Explicit`.

| Source | Responsibility |
| --- | --- |
| `constants.bas` | API identifiers, settings defaults, geometry and layout constants |
| `platform.bas` | Windows API declarations and window bounds type |
| `state.bas` | Shared runtime and drawing state |
| `background.bas` | Background SolidWorks windows and shutdown |
| `snapshots.bas` | Pack and Go targets, source restoration, backup cleanup |
| `inventory.bas` | Assembly traversal, bodies, quantities, drawing jobs |
| `jobs.bas` | Per-job drawing creation and artifact export |
| `sheets.bas` | Paper formats, templates, properties, title blocks |
| `views.bas` | Model views, flat views, fitting and centering |
| `six_views.bas` | Opposite views, two-sheet scale normalization, and dimension ownership across sheets |
| `semantic_dimensions.bas` | Source feature dimensions and geometric ownership of design constraints |
| `hole_patterns.bas` | Hole Wizard profile grouping and shared native hole callouts |
| `pattern_dimensions.bas` | Circular and rectangular hole locations, offsets, and pitch-circle graphics |
| `feature_constraints.bas` | Covered half-spans, coaxial diameters, and circular view quality |
| `sketch_symmetry.bas` | Sketch geometry evidence and local rotated-reference matching |
| `manufacturing_references.bas` | Model-topology circle classification, direct diameter references, profile view preference, and covered construction constraints |
| `manufacturing_callouts.bas` | Combined distance-angle chamfer labels and profile dimension placement beside hole-pattern datums |
| `profile_planner.bas` | Topological outline/slot classification and independent coordinate-constraint planning |
| `profile_drawing.bas` | Geometry-based view roles, native references, dimension lanes, and space allocation |
| `profile_routing.bas` | Filled-silhouette tests, material-free dimension bases, and leaders through open cuts |
| `planar_constraints.bas` | Whole-profile and local feature symmetries; independent contour length variables and checked vector closure |
| `planar_features.bas` | Constant-thickness prismatic recognition, circle loops and visible geometric references |
| `planar_drawing.bas` | Native aligned lengths, inclined angles, shared hole sizes and single-view ownership |
| `planar_routing.bas` | Outward dimension bases and crossing-aware placement for inclined planar edges |
| `planar_layout.bas` | Joint native candidate placement, rotated text and arc geometry, backtracking and compactness improvement |
| `planar_text.bas` | Cached Windows font measurements for native dimension text and safe GDI resource cleanup |
| `open_bore_slots.bas` | Open circular keyways, exposed-geometry dimension references, and directed label placement |
| `dimensions.bas` | Overall and diameter dimensions, validation, GOST styling |
| `native_dimensions.bas` | Native Autodimension entry point |
| `native_labels.bas` | Native dimension bounds, collision checks, label placement |
| `symmetry.bas` | Geometric symmetry evidence and dimension deduplication |
| `native_layout.bas` | View regions, native layout orchestration, centerlines |
| `coordinate_labels.bas` | Coordinate marker and label policies |
| `annotations.bas` | Drawing notes and assembly lists |
| `properties.bas` | Model properties, filenames, scales, text helpers |
| `runtime.bas` | Settings, logging, progress, document interfaces |
| `checkpoint.bas` | Queue persistence, retries, artifact validation, manifest |
| `orientation.bas` | Upright views and imported model dimensions |
| `feature_dimensions.bas` | Coordinate ordering and feature/radius dimensions |
| `flat_dimensions.bas` | Flat-pattern dimension sheets |
| `coordinate_requests.bas` | References, local datums, dimension requests and coverage |
| `detail_views.bas` | Same-sheet enlarged views and layout coverage |
| `coordinate_geometry.bas` | Visible coordinate collection and reference selection |
| `coordinate_layout.bas` | Dimension lanes, balancing and coordinate placement |
| `main.bas` | Macro entry point and batch orchestration |

## Native output

`dimension_appearance.bas` applies the common contour/dimension line weights and
the selected gray or black dimension color to every generated drawing.

These are source modules for a unity build, not independently imported VBA
components. `tools/vba_sources.py` strips their individual headers and combines
their bodies in manifest order. The native builder embeds that source in one
`Macro11` module inside `generator.swp`. This keeps the existing private scope,
shared state, library references, and `Macro11.main` launcher contract intact.
SolidWorks compiles the embedded VBA when it loads the project.

The build rejects missing, duplicate, external, or unlisted BAS files, mismatched
headers, duplicate procedures, missing entry points, and declarations placed after
procedures. Tests and external harnesses read the same composed source through
`read_source()`; they do not maintain their own assembly logic.
