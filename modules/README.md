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
