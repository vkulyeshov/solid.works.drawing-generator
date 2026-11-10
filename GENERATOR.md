# SolidWorks Drawing Generator

Generate editable SolidWorks drawings and PDFs from a saved assembly or part. The launcher displays progress and automatically continues interrupted generation from a checkpoint.

## Requirements

- Windows with licensed SolidWorks and its VBA component installed. Generation has been tested with SolidWorks 2025 SP5.
- Windows Script Host (`cscript.exe`).
- The `GOST Type BU` font used by the supplied templates.
- A saved `.SLDASM` or `.SLDPRT` file with accessible dependencies.

Keep `launcher.vbs`, `generator.swp`, `generator.ini`, and the `templates/` directory together. Extract the entire ZIP before running the launcher.

## Quick start

Open a terminal in the directory containing `launcher.vbs` and run:

```bat
cscript //nologo launcher.vbs /input:"C:\Models\Assembly.SLDASM" /out:"D:\Drawings\Assembly" /format:A4 /background
```

For A3 landscape:

```bat
cscript //nologo launcher.vbs /input:"C:\Models\Assembly.SLDASM" /out:"D:\Drawings\Assembly_A3" /format:A3 /orientation:landscape /background
```

For an explicit template:

```bat
cscript //nologo launcher.vbs /input:"C:\Models\Part.SLDPRT" /out:"D:\Drawings\Part" /format:A4 /template:"templates\A4.portrait.drwdot" /background
```

Quote paths containing spaces. Use `/help` for the launcher reference. Positional arguments are also supported:

```bat
cscript //nologo launcher.vbs "C:\Models\Assembly.SLDASM" "C:\Settings\drawing.ini" "D:\Drawings\Assembly"
```

## Options

| Option | Purpose |
| --- | --- |
| `/input:FILE` | Required saved assembly or part |
| `/ini:FILE` | Settings file; defaults to `generator.ini` beside the launcher |
| `/out:DIRECTORY` | Drawing output; defaults to a timestamped `DrawingPack_YYYYMMDD_HHMMSS` folder beside the input model |
| `/format:A4` or `/format:A3` | Paper format |
| `/orientation:portrait` or `/orientation:landscape` | Automatic template orientation |
| `/template:FILE.drwdot` | Explicit drawing template |
| `/templates:DIRECTORY` | Directory containing the four automatically selected templates |
| `/excludesuppressed:1` or `/excludesuppressed:0` | Enable or disable exclusion of suppressed components and their descendants; default: `0` |
| `/background` | Use a separate SolidWorks process with its window outside the screen; close that process when finished |
| `/reset` | Delete the output checkpoint and regenerate the queue |
| `/macro:FILE.swp` | Use another existing native VBA project; defaults to `generator.swp` beside the launcher |

Command-line options override INI settings. Command-line relative paths use the terminal's current directory. Relative paths inside INI use the INI directory. Preserve the supplied INI's UTF-16 encoding when editing it.

The supplied INI selects A4 portrait. If the paper settings are absent, the fallback is A3 landscape. An explicit template determines orientation; its sheet size must match the selected A4/A3 format. The four supplied templates are:

- `templates/A4.portrait.drwdot`
- `templates/A4.landscape.drwdot`
- `templates/A3.portrait.drwdot`
- `templates/A3.landscape.drwdot`

`/background` still requires a licensed interactive Windows session and a graphics context. Without this flag, the launcher can use an existing SolidWorks instance and leaves it open. A source that already has unsaved changes is rejected before activation when creating a new queue. A clean source that SolidWorks marks dirty during activation or template validation can proceed: the launcher records the pre-activation check for the same source path and SolidWorks PID, and the macro captures that result before validating the template. The input is not saved or discarded. Alternatively, `/background` reads the saved files into a separate process and leaves an unsaved document in the user's process alone.

## Generation settings

Edit `generator.ini` to adjust generation:

| Setting | Purpose |
| --- | --- |
| `AutoDimensions=1` | Add graphical dimensions using geometry and feature planning |
| `DimensionPlanning=roles` | Classify supported flat profiles into outline and slot dimensions. Panels separate overall sizes, large edge cutouts and corner spans from local notch sizes and locations. The overview always retains the complete model contour, including small cutouts. Equal local sizes share a counted dimension. `legacy` retains the previous flat-profile workflow. |
| `ProfileSheets=auto` | Group supported straight-edged profiles by geometry. Rectangular, clipped and stepped panels separate the complete outline from local notches and rectangular openings. Elements on the same supporting edge share an enlarged view; equal local sizes share one counted dimension. Widths and locating gaps use separate rows. The page allocator moves whole families to another sheet only when needed. The two-recess A4 portrait layout stays on two sheets. Inclined prismatic profiles separate general dimensions from dense local end details. `legacy` restores the earlier fixed two-view layout. |
| `DimensionColor=gray` | Dark-gray dimensions and leaders, with color preserved in PDF. Use `black` for monochrome output. |
| `NativeDimensions=1` | Use source-feature dimensions and bounded native layout for ordinary solids; unparameterized solids fall back to native geometry dimensions |
| `OrthographicViews=auto` | Begin with the normal views and add opposite views automatically when feature geometry is hidden or an unparameterized solid needs wider coverage. `3` and `6` explicitly select the previous fixed layouts. Flat profiles use their role-based layout; folded parts and assemblies retain their normal overview. |
| `DimensionStrategy=features` | Prefer source sketch/feature dimensions and recognized hole patterns in three- and six-view solid drawings. `chain` explicitly selects native Autodimension chains. |
| `StrictDimensionLayout=1` | Stop when the legacy coordinate planner cannot establish complete collision-free placement; `0` retains its compatibility behavior |
| `ImportModelDimensions=0` | Disable model-dimension import; alternatives are `marked` or `all` |
| `UpAxis=+Y` | Model up direction; also accepts `-Y`, `+X`, `-X`, `+Z`, or `-Z` |
| `JobsPerBatch=8` | Maximum job attempts per batch; `0` disables the limit |
| `MaxJobAttempts=3` | Persistent total attempt limit per drawing job |
| `MaxJobs=0` | Select all jobs; a positive value limits a new queue |
| `JobFilter=` | Optional path or body-name substring for selecting jobs |
| `ExcludeSuppressed=0` | Suppression filter off by default. Set `1` to skip suppressed component occurrences and their descendants. `/excludesuppressed:0|1` overrides this setting. |
| `Organization=` | Optional organization text for drawing notes |

The suppression filter is optional and disabled by default. To exclude suppressed component occurrences and all descendants of suppressed subassemblies for one run:

```bat
cscript //nologo launcher.vbs /input:"C:\Models\Assembly.SLDASM" /out:"D:\Drawings\Assembly" /excludesuppressed:1
```

Use `/excludesuppressed:0` to disable the filter, including when the INI enables it. Alternatively, set `ExcludeSuppressed=1` in `generator.ini`. With the filter enabled, each occurrence and its parents are checked in the active assembly configuration, including configurations referenced by subassemblies. Hidden components remain eligible. If the same file/configuration has an active occurrence elsewhere, it still receives one drawing with only the active quantity. The root input assembly is always eligible.

With the filter disabled, Pack and Go includes suppressed references. The generator walks saved top-level references separately for each assembly/configuration and opens their copied model files from `Models/`, including documents unavailable through `GetModelDoc2` while suppressed. Already loaded documents are reused only when their paths match the owned copies; unavailable or stale pointers fall back to opening those copies. It does not call `SetSuppression2` or depend on unsuppressing pattern instances. Child references are captured as plain data before other configurations are activated, and each assembly/configuration plan is cached while repeated occurrences still contribute to quantities. Suppression states and source files are not saved or changed by this inventory step. Missing copied files, missing configurations and cyclic assembly references fail planning instead of silently excluding requested models.

The resolved filter value is printed as `Exclude suppressed: 0` or `1` and stored in the run settings. Use a new output directory or `/reset` after changing the filter or model suppression, so the queue is rebuilt instead of resuming its previous snapshot. A different filter value cannot resume an existing queue.

Repeated instances of the same file/configuration share one drawing and a quantity. Flat prismatic profiles use a role-planned sheet; folded sheet-metal parts retain their folded views and a flat-pattern sheet. Drawing properties such as `Designation`, `Description`, `Material`, and `Mass` populate the title block. Review manufacturing dimensions, tolerances, datums, and drawing content before release.

The default planner runs for every part without filename or job-number exceptions. It currently recognizes axis-aligned constant-thickness prismatic profiles with orthogonal straight boundaries and polygonal openings. It orders the boundary, identifies inward rectangular recesses, and assigns overall/profile dimensions to the upper view and slot dimensions to the lower view. A profile without slots gets a single face view. Thickness is stated with the technical notes. The same classifier is also attempted on developed sheet-metal geometry. Other solids use source-feature planning followed by native geometry fallback; unsupported developed contours retain the bounded coordinate planner.

For each planar coordinate axis, the planner builds a graph of model references and proposed dimensions. It keeps a spanning set of independent dimensions, prioritizing overall sizes and local slot sizes. A redundant closing interval is omitted only when its endpoints are already connected by retained dimensions. Equal numbers at different coordinates are not treated as duplicates. Simplification is limited to 0.005 mm, below the 0.01 mm displayed precision. `dimension-plan.tsv` records each retained dimension's role, references and native name; `run.log` records both selected and dependent source intervals. Native save/reopen tests verify the measured values and coverage.

Constant-thickness planar parts with inclined straight boundaries and circular through holes use a feature constraint planner when the orthogonal profile classifier does not apply. It checks complete boundary reflections together with the hole centers, and recognizes equal opposite sides of local U features. Coincident numeric values alone do not establish a repeated dimension. The contour uses true aligned edge lengths and one angle per inclined direction family instead of chains of X/Y projections. Two independent contour lengths may be omitted only when the closed vector loop reconstructs both to within 0.005 mm. The original lengths, closure coefficients and selected state are recorded in `planar-constraints.tsv`; native dimension ownership and measured values are recorded in `planar-dimensions.tsv`. Hole diameters are grouped on one view, with center positions and pitch dimensioned separately. The three views on one sheet distribute the contour and hole measurements, and each selected length group has one owner. This recognition currently requires a straight outer loop and circular inner loops; curved outer contours and noncircular internal openings retain the other planners.

Inclined planar profiles use an overview and enlarged native end details on the same sheet. Their dimension placement considers the full native annotation, including rotated text, arrowheads, extension lines and angular arcs. Text widths come from cached Windows measurements of the actual annotation font. A backtracking search rejects overlapping labels, lines through labels, intersecting dimension baselines and graphics outside their owning view region. Inside and outside arrow placements are evaluated explicitly, with a strong preference for inside. A label can move outside the measured interval while the arrowheads stay inside. The actual paper-space arrowhead sizes and directions must fit, and their triangle outlines participate in text-collision checks. Outside arrows remain available when the inside geometry does not fit or conflicts with other annotations. Extension lines may cross; compactness scoring favors shorter extensions and leaders. Angular labels remain on their measured arcs; the planner can use an equivalent supplementary angle to avoid a long circular text leader. Local principal-direction construction lines define those angles from the generated model geometry. Detail identifiers and captions use Ukrainian Cyrillic letters and a compact letter/scale notation. Regenerate the drawing after modifying the model. The current native and PDF integration fixture validates this layout on A4 portrait; review each other model and format in its generated PDF.

Text that cannot fit between the references gets a free leader position. Lane demand reserves space before label placement. The final native pass checks paper bounds and prevents graphics from crossing another view. A failed layout is reported as a job failure instead of exporting that failed result. `DimensionLayoutAudit=1` optionally records native planar candidate geometry in `planar-layout-candidates.tsv` for diagnosis.

Profile dimension placement also checks the filled silhouette, including polygonal openings. Dimension baselines and text leaders must avoid material. Bottom-opening slot widths are placed below the view. A shared vertical slot depth is placed outside the nearer side of the profile, with thin solid extension lines at the two measured model coordinates. These declared extension lines may pass over the profile; `profile-witnesses.tsv` records their actual reference points for independent verification. Other depths use the empty cut when no external base is clear, and their labels exit through the opening. Witness overruns are included in base selection. `profile-contours.tsv` records the projected loops for an independent audit after saving and reopening the drawing.

Dimension styling uses the supplied GOST 2.307 rules: closed arrowheads with a 2.5 mm head length and a 20 degree angle, a 2 mm extension beyond the dimension line, 7 mm row spacing, and at least 10 mm outside offset from the view. Short chains can use separate clear dots where they fit. Insignificant trailing zeroes are removed without reducing the configured measurement precision.

Visible contours use a 0.6 mm line; dimension and extension lines use 0.25 mm, within one third to one half of the contour width under GOST 2.303. `DimensionColor=gray` adds RGB 96,96,96 as a reading aid to the dimension text, arrowheads, dimension lines and extension lines. Geometry and title-block text keep their existing colors. The exported PDF preserves this gray; `DimensionColor=black` restores monochrome dimensions and PDF export. Color differentiation is not claimed as proof of DSTU compliance. Use the black mode when the release specification requires monochrome drawings.

Arrow geometry alone does not establish compliance of the complete layout. The default strict coordinate mode rejects a layout when it cannot place all required dimensions. It can rebuild interleaved coordinates as local datum graphs and try enlarged views on the same sheet, checking that every source coordinate remains dimensioned. `StrictDimensionLayout=0` enables the older coordinate compatibility mode with `WARN LAYOUT REVIEW` diagnostics. The final native pass still checks view isolation, text collisions and paper bounds. Dense LeftSide A4 layouts remain a known limitation; a successful package build does not establish manufacturing compliance.

With `NativeDimensions=1`, ordinary solids prefer source-feature planning and use native Autodimension chains when design dimensions are unavailable. After native Auto Arrange, the postprocessor removes repeated physical intervals and dimensions of the same circular/cylindrical geometry. Reflected intervals are grouped only when the selected body's vertices and sampled edge geometry verify the symmetry; equal numbers alone are insufficient. Each dimensioned view has a separate rectangular dimension region. Dimensionless isometric views remain geometry obstacles, and dimension graphics cannot cross another view or another dimension's text. The upper orthographic row can move together to reserve lower-view space, without changing scales. Readable native labels remain on their dimension lines; conflicting labels use inclined leaders and can use 2.5 mm text when the standard text does not fit. Multi-line hole callouts use 2.5 mm text. Verified mirror planes receive centerlines. Removed dimensions and their retained representatives are recorded in `native-dimension-audit.tsv`; `native-view-regions.tsv` and `native-symmetry-lines.tsv` record layout boundaries and centerlines. Assemblies and sheet-metal parts retain their existing dimension sources and receive the common final layout validation. Review manufacturing completeness before release.

With `OrthographicViews=6`, dimension ownership is assigned across both sheets before placement. A geometric measurement has one retained owner; repeated views do not repeat that measurement. Equivalent radius and diameter annotations of the same circular geometry share an owner, while distinct hole/thread specifications remain separate. The most complete repeated-pattern callout takes priority. `six-view-dimensions.tsv` records retained and removed annotations and their normalized measurements; `six-view-transforms.tsv` records the final view directions and scales. Both sheets use the same orthographic scale. This mode currently supports ordinary solid parts only.

The default `features` strategy imports the part's design dimensions instead of generating chains between every visible vertex. It recognizes equally spaced circular hole patterns and rectangular rows, supplies their location dimensions and a boundary offset, and combines Hole Wizard holes of the same bore and specification into one native size/thread callout with the total quantity across the part. Pattern locations remain separate. Coaxial closed circles missing from the source dimensions receive a diameter. Rotational feature radii display as diameters. A source dimension has one owner across the available sheets; circular dimensions prefer a view showing the corresponding geometry clearly. In `auto` mode, hidden required geometry triggers an opposite-view sheet and replanning. A reverse view can remain dimensionless when its measurements are already covered.

Redundant half-spans are removed only when a full span in the same source sketch and a verified body mirror plane cover them. Local repeated sketch dimensions are combined only when their actual reference pairs match under rotation about that sketch's center and their measured values agree; equal numeric values alone do not qualify. `feature-dimensions.tsv` records the imported candidates, and `run.log` identifies each omitted constraint and its coverage. Source-less imported solids fall back to native geometry chains; `DimensionStrategy=chain` explicitly selects those chains. The planner does not infer manufacturing tolerances or datum choices.

After updating the dimension strategy or generator layout, use a new output directory or `/reset` to regenerate existing completed drawings.

Diameters are created as drawing references to visible model edges; imported model-radius annotations are not reformatted in place. Circular hole patterns receive an associated construction circle and a true diameter dimension with one leader arrow. Profile dimensions prefer the source feature's visible face, so a fork width is not placed over its hidden reverse projection. Cylindrical bores prefer the front view. Circle classification uses the original model topology because drawing-edge proxies can omit adjacent-face information. A chamfer rim without an adjacent cylindrical face is excluded from supplemental diameters. A radial wall thickness is omitted only when its collinear references and value are covered by two planned coaxial diameters.

Construction half-widths are omitted only when their datum passes through the profile center, is parallel to the material boundary, the material contour has verified rotational symmetry, and a full width references the same boundary orbit. Construction lines are excluded from the material-symmetry comparison. For Matochina4x100 this replaces the 13 mm construction half-width with the full repeated rib width, distinguishes the 98 mm fork span from the 100 mm pitch circle, and identifies the central boss by its 58.8 mm outer diameter and 3 mm height, with a 38.8 mm bore.

Centered rectangular cuts opening into cylindrical bores use the exposed keyway width and depth, measured between actual model vertices at the bore intersection and the keyway floor. The rectangle's center offset and its hidden portion inside the bore are omitted. Width labels stay above the view; nearby diameters move to the left. A circular bolt-pattern leader on the same view points to the right-hand part of the pitch circle, away from the hole centers. The bore diameter remains on its existing owning view.

Rectangular hole patterns retain the center-to-center pitch above the view and a boundary offset below it. The axial profile sizes stay together on the first sheet; the overall length is not repeated beside the hole pattern. For Matochina4x100 these are a 77 mm overall length on the first sheet, a 28 mm longitudinal hole pitch, and a 7 mm offset from the free end on the second sheet. The 6 mm flange step is shown with the first-sheet profile. Dimensions always follow the model: a symmetric offset is not inferred at the other end. Distance-angle chamfers combine the model distance into the angular annotation with a localized chamfer prefix, for example a 1 mm x 45 degree chamfer, and omit the separate linear chamfer dimension.

The longest axial profile dimension is placed below its first-sheet view; the shorter profile steps are placed above it. This choice uses the source geometry rather than a fixed part name or nominal value. Centered labels below the view retain their normal text height when they fit between the references; small spans use the existing smaller text height. For Matochina4x100 the 77 mm length is below the lower-left view and the 3 mm, 6 mm, and 9 mm steps are above it.

For a vertically projected axial profile, the same policy uses the left side for the overall length and the right side for the steps. Rectangular hole pitches and boundary offsets also use opposite sides when their dimension direction is vertical. Boundary selection examines both projected axes and both end directions, so rotating a view does not substitute an unrelated transverse offset for the nearby end datum. Construction circles and generated centerlines stay black in color PDF output.

## Output and progress

Keep drawing output separate from the program directory. The console displays the current operation, completed/total jobs, and overall progress. The output folder contains:

- `Drawings/`: editable `.SLDDRW` drawings.
- `PDF/`: exported drawings.
- `Models/`: the model snapshot used for generation.
- `generation-state.tsv`: the persistent job queue and checkpoint.
- `failed-jobs.tsv`: failed jobs, attempt counts, and last errors.
- `manifest.tsv` and `dimension-coverage.tsv`: generation diagnostics.
- `run.log`, `launcher.log`, and other diagnostic files.
- `drawing-format.slddrt`: the selected template's embedded sheet format.

Temporary `SourceBackup/` files protect the original models during snapshot creation. The current snapshot's backups are removed after the snapshot is validated and source restoration succeeds. Failures retain those backups.

Linked external CAD files, including STEP dependencies, are included in the model snapshot. Pack and Go applies the destination folder and filename prefix to the complete dependency set; every returned destination is checked before saving.

## Resume and retry

The launcher automatically continues between batches and after drawing errors or an interrupted SolidWorks process. Completed drawings whose artifacts match the checkpoint are skipped. An interrupted drawing restarts from its beginning; missing artifacts are regenerated.

If the launcher itself was closed or Windows restarted, repeat the same command with the same `/out`, INI, and template path. Use a new output folder when changing paper format, template, or rendering settings.

Pending jobs run before deferred retries of failed jobs. Once a job reaches `MaxJobAttempts`, it remains in `failed-jobs.tsv` and other jobs continue. After fixing its cause, increase `MaxJobAttempts` and repeat the command without `/reset` to retry it while preserving valid completed drawings.

To regenerate the entire queue, add `/reset` or delete `generation-state.tsv`. Matching artifact filenames can then be overwritten; unrelated files are retained. Only one launcher can generate drawings at a time.

## Exit codes

| Code | Meaning |
| --- | --- |
| `0` | All selected drawings completed successfully |
| `1` | Launcher, COM connection, or process failure |
| `2` | Model validation or drawing-generation failure; some jobs may remain failed after all eligible work was attempted |
