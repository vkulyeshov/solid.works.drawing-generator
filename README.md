# SolidWorks Drawing Generator

Generate editable SolidWorks drawings and PDFs from saved assemblies and parts.

For instructions on using the generated package, see the [runtime user guide](GENERATOR.md).

## Build

Requirements:

- Windows and Windows Script Host (`cscript.exe`).
- Python 3 with the Windows `py` launcher, or `SW_DRAWING_PYTHON` pointing to `python.exe`. The package builder uses only the Python standard library.

Run from the repository directory:

```bat
cscript //nologo build.vbs
```

The output directory defaults to `build/` inside this checkout, regardless of the terminal's current directory. The entire folder is excluded from Git by `.gitignore`. To choose another directory, pass it as the single positional argument:

```bat
cscript //nologo build.vbs "D:\Releases\DrawingGenerator"
```

If Python is installed without `py`, set its executable path in PowerShell:

```powershell
$env:SW_DRAWING_PYTHON = 'C:\Python313\python.exe'
```

The resulting directory contains:

```text
build/
  bin/
    launcher.vbs
    generator.swp
    generator.ini
    README.md
    templates/
  bin.zip
  generator.build.json
```

`bin.zip` contains the eight files from `bin/`. The root `GENERATOR.md` is copied into the package as `README.md`. The checksum report stays in the build directory, outside `bin/` and the archive. An explicit relative output directory is resolved from the terminal's current directory. Build outputs inside this repository must stay under `build/`; external output directories are also supported.

The editable VBA sources live in [`modules/`](modules/README.md). Save them as UTF-8. [`modules/project.json`](modules/project.json) lists every source module in build order, with declarations first. The builder validates the list and combines the modules in memory into the existing `Macro11` host module inside one `generator.swp`. This unity build preserves shared `Private` declarations and the `main` entry point. It does not generate a second editable BAS file. Windows-1251 encoding in the native project preserves localized drawing labels. The checksum report includes the composed-source checksum and the path and checksum of every input module and the manifest.

Generated SWP files and release packages stay outside Git. SolidWorks is not started during package creation; it compiles the VBA project when the project is first loaded.

The equivalent Python packaging command is:

```bat
python -B tools/package_launcher.py
```

For a native project without the full package, use `python -B tools/rebuild_native.py`. It reads `modules/` by default and writes `../tmp/solid.works.drawing-generator/build/generator.swp`. Use `--source DIRECTORY --output FILE.swp` for another module project. A standalone `--source FILE.bas` remains supported for external test harnesses.

## Tests

Run Python with `-B` to prevent bytecode caches in the checkout. Test fixtures, reports, copied models, and generated drawings use `../tmp/solid.works.drawing-generator`. Set `SW_DRAWING_WORK_DIR` to another external directory if needed. Test-data paths must remain external; release artifacts can use the ignored `build/` tree. Repository checks skip that tree while verifying it is ignored and contains no tracked files.

### Checks without SolidWorks

```bat
python -B tests/test_repository.py
python -B tests/test_modular_sources.py
python -B tests/test_build.py
python -B tests/test_launcher_contract.py
python -B tests/test_controller_recovery.py
python -B tests/test_snapshot_cleanup.py
python -B tests/test_snapshot_targets.py
python -B tests/test_inventory_suppression.py
python -B tests/test_inventory_saved.py
python -B tests/test_source_validation.py
python -B tests/test_dimension_layout.py
python -B tests/test_profile_pages.py
python -B tests/test_compact_profiles.py
python -B tests/test_panel_text_bounds.py
```

These check source/native consistency, output paths, repository hygiene, package contents, VBS entry points and arguments, build failures, checkpoint recovery, backup cleanup, snapshot targets including linked STEP files, and both suppression-filter modes with CLI/INI precedence, descendant inventory and active-instance quantities. WSH tests require Windows. Launcher contract tests use a disposable package with a private sentinel lock and never start CAD. Run controller recovery tests when no generator holds the shared lock.

The dimension-layout tests execute the VBA marker, spacing, row-placement, and local-datum graph routines under WSH. They check that overflow fails instead of losing dimensions. An optional external `dimension_layout/geometry.json` fixture, captured by `InspectDimensionLayout.ps1 -IncludeGeometry`, runs a capacity audit against native projected vertices and writes an external `geometry-capacity.json` report. A reported capacity limitation is not a successful drawing-layout verification.

### SolidWorks integration tests

To verify inventory against an existing saved snapshot without generating drawings, clone it outside the repository and run the native inspection entry point:

```bat
python -B tests/prepare_saved_inventory_test.py --snapshot-model C:\Drawings\Models\AD_YYYYMMDD_HHMMSS_Assembly.SLDASM --prefix AD_YYYYMMDD_HHMMSS_ --work-directory C:\external\inventory-test
powershell -NoProfile -ExecutionPolicy Bypass -File tests/RunSixViewGeneration.ps1 -WorkDirectory C:\external\inventory-test -ModelFileName AD_YYYYMMDD_HHMMSS_Assembly.SLDASM -InventoryOnly -RemapInventoryReferences
python -B tests/test_saved_inventory_report.py --directory C:\external\inventory-test --prefix AD_YYYYMMDD_HHMMSS_ --source-models C:\Drawings\Models
```

The fixture copies only the selected snapshot's prefixed files. `-RemapInventoryReferences` updates stored document paths only inside the fixture before opening it, preventing the clone from loading references from the original Models folder. It records file hashes after remapping; the report check verifies that inventory does not modify those bytes. Native inventory writes `inventory.tsv`, `progress.tsv` and `run.log` in that external directory. It runs in a dedicated PID-bound process with a fifteen-minute macro watchdog; drawing tests retain their eight-minute watchdog. Verify the expected components/configurations and quantities in the report; the test does not change component suppression.

`python -B tests/test_source_validation_native.py` runs two isolated native dirty-flag regressions: a source marked dirty after validation must complete, while a source already dirty at validation must fail before snapshot planning. Both cases verify unchanged input bytes. The CAD harness explicitly starts a new process and binds its PID-specific application object instead of reusing an existing user session.

Create the package first with `build.vbs`. Native tests use `build/bin/launcher.vbs` and require installed SolidWorks with VBA, the `GOST Type BU` font, and source models in the sibling `../model` directory. Test models are copied into the external working directory. PDF assertions require `pdfplumber`; template verification also uses `pypdf`, `pypdfium2`, and Pillow.

```bat
python -B tests/test_paper_formats.py --case A4.landscape
python -B tests/test_paper_formats.py --case resume
python -B tests/test_paper_formats.py --case A4.portrait
python -B tests/test_paper_formats.py --case A3.landscape
python -B tests/test_paper_formats.py --case A3.portrait
python -B tests/test_paper_formats.py --case assembly
python -B tests/test_paper_formats.py --case mismatch
```

The format tests verify generated sheets, PDF dimensions and frames, title blocks, graphical dimension coverage, and localized labels. Run `A4.landscape` before `resume`. To test a package in another location, add `--launcher "D:\Releases\DrawingGenerator\bin\launcher.vbs"` to a format test. Recreate the package after editing BAS so its source checksum matches the checkout.

After generating all five drawing cases, inspect native sheet notes, scale, and numbering:

```bat
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/InspectPaperResults.ps1
```

For native crash recovery and failed-job retries, use the same external working directory:

```bat
python -B tests/test_incremental_integration.py
python -B tests/test_failed_job_recovery.py
```

The crash test interrupts only its own background SolidWorks process. The failed-job test extends the disposable single-job A4 landscape queue to check version 2 migration, continuation after a failed batch, bounded retries, and preservation of completed drawings. These tests modify their fixture queue; use a fresh `A4.landscape` case for another independent run.

`test_compact_dimensions.py --before ORIGINAL_PACK --after UPDATED_PACK` compares external LeftSide drawing packs for graphical dimension coverage, deduplication, text overlap, page boundaries, and preservation of the flat dimensions on one sheet. Optional `--pattern-pack FOUR_HOLE_PACK` checks repeated-hole callouts and equal-pitch labels. All comparison paths must point outside the repository.

`InspectDimensionLayout.ps1 -DrawingPath EXTERNAL_DRAWING -OutputPath EXTERNAL_JSON` reads native dimension text, arrowheads, extension lines, and view scales in its own SolidWorks process. Add `-IncludeGeometry` only when projected vertex data is needed. LeftSide on A4 remains a known limitation of the strict same-sheet coordinate solver. The default rejects layouts that fail placement checks; the older coordinate compatibility mode can report `LAYOUT REVIEW` warnings. Package or unit-test success must not be presented as full GOST compliance.

To test the default geometry planner on a saved part, use a fresh external directory:

```powershell
python -B tests/test_planar_constraints.py
python -B tests/prepare_model_test.py --model C:\models\Panel.SLDPRT --work-directory C:\external\panel-test
powershell -NoProfile -ExecutionPolicy Bypass -File tests/RunSixViewGeneration.ps1 -WorkDirectory C:\external\panel-test -ModelFileName Panel.SLDPRT
python -B tests/test_profile_layout.py --directory C:\external\panel-test\output
python -B tests/test_profile_material.py --directory C:\external\panel-test\output
python -B tests/test_dimension_appearance.py --directory C:\external\panel-test\output --color gray
```

The historical `RunSixViewGeneration` harness also supports the automatic profile layout. It owns its SolidWorks process, closes its documents before reopening the saved drawing, and records persisted dimensions and graphical segments. The profile audit checks role ownership, unique physical intervals, source-coordinate connectivity, measured values, isolated view regions, paper bounds, and exported PDF glyphs against dimension lines. Use that audit for supported orthogonal planar profiles; use the six-view audit below for the solid-part fixture. Render and inspect the PDF after the checks.

For clipped panels, `RunSixViewGeneration.ps1 -ProfileGeometryOnly` captures the largest planar cap of the copied input in an external `profile-geometry.json`. `python -B tests/test_panel_planner.py --geometry EXTERNAL/profile-geometry.json` executes the actual VBA recognizer against independent contours and the captured boundary. It checks the separation of large cutouts and clipped corners from shallow boundary excursions. Panel grouping uses actual local sizes, not whole-body symmetry or equal rounded labels. Ordinary coordinate-profile audits do not describe this feature-based panel plan.

Use `python -B tests/test_panel_layout.py --directory EXTERNAL_OUTPUT` to audit a saved panel drawing. It checks every planned native measurement, separate ownership of overview and notch dimensions, model references on the printed contour, view regions, PDF markers, and dimension lines against printed text and material. The audit supports an independently scaled and rotated notch view. Render the resulting PDF for a separate readability review.

`ProfileSheets=auto` groups geometry before allocating pages. Rectangular, clipped and stepped panels separate frame dimensions from local boundary notches and rectangular openings. Supporting edges define feature families; equal local sizes share counted dimensions. The frame constraint graph omits derived lengths. Families rotate together and keep separate rows for widths and locating gaps. Whole families move to continuation sheets only when needed, with a complete reference outline on every sheet. The recognized two-recess panel retains its two-sheet A4 portrait layout. Dense inclined prismatic profiles use a general-dimension sheet and a local-feature sheet. `ProfileSheets=legacy` restores the earlier fixed two-view workflow.

The common native label pass also runs for ordinary solids, including curved geometry. It preserves source measurements and ownership, fits arrowheads inside when space permits, measures text using Windows font metrics, includes printed-ink clearance, and routes conflicting labels within their view regions. Own witnesses, other labels, model material, view notes, title blocks and other views are obstacles. A bounded failure delegates to the checked native candidate solver; it never removes a measurement just to fit a page. Assemblies and sheet-metal parts retain their dimension sources. Geometric recognition is deliberately conservative: unsupported geometry keeps its existing planner and strict validation.

Imported dimensions may omit added prefixes and suffixes from `DisplayData`. Printed-ink checks therefore compose missing text exactly once, apply the drawing font height and normalize GDI measurements to the font ascent used by SolidWorks. Width checks include occurrence counts and chamfer labels. Repeated source measurements use quantity prefixes such as `2x 14`. Circular leaders place the complete label outside the chosen side of the view; directed leaders retain their source planner's side. Linear retries examine both normal directions. Standard 2.5 mm font preparation remains active if the first search delegates to the fallback solver. Generated annotations and shared through-hole wording use Ukrainian; hole-callout parameter tokens remain linked to the native model.

Ordinary solid drawings are saved and reopened for a final label-only pass against persisted text data. That pass does not run Auto Arrange or prepare fonts again. With `StrictDimensionLayout=1`, unresolved conflicts reject the job before PDF export.

Native COM Boolean results are compared explicitly with `False`: some methods return a true value of `1`, on which VBA's bitwise `Not` does not produce false. The text-bounds regression includes this case for hole callouts.

`test_profile_pages.py` executes the VBA shelf allocator and polygon clipping under WSH without starting CAD. It checks full-page overflow, continuation reuse, oversized cells, local group splitting without lost ownership, axis-specific bands, text clearance inside cutouts, and preservation of an open U-shaped recess after clipping. `test_panel_planner.py --geometry EXTERNAL/profile-geometry.json --expect-rightside` also executes the compact VBA family planner and verifies all 64 notch measurements remain in nine families on two sheets. `test_panel_text_bounds.py` checks measured panel text independently of native dimension-line tails. Native drawing checks and visual review remain necessary.

Prepare an external single-part fixture and run the native test. The optional timeout allows a larger paginated drawing while preserving the shorter default for ordinary solid-part tests. Compare planned ownership against the reopened drawing, verify that references lie inside the actual crop rectangles, check distinct width/gap rows and fitting inside arrowheads, require exactly two A4 sheets, match every printed arrow to a labeled native dimension, and check native lines against exported PDF glyphs:

```powershell
python -B tests/prepare_model_test.py --model "..\model\RightSide.SLDPRT" --work-directory "..\tmp\solid.works.drawing-generator\profile_pages\run01"
powershell -NoProfile -ExecutionPolicy Bypass -File tests/RunSixViewGeneration.ps1 -WorkDirectory "..\tmp\solid.works.drawing-generator\profile_pages\run01" -ModelFileName RightSide.SLDPRT -MacroTimeoutSeconds 900
python -B tests/test_profile_pages_native.py --directory "..\tmp\solid.works.drawing-generator\profile_pages\run01\output" --expected-sheets 2
```

Use `-InspectOnly` on the native test command to refresh the saved-dimension report from a completed fixture without running generation again. The inspector keeps the input open while reopening the drawing, avoiding stale references from closing every document at once.

The panel audit verifies text, markers, native references, and material clearance. It does not yet reject every leader-to-leader intersection; dense notch groups still require a readability review.

`test_planar_constraints.py` executes the actual VBA solver through Windows Script Host. Its independent fixtures check complete boundary symmetry, asymmetric holes, local U features, a 0.05 mm mismatch that must remain visible, contour closure, and proper diagonal crossings. The planar feature planner records its selected physical length variables separately from the coordinate-profile audit. A successful source or solver test does not establish that the generated drawing is readable; the saved native drawing and rendered PDF must also be checked.

For an inclined planar feature drawing, run `python -B tests/test_planar_layout.py --directory EXTERNAL_OUTPUT` instead of the orthogonal profile audits. It verifies unique physical length ownership, reconstruction of the omitted closure variables, persisted measurements, printed arrow ownership, paper regions and native lines against PDF text. `tests/analyze_planar_layout.py AUDIT --view VIEW_NAME` independently inspects the optional native candidate audit. The same generation fixture can be rebuilt with `prepare_model_test.py --reuse`; `--diagnostic-failures` exports failed drafts only inside that test fixture. Production generation does not export rejected layouts.

Use `python -B tests/test_native_readability.py --directory EXTERNAL_OUTPUT --expected-sheets 2` for a planner-independent PDF check. It matches printed native markers and checks lines, angular arcs and arrowhead edges against rotated PDF glyphs on each sheet. It supplements geometry-specific coverage checks and visual review. `prepare_model_test.py --setting KEY=VALUE` accepts repeatable fixture-only settings, such as `--setting OrthographicViews=6`; protected input and output paths remain isolated.

The material audit checks reopened dimension segments against the filled projected silhouette and verifies the audited vertices against black model contours in the PDF. Only declared extension lines at verified model coordinates may cross material when a shared depth is dimensioned outside the profile; its baseline and label must remain on that side. The layout audit matches every filled PDF arrow or dot to a saved native dimension with nonempty text, detecting orphaned or missing markers. Add `--before C:\external\previous-panel-test\output` to reproduce material crossings in an older drawing of the same part at the same scale, allowing for moved views. Geometry unit tests include open notches, boundary-aligned witnesses, diagonal leaders, witness overruns, and narrow material strips that a sampling-only check can miss.

The appearance audit checks saved native annotation colors and the actual PDF strokes: 0.6 mm black contours and 0.25 mm dimensions in the selected color. Set `DimensionColor=black` in the external fixture INI and use `--color black` to check monochrome output. `ProbeDimensionAppearance.ps1 -DrawingPath DRAWING -OutputDirectory EXTERNAL_DIRECTORY` compares native thickness settings using read-only test documents and exports diagnostic PDF variants without saving the drawing.

`CompareDimensionArrangement.ps1 -DrawingPath EXTERNAL_DRAWING -OutputDirectory EXTERNAL_DIRECTORY -Cases placement|creation|postprocess|all` compares native SolidWorks dimension tools on a drawing referencing exactly one part. It copies the drawing and part into a unique external directory, replaces the copied reference, and uses its own background SolidWorks process. Placement cases compare sheet-wide Auto Arrange, bent radial/diameter leaders, and Auto Arrange per view; every existing dimension and value must remain unchanged. Creation cases replace dimensions in the copied drawing using native Autodimension with chain or baseline schemes, then Auto Arrange. Their dimension counts can differ and require a geometry coverage review. `all` runs placement and creation. The default `postprocess` case requires a previously created `autodimension_chain.slddrw` in the output directory and runs the canonical VBA leader postprocessor on an isolated copy. Build its external harness first with `python -B tests/build_leader_fixture.py`; after the CAD run, check `python -B tests/test_native_leaders.py`. The harness and check use the external `hub_leaders/comparison` fixture directory. The script exports native drawings, PDFs, and a JSON report. Failed layouts receive a `_failed` filename for diagnosis. Only Matochina4x100 is currently used to validate native postprocessing. Inspect crossings and manufacturing completeness before release; the runtime mode is opt-in through `NativeDimensions=1`.

For the six-side Matochina4x100 test, choose a new external directory and run:

```powershell
python -B tests/prepare_six_view_test.py --work-directory C:\external\six-view-test
powershell -NoProfile -ExecutionPolicy Bypass -File tests/RunSixViewGeneration.ps1 -WorkDirectory C:\external\six-view-test
python -B tests/test_six_views.py --directory C:\external\six-view-test\output
```

Preparation copies the test part, builds the canonical modules, and configures six views. The CAD harness owns its SolidWorks process and has an eight-minute macro watchdog. Artifact checks require two A4 pages, all six orthographic directions, opposite isometrics, a common scale, and exactly one retained owner for every measurement group across both sheets. Every removed annotation must resolve to a retained representative with the same normalized measurement. Render and inspect both exported pages as well.

The fixture uses the `features` dimension strategy. The harness reopens the saved native drawing and writes `saved-dimensions.tsv` outside the repository. Checks compare the persisted dimensions to the ownership audit, check their line coordinates against the view regions, and require one shared eight-hole callout with its native thread specification, one 100 mm pitch-circle dimension, 28 x 68 mm row spacings, and a 7 mm boundary offset. The repeated 35.36 mm coordinate chains must be absent. `InspectFeatureDimensions.ps1 -ModelPath MODEL -OutputDirectory EXTERNAL_DIRECTORY` reads source feature parameters and analytic geometry using an isolated copy and process for diagnosis.

The saved-drawing checks also require one pitch-circle arrow whose tip lies on that circle, a linear 98 mm fork span on a different view, front-view diameters of 38.8 and 58.8 mm, and the 3 mm boss height on the first sheet. The derived 57 mm chamfer-rim diameter and the 13 mm construction half-width must be absent. Persisted diameter values and symbols must remain correct after reopening. `ProbePitchCircle.ps1 -DrawingPath DRAWING -OutputDirectory EXTERNAL_DIRECTORY` isolates the drawing and its part to diagnose view-sketch coordinates and dimension ownership.

The native postprocess check audits every removal against an unchanged retained measurement and the original model-coordinate interval. It checks disjoint regions for dimensioned views, graphical bounds, compact text, symmetry centerlines, and inclined leader geometry. Dimensionless isometric views remain geometry obstacles; their blank surrounding space can be used by the lower view. The CAD harness limits macro execution to two minutes and stops only its own process if a VBA dialog blocks the run. `-Cases inspect` exports native reference points and view geometry without running the postprocessor. Delete `input-geometry.json` after replacing the chain fixture so the next postprocess run captures fresh source geometry.
