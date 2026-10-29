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
python -B tests/test_dimension_layout.py
```

These check source/native consistency, output paths, repository hygiene, package contents, VBS entry points and arguments, build failures, checkpoint recovery, backup cleanup, and snapshot targets including linked STEP files. WSH tests require Windows. Run the launcher contract and controller recovery tests sequentially when no generator holds the shared lock.

The dimension-layout tests execute the VBA marker, spacing, row-placement, and local-datum graph routines under WSH. They check that overflow fails instead of losing dimensions. An optional external `dimension_layout/geometry.json` fixture, captured by `InspectDimensionLayout.ps1 -IncludeGeometry`, runs a capacity audit against native projected vertices and writes an external `geometry-capacity.json` report. A reported capacity limitation is not a successful drawing-layout verification.

### SolidWorks integration tests

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

`InspectDimensionLayout.ps1 -DrawingPath EXTERNAL_DRAWING -OutputPath EXTERNAL_JSON` reads native dimension text, arrowheads, extension lines, and view scales in its own SolidWorks process. Add `-IncludeGeometry` only when projected vertex data is needed. The strict same-sheet layout solver remains experimental; LeftSide on A4 is a known failing case. The default runtime keeps compatible generation and reports unresolved placement through `LAYOUT REVIEW` warnings. Package or unit-test success must not be presented as full GOST compliance.

`CompareDimensionArrangement.ps1 -DrawingPath EXTERNAL_DRAWING -OutputDirectory EXTERNAL_DIRECTORY -Cases placement|creation|postprocess|all` compares native SolidWorks dimension tools on a drawing referencing exactly one part. It copies the drawing and part into a unique external directory, replaces the copied reference, and uses its own background SolidWorks process. Placement cases compare sheet-wide Auto Arrange, bent radial/diameter leaders, and Auto Arrange per view; every existing dimension and value must remain unchanged. Creation cases replace dimensions in the copied drawing using native Autodimension with chain or baseline schemes, then Auto Arrange. Their dimension counts can differ and require a geometry coverage review. `all` runs placement and creation. The default `postprocess` case requires a previously created `autodimension_chain.slddrw` in the output directory and runs the canonical VBA leader postprocessor on an isolated copy. Build its external harness first with `python -B tests/build_leader_fixture.py`; after the CAD run, check `python -B tests/test_native_leaders.py`. The harness and check use the external `hub_leaders/comparison` fixture directory. The script exports native drawings, PDFs, and a JSON report. Failed layouts receive a `_failed` filename for diagnosis. Only Matochina4x100 is currently used to validate native postprocessing. Inspect crossings and manufacturing completeness before release; the runtime mode is opt-in through `NativeDimensions=1`.

The native postprocess check audits every removal against an unchanged retained measurement and the original model-coordinate interval. It checks disjoint regions for dimensioned views, graphical bounds, compact text, symmetry centerlines, and inclined leader geometry. Dimensionless isometric views remain geometry obstacles; their blank surrounding space can be used by the lower view. The CAD harness limits macro execution to two minutes and stops only its own process if a VBA dialog blocks the run. `-Cases inspect` exports native reference points and view geometry without running the postprocessor. Delete `input-geometry.json` after replacing the chain fixture so the next postprocess run captures fresh source geometry.
