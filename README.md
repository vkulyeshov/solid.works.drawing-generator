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

`generator.bas` is the only editable VBA source. Save it as UTF-8; the builder embeds it in the native project using Windows-1251 to preserve localized drawing labels. Generated SWP files and release packages stay outside Git. SolidWorks is not started during package creation; it compiles the VBA project when the project is first loaded.

The equivalent Python packaging command is:

```bat
python -B tools/package_launcher.py
```

For a native project without the full package, use `python -B tools/rebuild_native.py`. Its default output is `../tmp/solid.works.drawing-generator/build/generator.swp`; `--source FILE.bas --output FILE.swp` selects explicit paths.

## Tests

Run Python with `-B` to prevent bytecode caches in the checkout. Test fixtures, reports, copied models, and generated drawings use `../tmp/solid.works.drawing-generator`. Set `SW_DRAWING_WORK_DIR` to another external directory if needed. Test-data paths must remain external; release artifacts can use the ignored `build/` tree. Repository checks skip that tree while verifying it is ignored and contains no tracked files.

### Checks without SolidWorks

```bat
python -B tests/test_repository.py
python -B tests/test_build.py
python -B tests/test_launcher_contract.py
python -B tests/test_controller_recovery.py
python -B tests/test_snapshot_cleanup.py
```

These check source/native consistency, output paths, repository hygiene, package contents, VBS entry points and arguments, build failures, checkpoint recovery, and backup cleanup. WSH tests require Windows. Run the launcher contract and controller recovery tests sequentially when no generator holds the shared lock.

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
