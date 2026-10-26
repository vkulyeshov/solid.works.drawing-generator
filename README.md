# SolidWorks Drawing Generator

Generate editable SolidWorks drawings and PDFs from a saved assembly or part. The VBS launcher opens the model, runs the VBA macro, reports progress, and resumes interrupted jobs from a durable checkpoint.

`generator.bas` is the only editable generator source. `launcher.vbs` builds a temporary native VBA project from that file, then runs it using `generator.ini`. Neither a second text source nor a generated `generator.swp` is stored in Git.

This repository is the canonical source directory. Drawings, PDFs, logs, reports, backups, test model copies, build intermediates, and release packages belong outside it.

## Requirements

- Windows with SolidWorks and its VBA component installed. Native generation has been tested with SolidWorks 2025 SP5.
- Windows Script Host (`cscript.exe`).
- Source checkout: Python 3 with the Windows `py` launcher, or the `SW_DRAWING_PYTHON` environment variable pointing to `python.exe`. The build uses only Python's standard library.
- Prebuilt runtime kit: Python is not required; the kit contains the generated VBA project.
- The `GOST Type BU` font used by the supplied templates.
- Accessible assembly dependencies. Different source files with identical basenames cannot be flattened into one snapshot.

## Quick start

Run from this repository directory. In the current workspace, source models are in the sibling `../model` directory.

```bat
cscript //nologo launcher.vbs /input:"..\model\8MM4E.SLDASM" /out:"..\output_A4" /format:A4 /background
```

A3 with an explicit portrait template:

```bat
cscript //nologo launcher.vbs /input:"..\model\8MM4E.SLDASM" /out:"..\output_A3_portrait" /format:A3 /template:"templates\A3.portrait.drwdot" /background
```

For automatic portrait selection, use `/orientation:portrait`. The default is A3 landscape, using `templates/A3.landscape.drwdot`.

Positional arguments remain supported:

```bat
cscript //nologo launcher.vbs "C:\Models\Assembly.SLDASM" "C:\Settings\drawing.ini" "D:\Drawings\Assembly"
```

Use `/help` for the launcher reference. Quote paths containing spaces.

The source launcher builds `generator.bas` on every run into its unique Windows temporary session. It then starts SolidWorks with that generated project. This always uses the current BAS source and avoids stale executable copies.

If Python is installed without the Windows `py` launcher, set its executable path in PowerShell before starting:

```powershell
$env:SW_DRAWING_PYTHON = 'C:\Python313\python.exe'
```

The normal drawing-generation command remains the same. In a prebuilt runtime kit, the launcher uses the supplied `generator.swp` directly and does not invoke Python.

## Options and settings

| Launcher option | INI setting | Default |
| --- | --- | --- |
| `/input:FILE` | Required launcher argument | Saved `.SLDASM` or `.SLDPRT` |
| `/ini:FILE` | Settings file | `generator.ini` beside the launcher |
| `/out:DIRECTORY` | Output directory | Timestamped `DrawingPack_YYYYMMDD_HHMMSS` beside the source model |
| `/format:A4` or `/format:A3` | `PaperFormat` | `A3` |
| `/orientation:landscape` or `/orientation:portrait` | `PaperOrientation` | `landscape` |
| `/template:FILE.drwdot` | `DrawingTemplate` | Automatic template selection |
| `/templates:DIRECTORY` | `TemplatesDirectory` | `templates` beside the launcher |
| `/background` | Launcher flag | Separate SolidWorks process when present |
| `/reset` | Launcher flag | Remove the existing checkpoint when present |

Command-line settings override INI settings. An explicit template overrides automatic orientation selection: its sheet size and orientation are preserved. The selected A4/A3 format must match the template dimensions; a mismatch stops generation before creating a snapshot or job queue.

Command-line relative paths use the current directory. Relative paths in INI use the INI directory. INI supports UTF-8, UTF-8 with BOM, or UTF-16 with BOM; preserve the supplied INI's UTF-16 encoding when editing it.

```ini
PaperFormat=A3
PaperOrientation=landscape
TemplatesDirectory=templates
DrawingTemplate=
Font=GOST Type BU
AutoDimensions=1
ImportModelDimensions=0
UpAxis=+Y
JobsPerBatch=8
MaxJobAttempts=3
MaxJobs=0
JobFilter=
Organization=
ShowSummary=0
```

`MaxJobs=0` selects all jobs; a positive number limits a new queue. `JobFilter` selects jobs by a path/body-name substring. `UpAxis` accepts `+X`, `-X`, `+Y`, `-Y`, `+Z`, or `-Z`. `ImportModelDimensions` accepts `0`, `marked`, or `all`; imported sketch dimensions may require layout review. `AutoDimensions=0` disables additional chains, circles, radii, and the flat-dimension sheet; overall dimensions are still generated.

## Templates and drawing content

The four runtime templates are `A4.landscape.drwdot`, `A4.portrait.drwdot`, `A3.landscape.drwdot`, and `A3.portrait.drwdot` in `templates/`. They contain native sheet-format geometry and title-block notes based on the supplied `rules.pdf`, DSTU GOST 2.104:2006, form 1. A4 landscape uses the horizontal bottom-right title block requested for this project.

The macro creates a document from the selected template and retains its embedded format. Additional sheets use `drawing-format.slddrt`, extracted into the output directory, so relocated templates do not depend on an old absolute external format path. View placement and scale use the actual sheet dimensions.

Drawing properties `Designation`, `Description`, `Material`, and `Mass` populate the supplied title block. Existing localized model-property aliases remain supported. Ukrainian drawing labels remain unchanged; source code and documentation are English. Optional organization and approval fields can be supplied through INI. Native sheet properties provide scale, page number, and total page count.

The generator groups repeated file/configuration pairs, creates separate drawings for visible solid bodies, adds supported sheet-metal flat patterns, and adds native assembly component tables. A flat pattern has its own dimensioned sheet, with all flat dimensions kept together on that sheet. Long technical notes may continue on additional formatted sheets. Drawings remain automatic drafts requiring engineering review of manufacturing dimensions, tolerances, datums, sections, bends, and release requirements.

All independent dimension intervals remain native graphical dimensions on the dimensioned view. Dense contours do not switch to coordinate tables or lose dimensions through a numeric limit. Labels are packed into outside bands above, below, left, and right of the contour. Each band uses as many lanes as the available paper permits; a forward/backward packing pass keeps text inside its band without overlaps. Short dimensions use outside arrows. Native smart extension-line display limits witness-line clutter without changing measured values or model references. The flat view and all its dimensions stay on one detail sheet; the generator does not add enlarged views to spread this layout over more sheets. If the selected format cannot accommodate the complete layout, generation reports a layout error instead of silently omitting dimensions.

Repeated equal pitches in short chains use one `count x pitch = total` dimension. Circular features with the same diameter and orientation use one diameter callout with the count of distinct projected centres; front/back edges at the same projected centre are counted once. Overall dimensions are deduplicated by their physical model direction and span across orthographic views, so equal numeric dimensions in different directions remain independent. Folded and flat overall sizes remain separate. Feature dimensions of sheet-metal parts are concentrated on the flat detail sheet when available. Repeated assembly instances of the same file/configuration continue to share one drawing and a quantity; geometrically different bodies are not merged merely because a dimension matches. The automatic-review banner is no longer inserted.

This graphical layout changes rendering compatibility. Generate a new output pack, or explicitly use `/reset` to regenerate an existing pack; a checkpoint made with an earlier dimension layout cannot be resumed with these rendering settings. The former `MaxCoordinateDimensionsPerAxis` setting is no longer used.

## Output and recovery

Use an output directory outside this repository. Each pack includes:

- `Models/`: the Pack and Go snapshot used for generation.
- `SourceBackup/` and `source-backup.tsv`: temporary exact source-file backups made before Pack and Go. After all expected files exist in `Models/`, are nonempty, and source restoration has completed successfully, the current snapshot's backup and its index are removed before drawing generation. The empty backup root is also removed. Backups from other failed runs are retained. Snapshot or restoration failures retain the backup; cleanup failures are logged as warnings and do not stop drawing generation.
- `Drawings/` and `PDF/`: editable drawings and exported PDFs.
- `generation-state.tsv`: the durable job queue and checkpoint.
- `failed-jobs.tsv`: remaining failed jobs, their attempt counts, retry eligibility, and last errors.
- `manifest.tsv` and `dimension-coverage.tsv`: job and dimension diagnostics.
- `run.log`, `launcher.log`, `progress.tsv`, and worker/settings diagnostics.
- `generator.build.json` for source launches: checksums of the BAS and generated project.
- `drawing-format.slddrt`: the selected template's embedded sheet format.

The launcher automatically continues between batches, after individual drawing errors, and after an interrupted native process. Completed drawings are skipped when both artifacts match the checkpoint. An interrupted job restarts from its beginning. Missing artifacts are regenerated. Repeat the same command with the same `/out` only if the launcher itself was closed or Windows restarted.

Existing version 2 checkpoints migrate automatically to version 3 without regenerating valid completed artifacts. The new version stores each job's attempt count. Legacy completed, failed, and interrupted jobs start with one recorded attempt; pending jobs start with zero. Use a new output directory when changing format, template, or rendering settings. `/reset` removes only the checkpoint; matching output filenames can then be overwritten during regeneration. Old unrelated files are not removed.

`JobsPerBatch=8` bounds job attempts per batch; `0` disables the limit. `MaxFailures` optionally ends a batch after that many failures, then the launcher continues with another batch. Background mode restarts SolidWorks between batches. A batch with failed jobs still advances the queue, including a batch with no successful drawings.

`MaxJobAttempts=3` permits three total attempts per job across launcher restarts. Attempts are committed before rendering so a native crash also consumes an attempt. Pending jobs run before deferred failed-job retries; an interrupted job is retried immediately while its budget remains. After exhausting its budget, a job is skipped and recorded in `failed-jobs.tsv`; other jobs continue. To retry exhausted jobs after fixing their cause, increase `MaxJobAttempts` and run the same command without `/reset`. Valid completed drawings remain untouched. Startup crashes that occur before a job attempt are retried up to twice without progress; validation errors stop immediately.

Dirty source documents are rejected for a new queue. Resume uses the saved snapshot without saving or discarding a user's unsaved changes. Final exit code `2` indicates that some jobs still failed after all eligible work was attempted; it does not abandon later pending jobs.

`/background` runs desktop SolidWorks in a separate process with its graphics window offscreen. It still requires a licensed interactive Windows session and graphics context. The launcher closes only its own process. Without this flag, it can use an existing SolidWorks instance and leaves it open.

Exit codes: `0` complete success, `1` launcher/COM/process failure, `2` validation or drawing-generation failure. A shared temporary lock prevents concurrent launchers; its existing name `SolidWorks_AirDuctDrawings.lock` is retained for compatibility.

## Repository layout

```text
launcher.vbs          Builds and runs the generator, or runs a prebuilt kit
generator.bas         The only editable VBA generator source
generator.ini         Runtime defaults
templates/            Four native runtime templates
tools/                Native project builder and runtime packaging
tools/templates/      Native template builder and verifier
tests/                Test source code; no generated fixtures/results
```

Historical reports, examples, backups, and earlier one-off diagnostics remain outside Git in the sibling `../sw.standart.generator` directory.

## Development and tests

Edit `generator.bas` directly and save it as UTF-8. It contains the `main` entry point; there is no source synchronization step. Write identifiers, comments, log messages, and documentation in English. Ukrainian drawing labels and localized model-property aliases are readable Cyrillic string literals. The builder converts the UTF-8 text to Windows-1251 inside the native project and declares that code page in its VBA metadata, preserving the displayed text. No separate localization file or hexadecimal string helper is needed for the current labels. For manual use in the VBA editor, open the generated SWP rather than importing the UTF-8 BAS as an ANSI file.

Use multiline `If ... Then`, `Else`, and `End If` blocks in VBA and VBScript; put each branch statement on its own line. Run Python with `-B` to prevent `__pycache__` directories.

### Named constants

`generator.bas` and `launcher.vbs` have grouped constants sections immediately after `Option Explicit`. They centralize SolidWorks and Windows API codes, file/protocol values, field indices, geometric tolerances, scale factors, title-block dimensions, view/text/dimension layout, and polling/recovery limits. Constants are named by purpose; `_M` means metres, `_PX` means pixels, and timing names state seconds or milliseconds. Values that happen to be equal but describe different settings have separate names.

Supported INI settings continue to override their defaults. API identifiers, checkpoint/progress field indices, and wire-format tags describe existing external contracts; their values are preserved by this refactoring. `MAX_RETRIES_WITHOUT_PROGRESS` also controls the retry total displayed in the launcher log.

Build intermediates and test data default to `../tmp/solid.works.drawing-generator`. Set `SW_DRAWING_WORK_DIR` to another external directory if needed. Python build and test helpers reject working paths inside this repository.

```bat
python -B tools/rebuild_native.py
python -B tests/test_repository.py
python -B tests/test_launcher_contract.py
python -B tests/test_controller_recovery.py
python -B tests/test_snapshot_cleanup.py
python -B tests/test_paper_formats.py --case A4.landscape
```

The build writes `generator.swp` and its checksum report into the external working directory, normally `../tmp/solid.works.drawing-generator/build/`. Use `--source FILE.bas --output FILE.swp` for an explicit external destination. No existing SWP file is needed: `tools/vba_project.py` supplies only empty SolidWorks host metadata and library references, with a zero-byte executable module; all generator code comes from BAS. SolidWorks compiles the resulting VBA project on first load.

The repository tests check source/native consistency, source freshness, empty-project metadata, output-path guards, syntax, and repository hygiene. The 32 WSH contract checks validate arguments, encodings, templates, locking, and checkpoint preservation after a build failure without starting CAD. The snapshot cleanup test runs the generator's filesystem routine in disposable external folders, checking successful deletion, preservation of older backups, and rejection of paths outside the backup root. Its WSH harness adapts VBA's labelled exception handler; it does not start CAD. PDF integration checks require `pdfplumber`; template verification also uses `pypdf`, `pypdfium2`, and Pillow.

Native format cases are `A4.landscape`, `A4.portrait`, `A3.landscape`, `A3.portrait`, `assembly`, `resume`, and `mismatch`. Run `A4.landscape` before `resume`. Test models are copied from the sibling `../model` directory into the external working directory. Integration tests require installed SolidWorks and may take several minutes.

`test_compact_dimensions.py --before ORIGINAL_PACK --after UPDATED_PACK` checks a native LeftSide comparison: folded-view overall deduplication, preservation of the full graphical flat dimension set on one sheet, interval coverage for both axes, removal of coordinate-table sheets and the banner, text overlap, and page boundaries. Both pack paths must refer to external generated output. Optional `--pattern-pack FOUR_HOLE_PACK` also checks a native four-hole plate: one diameter quantity, one equal-pitch label, and reuse of overall dimensions without duplicate total spans.

After generating all five drawing cases, run `powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/InspectPaperResults.ps1` to inspect native sheet notes, scale, and numbering. The optional `test_incremental_integration.py` injects a crash only into its own copied A4 fixture and checks recovery and completed artifact preservation. `test_controller_recovery.py` replaces only the worker in an external launcher copy to test immediate validation failure, bounded startup retries, and retry continuation when persistent attempt counts advance. Run it when no other launcher holds the shared lock.

After the native crash test, `python -B tests/test_failed_job_recovery.py` extends that disposable A4 landscape queue with a missing body followed by a valid body. It verifies version 2 migration, continuation after a wholly failed batch, deferred bounded retries, the final failure report, and preservation of the previously completed drawing. Run these native tests in the same external `SW_DRAWING_WORK_DIR`; they never use production output.

Create a distributable kit outside Git:

```bat
python -B tools/package_launcher.py --output-dir "..\output\releases"
```

This builds the current BAS and writes `Ready/` and `SolidWorks_DrawingMacro.zip` into the external output directory. The runtime kit contains the launcher, generated `generator.swp`, INI, four templates, and README. It does not contain another BAS or TXT source and does not require Python on the receiving computer. `generator.build.json` beside the staged macro records its source and output checksums.

To test the prebuilt kit with a copied model, pass `--launcher "..\output\releases\Ready\launcher.vbs"` to `tests/test_paper_formats.py`. Source launches and runtime-kit launches use the same model and PDF checks. Template development outputs also stay outside Git; see the output parameters of `tools/templates/BuildTemplates.ps1` and `tools/templates/verify_templates.py`.
