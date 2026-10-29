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
| `/background` | Use a separate SolidWorks process with its window outside the screen; close that process when finished |
| `/reset` | Delete the output checkpoint and regenerate the queue |
| `/macro:FILE.swp` | Use another existing native VBA project; defaults to `generator.swp` beside the launcher |

Command-line options override INI settings. Command-line relative paths use the terminal's current directory. Relative paths inside INI use the INI directory. Preserve the supplied INI's UTF-16 encoding when editing it.

The supplied INI selects A4 portrait. If the paper settings are absent, the fallback is A3 landscape. An explicit template determines orientation; its sheet size must match the selected A4/A3 format. The four supplied templates are:

- `templates/A4.portrait.drwdot`
- `templates/A4.landscape.drwdot`
- `templates/A3.portrait.drwdot`
- `templates/A3.landscape.drwdot`

`/background` still requires a licensed interactive Windows session and a graphics context. Without this flag, the launcher can use an existing SolidWorks instance and leaves it open. Save pending changes before starting a new queue; the launcher does not save or discard unsaved source documents.

## Generation settings

Edit `generator.ini` to adjust generation:

| Setting | Purpose |
| --- | --- |
| `AutoDimensions=1` | Add graphical dimension chains and supported feature dimensions |
| `NativeDimensions=0` | Set to `1` to use native Autodimension and Auto Arrange, followed by inclined text leaders and in-frame placement, for ordinary solid parts; currently tested on Matochina4x100 only |
| `StrictDimensionLayout=0` | Keep generation compatible with dense models; `1` enables the experimental collision and same-sheet detail layout checks |
| `ImportModelDimensions=0` | Disable model-dimension import; alternatives are `marked` or `all` |
| `UpAxis=+Y` | Model up direction; also accepts `-Y`, `+X`, `-X`, `+Z`, or `-Z` |
| `JobsPerBatch=8` | Maximum job attempts per batch; `0` disables the limit |
| `MaxJobAttempts=3` | Persistent total attempt limit per drawing job |
| `MaxJobs=0` | Select all jobs; a positive value limits a new queue |
| `JobFilter=` | Optional path or body-name substring for selecting jobs |
| `Organization=` | Optional organization text for drawing notes |

Repeated instances of the same file/configuration share one drawing and a quantity. Supported sheet-metal parts include a flat-pattern sheet. All graphical flat dimensions remain together on that sheet. Drawing properties such as `Designation`, `Description`, `Material`, and `Mass` populate the title block. Review manufacturing dimensions, tolerances, datums, and drawing content before release.

Dimension styling uses the supplied GOST 2.307 rules: closed arrowheads with a 2.5 mm head length and a 20 degree angle, a 2 mm extension beyond the dimension line, 7 mm row spacing, and at least 10 mm outside offset from the view. Short chains can use separate clear dots where they fit. Insignificant trailing zeroes are removed without reducing the configured measurement precision.

Arrow geometry alone does not establish compliance of the complete layout. The default mode preserves graphical dimensions when the strict placement solver cannot find a collision-free arrangement. Such views produce `WARN LAYOUT REVIEW` in `run.log` and need review; no review label is added to the drawing itself. The experimental strict mode rejects that layout instead of silently dropping dimensions. It can rebuild interleaved coordinates as local datum graphs and try enlarged views on the same sheet, with checks that every source coordinate remains dimensioned. Dense LeftSide A4 layouts are still a known limitation of this mode. Do not treat its implementation as verified for production release.

With `NativeDimensions=1`, solid parts use native chain dimensions instead of the custom coordinate chains. After native Auto Arrange, the postprocessor removes repeated physical intervals and dimensions of the same circular/cylindrical geometry. Reflected intervals are grouped only when the selected body's vertices and sampled edge geometry verify the symmetry; equal numbers alone are insufficient. Each dimensioned view has a separate rectangular dimension region. Dimensionless isometric views remain geometry obstacles, and no dimension graphics may cross another view. The upper orthographic row can move together to reserve lower-view space, without changing scales. Readable native labels remain on their dimension lines; conflicting labels use inclined leaders and can use 2.5 mm text when the standard text does not fit. Multi-line hole callouts use 2.5 mm text. Verified mirror planes receive centerlines. Removed dimensions and their retained representatives are recorded in `native-dimension-audit.tsv`; `native-view-regions.tsv` and `native-symmetry-lines.tsv` record the layout boundaries and centerlines. Review manufacturing completeness before release. Assemblies and sheet-metal parts continue to use the existing dimensioning path. This mode is currently tested only on Matochina4x100.

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
