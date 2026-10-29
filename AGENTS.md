# Repository instructions

This directory is the canonical working copy of the SolidWorks drawing generator.

- Write source code, comments, identifiers, log messages, and documentation in English.
- Preserve Ukrainian drawing annotations and existing localized model-property aliases as readable Cyrillic string literals in the UTF-8 BAS source. The builder encodes the native module in Windows-1251 to match the VBA project. Do not replace readable labels with hexadecimal helpers.
- Use multiline `If ... Then`, `Else`, and `End If` blocks in VBA and VBScript. Keep each branch statement on its own line.
- Keep API identifiers, protocol indices, filenames, geometry tolerances, layout dimensions, scale factors, and polling/recovery limits in the grouped constants sections after `Option Explicit`. Name constants by purpose and include units for physical quantities. Preserve external API and protocol values during refactoring.
- Keep only reusable source files, configuration, native drawing templates, documentation, and test source code in this repository. Generated VBA macro projects belong outside Git.
- Keep release artifacts in the ignored top-level `build/` directory or an external output directory. Place test fixtures, temporary files, generated drawings, PDFs, logs, and backups outside this repository.
- Use `tools/project_paths.py` for Python output paths. `build_path` permits the ignored repository `build/` tree or external paths; test fixtures use the strictly external `work_path`. The default test working directory is `../tmp/solid.works.drawing-generator`; `SW_DRAWING_WORK_DIR` may override it with an external path.
- Run Python with `-B` to prevent bytecode caches. Do not put test fixtures containing native models or test results in this repository.
- Edit the canonical VBA sources in `modules/`; `modules/project.json` defines their build order. Keep declarations before procedures. Use `tools/vba_sources.py` to read the composed source in builders and tests. Do not restore a tracked monolithic `generator.bas` or another source copy. The builder combines the source modules into one native host module, preserving Private visibility and Macro11.main. Generated projects inside this repository belong only in the ignored `build/` tree. Use `build.vbs [DIRECTORY]` for package creation and `launcher.vbs` only for running a prebuilt project. Keep build and test instructions in the root README; keep runtime usage in the root `GENERATOR.md`, linked from the README and packaged as `bin/README.md`. The low-level `python -B tools/rebuild_native.py` builder remains available.
- Keep executable module-level code in each VBS script's `Main` routine, with a single `Main` call at the end; constants and shared declarations remain at module scope.
- The small `tools/vba_project.py` template contains only empty SolidWorks host metadata and no drawing-generator code. Do not embed the generator source there.
- Do not save or discard a user's unsaved SolidWorks document. CAD tests must use external copies of model files and operate only on their own background process when injecting a crash.
- Keep the existing shared launcher lock name: `SolidWorks_AirDuctDrawings.lock`.
