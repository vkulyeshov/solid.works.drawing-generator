# Repository instructions

This directory is the canonical working copy of the SolidWorks drawing generator.

- Write source code, comments, identifiers, log messages, and documentation in English.
- Preserve Ukrainian drawing annotations and existing localized model-property aliases as readable Cyrillic string literals in the UTF-8 BAS source. The builder encodes the native module in Windows-1251 to match the VBA project. Do not replace readable labels with hexadecimal helpers.
- Use multiline `If ... Then`, `Else`, and `End If` blocks in VBA and VBScript. Keep each branch statement on its own line.
- Keep API identifiers, protocol indices, filenames, geometry tolerances, layout dimensions, scale factors, and polling/recovery limits in the grouped constants sections after `Option Explicit`. Name constants by purpose and include units for physical quantities. Preserve external API and protocol values during refactoring.
- Keep only reusable source files, configuration, native drawing templates, documentation, and test source code in this repository. Generated VBA macro projects belong outside Git.
- Place temporary files, generated drawings, PDFs, logs, reports, build intermediates, release ZIP files, and backups outside this repository.
- Use `tools/project_paths.py` for Python output paths. The default working directory is `../tmp/solid.works.drawing-generator`; `SW_DRAWING_WORK_DIR` may override it with an external path.
- Run Python with `-B` to prevent bytecode caches. Do not put test fixtures containing native models or test results in this repository.
- Edit only `generator.bas` as the generator source. Do not create another source copy or put `generator.swp` in this repository. Build with `python -B tools/rebuild_native.py`; the VBS launcher builds automatically in its external session.
- The small `tools/vba_project.py` template contains only empty SolidWorks host metadata and no drawing-generator code. Do not embed the generator source there.
- Do not save or discard a user's unsaved SolidWorks document. CAD tests must use external copies of model files and operate only on their own background process when injecting a crash.
- Keep the existing shared launcher lock name: `SolidWorks_AirDuctDrawings.lock`.
