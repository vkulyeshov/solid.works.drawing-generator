"""Build an isolated single-model generation fixture using the normal defaults."""
from pathlib import Path
import argparse
import shutil
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import REPOSITORY, external_path
from rebuild_native import rebuild
from vba_sources import read_source


def prepare(model, directory, diagnostic_failures=False, reuse=False):
    directory = external_path(directory)
    if directory.exists() and not reuse:
        raise ValueError(f'Use a fresh test directory: {directory}')
    model = model.resolve(strict=True)
    target = directory / 'input' / model.name
    if reuse and not (target.is_file() and (directory / 'generator.ini').is_file()):
        raise ValueError('Only an existing single-model test fixture can be reused')
    target.parent.mkdir(parents=True, exist_ok=reuse)
    shutil.copy2(model, target)
    settings = {}
    for line in (REPOSITORY / 'generator.ini').read_text(encoding='utf-16').splitlines():
        if '=' in line and not line.lstrip().startswith(';'):
            key, value = line.split('=', 1)
            settings[key.strip()] = value.strip()
    settings.update(MaxJobs='1', TemplatesDirectory=str(REPOSITORY / 'templates'),
                    LauncherInput=str(target), LauncherOutput=str(directory / 'output'),
                    LauncherProgress=str(directory / 'progress.tsv'))
    (directory / 'generator.ini').write_text(''.join(f'{k}={v}\n' for k, v in settings.items()), encoding='utf-16')
    macro_source = REPOSITORY / 'modules'
    if diagnostic_failures:
        settings['DimensionLayoutAudit'] = '1'
        settings['MaxJobAttempts'] = '20'
        (directory / 'generator.ini').write_text(''.join(f'{k}={v}\n' for k, v in settings.items()), encoding='utf-16')
        source = read_source()
        marker = '    If Not d Is Nothing Then\n        dimensionKeys.RemoveAll'
        assert source.count(marker) == 1, 'The isolated failure diagnostic needs an unambiguous insertion point'
        diagnostic = '''    If Not d Is Nothing Then
        d.GraphicsRedraw2
        d.Save3 SW_SAVE_SILENT, er, wr
        Set pdf = sw.GetExportFileData(SW_EXPORT_PDF)
        pdf.ViewPdfAfterSaving = False
        pdf.SetSheets SW_EXPORT_ALL_SHEETS, Empty
        d.Extension.SaveAs fs.BuildPath(runDir, "diagnostic-failed.pdf"), SW_SAVE_CURRENT_VERSION, SW_SAVE_SILENT, pdf, er, wr
        dimensionKeys.RemoveAll'''
        macro_source = directory / 'diagnostic-generator.bas'
        macro_source.write_text(source.replace(marker, diagnostic), encoding='utf-8')
    rebuild(macro_source, directory / 'generator.swp')
    print(f'Test directory: {directory}; model: {model.name}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--model', type=Path, required=True)
    parser.add_argument('--work-directory', type=Path, required=True)
    parser.add_argument('--diagnostic-failures', action='store_true', help='Export failed layouts only in the isolated test fixture')
    parser.add_argument('--reuse', action='store_true', help='Rebuild an existing single-model fixture while preserving its checkpoint')
    args = parser.parse_args()
    prepare(args.model, args.work_directory, args.diagnostic_failures, args.reuse)
