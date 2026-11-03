"""Prepare an isolated Matochina4x100 six-view generation test."""
from datetime import datetime
from pathlib import Path
import argparse
import shutil
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import REPOSITORY, external_path, work_path
from rebuild_native import rebuild


def prepare(model, directory):
    directory = external_path(directory)
    if directory.exists():
        raise ValueError(f'Use a new test directory: {directory}')
    model = model.resolve(strict=True)
    if model.name.lower() != 'matochina4x100.sldprt':
        raise ValueError('This CAD fixture currently validates Matochina4x100 only')
    input_directory = directory / 'input'
    input_directory.mkdir(parents=True)
    shutil.copy2(model, input_directory / 'Matochina4x100.SLDPRT')
    settings = {}
    for line in (REPOSITORY / 'generator.ini').read_text('utf-16').splitlines():
        if '=' in line and not line.lstrip().startswith(';'):
            key, value = line.split('=', 1)
            settings[key.strip()] = value.strip()
    settings.update(OrthographicViews='6', NativeDimensions='1', MaxJobs='1',
                    TemplatesDirectory=str(REPOSITORY / 'templates'),
                    LauncherInput=str(input_directory / 'Matochina4x100.SLDPRT'),
                    LauncherOutput=str(directory / 'output'),
                    LauncherProgress=str(directory / 'progress.tsv'))
    (directory / 'generator.ini').write_text(
        ''.join(f'{key}={value}\n' for key, value in settings.items()), 'utf-16')
    rebuild(REPOSITORY / 'modules', directory / 'generator.swp')
    print(f'Test directory: {directory}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--model', type=Path,
                        default=REPOSITORY.parent / 'model/Matochina4x100.SLDPRT')
    parser.add_argument('--work-directory', type=Path,
                        default=work_path('six_views', datetime.now().strftime('%Y%m%d_%H%M%S')))
    args = parser.parse_args()
    prepare(args.model, args.work_directory)
