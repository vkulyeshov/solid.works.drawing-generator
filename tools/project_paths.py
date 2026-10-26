"""Locations for ignored release artifacts and external test data."""
from pathlib import Path
import os
import sys

sys.dont_write_bytecode = True
REPOSITORY = Path(__file__).resolve().parents[1]
WORKSPACE = REPOSITORY.parent
BUILD_DIRECTORY = REPOSITORY / 'build'


def build_path(value=BUILD_DIRECTORY):
    """Allow release artifacts in build/ or outside the source checkout."""
    path = Path(value).expanduser().resolve()
    if path == REPOSITORY or (REPOSITORY in path.parents
                             and path != BUILD_DIRECTORY and BUILD_DIRECTORY not in path.parents):
        raise ValueError(f'Build output inside the repository must stay under build/: {path}')
    return path


def external_path(value):
    """Resolve an output path and reject the repository and its descendants."""
    path = Path(value).expanduser().resolve()
    if path == REPOSITORY or REPOSITORY in path.parents:
        raise ValueError(f'Generated files must stay outside the repository: {path}')
    return path


def work_path(*parts):
    """Use a sibling working directory, optionally overridden by the environment."""
    base = os.environ.get('SW_DRAWING_WORK_DIR', str(WORKSPACE/'tmp/solid.works.drawing-generator'))
    return external_path(Path(base).joinpath(*parts))
