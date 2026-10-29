"""Package the runtime in the ignored build directory or an external directory."""
import argparse
import shutil
import sys
import zipfile
sys.dont_write_bytecode = True
from project_paths import REPOSITORY, BUILD_DIRECTORY, build_path
from rebuild_native import rebuild
from vba_sources import DEFAULT_SOURCE

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('output_dir', nargs='?', default=str(BUILD_DIRECTORY), help='Output directory; default: repository build/')
args = parser.parse_args()
base = REPOSITORY
output = build_path(args.output_dir)
output.mkdir(parents=True, exist_ok=True)
bin_dir = output/'bin'
legacy_dir = output/'Ready'
if legacy_dir.is_dir() and not bin_dir.exists():
    legacy_dir.rename(bin_dir)
bin_dir.mkdir(exist_ok=True)
native = rebuild(DEFAULT_SOURCE, bin_dir/'generator.swp')
native.with_suffix('.build.json').replace(output/'generator.build.json')
# A runtime kit deliberately has no BAS file, so the launcher needs no Python.
for obsolete in ['generator.bas','generator.utf8.txt','RunDrawings.vbs',
                 'AirDuctDrawingGenerator.bas','AirDuctDrawingGenerator.utf8.txt',
                 'AirDuctDrawingGenerator.swp','AirDuctDrawingGenerator.ini']:
    path = bin_dir/obsolete
    if path.exists():
        path.unlink()
files = ['launcher.vbs', 'generator.swp', 'generator.ini', 'README.md']
files += ['templates/'+name for name in ['A3.landscape.drwdot','A3.portrait.drwdot','A4.landscape.drwdot','A4.portrait.drwdot']]
for name in files:
    (bin_dir/name).parent.mkdir(parents=True, exist_ok=True)
    if name != 'generator.swp':
        source = base/'GENERATOR.md' if name == 'README.md' else base/name
        shutil.copy2(source, bin_dir/name)
archive = output/'bin.zip'
with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
    for name in files:
        z.write(bin_dir/name, name)
with zipfile.ZipFile(archive) as z:
    assert set(z.namelist()) == set(files)
    for name in files:
        assert z.read(name) == (bin_dir/name).read_bytes()
legacy_archive = output/'SolidWorks_DrawingMacro.zip'
if legacy_archive.is_file():
    legacy_archive.unlink()
print(f'Updated kit: {archive} ({archive.stat().st_size} bytes, {len(files)} files)')
