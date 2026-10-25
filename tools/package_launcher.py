"""Package runtime sources and templates outside the source repository."""
import argparse
import shutil
import sys
import zipfile
sys.dont_write_bytecode = True
from project_paths import REPOSITORY, WORKSPACE, external_path
from rebuild_native import rebuild

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--output-dir', default=str(WORKSPACE/'output/releases'))
args = parser.parse_args()
base = REPOSITORY
output = external_path(args.output_dir)
output.mkdir(parents=True, exist_ok=True)
ready = output/'Ready'
ready.mkdir(exist_ok=True)
native = rebuild(base/'generator.bas', ready/'generator.swp')
# A runtime kit deliberately has no BAS file, so the launcher needs no Python.
for obsolete in ['generator.bas','generator.utf8.txt','RunDrawings.vbs',
                 'AirDuctDrawingGenerator.bas','AirDuctDrawingGenerator.utf8.txt',
                 'AirDuctDrawingGenerator.swp','AirDuctDrawingGenerator.ini']:
    path = ready/obsolete
    if path.exists():
        path.unlink()
files = ['launcher.vbs', 'generator.swp', 'generator.ini', 'README.md']
files += ['templates/'+name for name in ['A3.landscape.drwdot','A3.portrait.drwdot','A4.landscape.drwdot','A4.portrait.drwdot']]
for name in files:
    (ready/name).parent.mkdir(parents=True, exist_ok=True)
    if name != 'generator.swp':
        shutil.copy2(base/name, ready/name)
archive = output/'SolidWorks_DrawingMacro.zip'
with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
    for name in files:
        z.write(ready/name, name)
with zipfile.ZipFile(archive) as z:
    assert set(z.namelist()) == set(files)
    for name in files:
        assert z.read(name) == (ready/name).read_bytes()
print(f'Updated kit: {archive} ({archive.stat().st_size} bytes, {len(files)} files)')
