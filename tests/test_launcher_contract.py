"""Exercise the real WSH argument/encoding/locking contract without starting CAD."""
from pathlib import Path
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1]/"tools"))
from project_paths import work_path
import atexit
import json
import os
import subprocess
import tempfile
import shutil

base = Path(__file__).resolve().parents[1]
host = Path(os.environ['WINDIR']) / 'System32/cscript.exe'
folder = work_path('launcher_contract')
folder.mkdir(parents=True, exist_ok=True)
# Exercise a disposable launcher with a private lock, so a user's running
# generator is never blocked, stopped, or mistaken for a stale test sentinel.
package = tempfile.TemporaryDirectory(prefix='package_', dir=folder)
atexit.register(package.cleanup)
package_path = Path(package.name)
lock_name = 'GeneratorContract_' + package_path.name + '.lock'
script = package_path / 'launcher.vbs'
launcher_source = (base / 'launcher.vbs').read_text('utf8')
lock_declaration = 'Const SHARED_LOCK_FILENAME = "SolidWorks_AirDuctDrawings.lock"'
assert launcher_source.count(lock_declaration) == 1
script.write_text(launcher_source.replace(lock_declaration,
                  'Const SHARED_LOCK_FILENAME = "' + lock_name + '"'), encoding='utf8')
shutil.copy2(base / 'generator.ini', package_path / 'generator.ini')
shutil.copytree(base / 'templates', package_path / 'templates')
model = base.parent / 'model/RoofFrontLeftFrontArch.SLDPRT'
macro = folder / 'fixture.swp'
macro.write_bytes(b'Unused macro fixture; the lock prevents CAD startup')
results = []


def case(name, args, expected_code, expected_text, environment=None):
    if not any(str(arg).lower().startswith('/macro:') for arg in args):
        args = [*args, '/macro:' + str(macro)]
    run = subprocess.run([str(host), '//nologo', str(script), *map(str, args)],
                         capture_output=True, timeout=15, env=environment)
    output = run.stdout + run.stderr
    assert run.returncode == expected_code, (name, run.returncode, output)
    assert expected_text.encode('ascii') in output, (name, output)
    results.append({'test': name, 'exit_code': run.returncode, 'passed': True})


case('help', ['/help'], 0, 'Default output: DrawingPack_')
case('missing input', [], 1, 'Input model is required')
case('missing model', ['/input:' + str(folder/'absent.sldasm')], 1, 'Input file not found')
case('unsupported model', ['/input:' + str(base/'README.md')], 1, '.SLDASM or .SLDPRT')
case('missing INI', [model, folder/'absent.ini'], 1, 'INI file not found')
case('unknown option', ['/unknown:value'], 1, 'Unknown option')
case('background flag value', ['/background:false'], 1, 'flag and takes no value')
case('reset flag value', ['/reset:false'], 1, 'flag and takes no value')
case('duplicate argument', [model, '/input:' + str(model)], 1, 'both by name and by position')
(folder/'not-empty').mkdir(exist_ok=True)
(folder/'not-empty/keep.txt').write_text('existing user data', encoding='utf8')

# An exclusive sentinel guarantees these valid INIs stop before touching COM.
lock = Path(tempfile.gettempdir())/lock_name
try:
    with lock.open('x', encoding='utf-16') as stream:
        stream.write('launcher-contract-test')
except FileExistsError:
    raise RuntimeError('Another launcher holds the lock; run contract tests after it finishes.')
try:
    for encoding in ['utf-8', 'utf-8-sig', 'utf-16', 'utf-16-be']:
        ini = folder / (encoding + '.ini')
        data = 'Font=GOST Type BU\r\nOrganization=\u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430\r\n'.encode(encoding)
        if encoding == 'utf-16-be':
            data = b'\xfe\xff' + data
        ini.write_bytes(data)
        case('INI ' + encoding + ' and busy lock', [model, ini], 1, 'Another launcher is running')
        assert lock.read_text('utf-16') == 'launcher-contract-test'
    case('background accepts parameters before lock', [model, ini, '/background'],
         1, 'Another launcher is running')
    for value in ['0', '1']:
        case('suppression filter ' + value, [model, ini, '/excludesuppressed:' + value],
             1, 'Another launcher is running')
    for value in ['', 'true', 'false', '2']:
        case('invalid suppression filter ' + repr(value), [model, ini, '/excludesuppressed:' + value],
             1, 'ExcludeSuppressed must be 0 or 1')
    filter_ini = folder/'suppression.ini'
    filter_ini.write_text('ExcludeSuppressed=invalid\r\n', encoding='utf-16')
    case('invalid INI suppression filter', [model, filter_ini], 1, 'ExcludeSuppressed must be 0 or 1')
    case('CLI overrides invalid INI suppression filter', [model, filter_ini, '/excludesuppressed:0'],
         1, 'Another launcher is running')
    case('runtime does not require Python', [model, ini], 1, 'Another launcher is running',
         dict(os.environ, SW_DRAWING_PYTHON=str(folder/'missing-python.exe')))
    case('existing output accepted before lock', [model, ini, folder/'not-empty'],
         1, 'Another launcher is running')
    state = folder/'not-empty/generation-state.tsv'
    state.write_text('checkpoint sentinel', encoding='utf-16')
    case('reset cannot delete checkpoint while busy', [model, ini, folder/'not-empty', '/reset'],
         1, 'Another launcher is running')
    assert state.read_text('utf-16') == 'checkpoint sentinel'
    assert (folder/'not-empty/keep.txt').read_text('utf8') == 'existing user data'
    case('unsupported paper', [model, '/format:A2'], 1, 'Format must be A4 or A3')
    case('empty paper', [model, '/format:'], 1, 'Format must be A4 or A3')
    case('unsupported orientation', [model, '/orientation:diagonal'], 1, 'Orientation must be landscape or portrait')
    case('empty template', [model, '/template:'], 1, 'Template file is required')
    case('wrong template extension', [model, '/template:'+str(base/'README.md')], 1, '.drwdot file')
    case('missing template', [model, '/template:'+str(folder/'absent.drwdot')], 1, 'Drawing template not found')
    for paper in ['A3','A4']:
        for orientation in ['landscape','portrait']:
            case(paper+' '+orientation+' template selection', [model,'/format:'+paper,'/orientation:'+orientation], 1, 'Another launcher is running')
    case('case insensitive paper', [model,'/format:a4'],1,'Another launcher is running')
    case('explicit portrait template overrides auto-selection', [model,'/format:A4','/template:'+str(base/'templates/A4.portrait.drwdot')],1,'Another launcher is running')
    custom = folder/'\u0428\u0430\u0431\u043b\u043e\u043d\u044b \u0441 \u043f\u0440\u043e\u0431\u0435\u043b\u0430\u043c\u0438'
    custom.mkdir(exist_ok=True)
    shutil.copy2(base/'templates/A4.landscape.drwdot',custom/'A4.landscape.drwdot')
    case('custom template directory', [model,'/format:A4','/orientation:landscape','/templates:'+str(custom)],1,'Another launcher is running')
    ini2 = folder/'template-settings.ini'
    ini2.write_bytes(('PaperFormat=A4\r\nPaperOrientation=portrait\r\nDrawingTemplate='+str(base/'templates/A4.portrait.drwdot')+'\r\n').encode('utf-16'))
    case('INI paper and explicit template', [model,ini2],1,'Another launcher is running')
    ini2.write_bytes(('PaperFormat=A4\r\nTemplatesDirectory='+custom.name+'\r\n').encode('utf-16'))
    case('INI relative template directory', [model,ini2],1,'Another launcher is running')
finally:
    assert lock.resolve().parent == Path(tempfile.gettempdir()).resolve()
    if lock.read_text('utf-16') == 'launcher-contract-test':
        lock.unlink()

environment=dict(os.environ,SW_DRAWING_PYTHON=str(folder/'missing-python.exe'))
case('missing runtime macro preserves reset checkpoint',
     [model, '/macro:'+str(folder/'missing.swp'), '/out:'+str(folder/'not-empty'), '/reset'],
     1, 'Packaged macro not found', environment)
case('empty macro', [model, '/macro:'], 1, 'A compiled VBA file is required')
case('wrong macro extension', [model, '/macro:'+str(base/'modules/main.bas')],
     1, 'Macro must be a compiled .swp file')
assert state.read_text('utf-16')=='checkpoint sentinel'
assert (folder/'not-empty/keep.txt').read_text('utf8')=='existing user data'
(folder/'results.json').write_text(json.dumps(results, indent=2), encoding='utf8')
print(f'{len(results)} launcher contract checks passed; report: {folder / "results.json"}')
