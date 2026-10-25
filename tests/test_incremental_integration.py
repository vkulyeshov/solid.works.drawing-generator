"""Exercise the final launcher against a real SolidWorks process in a dedicated test pack."""
from pathlib import Path
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1]/'tools'))
from project_paths import work_path
import hashlib
import json
import os
import subprocess
import time

base=Path(__file__).resolve().parents[1]
workspace=base.parent
fixture=work_path('paper_formats/A4.landscape')
pack=fixture/'out'
audit=work_path('incremental_audit')
audit.mkdir(parents=True,exist_ok=True)
if not (pack/'generation-state.tsv').exists():
    subprocess.run([sys.executable,'-B',str(base/'tests/test_paper_formats.py'),'--case','A4.landscape'],check=True)
host=Path(os.environ['WINDIR'])/'System32/cscript.exe'
args=[str(host),'//nologo',str(base/'launcher.vbs'),
      '/input:'+str(fixture/'input/CarFenderGusset.SLDPRT'),
      '/ini:'+str(fixture/'test.ini'),'/format:A4','/out:'+str(pack),'/background']
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
def artifacts():
    return {str(p.relative_to(pack)):{'sha256':sha(p),'mtime_ns':p.stat().st_mtime_ns}
            for sub in ['Drawings','PDF'] for p in (pack/sub).iterdir() if p.is_file()}
def execute(name, interrupt=False):
    started=time.monotonic()
    environment=dict(os.environ)
    environment.setdefault('SW_DRAWING_PYTHON',sys.executable)
    process=subprocess.Popen(args,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,env=environment)
    lines=[]; session=None; interrupted=False; interruption=None
    for raw in iter(process.stdout.readline,b''):
        line=raw.decode('cp866',errors='replace').rstrip()
        lines.append(line)
        if line.startswith('Diagnostics: '):session=line[len('Diagnostics: '):]
        if interrupt and not interrupted and ' | Creating drawing' in line:
            assert session
            run=subprocess.run(['powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-File',
                str(base/'tests/InterruptOwnedTest.ps1'),'-SessionDir',session],capture_output=True,timeout=30)
            assert run.returncode==0,run.stderr
            interruption=json.loads(run.stdout)
            interrupted=True
            print('Injected native crash in owned PID',interruption['pid'],flush=True)
        if 'Recovering interrupted job' in line or line.startswith('Exit code:'):
            print(name+': '+line,flush=True)
        if time.monotonic()-started>600:
            raise TimeoutError('Integration test exceeded 10 minutes; inspect the dedicated test session')
    code=process.wait(timeout=10)
    (audit/(name+'.log')).write_text('\n'.join(lines),encoding='utf8')
    assert code==0,(name,code,lines[-12:])
    if interrupt:
        assert interrupted and any('Recovering interrupted job' in x for x in lines)
    return {'test':name,'passed':True,'seconds':round(time.monotonic()-started,1),
            'session':session,'interruption':interruption,'log':str(audit/(name+'.log'))}

results=[]
before=artifacts()
runlog=(pack/'run.log').read_text('utf-16')
results.append(execute('completed_resume'))
assert artifacts()==before,'A completed artifact was rewritten during no-op resume'
assert (pack/'run.log').read_text('utf-16').count(' JOB ')==runlog.count(' JOB ')
pdf=next((pack/'PDF').glob('*.pdf'))
missing=audit/'gusset-before-crash-test.pdf'
assert not missing.exists()
pdf.rename(missing)
try:
    results.append(execute('missing_pdf_and_native_crash',interrupt=True))
finally:
    if not pdf.exists():missing.replace(pdf)
rows=[x.split('\t') for x in (pack/'generation-state.tsv').read_text('utf-16').splitlines() if x.startswith('JOB\t')]
assert len(rows)==1 and all(x[10]=='DONE' for x in rows)
assert pdf.stat().st_size==int(rows[0][14])
assert (pack/'Drawings'/(rows[0][11]+'.slddrw')).stat().st_size==int(rows[0][13])
results.append({'test':'checkpoint_commits_both_artifacts','passed':True})
data={'source_sha256':sha(base/'generator.bas'),
      'launcher_sha256':sha(base/'launcher.vbs'),'checks':results}
(audit/'incremental-integration.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8')
print('Final incremental integration checks passed',flush=True)
