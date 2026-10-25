"""Exercise real VBS -> SolidWorks -> VBA generation for the four supplied templates."""
from pathlib import Path
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1]/"tools"))
from project_paths import work_path
import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import time
import pdfplumber

base=Path(__file__).resolve().parents[1]
root=base.parent
audit=work_path('paper_formats')
audit.mkdir(parents=True,exist_ok=True)
parser=argparse.ArgumentParser()
parser.add_argument('--case',default='A4.landscape',choices=['A4.landscape','A4.portrait','A3.landscape','A3.portrait','assembly','resume','mismatch'])
parser.add_argument('--launcher',type=Path,default=base/'launcher.vbs',help='Use a packaged launcher to verify the runtime kit')
args=parser.parse_args()
name=args.case
host=Path(os.environ['WINDIR'])/'System32/cscript.exe'
source=root/'model/CarFenderGusset.SLDPRT'
if name=='assembly':source=root/'model/SnorkelCaseHolder.SLDASM'
case=audit/name
case.mkdir(exist_ok=True)
inputdir=case/'input';inputdir.mkdir(exist_ok=True)
inputfile=inputdir/source.name
if not inputfile.exists():shutil.copy2(source,inputfile)
if name=='assembly':
    for dependency in (root/'model').iterdir():
        if dependency.suffix.lower() in ['.sldprt','.sldasm'] and not (inputdir/dependency.name).exists():
            shutil.copy2(dependency,inputdir/dependency.name)
ini=case/'test.ini'
settings=(base/'generator.ini').read_text('utf-16')
# All tests select a single drawing; the assembly still uses the real native BOM path.
settings+='\nMaxJobs=1\nJobsPerBatch=1\nTemplatesDirectory='+str(base/'templates')+'\n'
ini.write_bytes(settings.replace('\r\n','\n').replace('\n','\r\n').encode('utf-16'))
out=case/'out'
paper='A4' if name.startswith('A4') or name in ['assembly','resume','mismatch'] else 'A3'
orientation='portrait' if name.endswith('portrait') else 'landscape'
options=['/format:'+paper]
if name=='A4.portrait':
    custom=case/'\u0428\u0430\u0431\u043b\u043e\u043d \u0441 \u043f\u0440\u043e\u0431\u0435\u043b\u0430\u043c\u0438';custom.mkdir(exist_ok=True)
    shutil.copy2(base/'templates/A4.portrait.drwdot',custom/'\u043c\u043e\u0439 \u0448\u0430\u0431\u043b\u043e\u043d.drwdot')
    options+=['/template:'+str(custom/'\u043c\u043e\u0439 \u0448\u0430\u0431\u043b\u043e\u043d.drwdot')]
else:options+=['/orientation:'+orientation]
if name=='resume':
    case=audit/'A4.landscape';inputfile=case/'input'/source.name;ini=case/'test.ini';out=case/'out'
if name=='mismatch':options=['/format:A4','/template:'+str(base/'templates/A3.landscape.drwdot')]
if name not in ['resume','mismatch']:options+=['/reset']
native_before={str(p):hashlib.sha256(p.read_bytes()).hexdigest() for folder in ['Drawings','PDF'] for p in (out/folder).glob('*') if p.is_file()}
command=[str(host),'//nologo',str(args.launcher),'/input:'+str(inputfile),'/ini:'+str(ini),'/out:'+str(out),'/background',*options]
started=time.monotonic()
with (audit/(name+'.log')).open('wb') as log:
    environment=dict(os.environ)
    environment.setdefault('SW_DRAWING_PYTHON',sys.executable)
    run=subprocess.Popen(command,stdout=log,stderr=subprocess.STDOUT,env=environment)
    try:code=run.wait(timeout=480)
    except subprocess.TimeoutExpired:raise RuntimeError('CAD generation timed out; inspect the launcher-owned session before terminating it')
text=(audit/(name+'.log')).read_bytes().decode('cp866',errors='replace')
print(text[-5000:])
if name=='mismatch':
    assert code==2 and 'does not match A4' in text,(code,text)
    assert not (out/'generation-state.tsv').exists(),'Invalid template created a queue'
else:
    assert code==0,(code,text[-2500:])
    rows=[line.split('\t') for line in (out/'generation-state.tsv').read_text('utf-16').splitlines() if line.startswith('JOB\t')]
    assert len(rows)==1 and rows[0][10]=='DONE'
    width,height=(210,297) if orientation=='portrait' and paper=='A4' else (297,420) if orientation=='portrait' else (297,210) if paper=='A4' else (420,297)
    pdfs=list((out/'PDF').glob('*.pdf'));assert len(pdfs)==1
    for path in pdfs:
        with pdfplumber.open(path) as pdf:
            assert len(pdf.pages)>=2,'Expected folded/flat sheets or native BOM'
            for page in pdf.pages:
                assert abs(page.width*25.4/72-width)<.05 and abs(page.height*25.4/72-height)<.05
                textpage=page.extract_text() or ''
                assert '\u0424\u043e\u0440\u043c\u0430\u0442 '+paper in textpage,'Missing template footer'
                if name!='assembly':assert source.stem in textpage.replace('\n',''),'Missing populated title block'
                frame=[20,5,width-5,height-5]
                for x1,y1,x2,y2 in [(frame[0],frame[1],frame[2],frame[1]),(frame[0],frame[3],frame[2],frame[3]),(frame[0],frame[1],frame[0],frame[3]),(frame[2],frame[1],frame[2],frame[3])]:
                    assert any(max(abs(a-b) for a,b in zip([e['x0']*25.4/72,e['y0']*25.4/72,e['x1']*25.4/72,e['y1']*25.4/72],[x1,y1,x2,y2]))<.1 for e in page.edges),'Missing outer frame'
                assert all(c['x0']>=-1 and c['x1']<=page.width+1 and c['y0']>=-1 and c['y1']<=page.height+1 for c in page.chars),'Text outside the page'
            # Check labels created by VBA, independently of the embedded template.
            # This detects project-code-page mistakes that corrupt Cyrillic at runtime.
            vba=(base/'generator.bas').read_text('utf8')
            marker=re.search(r'AddText d, "([^"\n]+)", DRAFT_MARKER_X_M, sheetHeight - DRAFT_MARKER_TOP_MARGIN_M',vba)[1]
            compact=lambda value: re.sub(r'\s+','',value)
            assert compact(marker) in compact(pdf.pages[0].extract_text()),'Cyrillic review label was corrupted'
            font_size='TEXT_HEIGHT_STANDARD_M' if name=='assembly' else 'TEXT_HEIGHT_SMALL_M'
            detail=re.search(r'AddText d, "([^"\n]+)", CONTENT_LEFT_M, sheetHeight - SECTION_HEADER_TOP_MARGIN_M, '+font_size,vba)[1]
            assert compact(detail) in compact(pdf.pages[1].extract_text()),'Cyrillic detail-sheet label was corrupted'
    if name!='assembly':
        coverage=[line.split('\t') for line in (out/'dimension-coverage.tsv').read_text('utf-16').splitlines()[1:]]
        assert max(int(row[4]) for row in coverage)>=60,'Flat dimension coverage regressed'
    if name=='resume':
        assert all(hashlib.sha256(Path(p).read_bytes()).hexdigest()==h for p,h in native_before.items()),'Completed artifacts changed during resume'
result={'test':name,'passed':True,'seconds':round(time.monotonic()-started,1),'exit_code':code,'output':str(out),
        'source_sha256':hashlib.sha256((base/'generator.bas').read_bytes()).hexdigest(),
        'launcher_sha256':hashlib.sha256(args.launcher.read_bytes()).hexdigest()}
if (out/'generator.build.json').exists():
    build=json.loads((out/'generator.build.json').read_text('ascii'))
    assert build['source_sha256']==result['source_sha256'],'The launcher did not build the current BAS source'
    result['macro_sha256']=build['macro_sha256']
(audit/(name+'.json')).write_text(json.dumps(result,indent=2),encoding='utf8')
print(json.dumps(result,indent=2))
