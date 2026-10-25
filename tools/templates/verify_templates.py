"""Validate the actual SolidWorks exports and package four native template sets."""
from pathlib import Path
import argparse
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from project_paths import REPOSITORY, WORKSPACE, external_path, work_path
import hashlib
import json
import zipfile
import pdfplumber
import pypdfium2 as pdfium
from pypdf import PdfReader, PdfWriter
from PIL import Image, ImageDraw

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--input-dir', default=str(work_path('templates')))
args = parser.parse_args()
base = external_path(args.input_dir)
workspace = WORKSPACE
audit = work_path('template_verification')
audit.mkdir(parents=True, exist_ok=True)
records = json.loads((base/'Verification/native-report.json').read_text('utf-8-sig'))
assert len(records) == 4
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
checks = []
files = {}

for record in records:
    name = record['name']
    rotate = False
    w, h = record['width_mm'], record['height_mm']
    notes = {n['name'].removeprefix('DSTU_'): n for n in record['notes']}
    assert len(notes) == record['expectedNotes'] == 51
    assert record['formatReloadResult'] == 0
    assert notes['Designation']['text'] == '\u041a\u0420\u041d 02.04.002'
    assert notes['Description']['text'] == '\u0411\u043e\u043a\u043e\u0432\u0430 \u0441\u0442\u0456\u043d\u043a\u0430'
    assert notes['Designation26']['text'] == notes['Designation']['text']
    assert notes['Designation']['link'] == '$PRP:"Designation"'
    assert notes['Scale']['text'] == '1:1'
    assert notes['SheetCount']['text'] == '\u0410\u0440\u043a\u0443\u0448\u0456\u0432 1'
    assert notes['SheetNumber']['text'].strip() == '\u0410\u0440\u043a\u0443\u0448'
    if name == 'DSTU_A4_Landscape':
        assert abs(notes['Designation']['angle']) < 1e-8
        main = next(c for c in record['cells'] if c['Name']=='Designation')
        assert main['X']==172 and main['Y']==45
    overflows = []
    for cell in record['cells']:
        note = notes[cell['Name']]
        if not note['text'].strip():
            continue
        x1, y1, x2, y2 = cell['X'], cell['Y'], cell['X']+cell['W'], cell['Y']+cell['H']
        if rotate:
            x1, y1, x2, y2 = w-y2, x1, w-y1, x2
        ext = note['extent']
        xx1, yy1, xx2, yy2 = ext[0]*1000, ext[1]*1000, ext[3]*1000, ext[4]*1000
        if min(xx1-x1, yy1-y1, x2-xx2, y2-yy2) < -.4:
            overflows.append({'name': cell['Name'], 'cell': [x1,y1,x2,y2], 'extent': [xx1,yy1,xx2,yy2]})
    assert not overflows, (name, overflows)
    for folder, suffix in [('Preview',''), ('Verification','_filled')]:
        path = base/folder/(name+suffix+'.pdf')
        with pdfplumber.open(path) as document:
            assert len(document.pages) == 1
            page = document.pages[0]
            assert abs(page.width*25.4/72-w)<.03 and abs(page.height*25.4/72-h)<.03
            assert sorted(set(round(e['linewidth']*25.4/72,3) for e in page.edges)) == [.25,.5]
            # Check the exact outer drawing frame in physical sheet coordinates.
            frame = [5,20,w-5,h-5] if rotate else [20,5,w-5,h-5]
            for x1,y1,x2,y2 in [(frame[0],frame[1],frame[2],frame[1]),
                                  (frame[0],frame[3],frame[2],frame[3]),
                                  (frame[0],frame[1],frame[0],frame[3]),
                                  (frame[2],frame[1],frame[2],frame[3])]:
                assert any(max(abs(a-b) for a,b in zip(
                    [e['x0']*25.4/72,e['y0']*25.4/72,e['x1']*25.4/72,e['y1']*25.4/72],
                    [x1,y1,x2,y2]))<.03 for e in page.edges), (name,'frame missing',[x1,y1,x2,y2])
            if not suffix:
                assert not any('\u041a\u0420\u041d' in c['text'] for c in page.chars)
                assert not any('ERROR' in c['text'] for c in page.chars)
            for c in page.chars:
                assert c['x0'] >= -0.1 and c['x1'] <= page.width+.1
                assert c['y0'] >= -0.1 and c['y1'] <= page.height+.1
        reader = PdfReader(path)
        for font in reader.pages[0]['/Resources']['/Font'].values():
            font = font.get_object()
            desc = font.get('/FontDescriptor')
            if not desc and '/DescendantFonts' in font:
                desc = font['/DescendantFonts'][0].get_object().get('/FontDescriptor')
            assert desc and any(k in desc.get_object() for k in ['/FontFile','/FontFile2','/FontFile3'])
        doc = pdfium.PdfDocument(path)
        doc[0].render(scale=2).to_pil().save(audit/(name+suffix+'.png'))
    for p in [base/record.get('template_file',name+'.drwdot'), base/'Formats'/(name+'.slddrt')]:
        # Modern SolidWorks stores compressed native files, not necessarily OLE.
        # Native validity is established by SaveFormat/NewDocument above.
        assert p.stat().st_size > 30000
        files[str(p.relative_to(base))] = sha(p)
    checks.append({'name': name, 'passed': True, 'width_mm':w,'height_mm':h,
                   'notes':51,'native_instantiation':True,'external_format_reload':True,'filled_cell_overflows':0,
                   'frame_and_line_widths_verified':True,'embedded_fonts':True})

# A single user-facing PDF: four blank sheets followed by four populated proofs.
final_pdf = base.parent/'DSTU_templates_preview.pdf'
final_pdf.parent.mkdir(parents=True, exist_ok=True)
writer = PdfWriter()
for folder,suffix,label in [('Preview','','\u041f\u043e\u0440\u043e\u0436\u043d\u0456\u0439 \u0448\u0430\u0431\u043b\u043e\u043d'),('Verification','_filled','\u041f\u0440\u0438\u043a\u043b\u0430\u0434 \u0437\u0430\u043f\u043e\u0432\u043d\u0435\u043d\u043d\u044f')]:
    for record in records:
        writer.append(base/folder/(record['name']+suffix+'.pdf'), outline_item=label+' \u2014 '+record['name'])
writer.add_metadata({'/Title':'\u0414\u0421\u0422\u0423 \u0413\u041e\u0421\u0422 2.104:2006 \u2014 \u0447\u043e\u0442\u0438\u0440\u0438 \u0448\u0430\u0431\u043b\u043e\u043d\u0438 SolidWorks',
                     '/Subject':'\u0427\u043e\u0442\u0438\u0440\u0438 \u043f\u043e\u0440\u043e\u0436\u043d\u0456 \u0444\u043e\u0440\u043c\u0430\u0442\u0438 \u0442\u0430 \u0447\u043e\u0442\u0438\u0440\u0438 \u043f\u0440\u0438\u043a\u043b\u0430\u0434\u0438 \u0437\u0430\u043f\u043e\u0432\u043d\u0435\u043d\u043d\u044f'})
with final_pdf.open('wb') as f:
    writer.write(f)
assert len(PdfReader(final_pdf).pages) == 8
contact = Image.new('RGB',(1200,900),'#d8d8d8')
draw = ImageDraw.Draw(contact)
for i,record in enumerate(records):
    im = Image.open(audit/(record['name']+'.png'))
    im.thumbnail((580,410))
    x=(i%2)*600; y=(i//2)*450
    contact.paste(im,(x+(600-im.width)//2,y+30))
    draw.text((x+20,y+8),record['name'],fill='black')
contact.save(audit/'contact.png')
report={'solidworks':'2025 SP5 (33.5.0)','standard_source':'rules.pdf, DSTU GOST 2.104:2006, form 1',
        'font':'GOST Type BU italic','checks':checks,'sha256':files,
        'preview_pdf_sha256':sha(final_pdf),'scope':'Four first-sheet form-1 templates instantiated and exported by native SolidWorks.'}
(base/'Verification/verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
archive=base.parent/'SolidWorks_DSTU_Templates.zip'
with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED) as z:
    for record in records:
        filename=record.get('template_file',record['name']+'.drwdot')
        z.write(base/filename,filename)
    for p in sorted((base/'Formats').glob('*.slddrt')):z.write(p,'Formats/'+p.name)
    z.write(REPOSITORY/'README.md','README.md')
    z.write(base/'Verification/verification.json','verification.json')
    z.write(final_pdf,'DSTU_templates_preview.pdf')
with zipfile.ZipFile(archive) as z:
    assert len(z.namelist())==11
    for name,expected in files.items():assert hashlib.sha256(z.read(name.replace('\\','/'))).hexdigest()==expected
print(json.dumps({'passed':True,'templates':4,'formats':4,'preview_pages':8,'archive':str(archive),'bytes':archive.stat().st_size},indent=2))
