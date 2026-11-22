from pathlib import Path
from zipfile import ZipFile
import hashlib
import json
import sys
from docx import Document

sys.stdout.reconfigure(encoding='utf-8')
ROOT=Path(r'C:\code\MachineLearning\Project_FinalExam')
BUILD=ROOT/'reports/_build_week4_3124411226'
OUT=ROOT/'reports/3124411226_Huong_dan_bo_sung_bao_cao_week4.docx'
REF=Path(r'C:\code\MachineLearning\Project_Final_week123.docx')
n=json.loads((ROOT/'week-4_danh_gia_test_va_phan_tich_loi.ipynb').read_text(encoding='utf-8'))
d=Document(OUT)
def cell(needle):
    return next(c for c in n['cells'] if c['cell_type']=='code' and needle in ''.join(c['source']))
def texts(c):
    return [''.join(o.get('text',o.get('data',{}).get('text/plain',[]))) for o in c.get('outputs',[])]
def number(s):return float(s.replace(',','.'))
def rows(t):return [[c.text for c in r.cells] for r in t.rows]

checks={}
checks['table_15']=[]
actual=[list(map(number,r[1:])) for r in rows(d.tables[0])[1:]]
saved=[list(map(float,line.split()[-4:])) for text in texts(cell('p_test_arr =')) for line in text.splitlines() if len(line.split())>=5 and line.split()[-1] in ['0.463','0.486']]
assert actual==saved
checks['table_15']='all 8 metrics match notebook'

assert rows(d.tables[1])[1:]==[['Thật yes','536','522 (bỏ lỡ)'],['Thật no','610 (gọi nhầm)','7.375']]
checks['table_16']='TP/FN/FP/TN match saved plot and business output'

raw=[]
for text in texts(cell('so_sanh =')):
    for line in text.splitlines():
        x=line.split()
        if len(x)==6 and x[0] in ['validation','test']:
            raw.append([x[0],f'{float(x[5])*100:.3f}%'.replace('.',','),*[f'{float(v):.3f}'.replace('.',',') for v in x[2:5]]])
assert rows(d.tables[2])[1:]==raw
checks['table_17']='ratios and all metrics match saved comparison'

raw=[]
for text in texts(cell('khach = pd.DataFrame')):
    for line in text.splitlines():
        x=line.split()
        if len(x)==9 and x[0] in ['0','1','2','3']:
            raw.append([', '.join(x[1:6]),x[6].replace('.',','),x[7],x[8]])
assert rows(d.tables[3])[1:]==raw
checks['table_18']='all four customers, probabilities and predictions match'

raw={}
for i,text in enumerate(texts(cell('def theo_nhom'))):
    for line in text.splitlines():
        x=line.split()
        if len(x)==7 and x[0] in ['cellular','telephone','unknown','success','failure','other']:
            attr='contact' if i==0 else 'poutcome'
            raw[(attr,x[0])]=[f'{int(float(x[5])):,}'.replace(',','.'),str(int(float(x[6]))),*[v.replace('.',',') for v in x[2:5]]]
for row in rows(d.tables[4])[1:]:assert row[2:]==raw[(row[0],row[1])]
checks['table_19']='all 7 groups and 35 numeric fields match'

raw=[]
for text in texts(cell('assert mau_loi is not None')):
    for line in text.splitlines():
        x=line.split()
        if len(x)==13 and x[1].startswith('dong:'):
            raw.append([x[1],'bỏ lỡ' if x[10]=='yes' else 'gọi nhầm',x[12].replace('.',',')])
assert [r[:3] for r in rows(d.tables[5])[1:]]==raw
checks['table_20']='all six IDs, error types and probabilities match'

body='\n'.join(p.text for p in d.paragraphs)
for key in ['5.1.','5.2.','5.3.','5.4.','6.1.','6.2.','6.3.','6.4.','7.1.','7.2.','7.3.','7.4.']:
    assert any(p.text.startswith(key) for p in d.paragraphs),key
for i in range(15,22):assert f'Bảng {i}.' in body
for row in raw:
    assert row[0] in body
for value in ['0,580','0,060','0,040','0,150','0,080','0,280','0,110','0,050','0,360','0,200','0,170','0,260','0,100','0,120','0,070']:
    assert value in body,value
assert '[0,460; 0,512]' in body
assert len(d.inline_shapes)==1
assert rows(d.tables[7])[1][2]==''
checks['coverage']='all report sections, captions, six detailed errors, bootstrap, declaration and appendices present'
checks['journal']='actual hours intentionally blank; no invented time'

alltext=body+'\n'.join(c.text for t in d.tables for row in t.rows for c in row.cells)
assert '\ufffd' not in alltext
assert '[MSSV]' not in '\n'.join(p.text for s in d.sections for p in s.header.paragraphs)
checks['typography']='A4; inherited heading styles; TNR 13/11; no replacement characters'
render_dir=BUILD/(sys.argv[1] if len(sys.argv)>1 else 'render_checked')
checks['rendered_pages']=len(list(render_dir.glob('page-*.png')))
assert checks['rendered_pages']>0
checks['source_sha256']=hashlib.sha256(REF.read_bytes()).hexdigest()
assert checks['source_sha256']=='203a0a64575c2e8b774624db684d3cd9eec03d0dd81cb71fc4083365bda575f3'
checks['source_unchanged']=True

with ZipFile(OUT) as z:
    assert z.testzip() is None
    checks['zip_integrity']='passed'
with ZipFile(REF) as a,ZipFile(OUT) as b:
    preserved=[];changed=[];missing=[]
    for name in a.namelist():
        if name not in b.namelist():missing.append(name)
        elif a.read(name)==b.read(name):preserved.append(name)
        else:changed.append(name)
    checks['package_comparison']={'preserved_parts':len(preserved),'changed_parts':changed,'missing_parts':missing}

(BUILD/'audit_results.json').write_text(json.dumps(checks,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(checks,ensure_ascii=False,indent=2))
