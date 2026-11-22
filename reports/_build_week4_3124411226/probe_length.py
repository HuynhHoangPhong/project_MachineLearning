from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.oxml.ns import qn
import sys

ROOT=Path(r'C:\code\MachineLearning\Project_FinalExam')
BUILD=ROOT/'reports/_build_week4_3124411226'
source=ROOT/'reports/3124411226_Huong_dan_bo_sung_bao_cao_week4.docx'
tag=sys.argv[1]
if len(sys.argv)>2:source=Path(sys.argv[2])
original=Document(source)
active=False
blocks=[]
for e in original._element.body:
    text=''.join(e.itertext()) if False else ''.join(e.xpath('.//w:t/text()'))
    if text=='5. Kết quả trên test':active=True
    if text=='Bổ sung các mục cuối báo cáo':active=False
    if not active or e.tag==qn('w:sectPr'):continue
    if text.startswith('Hướng dẫn đặt nội dung:') or text.startswith('Nội dung để chép'):continue
    blocks.append(deepcopy(e))
out=Document(source)
for e in list(out._element.body):
    if e.tag!=qn('w:sectPr'):out._element.body.remove(e)
for e in blocks:out._element.body.insert(len(out._element.body)-1,e)
out.save(BUILD/f'report_content_{tag}.docx')
print('Report-only probe:',BUILD/f'report_content_{tag}.docx')
