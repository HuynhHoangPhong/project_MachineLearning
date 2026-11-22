from pathlib import Path
from copy import deepcopy
from zipfile import ZipFile
import ast
import base64
import hashlib
import json
import re
import shutil
import sys
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont

sys.stdout.reconfigure(encoding='utf-8')
ROOT = Path(r'C:\code\MachineLearning\Project_FinalExam')
BUILD = ROOT / 'reports/_build_week4_3124411226'
REF = Path(r'C:\code\MachineLearning\Project_Final_week123.docx')
OUT = ROOT / 'reports/3124411226_Huong_dan_bo_sung_bao_cao_week4.docx'
SKILL = Path(r'C:\Users\ADMIN\.codex\plugins\cache\openai-primary-runtime\documents\26.1119.11507\skills\documents')
BUILD.mkdir(parents=True, exist_ok=True)
reference_hash = hashlib.sha256(REF.read_bytes()).hexdigest()
shutil.copy2(REF, BUILD / REF.name)
nb = json.loads((ROOT / 'week-4_danh_gia_test_va_phan_tich_loi.ipynb').read_text(encoding='utf-8'))
w3 = json.loads((ROOT / 'week-3_mlp_qua_khop_va_chon_mo_hinh.ipynb').read_text(encoding='utf-8'))
source = [''.join(c['source']) for c in nb['cells']]
lock = ast.literal_eval(ast.parse(next(s for s in source if s.startswith('QUYET_DINH ='))).body[0].value)
assert lock == {'mssv':'3124411226','mo_hinh':'E5','so_no_ron':256,'nguong':0.6,'f1_yes_val':0.4806,'thoi_diem':'2026-11-18 09:09'}
assert not any(o.get('output_type')=='error' for c in nb['cells'] for o in c.get('outputs', []))

def plain(c):
    return '\n'.join(''.join(o.get('text', o.get('data', {}).get('text/plain', []))) for o in c.get('outputs', []))

def output_cell(needle):
    return next(c for c in nb['cells'] if c['cell_type']=='code' and needle in ''.join(c['source']))

conf = output_cell('ConfusionMatrixDisplay.from_predictions(test')
confirmation = plain(output_cell('f1_lai ='))
assert '0.4806' in confirmation and 'E5' in confirmation
assert '[0.46, 0.512]' in plain(output_cell('kq = bootstrap'))
assert all(v in plain(output_cell('khach = pd.DataFrame')) for v in ['0.851','0.103','0.951','0.431'])
for o in conf['outputs']:
    if 'image/png' in o.get('data', {}):
        (BUILD / 'confusion_matrix.png').write_bytes(base64.b64decode(''.join(o['data']['image/png'])))

# Figure B.4: recompose saved output values without training or predicting again.
canvas = Image.new('RGB', (1800, 1940), 'white')
draw = ImageDraw.Draw(canvas)
font_path = r'C:\Windows\Fonts\times.ttf'
bold_path = r'C:\Windows\Fonts\timesbd.ttf'
f = ImageFont.truetype(font_path, 41)
fb = ImageFont.truetype(bold_path, 43)
draw.text((35,20), '(a) Xác nhận đúng mô hình đã khóa — mục 4.1', font=fb, fill='black')
draw.text((35,85), 'ĐẠT: đúng mô hình đã khóa: E5, tau* = 0.6,', font=f, fill='black')
draw.text((35,140), 'F1-yes validation = 0.4806', font=f, fill='black')
draw.text((35,240), '(b) Kết quả test ở hai ngưỡng — mục 4.2', font=fb, fill='black')
draw.text((35,302), 'ĐẠT: đã đánh giá trên test.', font=f, fill='black')
cols=[35,535,825,1120,1430,1765]
rows=[390,525,630,735]
headers=['Ngưỡng','accuracy','precision_yes','recall_yes','f1_yes']
for j,h in enumerate(headers):
    if j in [2,3]:
        a,b=h.split('_'); draw.text((cols[j]+15,410), a, font=f, fill='black');draw.text((cols[j]+15,460),'_'+b,font=f,fill='black')
    else: draw.text((cols[j]+15,435),h,font=f,fill='black')
for k,row in enumerate([['0,5','0.839','0.380','0.595','0.463'],['tau* = 0.6','0.875','0.468','0.507','0.486']]):
    for j,v in enumerate(row): draw.text((cols[j]+15,rows[k+1]+25),v,font=f,fill='black')
for x in cols: draw.line((x,390,x,735), fill='#777777', width=2)
for y in rows: draw.line((35,y,1765,y),fill='#777777',width=2)
draw.text((35,790), '(c) Ma trận nhầm lẫn tại tau* = 0.6 — mục 4.2',font=fb,fill='black')
im=Image.open(BUILD / 'confusion_matrix.png').convert('RGB')
im=im.resize((1408,1056),Image.Resampling.LANCZOS)
canvas.paste(im,(185,860))
canvas.save(BUILD / 'Hinh_B4.png',dpi=(300,300))

# Retain actual styles and package components from the user's document.
doc = Document(REF)
body=doc._element.body
for element in list(body):
    if element.tag != qn('w:sectPr'):
        body.remove(element)
section=doc.sections[0]
section.page_width=Cm(21);section.page_height=Cm(29.7)
section.left_margin=Cm(3);section.right_margin=Cm(2)
section.top_margin=Cm(2);section.bottom_margin=Cm(2)
for sec in doc.sections:
    for container in [sec.header,sec.first_page_header,sec.even_page_header]:
        for par in container.paragraphs:
            if any(token in par.text for token in ['[MSSV]','Báo cáo đồ án Máy học']):
                for run in par.runs:
                    run.text=''
                run=par.add_run('Hướng dẫn báo cáo tuần 4 — 3124411226')
                run.font.name='Times New Roman';run.font.size=Pt(9)
                run.font.color.rgb=RGBColor(100,100,100)
normal=doc.styles['Normal']
normal.font.name='Times New Roman'; normal.font.size=Pt(13);normal.font.color.rgb=RGBColor(0,0,0)
normal.paragraph_format.line_spacing=1.5
normal.paragraph_format.space_after=Pt(6)
normal.paragraph_format.widow_control=True
for name in ['Title','Heading 1','Heading 2']:
    doc.styles[name].font.color.rgb=RGBColor(0,0,0)

def p(text='', bold=False, size=None, italic=False, keep=False):
    par=doc.add_paragraph()
    par.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    par.paragraph_format.keep_with_next=keep
    previous=par._p.getprevious()
    if previous is not None and previous.tag==qn('w:tbl'):
        par.paragraph_format.space_before=Pt(6)
    run=par.add_run(text);run.bold=bold;run.italic=italic
    if size:run.font.size=Pt(size)
    return par

def note(text):
    par=p('Hướng dẫn đặt nội dung: '+text,size=11,italic=True,keep=True)
    par.paragraph_format.line_spacing=1.15
    par.paragraph_format.space_after=Pt(7)
    return par

def label(text='Nội dung để chép'):
    par=p(text,bold=True,size=11,keep=True)
    par.paragraph_format.line_spacing=1.0

def heading(text,level=2):
    par=doc.add_paragraph(text,style=f'Heading {level}')
    par.paragraph_format.keep_with_next=True
    return par

def caption(text):
    par=doc.add_paragraph(text)
    par.alignment=WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.line_spacing=1.0
    par.paragraph_format.space_before=Pt(6)
    par.paragraph_format.space_after=Pt(6)
    par.paragraph_format.keep_with_next=True
    for r in par.runs:r.font.size=Pt(11);r.italic=True
    return par

def table(headers,rows,widths):
    t=doc.add_table(rows=1,cols=len(headers))
    t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
    pr=t._tbl.tblPr
    borders=OxmlElement('w:tblBorders')
    # Reuse the restrained horizontal rules of the reference report.
    for side in ['top','left','bottom','right','insideH','insideV']:
        el=OxmlElement('w:'+side)
        el.set(qn('w:val'),'single' if side in ['top','bottom','insideH'] else 'nil')
        el.set(qn('w:sz'),'4');el.set(qn('w:color'),'000000')
        borders.append(el)
    pr.append(borders)
    margins=OxmlElement('w:tblCellMar')
    for side,value in [('top',70),('bottom',70),('left',65),('right',65)]:
        el=OxmlElement('w:'+side);el.set(qn('w:w'),str(value));el.set(qn('w:type'),'dxa');margins.append(el)
    pr.append(margins)
    for j,h in enumerate(headers):t.rows[0].cells[j].text=h
    for row in rows:
        cells=t.add_row().cells
        for j,value in enumerate(row):cells[j].text=str(value)
    for col,w in zip(t.columns,widths):col.width=Cm(w)
    for row in t.rows:
        trpr=row._tr.get_or_add_trPr()
        cant=OxmlElement('w:cantSplit');trpr.append(cant)
        for j,c in enumerate(row.cells):
            c.width=Cm(widths[j]);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for par in c.paragraphs:
                par.paragraph_format.line_spacing=1.05
                par.paragraph_format.space_before=Pt(1);par.paragraph_format.space_after=Pt(1)
                par.alignment=WD_ALIGN_PARAGRAPH.LEFT if len(str(c.text))>18 else WD_ALIGN_PARAGRAPH.CENTER
                for r in par.runs:r.font.name='Times New Roman';r.font.size=Pt(11)
    repeat=OxmlElement('w:tblHeader');t.rows[0]._tr.get_or_add_trPr().append(repeat)
    for c in t.rows[0].cells:
        for par in c.paragraphs:
            for r in par.runs:r.bold=True
    return t

def block(text):
    for par in text.strip().split('\n\n'):p(par.strip())

title=doc.add_paragraph('Hướng dẫn bổ sung báo cáo tuần 4',style='Title')
title.alignment=WD_ALIGN_PARAGRAPH.CENTER
p('Bank Marketing — Huỳnh Hoàng Phong — MSSV 3124411226',bold=True)
p('Bản hướng dẫn này dùng cùng báo cáo Project_Final_week123.docx. Nội dung bổ sung chỉ gồm Mục 5–7 và phần tuần 4 trong các mục cuối báo cáo. Mô hình đã khóa: E5; ngưỡng 0,6; F1-yes validation ghi trong quyết định: 0,4806.')
p('Cách sử dụng',bold=True,keep=True)
p('Dùng Ctrl+F tìm đúng tiêu đề hoặc chú thích được nêu trong từng phần. Giữ tiêu đề đang có trong báo cáo; điền bảng trống và thay dấu “…” bằng nội dung ở dưới nhãn “Nội dung để chép”. Các dòng “Hướng dẫn đặt nội dung” là chỉ dẫn thao tác, không chép vào bản báo cáo nộp.')
p('Định dạng kế thừa báo cáo hiện tại: A4; lề trái 3 cm, các lề còn lại 2 cm; Times New Roman 13 pt, giãn dòng 1,5; bảng và chú thích 11 pt. Giữ số và tên Bảng 15–21, Hình B.4 theo mẫu. Số thập phân dùng dấu phẩy, chỉ số làm tròn ba chữ số; dòng QUYET_DINH giữ nguyên JSON đã khóa.')
p('Thứ tự bổ sung',bold=True,keep=True)
for text in ['Mục 5.1–5.4: Bảng 15–18, diễn giải test, so sánh hai tập, khách giả định và bootstrap.',
             'Mục 6.1–6.4: Bảng 19–20, phân tích đủ sáu lỗi và đề xuất kiểm chứng.',
             'Mục 7.1–7.4: Bảng 21, kết luận, giới hạn, hướng phát triển và bài học.',
             'Các mục cuối: nguồn bổ sung, khai báo công cụ, dòng tuần 4 của Phụ lục A, Hình B.4 và câu tuần 4 trong Phụ lục C.']:
    p('• '+text)
p('Tình trạng cá nhân hóa',bold=True,keep=True)
p('Số liệu được đối chiếu từ notebook đã lưu: E5, ngưỡng 0,6, seed lấy mẫu 226, sáu mã dòng riêng và các đầu ra giải thích thực tế. Các phép phân tích không dùng để đổi mô hình hoặc chọn lại ngưỡng trên test. Số giờ thực tế trong nhật ký để người thực hiện xác nhận sau.')

heading('5. Kết quả trên test',1)
note('Tìm “5. Kết quả trên test”, nằm ngay sau Mục 4.3 Quyết định đã khóa. Bổ sung lần lượt bốn tiểu mục dưới đây; số mục này thuộc báo cáo, không phải notebook.')
heading('5.1. Kết quả của mô hình đã khóa')
note('Tại Mục 5.1, điền Bảng 15 và Bảng 16 đúng các ô trống sẵn có. Đặt hai đoạn giới thiệu trước Bảng 15; thay dấu “…” sau Bảng 16 bằng phần diễn giải nghiệp vụ bên dưới. Không cần thêm một hình ma trận ở Mục 5.1 vì Hình B.4 sẽ lưu đầu ra đó trong phụ lục.')
label()
block('''Mô hình được đánh giá là rừng ngẫu nhiên E5 với 200 cây, min_samples_leaf=5, class_weight="balanced" và random_state=42. Ngưỡng đã khóa là τ* = 0,6. Ô 4.1 của notebook tuần 4 dựng lại mô hình và xác nhận F1-yes trên validation khớp với quyết định tuần 3.

Bảng 15 trình bày kết quả tại ngưỡng đã khóa và ngưỡng 0,5 như mốc đối chiếu được quy định trước. Không thay đổi mô hình hoặc ngưỡng theo kết quả test.''')
caption('Bảng 15. Kết quả trên test ở hai ngưỡng')
table(['Ngưỡng','Accuracy','Precision-yes','Recall-yes','F1-yes'],[
 ['0,5','0,839','0,380','0,595','0,463'],['τ* = 0,6','0,875','0,468','0,507','0,486']],[3.1,2.6,3.7,3.5,3.1])
caption('Bảng 16. Ma trận nhầm lẫn trên test tại τ* (hàng là nhãn thật)')
table(['','Dự đoán yes (được gọi)','Dự đoán no (không gọi)'],[
 ['Thật yes','536','522 (bỏ lỡ)'],['Thật no','610 (gọi nhầm)','7.375']],[3.2,6.4,6.4])
block('''Nếu chỉ gọi những khách được mô hình dự đoán là yes, tại ngưỡng 0,6 ngân hàng sẽ chọn 1.146 trong 9.043 khách, khoảng 12,7% danh sách. Trong nhóm được chọn có 536 khách thực tế đồng ý và 610 khách không đồng ý, tương ứng tỉ lệ thành công khoảng 46,8% trên test.

Để so sánh công bằng, phương án gọi ngẫu nhiên cũng sử dụng 1.146 cuộc gọi. Với tỉ lệ đồng ý của test là 1.058/9.043, số khách đồng ý kỳ vọng của gọi ngẫu nhiên là khoảng 134 khách. Tỉ lệ đồng ý trong nhóm do mô hình chọn cao khoảng 4 lần mức nền 11,7%. Kết luận này được rút ra từ Precision-yes và số cuộc gọi, không chỉ từ F1-yes.

Mô hình phát hiện 536/1.058 khách đồng ý, đạt Recall-yes khoảng 50,7%, và vẫn bỏ lỡ 522 khách. Khi tăng ngưỡng từ 0,5 lên 0,6, Precision-yes tăng từ 0,380 lên 0,468 nhưng Recall-yes giảm từ 0,595 xuống 0,507. Việc tập trung vào nhóm có tỉ lệ đồng ý cao hơn đi kèm việc bỏ lỡ thêm khách tiềm năng; các tỉ lệ trên dữ liệu lịch sử chưa bảo đảm kết quả của một chiến dịch tương lai.''')

heading('5.2. So sánh kết quả trên validation và test')
note('Điền Bảng 17 sẵn có tại Mục 5.2. Thay dấu “…” dưới bảng bằng hai đoạn sau; đoạn đầu so sánh số liệu, đoạn sau trả lời câu suy ngẫm. Tỉ lệ yes được giữ dưới dạng phần trăm để nhất quán với Bảng 1 của báo cáo hiện tại.')
label()
caption('Bảng 17. Cùng mô hình, cùng ngưỡng τ* trên validation và test')
table(['Tập','Tỉ lệ yes','Precision-yes','Recall-yes','F1-yes'],[
 ['validation','11,701%','0,455','0,509','0,481'],['test','11,700%','0,468','0,507','0,486']],[3.0,3.0,3.7,3.4,2.9])
block('''Với cùng mô hình E5 và ngưỡng 0,6, F1-yes trên test cao hơn validation khoảng 0,006 và Precision-yes cao hơn khoảng 0,013, tính từ các giá trị chưa làm tròn. Recall-yes giảm khoảng 0,003, tương ứng số khách đồng ý được phát hiện giảm từ 539 xuống 536 trong khi mỗi tập đều có 1.058 khách yes. Chênh lệch nhỏ và phù hợp với dao động ngẫu nhiên; việc chọn trên validation có thể làm điểm validation lạc quan nhưng không có nghĩa điểm test luôn thấp hơn. Có thể báo cáo điểm test như kết quả đánh giá độc lập trên dữ liệu cùng nguồn, song phép chia ngẫu nhiên từ dữ liệu giai đoạn 2008–2010 chưa kiểm chứng khả năng áp dụng cho chiến dịch ở thời điểm khác.

Nếu F1-yes trên test thấp hơn validation 0,01, cũng không quay lại chọn mô hình khác dựa trên test. Việc đó sẽ làm test trở thành một tập phục vụ lựa chọn và mất vai trò đánh giá độc lập. Cần giữ nguyên quyết định đã khóa, báo cáo trung thực và giải thích giới hạn của kết quả.''')

heading('5.3. Thử mô hình với khách hàng giả định')
note('Điền bốn dòng có sẵn của Bảng 18. Xóa dòng tùy chọn “(khách do sinh viên tự đặt)” vì notebook hiện tại chỉ có bốn khách; không điền thêm dự đoán chưa chạy. Đặt phần nhận xét dưới bảng, thay dấu “…” đang có.')
label()
caption('Bảng 18. Dự đoán cho các khách hàng giả định')
table(['Khách (tuổi, nghề, kênh, tháng, poutcome)','p_yes','Ngưỡng 0,5','Ngưỡng τ*'],[
 ['22, student, cellular, mar, unknown','0,851','yes','yes'],
 ['45, blue-collar, unknown, may, unknown','0,103','no','no'],
 ['67, retired, telephone, oct, success','0,951','yes','yes'],
 ['35, technician, cellular, aug, unknown','0,431','no','no']],[8.3,2.0,2.85,2.85])
block('''Các dự đoán nhìn chung phù hợp với EDA tuần 1. Khách 22 tuổi liên lạc tháng 3 có xác suất 0,851, phù hợp với tỉ lệ đồng ý tương đối cao của nhóm trẻ và tháng 3 trong train. Khách nghỉ hưu 67 tuổi có lịch sử success đạt xác suất 0,951, phù hợp với xu hướng của nhóm trên 60 tuổi và nhóm từng có chiến dịch thành công. Khách 45 tuổi có contact = unknown, liên lạc tháng 5, chỉ đạt 0,103; các nhóm này cũng có tỉ lệ đồng ý thấp trong EDA.

Khách kỹ thuật viên 35 tuổi đạt 0,431 và được dự đoán no, phù hợp với tỉ lệ đồng ý của nhóm tuổi 30–40 và tháng 8 thấp hơn mức chung trong train. Cả bốn khách giữ nguyên nhãn ở hai ngưỡng. Đây chỉ là đối chiếu với xu hướng; các khách giả định không có nhãn thật nên chưa thể đánh giá đúng hoặc sai. Trường hợp khách 22 tuổi cũng cho thấy poutcome = unknown không đồng nghĩa với việc luôn bị dự đoán no.

Nếu chỉ gọi khách có xác suất cao, một số nhóm có thể gần như không được liên hệ. Ở ngưỡng 0,6, từ các chỉ số của nhóm contact = unknown tại Bảng 19 có thể suy ra mô hình chỉ chọn 4/2.634 khách, khoảng 0,15%. Điều này có thể làm giảm cơ hội tiếp cận sản phẩm và khiến dữ liệu chiến dịch sau ít phản hồi từ nhóm bị đánh giá thấp. Đây là giới hạn về khả năng tiếp cận và dữ liệu, chưa phải kết luận toàn diện về công bằng.''')

heading('5.4. Khoảng tin cậy bootstrap (nâng cao, không bắt buộc)')
note('Thay nguyên dòng mẫu “F1-yes trên test là …” tại Mục 5.4 bằng hai đoạn sau. Không cần thêm bảng hay hình mới cho bootstrap.')
label()
block('''F1-yes trên test là 0,486 (khoảng tin cậy 95%: [0,460; 0,512], bootstrap, B = 1.000). Notebook lấy mẫu có hoàn lại các cặp nhãn thật và dự đoán, giữ đúng cặp của mỗi khách, sử dụng phương pháp phân vị và random_state=42.

Khoảng có độ rộng 0,052, tương đối rộng so với các chênh lệch F1 chỉ vài phần nghìn. Các chênh lệch nhỏ giữa mô hình vì vậy chưa nên xem là chắc chắn nếu chỉ dựa vào điểm F1; bản thân khoảng này không thay thế một phép so sánh thống kê giữa hai mô hình. Khoảng chỉ phản ánh dao động do lấy mẫu cho mô hình cố định trên dữ liệu cùng nguồn, chưa đánh giá tác động của thay đổi phân phối khách hàng trong tương lai.''')

heading('6. Phân tích lỗi',1)
heading('6.1. Lỗi theo nhóm khách hàng')
note('Điền Bảng 19 dưới Mục 6.1 theo thứ tự dòng của mẫu: ba nhóm contact, rồi success, failure, other, unknown của poutcome. Thay dấu “…” dưới bảng bằng phần nhận xét. Không thêm Accuracy vì mẫu Bảng 19 không yêu cầu cột này.')
label()
caption('Bảng 19. Chỉ số trên test theo kênh liên lạc và kết quả chiến dịch trước')
table(['Thuộc tính','Giá trị','Số khách','Số khách yes','Precision-yes','Recall-yes','F1-yes'],[
 ['contact','cellular','5.820','888','0,479','0,553','0,513'],
 ['contact','telephone','589','67','0,368','0,642','0,467'],
 ['contact','unknown','2.634','103','0,500','0,019','0,037'],
 ['poutcome','success','306','199','0,650','1,000','0,788'],
 ['poutcome','failure','1.012','140','0,375','0,536','0,441'],
 ['poutcome','other','355','62','0,416','0,597','0,490'],
 ['poutcome','unknown','7.370','657','0,408','0,342','0,373']],[2.2,2.25,1.9,2.2,2.7,2.5,2.25])
block('''Nhóm contact = unknown có Recall-yes thấp nhất, đạt 0,019. Trong 103 khách thực tế đồng ý, mô hình chỉ phát hiện 2 khách và bỏ lỡ 101 khách. Một giả thuyết là mô hình học xu hướng ít đồng ý của nhóm này trong train, nơi EDA ghi nhận tỉ lệ đồng ý khoảng 4,3%, nên nhiều khách yes vẫn nhận xác suất dưới ngưỡng.

Recall thấp nhất thể hiện tỉ lệ bỏ lỡ cao nhất, không đồng nghĩa với số khách bỏ lỡ lớn nhất. Trong bảng contact, nhóm cellular bỏ lỡ 397 khách, nhiều hơn nhóm unknown, do có đến 888 khách yes.

Theo poutcome, nhóm unknown có Recall-yes thấp nhất là 0,342, với 657 khách yes và 432 khách bị bỏ lỡ. Ngược lại, nhóm success có Recall-yes bằng 1,000 nhưng Precision-yes chỉ đạt 0,650; việc phát hiện toàn bộ khách đồng ý trong nhóm này vẫn đi kèm các cuộc gọi nhầm.''')

heading('6.2. Bảng tổng hợp sáu lỗi')
note('Đặt đoạn giới thiệu trước Bảng 20, rồi điền đủ sáu dòng sẵn có. Giữ đúng mã dòng và thứ tự lấy mẫu của notebook. “Nhóm lỗi” là mô tả phân tích, không phải nhãn bổ sung của dữ liệu.')
label()
p('Trên toàn bộ test, mô hình bỏ lỡ 522 khách và gọi nhầm 610 khách. Sáu trường hợp được lấy ngẫu nhiên bằng cách chọn ba khách từ mỗi loại lỗi với random_state=S_SV. Seed cá nhân là S_SV=226, lấy từ ba chữ số cuối MSSV 3124411226; sáu mã dòng đều khác nhau.')
caption('Bảng 20. Sáu lỗi của mô hình đã khóa trên test')
table(['Mã dòng','Loại lỗi','p_yes','Đặc điểm chính của khách','Nhóm lỗi'],[
 ['dong:11712','bỏ lỡ','0,130','60 tuổi, retired; contact và poutcome unknown','Thiếu thông tin kênh và lịch sử'],
 ['dong:18051','bỏ lỡ','0,401','41 tuổi, blue-collar; số dư 0; pdays = −1','Chưa có lịch sử chiến dịch trước'],
 ['dong:21621','bỏ lỡ','0,272','43 tuổi, management; tháng 8; campaign = 6','Chưa có lịch sử; nhạy với thời điểm'],
 ['dong:40923','gọi nhầm','0,915','44 tuổi, self-employed; success; pdays = 100','Lịch sử thuận lợi, nhãn hiện tại ngược lại'],
 ['dong:39795','gọi nhầm','0,849','55 tuổi, technician; success; pdays = 91','Lịch sử thuận lợi, nhãn hiện tại ngược lại'],
 ['dong:41544','gọi nhầm','0,796','72 tuổi, retired; số dư 17.739; tháng 9','Giống nhóm thường đồng ý nhưng từ chối']],[2.5,2.0,1.5,5.6,4.4])
p('Không trường hợp nào trong mẫu có xác suất sát ngưỡng 0,6; trường hợp gần nhất là dong:18051 vẫn thấp hơn ngưỡng khoảng 0,199. Vì vậy, không gán nhóm lỗi “xác suất sát ngưỡng” cho sáu khách này.')

heading('6.3. Phân tích chi tiết từng lỗi')
note('Thay toàn bộ các dòng mẫu từ “Lỗi 1 (bỏ lỡ): mã dòng …” đến hết phần “Đề xuất” của Lỗi 6 bằng nội dung dưới đây. Giữ đủ đặc điểm, nhầm lẫn, bằng chứng, giả thuyết và đề xuất của từng khách. Có thể dùng tên lỗi in đậm như trong tài liệu này; không đánh lại số các mục lớn.')
label()
p('Với E5, hàm giai_thich thay riêng từng thuộc tính bằng giá trị đại diện trong train: giá trị xuất hiện nhiều nhất cho thuộc tính phân loại và trung vị cho thuộc tính số. Các mức thay đổi dưới đây được làm tròn theo đầu ra notebook, không cộng thành tổng đóng góp và không được diễn giải thành quan hệ nhân quả. Các đề xuất là hướng kiểm chứng trong lần phát triển sau, chưa được áp dụng lên test hiện tại.')

errors=[
('Lỗi 1 (bỏ lỡ): dong:11712, p_yes = 0,130',[
 'Đặc điểm và nhầm lẫn: Khách 60 tuổi, đã nghỉ hưu, số dư 1.690, liên lạc tháng 6 với campaign = 4; contact và poutcome đều unknown, pdays = −1. Khách thực tế đồng ý nhưng mô hình dự đoán no.',
 'Bằng chứng: Thay contact từ unknown thành cellular làm xác suất tăng khoảng 0,580; thay job từ retired thành blue-collar tăng 0,060; thay day từ 20 thành 16 tăng 0,040. Dự đoán đặc biệt nhạy với kênh liên lạc, phù hợp với Recall thấp của nhóm contact = unknown tại Bảng 19.',
 'Giả thuyết và đề xuất: Giá trị unknown có thể góp phần làm mô hình đánh giá thấp khách. Cần kiểm tra hồ sơ gốc để bổ sung kênh liên lạc nếu có thông tin xác thực, rồi so sánh cách xử lý thông tin thiếu trên validation. Không tự thay unknown thành cellular chỉ nhằm tăng xác suất.'
]),
('Lỗi 2 (bỏ lỡ): dong:18051, p_yes = 0,401',[
 'Đặc điểm và nhầm lẫn: Khách 41 tuổi, nghề blue-collar, số dư 0, liên lạc qua cellular vào tháng 7; pdays = −1 và poutcome = unknown. Mô hình dự đoán no dù khách thực tế đồng ý.',
 'Bằng chứng: Thay marital từ divorced thành married làm xác suất giảm khoảng 0,150; thay day từ 30 thành 16 tăng 0,080; thay education từ primary thành secondary giảm 0,040. Các phép thay cho thấy phản ứng với nhiều thuộc tính, chưa xác định được một nguyên nhân duy nhất của lỗi.',
 'Giả thuyết và đề xuất: Thông tin hiện có có thể chưa phản ánh đủ nhu cầu gửi tiết kiệm của khách. Có thể phân tích thêm các khách chưa có lịch sử trên train và validation; nếu thu thập dữ liệu mới, bổ sung nhu cầu và mức độ quan tâm gần đây rồi kiểm tra giá trị dự báo trên validation.'
]),
('Lỗi 3 (bỏ lỡ): dong:21621, p_yes = 0,272',[
 'Đặc điểm và nhầm lẫn: Khách 43 tuổi, nghề management, số dư 483, liên lạc qua cellular vào tháng 8, campaign = 6; pdays = −1 và poutcome = unknown. Khách thực tế đồng ý nhưng mô hình dự đoán no.',
 'Bằng chứng: Thay month từ aug thành may làm xác suất tăng khoảng 0,280; thay day từ 19 thành 16 tăng 0,110; thay campaign từ 6 thành 2 tăng 0,050.',
 'Giả thuyết và đề xuất: Tổ hợp thời điểm và số lần liên lạc có thể khiến mô hình đánh giá thấp trường hợp này. Đề xuất nghiên cứu cách biểu diễn ngày, tháng và tương tác với số lần liên lạc, rồi đối chiếu lỗi theo nhóm và F1-yes trên validation. Mức thay đổi cục bộ chưa chứng minh đổi tháng gọi sẽ làm khách đồng ý.'
]),
('Lỗi 4 (gọi nhầm): dong:40923, p_yes = 0,915',[
 'Đặc điểm và nhầm lẫn: Khách 44 tuổi, nghề self-employed, số dư 163, liên lạc qua cellular tháng 8; poutcome = success, pdays = 100 và previous = 3. Mô hình dự đoán yes với xác suất cao nhưng khách thực tế không đồng ý.',
 'Bằng chứng: Thay poutcome từ success thành unknown làm xác suất giảm khoảng 0,360; thay pdays từ 100 thành −1 giảm 0,200; thay previous từ 3 thành 0 giảm 0,170. Dự đoán nhạy với nhiều thông tin lịch sử chiến dịch.',
 'Giả thuyết và đề xuất: Mô hình có thể học xu hướng đồng ý cao của nhóm từng có chiến dịch thành công, trong khi nhu cầu hiện tại của khách đã khác. Đề xuất phân tích các trường hợp success nhưng nhãn no trên train và validation, đồng thời nghiên cứu tín hiệu về nhu cầu hiện tại nếu có dữ liệu bổ sung.'
]),
('Lỗi 5 (gọi nhầm): dong:39795, p_yes = 0,849',[
 'Đặc điểm và nhầm lẫn: Khách 55 tuổi, nghề technician, số dư 0, liên lạc qua cellular tháng 6; poutcome = success, pdays = 91 và previous = 1. Khách thực tế không đồng ý nhưng mô hình dự đoán yes.',
 'Bằng chứng: Thay poutcome từ success thành unknown làm xác suất giảm khoảng 0,260; thay pdays từ 91 thành −1 giảm 0,100; thay previous từ 1 thành 0 giảm 0,040.',
 'Giả thuyết và đề xuất: Mô hình có thể chưa phân biệt được các ngoại lệ trong nhóm từng có chiến dịch thành công. Đề xuất phân tích tương tác giữa lịch sử chiến dịch và tình trạng tài chính trên train và validation. Số dư bằng 0 là đặc điểm quan sát được, chưa đủ để khẳng định đó là nguyên nhân khách từ chối.'
]),
('Lỗi 6 (gọi nhầm): dong:41544, p_yes = 0,796',[
 'Đặc điểm và nhầm lẫn: Khách 72 tuổi, đã nghỉ hưu, số dư 17.739, liên lạc qua cellular tháng 9; pdays = −1, poutcome = unknown và campaign = 4. Mô hình dự đoán yes nhưng khách thực tế không đồng ý.',
 'Bằng chứng: Thay job từ retired thành blue-collar làm xác suất giảm khoảng 0,120; thay age từ 72 thành 39 giảm 0,100; thay month từ sep thành may giảm 0,070.',
 'Giả thuyết và đề xuất: Khách có thể mang các đặc điểm thường gắn với nhóm đồng ý nhưng có nhu cầu cá nhân khác xu hướng nhóm. EDA tuần 1 ghi nhận tỉ lệ đồng ý cao ở nhóm trên 60 tuổi, song tỉ lệ nhóm không quyết định kết quả từng người. Đề xuất phân tích lỗi trong nhóm lớn tuổi trên validation và nghiên cứu thông tin nhu cầu cá nhân để phân biệt tốt hơn.'
])]
for title,paras in errors:
    p(title,bold=True,keep=True)
    for text in paras:p(text)

heading('6.4. Tổng hợp nhóm lỗi và đề xuất')
note('Thay dấu “…” dưới Mục 6.4 bằng phần tổng hợp sau. Không cần thêm bảng mới hoặc thử thêm ngưỡng trên test.')
label()
block('''Sáu trường hợp minh họa ba nhóm vấn đề: thiếu thông tin kênh hoặc lịch sử liên hệ; dự đoán nhạy với tổ hợp thời điểm và số lần liên lạc; khách có đặc điểm giống nhóm thường đồng ý nhưng mang nhãn ngược lại. Các nhóm có thể giao nhau. Đây là mẫu được lấy từ các trường hợp sai, không đại diện cho toàn bộ test.

Thứ nhất, kiểm tra và cải thiện thông tin thiếu khi có hồ sơ gốc; so sánh cách xử lý unknown bằng Recall-yes theo nhóm và F1-yes toàn tập trên validation. Thứ hai, phân tích thêm các ngoại lệ trong nhóm từng có chiến dịch thành công và nghiên cứu biểu diễn thời điểm, số lần liên lạc cùng lịch sử; các thay đổi được học trên train rồi đánh giá trên validation. Thứ ba, theo dõi tỉ lệ được chọn gọi và số khách bị bỏ lỡ trong từng nhóm; nếu triển khai trong tương lai, cân nhắc một phần cuộc gọi khảo sát ngẫu nhiên để có thêm phản hồi từ nhóm ít được chọn.

Các đề xuất chưa được thực nghiệm trong đồ án này. Mô hình và ngưỡng đã khóa vẫn được giữ nguyên; không thử đề xuất trên test rồi báo cáo điểm mới.''')

heading('7. Kết luận',1)
heading('7.1. Trả lời các câu hỏi của đồ án')
note('Điền đủ năm dòng Bảng 21 tại Mục 7.1. Đặt đoạn mở đầu trước bảng và hai đoạn kết luận sau bảng, thay dấu “…” sẵn có. Các bằng chứng ở Bảng 10, 12–14 đã nằm trong phần tuần 1–3 của báo cáo; chỉ dẫn lại, không chép lại các phần đó.')
label()
p('Đồ án đã xây dựng và so sánh các mô hình trên dữ liệu Bank Marketing, chọn số nơ-ron, mô hình và ngưỡng trên validation rồi đánh giá mô hình đã khóa trên test. E5 được chọn tại ngưỡng 0,6 và đạt F1-yes 0,486 trên test. Bảng 21 tổng hợp các khẳng định chính và phạm vi bằng chứng.')
caption('Bảng 21. Các khẳng định chính, bằng chứng và mức chắc chắn')
table(['Khẳng định','Bằng chứng (bảng, mục)','Mức chắc chắn'],[
 ['Precision-yes trên test cao hơn mức nền của gọi ngẫu nhiên cùng số khách.','Bảng 15–16, Mục 5.1: 0,468 so với khoảng 0,117; cùng 1.146 cuộc gọi, mô hình tìm được 536 khách đồng ý, ngẫu nhiên kỳ vọng khoảng 134.','Quan sát trên test lịch sử; chưa bảo đảm chiến dịch tương lai.'],
 ['Số nơ-ron, mô hình và ngưỡng được chọn theo quy tắc định trước trên validation.','Bảng 12–14, Mục 4.3: chọn 256 nơ-ron cho E3; mô hình cuối E5, ngưỡng 0,6. Ô 4.1 tuần 4 tái tạo đúng quyết định.','Được xác nhận bằng mã và đầu ra đã lưu.'],
 ['Trên test, chọn gọi 1.146 khách thì có 536 khách thực tế đồng ý; bỏ lỡ 522 khách đồng ý.','Bảng 16 và diễn giải Mục 5.1.','Số đếm trực tiếp tại ngưỡng đã khóa.'],
 ['Bổ sung nhóm tuổi trong E4 cải thiện F1-yes validation so với E1.','Bảng 13: E4 đạt 0,458; E1 đạt 0,446 ở ngưỡng tốt nhất riêng, tăng khoảng 0,012.','Có cải thiện quan sát được; chưa khẳng định ổn định trên dữ liệu khác.'],
 ['Loại duration làm điểm thấp hơn trong phép thử E1 nhưng phù hợp dự đoán trước cuộc gọi.','Bảng 10, Mục 2.5: E1 đạt 0,375; thêm duration đạt 0,551 nhưng dùng thông tin chỉ biết sau khi gọi.','Có bằng chứng trong phép thử và căn cứ về thời điểm có thông tin.']],[5.1,6.6,4.3])
block('''Có bằng chứng cho thấy mô hình tập trung được nhóm có tỉ lệ đồng ý cao hơn mức nền trên test. Tuy nhiên, Recall-yes chỉ đạt 0,507 và vẫn bỏ lỡ 522 khách, nên chưa thể xem mô hình thay thế hoàn toàn việc tìm kiếm khách hàng tiềm năng.

E5 được chọn đúng theo quy tắc validation. Chênh lệch nhỏ giữa E5 và E3 trên validation chưa đủ để khẳng định E5 tốt hơn ổn định; đồ án không đánh giá lại các ứng viên trên test để thay đổi lựa chọn.''')

heading('7.2. Giới hạn')
note('Thay dấu “…” dưới Mục 7.2 bằng các đoạn dưới đây. Đây là nơi ghi đầy đủ giới hạn; Mục 5.3 chỉ nêu ngắn gọn hệ quả từ khách giả định và nhóm ít được chọn.')
label()
block('''Dữ liệu chỉ đến từ một ngân hàng Bồ Đào Nha trong giai đoạn 2008–2010. Ba tập được chia ngẫu nhiên từ cùng nguồn, nên kết quả chưa kiểm tra được khả năng dự đoán cho ngân hàng khác hoặc chiến dịch ở thời điểm khác. Phân tầng theo nhãn giữ tỉ lệ yes gần nhau nhưng không bảo đảm phân phối các thuộc tính đầu vào giống nhau.

Các mối liên hệ theo tuổi, tháng và kênh liên lạc là tương quan trong dữ liệu lịch sử. Đầu ra thay thuộc tính chỉ minh họa phản ứng cục bộ của mô hình; các thuộc tính có thể tương tác hoặc liên quan với nhau, nên chưa thể chuyển thành kết luận nhân quả.

Quá trình lựa chọn giới hạn ở các mô hình và cấu hình đã thử. MLP được đánh giá qua năm seed, ngưỡng được chọn trong chín giá trị từ 0,1 đến 0,9. Khoảng bootstrap chỉ phản ánh dao động do lấy mẫu cho mô hình cố định, chưa bao quát toàn bộ bất định do lựa chọn mô hình và thay đổi dữ liệu.

Khả năng phát hiện khách đồng ý khác nhau giữa các nhóm. Nhóm contact = unknown có Recall-yes chỉ 0,019 và gần như không được chọn gọi tại ngưỡng 0,6. Nếu chỉ liên hệ nhóm có xác suất cao, ngân hàng có thể bỏ qua khách tiềm năng và thu thập dữ liệu tương lai thiên lệch về nhóm đã được ưu tiên. Các kết quả hiện tại chưa đủ để kết luận toàn diện về công bằng.''')

heading('7.3. Hướng phát triển')
note('Thay dấu “…” dưới Mục 7.3 bằng hai đoạn sau. Trình bày là hướng tương lai, không mô tả như việc đã hoàn thành.')
label()
block('''Hướng tiếp theo là đánh giá trên chiến dịch mới hoặc phép chia theo thời gian để kiểm tra hiệu năng khi khách hàng khác dữ liệu huấn luyện. Việc lựa chọn cấu hình tiếp tục thực hiện trên dữ liệu phát triển, còn dữ liệu đánh giá mới được giữ độc lập. Có thể nghiên cứu thêm thông tin nhu cầu hiện tại, cải thiện dữ liệu thiếu và phân tích tương tác giữa lịch sử chiến dịch với thời điểm liên lạc; mỗi thay đổi cần được kiểm chứng bằng F1-yes toàn tập và chỉ số theo nhóm.

Nếu có điều kiện triển khai thử, cần theo dõi số cuộc gọi, tỉ lệ thành công, khách bỏ lỡ và chi phí nghiệp vụ. Một phần cuộc gọi khảo sát ngẫu nhiên có thể cung cấp thêm phản hồi từ nhóm ít được chọn, hỗ trợ phát hiện thiên lệch và cập nhật mô hình.''')

heading('7.4. Bài học rút ra')
note('Thay dấu “…” dưới Mục 7.4 bằng một đoạn sau; giữ phần bài học ngắn gọn theo mẫu.')
label()
p('Bài học chính là duy trì đúng vai trò của từng tập dữ liệu và đánh giá tương ứng với thời điểm sử dụng mô hình. Duration có thể làm điểm cao hơn nhưng không phù hợp với dự đoán trước cuộc gọi; mô hình và ngưỡng cần được chọn trên validation rồi khóa trước test. Phân tích lỗi phải tách số liệu quan sát được khỏi giả thuyết giải thích. Nếu làm lại đồ án, việc ghi phiên bản thư viện, lưu cấu hình và kết quả trung gian sẽ được thực hiện ngay từ đầu để giảm khó khăn khi tái tạo; đồng thời theo dõi lỗi theo nhóm sớm hơn trong quá trình phát triển.')

heading('Bổ sung các mục cuối báo cáo',1)
heading('Tài liệu tham khảo')
note('Tại mục “Tài liệu tham khảo”, giữ nguyên hai nguồn bắt buộc [1] và [2]. Thay dòng “[3] …” bằng nguồn SciPy dưới đây; nếu chưa có nguồn notebook, thêm dòng [4]. Kiểm tra lại số thứ tự nếu bạn đã bổ sung tài liệu khác.')
label()
p('[3] SciPy. scipy.stats.bootstrap. Tài liệu tham khảo thư viện SciPy. https://docs.scipy.org/doc/scipy/reference/generated/scipy.stats.bootstrap.html. Truy cập ngày 21/11/2026.')
p('[4] Tài liệu học phần Máy học do giảng viên cung cấp. Tuần 4 — Đánh giá cuối trên test, phân tích lỗi, hoàn thiện báo cáo. Notebook thực hành đồ án Bank Marketing.')

heading('Khai báo sử dụng công cụ hỗ trợ')
note('Điền đoạn sau vào phần trống dưới tiêu đề “Khai báo sử dụng công cụ hỗ trợ”. Đọc lại và điều chỉnh nếu bạn dùng thêm công cụ hoặc có phạm vi sử dụng khác với mô tả.')
label()
p('ChatGPT của OpenAI, thông qua Codex, được sử dụng để giải thích khái niệm, hỗ trợ hoàn thiện các ô mã, chẩn đoán lỗi môi trường, đối chiếu đầu ra notebook và hỗ trợ soạn, chỉnh diễn đạt báo cáo. Các phần liên quan gồm thực hành notebook và trình bày kết quả; riêng tuần 4 gồm tái tạo mô hình đã khóa, diễn giải đánh giá test và phân tích lỗi. Người thực hiện trực tiếp chạy notebook, kiểm tra kết quả và chịu trách nhiệm về toàn bộ nội dung báo cáo.')

heading('Phụ lục A. Nhật ký thực hiện')
note('Trong Bảng A.1 đang có, chỉ cập nhật dòng tuần 4; giữ các dòng tuần 1–3 để rà soát khi chốt nhật ký. Ô thời gian dưới đây cố ý để trống, cần điền số giờ thực tế. Việc hoàn thiện báo cáo và xuất tệp chỉ ghi nhận sau khi thực sự làm xong.')
label('Nội dung để điền vào dòng tuần 4 của Bảng A.1')
table(['Tuần','Việc đã làm','Thời gian (giờ)','Khó khăn và cách giải quyết'],[
 ['4','Tái tạo E5 đã khóa; đánh giá test ở hai ngưỡng; so sánh validation–test; phân tích theo nhóm và sáu lỗi với seed 226; bootstrap; thử khách giả định; rà soát notebook và chuẩn bị nội dung báo cáo.','','Gặp lỗi kết nối Colab trong VS Code, xử lý bằng kết nối lại. F1 tái tạo ban đầu 0,4702 khi dùng scikit-learn 1.6.1; đồng bộ 1.9.1 với tuần 3, khởi động lại kernel, kết quả khớp 0,4806.']],[1.2,6.3,2.0,6.5])

heading('Phụ lục B. Đầu ra then chốt của notebook')
note('Tìm chú thích “Hình B.4. Tuần 4: ô 4.1 xác nhận đúng mô hình đã khóa, bảng kết quả test và ma trận nhầm lẫn”. Thay ô “[Chèn hình]” ngay phía trên bằng hình ở trang kế tiếp. Giữ chú thích bên dưới ảnh; không chèn ảnh vào Mục 4.1 của báo cáo vì mục đó thuộc tuần 3.')
p('Hình dưới được bố trí lại từ các đầu ra đã lưu trong notebook: xác nhận mục 4.1, bảng test mục 4.2 và ảnh ma trận nhầm lẫn gốc. Các giá trị được giữ nguyên. Khi sử dụng ảnh này, có thể sao chép ảnh trực tiếp trong Word; không cần chụp mã hoặc chạy lại mô hình.',size=11,italic=True)
doc.add_page_break()
pic=doc.add_paragraph();pic.alignment=WD_ALIGN_PARAGRAPH.CENTER
pic.add_run().add_picture(str(BUILD/'Hinh_B4.png'),width=Cm(15.7))
pic.paragraph_format.space_after=Pt(4)
cap=p('Hình B.4. Tuần 4: ô 4.1 xác nhận đúng mô hình đã khóa,\nbảng kết quả test và ma trận nhầm lẫn',size=11,italic=True)
cap.alignment=WD_ALIGN_PARAGRAPH.CENTER;cap.paragraph_format.line_spacing=1.0

heading('Phụ lục C. Trả lời các câu hỏi suy ngẫm')
note('Ở cuối Phụ lục C, tìm đúng “Tuần 4, câu 1”. Giữ câu hỏi và thay dòng “Trả lời: …” bằng câu trả lời dưới đây. Các câu tuần 1–3 đã có trong báo cáo không thuộc phần bổ sung lần này.')
label()
p('Tuần 4, câu 1. Nếu F1-yes trên test thấp hơn validation 0,01, có nên quay lại tuần 3 để chọn mô hình khác không?',bold=True,keep=True)
p('Trả lời: Không nên chọn lại mô hình dựa trên test vì sẽ làm mất tính độc lập của phép đánh giá. Trong kết quả thực tế, F1-yes test đạt 0,486, cao hơn validation khoảng 0,006; dù chênh lệch theo hướng nào, vẫn giữ E5 và ngưỡng 0,6 đã khóa rồi báo cáo trung thực.')

heading('Kiểm tra sau khi đưa nội dung vào báo cáo',1)
for text in [
 'Điền đủ Bảng 15–21 tại các vị trí sẵn có; giữ tên bảng theo mẫu, chú thích bảng ở trên và chú thích hình ở dưới.',
 'Kiểm tra τ* = 0,6; E5; TP = 536, FN = 522, FP = 610, TN = 7.375; F1-yes test = 0,486.',
 'Giữ nguyên sáu mã dòng dong:11712, dong:18051, dong:21621, dong:40923, dong:39795 và dong:41544; các mức thay đổi xác suất là xấp xỉ theo đầu ra đã làm tròn.',
 'Chèn Hình B.4 tại Phụ lục B; không đánh số lại các hình tuần 1–3. Không cần thêm số hình mới cho ma trận trong Mục 5 vì Bảng 16 đã thể hiện số đếm.',
 'Điền số giờ thực tế và rà soát nhật ký Phụ lục A khi chốt. Đọc lại khai báo công cụ để phản ánh đúng việc đã sử dụng.',
 'Sau khi nhập xong, xóa các hộp HƯỚNG DẪN và dòng mẫu còn dư của phần đã hoàn thiện; không chép các chỉ dẫn thao tác trong tài liệu này vào bản nộp.',
 'Cập nhật mục lục và kiểm tra toàn bộ báo cáo đạt tối thiểu 20 trang từ Mục 1 đến hết phụ lục, đúng quy cách và không độn trang. Tài liệu hướng dẫn này không dùng để tính số trang báo cáo.',
 'Khi hoàn tất, xuất báo cáo thành 3124411226_ML_baocao.pdf; tên notebook tuần 4 khi nộp là 3124411226_tuan4.ipynb. Tệp ZIP cuối cùng là 3124411226_ML.zip theo yêu cầu của đề.'
]:
    p('□ '+text)

# Remove title decoration inherited through paragraph properties if present.
for par in doc.paragraphs:
    if par.style.name=='Title':
        pp=par._p.get_or_add_pPr()
        for x in list(pp):
            if x.tag in [qn('w:pBdr'),qn('w:shd')]:pp.remove(x)

doc.core_properties.title='Hướng dẫn bổ sung báo cáo tuần 4 Bank Marketing'
doc.core_properties.subject='Huỳnh Hoàng Phong — 3124411226 — E5 — ngưỡng 0,6'
doc.core_properties.author='Huỳnh Hoàng Phong'
doc.save(OUT)
assert hashlib.sha256(REF.read_bytes()).hexdigest()==reference_hash

contract=f'''# Week 4 document contract
Reference: {REF}
SHA256: {reference_hash}
Reference retained at {BUILD / REF.name}; unchanged.
Source evidence: current DOCX, Docs/template.docx, saved notebooks weeks 1–4.
Reference rendering: {BUILD / 'reference_render'} via Microsoft Word PDF export,
then packaged render_docx.py and bundled Poppler. Windows runtime has no LibreOffice.
Output: {OUT}

Page system: A4 portrait, 21 x 29.7 cm; left 3 cm; right/top/bottom 2 cm.
The original has two sections. The standalone guide intentionally uses one section.
Title: existing Title style, Times New Roman 22 pt bold black.
Heading 1: source Heading1, 15 pt bold black. Heading 2: source Heading2, 13.5 pt bold black.
Body: Times New Roman 13 pt, 1.5 spacing. Tables/captions and placement guidance: 11 pt.
Source style definitions and document package components are retained by starting from the reference.
Output body is replaced completely because user requested a separate guide only for week 4.
Existing report body content is not included. Table horizontal rules follow the restrained source style.
Tables have deliberate 16 cm total widths, repeated headers, no fixed heights, no split rows.
No reference content is edited in place. Heading numbers and table/figure names follow teacher template.

Slot map: report section 5.1 tables 15/16; 5.2 table 17; 5.3 table 18;
5.4 bootstrap; 6.1 table 19; 6.2 table 20; 6.3 six detailed cases; 6.4 proposals;
7.1 table 21; 7.2–7.4 prose; references; tool declaration; appendix A week-4 row;
appendix B placeholder immediately before Figure B.4 caption; appendix C week-4 answer.
Locations use heading and caption text, not guessed report page numbers.
Journal hours are intentionally blank, awaiting the user's real elapsed time.
All known notebook values must remain exact at their displayed precision.
Figure B.4 is explicitly described as recomposed saved output, including the original plot PNG.
QA: inspect every output page PNG; check numeric tables, six cases, layout and source hash.
'''
(BUILD/'artifact.md').write_text(contract,encoding='utf-8')
with ZipFile(REF) as z:
    inventory={i.filename:{'size':i.file_size,'sha256':hashlib.sha256(z.read(i.filename)).hexdigest()} for i in z.infolist()}
(BUILD/'reference_inventory.json').write_text(json.dumps(inventory,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'output':str(OUT),'paragraphs':len(doc.paragraphs),'tables':len(doc.tables),'reference_sha256':reference_hash},ensure_ascii=False))
