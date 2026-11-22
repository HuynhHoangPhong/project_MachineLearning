from pathlib import Path
from copy import deepcopy
import shutil
import sys
import json
from docx import Document
from docx.shared import Pt

sys.stdout.reconfigure(encoding='utf-8')
ROOT=Path(r'C:\code\MachineLearning\Project_FinalExam')
BUILD=ROOT/'reports/_build_week4_3124411226'
OUT=ROOT/'reports/3124411226_Huong_dan_bo_sung_bao_cao_week4.docx'
backup=BUILD/'guide_before_length_review.docx'
if not backup.exists():shutil.copy2(OUT,backup)
d=Document(OUT)
ops=[]

def replace(prefix,new):
    found=[p for p in d.paragraphs if p.text.startswith(prefix)]
    assert len(found)==1,(prefix,len(found))
    p=found[0]
    ops.append({'old':p.text,'new':new})
    if new is None:
        p._p.getparent().remove(p._p)
        return
    # Preserve the paragraph and first run properties of the current guide.
    if p.runs:
        p.runs[0].text=new
        for r in p.runs[1:]:r.text=''
    else:p.add_run(new)

replacements={
 'Mô hình được đánh giá là rừng ngẫu nhiên E5': 'Mô hình đánh giá là rừng ngẫu nhiên E5 với 200 cây, min_samples_leaf=5, class_weight="balanced", random_state=42 và ngưỡng đã khóa τ* = 0,6. Ô 4.1 tuần 4 tái tạo đúng F1-yes validation trong quyết định tuần 3.',
 'Bảng 15 trình bày kết quả tại ngưỡng đã khóa': 'Bảng 15 đối chiếu ngưỡng đã khóa với mốc 0,5 theo đề; kết quả test không được dùng để chọn lại mô hình hoặc ngưỡng.',
 'Nếu chỉ gọi những khách được mô hình dự đoán': 'Tại ngưỡng 0,6, mô hình đề xuất gọi 1.146/9.043 khách, khoảng 12,7% danh sách. Trong đó, 536 khách thực tế đồng ý và 610 khách không đồng ý, đạt tỉ lệ thành công khoảng 46,8% trên test.',
 'Để so sánh công bằng, phương án gọi ngẫu nhiên': 'Với cùng 1.146 cuộc gọi, gọi ngẫu nhiên có số khách đồng ý kỳ vọng là 1.146 × (1.058/9.043) ≈ 134 khách. Tỉ lệ đồng ý của nhóm do mô hình chọn cao khoảng 4 lần mức nền 11,7%. So sánh này dựa vào Precision-yes và cùng số cuộc gọi, không chỉ dựa vào F1-yes.',
 'Mô hình phát hiện 536/1.058 khách đồng ý': 'Mô hình tìm được 536/1.058 khách đồng ý, đạt Recall-yes khoảng 50,7%, và bỏ lỡ 522 khách. Tăng ngưỡng từ 0,5 lên 0,6 làm Precision-yes tăng từ 0,380 lên 0,468 nhưng Recall-yes giảm từ 0,595 xuống 0,507: danh sách gọi tập trung hơn nhưng mất thêm cơ hội. Các tỉ lệ lịch sử này chưa bảo đảm kết quả của chiến dịch tương lai.',
 'Với cùng mô hình E5 và ngưỡng 0,6, F1-yes trên test': 'Với cùng E5 và ngưỡng 0,6, F1-yes test cao hơn validation khoảng 0,006 và Precision-yes cao hơn khoảng 0,013, tính từ số chưa làm tròn. Recall-yes giảm khoảng 0,003, tương ứng phát hiện 536 thay vì 539 khách trong tổng 1.058 khách yes của mỗi tập. Chênh lệch nhỏ, phù hợp với dao động ngẫu nhiên; lựa chọn trên validation có thể tạo sự lạc quan nhưng không khiến điểm test luôn thấp hơn. Có thể báo cáo điểm test như đánh giá độc lập cùng nguồn, song phép chia ngẫu nhiên dữ liệu 2008–2010 chưa kiểm chứng được chiến dịch tương lai.',
 'Nếu F1-yes trên test thấp hơn validation 0,01, cũng': 'Nếu F1-yes test thấp hơn validation 0,01, vẫn giữ quyết định đã khóa và báo cáo trung thực; chọn lại dựa trên test sẽ làm mất tính độc lập của phép đánh giá.',
 'Các dự đoán nhìn chung phù hợp với EDA tuần 1. Khách': 'Dự đoán phù hợp với xu hướng EDA: khách 22 tuổi, tháng 3 đạt p_yes = 0,851; khách nghỉ hưu 67 tuổi, lịch sử success đạt 0,951. Khách 45 tuổi có contact = unknown, tháng 5 chỉ đạt 0,103, phù hợp với tỉ lệ đồng ý thấp của các nhóm này.',
 'Khách kỹ thuật viên 35 tuổi đạt 0,431': 'Khách kỹ thuật viên 35 tuổi đạt 0,431 và được dự đoán no; nhóm tuổi 30–40 và tháng 8 cũng có tỉ lệ đồng ý thấp hơn mức chung trong train. Cả bốn khách giữ nhãn ở hai ngưỡng; đây chỉ là đối chiếu xu hướng, chưa đánh giá đúng/sai vì không có nhãn thật. Khách 22 tuổi vẫn được chọn dù poutcome = unknown, nên thiếu lịch sử không đồng nghĩa với luôn bị loại.',
 'Nếu chỉ gọi khách có xác suất cao, một số nhóm': 'Ngược lại, tại ngưỡng 0,6 chỉ 4/2.634 khách contact = unknown được chọn, khoảng 0,15%, suy ra từ Bảng 19. Nếu duy trì cách chọn này, nhóm có thể ít có cơ hội tiếp cận sản phẩm và ít đóng góp phản hồi cho dữ liệu sau. Giới hạn về khả năng tiếp cận và thiên lệch dữ liệu được bàn rõ ở Mục 7.2.',
 'F1-yes trên test là 0,486 (khoảng tin cậy': 'F1-yes trên test là 0,486 (khoảng tin cậy 95%: [0,460; 0,512], bootstrap, B = 1.000). Notebook lấy mẫu có hoàn lại theo cặp nhãn thật–dự đoán, dùng phương pháp phân vị và random_state=42.',
 'Khoảng có độ rộng 0,052, tương đối rộng': 'Độ rộng 0,052 tương đối lớn so với chênh lệch F1 vài phần nghìn, nên chưa xem các chênh lệch nhỏ là chắc chắn chỉ từ điểm F1. Khoảng này không thay thế phép so sánh hai mô hình và chỉ phản ánh dao động lấy mẫu cùng nguồn, chưa đánh giá thay đổi phân phối trong tương lai.',
 'Nhóm contact = unknown có Recall-yes thấp nhất': 'Nhóm contact = unknown có Recall-yes thấp nhất là 0,019: chỉ tìm được 2/103 khách yes, bỏ lỡ 101 khách. Giả thuyết là mô hình học xu hướng ít đồng ý của nhóm trong train (EDA: khoảng 4,3%), nên nhiều khách yes vẫn dưới ngưỡng.',
 'Recall thấp nhất thể hiện tỉ lệ bỏ lỡ cao nhất': 'Đây là tỉ lệ bỏ lỡ cao nhất, không phải số khách bỏ lỡ lớn nhất: nhóm cellular bỏ lỡ 397 khách do có 888 khách yes.',
 'Theo poutcome, nhóm unknown có Recall-yes thấp nhất': 'Theo poutcome, nhóm unknown có Recall-yes thấp nhất là 0,342, với 657 khách yes và 432 khách bị bỏ lỡ. Nhóm success tìm được toàn bộ khách yes nhưng Precision-yes chỉ 0,650, nên vẫn có gọi nhầm.',
 'Trên toàn bộ test, mô hình bỏ lỡ 522 khách': 'Test có 522 khách bị bỏ lỡ và 610 khách bị gọi nhầm. Lấy ngẫu nhiên ba khách mỗi loại bằng random_state=S_SV, với S_SV=226 từ MSSV 3124411226; sáu mã dòng đều khác nhau.',
 'Không trường hợp nào trong mẫu có xác suất sát': 'Không khách nào trong mẫu có xác suất sát 0,6; gần nhất là dong:18051 vẫn cách khoảng 0,199. Vì vậy, không gán nhóm “sát ngưỡng” cho mẫu này.',
 'Với E5, hàm giai_thich thay riêng từng thuộc tính': 'Với E5, giai_thich thay riêng từng thuộc tính bằng mode của thuộc tính phân loại hoặc trung vị của thuộc tính số trong train. Các mức thay đổi dưới đây là xấp xỉ theo đầu ra notebook, không cộng thành tổng đóng góp và không thể diễn giải nhân quả.',
 'Đặc điểm và nhầm lẫn: Khách 60 tuổi': 'Đặc điểm và nhầm lẫn: Khách 60 tuổi, retired; contact và poutcome unknown, pdays = −1. Thực tế yes nhưng được dự đoán no.',
 'Bằng chứng: Thay contact từ unknown': 'Bằng chứng: Thay contact unknown → cellular làm xác suất tăng khoảng 0,580; job retired → blue-collar tăng 0,060; day 20 → 16 tăng 0,040. Dự đoán nhạy nhất với kênh liên lạc trong các phép thay.',
 'Giả thuyết và đề xuất: Giá trị unknown': 'Giả thuyết và đề xuất: Kênh unknown có thể góp phần làm khách bị đánh giá thấp, phù hợp Bảng 19. Kiểm tra hồ sơ để bổ sung kênh nếu có căn cứ; so sánh cách xử lý thiếu trên validation, không tự đổi kênh để tăng xác suất.',
 'Đặc điểm và nhầm lẫn: Khách 41 tuổi': 'Đặc điểm và nhầm lẫn: Khách 41 tuổi, blue-collar, số dư 0, cellular; pdays = −1 và poutcome unknown. Thực tế yes nhưng được dự đoán no.',
 'Bằng chứng: Thay marital từ divorced': 'Bằng chứng: Thay marital divorced → married làm xác suất giảm khoảng 0,150; day 30 → 16 tăng 0,080; education primary → secondary giảm 0,040. Chưa xác định được một nguyên nhân duy nhất từ các phép thay.',
 'Giả thuyết và đề xuất: Thông tin hiện có': 'Giả thuyết và đề xuất: Thiếu lịch sử có thể khiến thông tin hiện có chưa phản ánh đủ nhu cầu. Phân tích thêm nhóm chưa có lịch sử trên validation; nếu có dữ liệu nhu cầu mới, kiểm tra giá trị dự báo trước khi sử dụng.',
 'Đặc điểm và nhầm lẫn: Khách 43 tuổi': 'Đặc điểm và nhầm lẫn: Khách 43 tuổi, management, cellular, tháng 8, campaign = 6; chưa có lịch sử chiến dịch trước. Thực tế yes nhưng được dự đoán no.',
 'Bằng chứng: Thay month từ aug': 'Bằng chứng: Thay month aug → may làm xác suất tăng khoảng 0,280; day 19 → 16 tăng 0,110; campaign 6 → 2 tăng 0,050.',
 'Giả thuyết và đề xuất: Tổ hợp thời điểm': 'Giả thuyết và đề xuất: Tổ hợp thời điểm và số lần gọi có thể làm dự đoán thấp. Thử biểu diễn ngày, tháng và tương tác với campaign trên train, rồi so sánh lỗi nhóm và F1 trên validation. Phép thay chưa chứng minh đổi tháng gọi sẽ làm khách đồng ý.',
 'Đặc điểm và nhầm lẫn: Khách 44 tuổi': 'Đặc điểm và nhầm lẫn: Khách 44 tuổi, self-employed; poutcome success, pdays = 100, previous = 3. Thực tế no nhưng được dự đoán yes với xác suất cao.',
 'Bằng chứng: Thay poutcome từ success thành unknown làm xác suất giảm khoảng 0,360': 'Bằng chứng: Thay poutcome success → unknown làm xác suất giảm khoảng 0,360; pdays 100 → −1 giảm 0,200; previous 3 → 0 giảm 0,170. Các tín hiệu lịch sử ảnh hưởng rõ đến dự đoán cục bộ.',
 'Giả thuyết và đề xuất: Mô hình có thể học xu hướng': 'Giả thuyết và đề xuất: Lịch sử thành công có thể làm mô hình đánh giá cao dù nhu cầu hiện tại khác trước. Phân tích các ngoại lệ success nhưng nhãn no trên train/validation; nghiên cứu tín hiệu nhu cầu cập nhật nếu có dữ liệu.',
 'Đặc điểm và nhầm lẫn: Khách 55 tuổi': 'Đặc điểm và nhầm lẫn: Khách 55 tuổi, technician, số dư 0; poutcome success, pdays = 91, previous = 1. Thực tế no nhưng được dự đoán yes.',
 'Bằng chứng: Thay poutcome từ success thành unknown làm xác suất giảm khoảng 0,260': 'Bằng chứng: Thay poutcome success → unknown làm xác suất giảm khoảng 0,260; pdays 91 → −1 giảm 0,100; previous 1 → 0 giảm 0,040.',
 'Giả thuyết và đề xuất: Mô hình có thể chưa phân biệt': 'Giả thuyết và đề xuất: Mô hình có thể chưa phân biệt ngoại lệ trong nhóm từng thành công. Kiểm tra tương tác lịch sử và tình trạng tài chính trên validation. Số dư 0 là đặc điểm quan sát, chưa đủ kết luận đó là nguyên nhân từ chối.',
 'Đặc điểm và nhầm lẫn: Khách 72 tuổi': 'Đặc điểm và nhầm lẫn: Khách 72 tuổi, retired, số dư 17.739, tháng 9, chưa có lịch sử chiến dịch trước. Thực tế no nhưng được dự đoán yes.',
 'Bằng chứng: Thay job từ retired': 'Bằng chứng: Thay job retired → blue-collar làm xác suất giảm khoảng 0,120; age 72 → 39 giảm 0,100; month sep → may giảm 0,070.',
 'Giả thuyết và đề xuất: Khách có thể mang': 'Giả thuyết và đề xuất: Khách giống nhóm thường đồng ý nhưng nhu cầu cá nhân có thể khác. EDA ghi nhận nhóm trên 60 tuổi đồng ý cao, không bảo đảm từng người. Kiểm tra ngoại lệ nhóm lớn tuổi trên validation và giá trị của thông tin nhu cầu bổ sung.',
 'Sáu trường hợp minh họa ba nhóm vấn đề': 'Sáu lỗi gợi ý ba nhóm vấn đề: thiếu thông tin hoặc lịch sử; nhạy với thời điểm và số lần gọi; đặc điểm thường thuận lợi nhưng nhãn ngược lại. Nhóm có thể giao nhau; mẫu lỗi không đại diện toàn bộ test.',
 'Thứ nhất, kiểm tra và cải thiện thông tin thiếu': 'Ba hướng kiểm chứng: (1) kiểm tra dữ liệu thiếu bằng hồ sơ gốc, so sánh Recall nhóm và F1 trên validation; (2) thử biểu diễn thời điểm và tương tác với lịch sử trên train, đánh giá trên validation; (3) theo dõi tỉ lệ chọn gọi theo nhóm và cân nhắc khảo sát ngẫu nhiên khi triển khai để bổ sung phản hồi.',
 'Các đề xuất chưa được thực nghiệm trong đồ án': 'Đây là đề xuất cho lần phát triển sau; giữ nguyên mô hình đã khóa và không thử trên test hiện tại.',
 'Đồ án đã xây dựng và so sánh các mô hình': 'Đồ án chọn E5 và ngưỡng 0,6 trên validation, sau đó đạt F1-yes 0,486 trên test. Bảng 21 tổng hợp các kết luận và phạm vi bằng chứng.',
 'Có bằng chứng cho thấy mô hình tập trung được': 'Mô hình tập trung được nhóm có tỉ lệ đồng ý cao nhưng vẫn bỏ lỡ 522 khách. E5 được chọn đúng quy tắc; chênh lệch nhỏ với E3 trên validation chưa đủ khẳng định ưu thế ổn định và không được dùng để chọn lại trên test.',
 'E5 được chọn đúng theo quy tắc validation.': None,
 'Dữ liệu chỉ đến từ một ngân hàng Bồ Đào Nha': 'Dữ liệu của một ngân hàng Bồ Đào Nha giai đoạn 2008–2010 được chia ngẫu nhiên, nên chưa kiểm chứng ngân hàng khác hoặc chiến dịch tương lai. Phân tầng chỉ giữ tỉ lệ nhãn gần nhau. Tương quan theo tuổi, tháng, kênh và phép thay thuộc tính không chứng minh nhân quả.',
 'Các mối liên hệ theo tuổi, tháng và kênh liên lạc': None,
 'Quá trình lựa chọn giới hạn ở các mô hình': 'Phạm vi thử nghiệm chỉ gồm các cấu hình đã xét, năm seed MLP và chín ngưỡng; bootstrap chưa bao quát bất định do chọn mô hình hoặc đổi phân phối. Recall của nhóm contact = unknown chỉ 0,019; ít liên hệ nhóm này có thể giảm cơ hội tiếp cận và làm dữ liệu sau thiên lệch. Chưa đủ bằng chứng kết luận toàn diện về công bằng.',
 'Khả năng phát hiện khách đồng ý khác nhau giữa': None,
 'Hướng tiếp theo là đánh giá trên chiến dịch mới': 'Ưu tiên đánh giá theo thời gian hoặc chiến dịch mới với tập đánh giá độc lập; cải thiện thông tin thiếu và nghiên cứu nhu cầu hiện tại trên train/validation. Nếu triển khai thử, theo dõi F1, chi phí, tỉ lệ được gọi và khách bỏ lỡ theo nhóm; cân nhắc một phần liên hệ ngẫu nhiên để thu phản hồi từ nhóm ít được chọn.',
 'Nếu có điều kiện triển khai thử, cần theo dõi': None,
 'Bài học chính là duy trì đúng vai trò': 'Bài học chính là loại duration để phù hợp thời điểm dự đoán, chọn trên validation và khóa trước test; đồng thời phân biệt bằng chứng với giả thuyết khi phân tích lỗi. Nếu làm lại, sẽ ghi phiên bản thư viện và cấu hình ngay từ đầu, lưu kết quả trung gian để tái tạo, và theo dõi lỗi theo nhóm sớm hơn.',
}
for a,b in replacements.items():replace(a,b)

# Keep the teacher's five claims, while reducing repetition in the evidence table.
values=[
 ['Precision-yes test cao hơn mức nền gọi ngẫu nhiên cùng số khách.','Bảng 15–16: 0,468 so với khoảng 0,117; 1.146 cuộc gọi tìm được 536 khách, ngẫu nhiên kỳ vọng 134.','Quan sát trên test; chưa bảo đảm chiến dịch mới.'],
 ['Chọn số nơ-ron, mô hình và ngưỡng theo quy tắc trên validation.','Bảng 12–14, Mục 4.3: E3 chọn 256 nơ-ron; mô hình cuối E5, τ* = 0,6; tuần 4 tái tạo đúng.','Xác nhận bằng mã và đầu ra.'],
 ['Chọn gọi 1.146 khách, có 536 khách đồng ý và bỏ lỡ 522 khách.','Bảng 16, Mục 5.1.','Số đếm trên test tại τ*.'],
 ['E4 thêm nhóm tuổi cải thiện so với E1 trên validation.','Bảng 13: 0,458 so với 0,446, tăng khoảng 0,012 ở ngưỡng tốt nhất riêng.','Có cải thiện quan sát; chưa khẳng định ổn định.'],
 ['Loại duration giảm điểm E1 nhưng phù hợp dự đoán trước gọi.','Bảng 10, Mục 2.5: 0,375 so với 0,551 khi thêm thông tin chỉ biết sau gọi.','Có bằng chứng phép thử và căn cứ thời điểm thông tin.'],
]
for row,newrow in zip(d.tables[6].rows[1:],values):
    for cell,text in zip(row.cells,newrow):
        para=cell.paragraphs[0]
        if para.runs:
            para.runs[0].text=text
            for r in para.runs[1:]:r.text=''
        else:para.add_run(text)

# Update placement instructions to agree with the revised number of paragraphs.
replace('Hướng dẫn đặt nội dung: Tại Mục 5.1,','Hướng dẫn đặt nội dung: Tại Mục 5.1, điền Bảng 15 và Bảng 16 vào các ô trống sẵn có. Đặt hai đoạn giới thiệu trước Bảng 15; thay dấu “…” sau Bảng 16 bằng ba đoạn diễn giải dưới đây. Ảnh ma trận đặt ở Hình B.4 trong phụ lục, không cần lặp lại tại Mục 5.1.')
replace('Hướng dẫn đặt nội dung: Thay toàn bộ các dòng mẫu','Hướng dẫn đặt nội dung: Thay các dòng mẫu của Lỗi 1 đến hết Lỗi 6 bằng sáu phân tích dưới đây. Mỗi lỗi giữ đủ đặc điểm, nhầm lẫn, bằng chứng, giả thuyết và đề xuất; khoảng 110–130 từ theo cách tách bằng khoảng trắng, hướng tới khoảng 1/3 trang khi chép vào báo cáo. Chi tiết cơ bản đã có ở Bảng 20 nên không nhắc lại toàn bộ.')
replace('Hướng dẫn đặt nội dung: Điền đủ năm dòng Bảng 21','Hướng dẫn đặt nội dung: Điền năm dòng Bảng 21 tại Mục 7.1. Đặt đoạn mở đầu trước bảng, một đoạn tổng kết sau bảng. Dẫn lại Bảng 10, 12–14 của phần tuần trước; không chép lại thí nghiệm hoặc diễn giải đã có.')
replace('Hướng dẫn đặt nội dung: Thay dấu “…” dưới Mục 7.2','Hướng dẫn đặt nội dung: Thay dấu “…” tại Mục 7.2 bằng hai đoạn dưới đây. Giữ đủ giới hạn dữ liệu, cách chia, nhân quả, phạm vi chọn cấu hình và khả năng tiếp cận theo nhóm; không lặp lại toàn bộ phân tích Mục 5–6.')
replace('Hướng dẫn đặt nội dung: Thay dấu “…” dưới Mục 7.3','Hướng dẫn đặt nội dung: Thay dấu “…” tại Mục 7.3 bằng một đoạn sau. Đây là hướng tương lai, chưa phải việc đã thực hiện.')

# Add a short, explicit interpretation of the teacher's length guidance.
anchor=next(p for p in d.paragraphs if p.text=='5. Kết quả trên test')
def insert(text,bold=False):
    par=anchor.insert_paragraph_before(text)
    par.paragraph_format.line_spacing=1.15
    par.paragraph_format.space_after=Pt(6)
    for r in par.runs:r.font.size=Pt(11);r.bold=bold
    return par
insert('Cân đối độ dài khi chép vào báo cáo',True)
insert('Bảng tổng quan gợi ý Mục 5 khoảng 2 trang, Mục 6 khoảng 2,5 trang và Mục 7 khoảng 1 trang; hướng dẫn chi tiết lần lượt nêu khoảng 2,5; 2,5 và 1,5 trang. Đây là gợi ý phân bổ, không phải giới hạn cứng. Ưu tiên đủ yêu cầu, đúng số liệu và không lặp ý.')
insert('Bản đã rút gọn ưu tiên Mục 5 khoảng 2–2,5 trang, Mục 7 khoảng 1,5–2 trang. Mục 6 có thể khoảng 3,5–4 trang vì cần hai bảng và sáu lỗi, mỗi lỗi khoảng 1/3 trang theo hướng dẫn riêng. Không ép Mục 6 xuống 2,5 trang bằng cách bỏ bằng chứng hoặc thu nhỏ chữ. Đây là ước lượng phần nội dung để chép, không tính các dòng hướng dẫn, trang mở đầu và phụ lục trong tài liệu này; số trang thực tế còn phụ thuộc vị trí nối với phần trước.')
candidate=BUILD/'revision_candidate.docx'
d.save(candidate)
(BUILD/'revision_operations.json').write_text(json.dumps({'paragraphs':ops,'table_21':values,'insert_before':'5. Kết quả trên test','insertions':[p.text for p in d.paragraphs if p.text.startswith(('Cân đối độ dài khi chép','Bảng tổng quan gợi ý','Bản đã rút gọn ưu tiên'))]},ensure_ascii=False,indent=2),encoding='utf-8')
print('Prepared revision:',candidate)
print('Paragraph replacements:',len(replacements))
