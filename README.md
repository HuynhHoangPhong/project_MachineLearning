# Đồ án Máy học — Bank Marketing

Dự đoán khách hàng có đồng ý gửi tiết kiệm có kỳ hạn trước khi gọi điện.

## Thứ tự làm

| Tuần theo đề | Notebook | Nội dung |
| --- | --- | --- |
| 0 | [week-0_cong_cu_va_du_lieu.ipynb](week-0_cong_cu_va_du_lieu.ipynb) | Làm quen công cụ và dữ liệu |
| 1 | [week-1_eda_va_thuoc_do.ipynb](week-1_eda_va_thuoc_do.ipynb) | Khám phá dữ liệu và tính các thước đo |
| 2 | [week-2_tien_xu_ly_logistic_knn.ipynb](week-2_tien_xu_ly_logistic_knn.ipynb) | Tiền xử lý, hồi quy logistic và kNN |
| 3 | [week-3_mlp_qua_khop_va_chon_mo_hinh.ipynb](week-3_mlp_qua_khop_va_chon_mo_hinh.ipynb) | MLP, quá khớp, chọn mô hình và ngưỡng |
| 4 | [week-5_danh_gia_test_va_phan_tich_loi.ipynb](week-5_danh_gia_test_va_phan_tich_loi.ipynb) | Đánh giá trên test và phân tích lỗi |

## Chạy và lưu kết quả

- Khi chạy trên máy, đặt thư mục làm việc của phiên Python tại `Project_FinalExam/`. Bốn notebook đồ án dùng chung `du_lieu/bank-full.csv`. Nếu chưa có, ô thiết lập tải tệp từ UCI. Trên Colab, đặt tệp dự phòng vào `du_lieu/` của phiên chạy.
- Tuần 0 tạo `du_lieu/thu.csv` và `hinh_thu.png` khi chạy các ô thực hành tương ứng.
- Lưu bản nháp và nhật ký thực hiện trong `reports/`, hình trong `reports/figures/` và bảng lưu riêng trong `reports/tables/`.
- Khóa mô hình và ngưỡng trên validation ở Tuần 3, sau đó chép nguyên dòng `QUYET_DINH` vào báo cáo và ô 4.1 của Tuần 4 trước khi đánh giá test.

## Nộp bài

Lưu `MSSV_ML.zip` trong `submission/`. Tệp ZIP gồm bốn notebook đã chạy (`MSSV_tuan1.ipynb` đến `MSSV_tuan4.ipynb`) và `MSSV_ML_baocao.pdf`. Nộp một lần qua Moodle.

Xem [đề bài](Docs/ML_Project_DCT124C_6-9.pdf), [mẫu báo cáo](Docs/ML_Report-Template_DCT124_6-9.pdf) và [sổ tay thư viện](Docs/So_tay_thu_vien.pdf) trong `Docs/`.
