# Ghi chú Phát hành (Release Notes) — `ChiHoLuongTool` V1.0.0

- **Phiên bản:** `1.0.0` (Release Candidate V1.0)
- **Ngày phát hành:** `08/11/2026`
- **Nền tảng mục tiêu:** Windows 10 / Windows 11 (64-bit) + Microsoft Excel

---

## 1. Tóm tắt Phiên bản V1.0.0

`ChiHoLuongTool` V1.0.0 là công cụ tự động hóa đối chiếu và cập nhật bảng kê chi hộ lương ngân hàng dành cho bộ phận Kế toán – Tiền lương. Công cụ đọc dữ liệu từ file **Bảng lương tháng** (`.xlsx` / `.xls`), đối chiếu chính xác với **File biểu mẫu ngân hàng kỳ trước** (`.xls` / `.xlsx` — bao gồm mẫu **ACB** và mẫu **Liên ngân hàng Thông thường**), kiểm tra chốt chặn an toàn và điền cột **Số tiền** vào bản sao mới mà **giữ nguyên 100% định dạng, công thức, ô gộp, khung viền và các sheet ẩn/khóa**.

---

## 2. Các Tính năng Chính trong V1.0.0

### 2.1. Đọc & Chuẩn hóa Dữ liệu Đa Định dạng (`Module 1` & `Module 2`)
- Hỗ trợ đọc cả hai định dạng `.xlsx` (`openpyxl`) và `.xls` (`xlrd`).
- Tự động chuẩn hóa và bảo toàn **100% chữ số `0` đứng đầu** cho Số thẻ nhân viên (`0001`), Số tài khoản (`0587952606`), và Mã ngân hàng Citad 8 chữ số (`79203001`).
- Chuẩn hóa Họ và tên tiếng Việt (bỏ dấu, viết hoa, gom khoảng trắng) để kiểm tra chéo chính xác giữa bảng lương và biểu mẫu ngân hàng.
- Quản lý cấu hình cột linh hoạt qua `templates/*.json` (`source_bang_luong.json`, `target_acb.json`, `target_thong_thuong.json`), tự động nhận diện mẫu theo cấu trúc sheet và tiêu đề cột.

### 2.2. Đối chiếu Thông minh & Phân loại Ngoại lệ (`Module 3 — Reconciler`)
- Đối chiếu theo khóa chính **Số tài khoản** (từ các dòng hợp lệ có Số thẻ nhân viên) và kiểm tra chéo **Họ và tên**.
- Phát hiện và phân loại đầy đủ 8 mã ngoại lệ kế toán (`X1` → `X8`):
  - **`X1`:** Có trong bảng lương nhưng chưa có trong file ngân hàng đích.
  - **`X2`:** Có Số thẻ nhân viên nhưng trống Số tài khoản (nhận tiền mặt — tự động bỏ qua và ghi nhận vào báo cáo).
  - **`X3`:** Số tiền thực lãnh trống, bằng `0`, âm hoặc số thập phân lẻ.
  - **`X4`:** Trùng Số tài khoản nhưng lệch Họ và tên giữa nguồn và đích (cho phép kế toán xác nhận `ACCEPT` hoặc `SKIP`).
  - **`X5`:** Có trong biểu mẫu ngân hàng kỳ trước nhưng không có trong bảng lương kỳ này (bắt buộc chọn `SET_ZERO`, `CLEAR`, hoặc `KEEP`).
  - **`X6`:** Dòng trong mẫu Liên ngân hàng Thông thường thiếu Mã ngân hàng Citad.
  - **`X7`:** Cảnh báo nghi ngờ nhập lệch cột tiền trong bảng lương.
  - **`X8`:** Trùng lặp Số thẻ hoặc Số tài khoản (chặn cứng không cho phép ghi).

### 2.3. Chốt chặn An toàn & Ghi File Bảo toàn Định dạng (`Module 4` & `Module 5`)
- **Safety Layer:** Kiểm tra khóa file (`~$`), kiểm tra ngoại lệ chưa xử lý, kiểm tra trùng lặp định danh và cân đối tổng tiền (`K8`) trước khi cho phép ghi. Hỗ trợ đầy đủ cả trường hợp file gốc bật thuộc tính *Read-Only (Chỉ đọc)* (`UAT-15`).
- **Excel Writer Layer:**
  - Tự động sao lưu nguyên bản file lương nguồn và các file ngân hàng đích vào `data/backup/` kèm dấu thời gian.
  - Tạo file ngân hàng mới tại `data/output/` (không bao giờ ghi đè lên file gốc; kiểm tra SHA-256 của file gốc trước và sau khi chạy).
  - Ghi qua Excel COM để giữ nguyên 100% phông chữ, màu nền, viền ô, công thức, ô gộp, độ rộng cột, chiều cao dòng và toàn bộ các sheet phụ/ẩn/khóa (`Soạn lệnh`, `Hướng dẫn nhập liệu`, `DSid`, `dstinhNH`).

### 2.4. Hậu kiểm Toàn diện & Xuất Báo cáo Excel (`Module 6`, `Module 7` & `Module 8`)
- **Verifier Layer:** Đọc lại toàn bộ các sheet và từng ô của file kết quả mới để đối chiếu với file gốc; đảm bảo **chỉ duy nhất các ô trong `WritePlan` được thay đổi** (`unexpected_diffs = 0`).
- **Reporting Layer:** Tự động xuất báo cáo đối chiếu `.xlsx` gồm 4 sheet (`Tổng hợp`, `Chi tiết`, `Exception`, `Kiểm tra & Hậu kiểm`) vào `data/reports/`.

### 2.5. Giao diện Kế toán Trực quan 3 Bước (`GUI Modules 1–3`)
- **Main Window:** Thiết kế 1 màn hình – 3 bước rõ ràng (`① Chọn file & Template` → `② Phân tích & Preview` → `③ Xuất kết quả`) kèm 4 thẻ chỉ số thống kê thời gian thực.
- **Preview Table:** Bảng xem trước tô màu trực quan theo trạng thái, hỗ trợ tìm kiếm không dấu, lọc nhanh theo trạng thái/ngoại lệ, sắp xếp cột và xử lý ngoại lệ ngay trên bảng.
- **Mapping Dialog:** Cho phép xem danh sách cột nguồn/đích, tạo/sửa/xóa và kiểm tra hợp lệ các file cấu hình `templates/*.json`.

---

## 3. Danh mục Thành phần Đóng gói (`dist/ChiHoLuongTool/`)

- `ChiHoLuongTool.exe` — Tệp thực thi chính trên Windows (không yêu cầu cài sẵn Python trên máy người dùng).
- `_internal/` — Thư viện runtime Python 3.14, Tkinter/Tcl/Tk, `openpyxl`, `xlrd` và các gói phụ thuộc đã đóng gói sẵn.
- `templates/` — Bộ 3 cấu hình JSON chuẩn (`source_bang_luong.json`, `target_acb.json`, `target_thong_thuong.json`).
- `data/backup/`, `data/output/`, `data/reports/`, `data/logs/` — Cấu trúc thư mục dữ liệu vận hành.
- `README.md` — Tài liệu hướng dẫn sử dụng chi tiết cho kế toán.

---

## 4. Kết quả Kiểm thử Chấp nhận (UAT & Smoke Test)

- **Unit & Integration Tests:** `161/161` bài kiểm thử tự động đạt `100%` (bao gồm `UAT-01` → `UAT-15`).
- **Kiểm thử đối chiếu thực tế tháng 05/2026 (`UAT-01`):**
  - Tổng số bản ghi nguồn: `90` dòng (`86` dòng có tài khoản + `4` dòng nhận tiền mặt `X2`).
  - Khớp và cập nhật thành công: `86/86` nhân viên (`79` dòng ACB = `405,476,927 VND` + `7` dòng Thông thường = `39,661,674 VND`).
  - Tổng tiền chi hộ chuẩn bị ghi: **`445,138,601 VND`** (khớp tuyệt đối 100% với tổng tiền nguồn đủ điều kiện).
  - Đối chiếu kết quả chạy giữa `python app.py` và `ChiHoLuongTool.exe`: **Giống nhau 100%**.
