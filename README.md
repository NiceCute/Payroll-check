# Công cụ Tự động Đối chiếu & Chi hộ Lương Ngân hàng (`ChiHoLuongTool` V1)

Ứng dụng máy tính (Desktop Windows) hỗ trợ nhân viên Kế toán – Tiền lương tự động đọc **Bảng lương tháng** (`.xlsx` / `.xls`), đối chiếu với các **File biểu mẫu chi hộ ngân hàng** (ví dụ: mẫu ACB `.xls`, mẫu Liên ngân hàng Thông thường `.xls`) và điền chính xác cột **Số tiền** vào bản sao mới mà **giữ nguyên 100% định dạng, công thức, ô gộp, khung viền và các sheet ẩn/khóa** của ngân hàng.

---

## 1. Nguyên tắc An toàn Kế toán (Cam kết của Công cụ)

1. **Không bao giờ sửa file gốc:** Mọi file bảng lương nguồn và file ngân hàng đầu vào đều được giữ nguyên vẹn 100% (kiểm tra mã băm SHA-256 trước và sau khi chạy).
2. **Tự động sao lưu (Backup):** Trước khi tạo file kết quả, công cụ tự động sao lưu file gốc vào thư mục `data/backup/` kèm dấu thời gian.
3. **Chỉ cập nhật đúng cột Số tiền:** Tuyệt đối không sửa Số tài khoản, Họ tên, Mã ngân hàng, tiêu đề, công thức hay các sheet hướng dẫn/tham chiếu (`Soạn lệnh`, `Hướng dẫn nhập liệu`, `DSid`, `dstinhNH`).
4. **Giữ nguyên 100% số `0` ở đầu:** Số thẻ nhân viên (`0001`), Số tài khoản (`0587952606`), Mã ngân hàng (`79203001`) luôn được xử lý dưới dạng chuỗi văn bản.
5. **Chốt chặn an toàn & Hậu kiểm tự động:**
   - Tự động chặn ghi nếu còn ngoại lệ chưa xử lý, trùng Số thẻ / Số tài khoản, hoặc lệch tổng tiền.
   - Sau khi ghi xong, bộ **Verifier (Hậu kiểm)** đọc lại toàn bộ file mới để đối chiếu từng ô với file gốc: nếu bất kỳ ô nào ngoài danh sách dự kiến bị thay đổi, công cụ tự động hủy file lỗi và cảnh báo ngay.

---

## 2. Cấu trúc Thư mục

```text
ChiHoLuongTool/
├─ app.py                  # Điểm khởi chạy chính của ứng dụng
├─ requirements.txt        # Danh sách thư viện Python (openpyxl, xlrd)
├─ install.bat             # Tự động tạo môi trường ảo (.venv) và cài thư viện
├─ run.bat                 # Khởi chạy ứng dụng trực tiếp bằng Python
├─ build.bat               # Đóng gói thành file chạy độc lập (.exe) bằng PyInstaller
├─ pyinstaller.spec        # Cấu hình đóng gói PyInstaller
├─ README.md               # Tài liệu hướng dẫn sử dụng
│
├─ core/                   # Các module xử lý nghiệp vụ cốt lõi
├─ gui/                    # Giao diện người dùng (Main Window, Preview Table, Mapping Dialog)
├─ templates/              # Các file cấu hình Mapping cột lưu sẵn (.json)
│  ├─ source_bang_luong.json
│  ├─ target_acb.json
│  └─ target_thong_thuong.json
│
└─ data/                   # Dữ liệu sinh ra khi vận hành
   ├─ backup/              # Bản sao lưu tự động của file gốc (kèm ngày giờ)
   ├─ output/              # File ngân hàng kết quả mới (mặc định thêm hậu tố _KetQua)
   ├─ reports/             # Báo cáo đối chiếu & hậu kiểm (.xlsx, gồm 4 sheet)
   └─ logs/                # Nhật ký vận hành theo ngày (chi_ho_luong_YYYYMMDD.log)
```

---

## 3. Hướng dẫn Cài đặt & Khởi chạy

### Yêu cầu hệ thống
- Hệ điều hành: **Windows 10 / Windows 11**.
- Đã cài đặt **Microsoft Excel** (để giữ nguyên 100% định dạng và các sheet khóa của file `.xls` ngân hàng).
- Nếu chạy từ mã nguồn: cần **Python 3.10+**.

### Cách 1: Chạy trực tiếp bằng Python
1. Nhấp đúp vào **`install.bat`** (chỉ cần chạy 1 lần đầu tiên để tạo môi trường `.venv` và cài đặt thư viện).
2. Nhấp đúp vào **`run.bat`** để mở giao diện chương trình.

### Cách 2: Đóng gói thành file `.exe` (không cần cài Python trên máy kế toán khác)
1. Nhấp đúp vào **`build.bat`**.
2. Sau khi đóng gói xong, toàn bộ bộ cài đặt gọn nhẹ nằm tại thư mục:
   `dist\ChiHoLuongTool\ChiHoLuongTool.exe`

---

## 4. Quy trình Sử dụng 3 Bước trên Màn hình Chính

### Bước ①: Chọn File & Template
1. **Chọn file nguồn:** Bấm **`Chọn file nguồn...`** và chọn file Bảng lương tháng (ví dụ: `BẢNG LƯƠNG THÁNG 05.2026.xlsx`).
   - Công cụ tự động liệt kê các sheet hiển thị và tự nhận diện sheet lương (ví dụ: `HLCK05.26 (2)`).
2. **Chọn file đích:** Bấm **`Chọn file đích...`** và chọn 1 hoặc nhiều file biểu mẫu ngân hàng kỳ trước (ví dụ: cả 2 file `4 MauChiHoLuongQuaTaiKhoanACB.xls` và `5 MauChiHoLuongThongThuong.xls`).
3. **Template:** Mặc định để **`Tự động nhận diện (Auto)`**. Nếu biểu mẫu ngân hàng thay đổi vị trí cột, bấm **`Cấu hình Mapping...`** để chỉnh sửa và lưu lại.

### Bước ②: Phân tích (`Analyze`) & Xem trước (`Preview`)
1. Bấm **`🔍 Analyze (Phân tích)`**:
   - Công cụ đọc dữ liệu, đối chiếu theo **Số tài khoản** và kiểm tra chéo **Họ và tên**.
   - 4 thẻ thống kê hiển thị ngay: **Số lượng Record**, **Ngoại lệ (Exception)**, **Tổng tiền ghi (VND)** và **Trạng thái An toàn**.
2. Bấm **`👁 Preview (Xem trước)`**:
   - Xem chi tiết từng dòng nhân viên với màu sắc trực quan:
     - 🟢 **Xanh lá:** Khớp hoàn toàn, sẵn sàng cập nhật số tiền mới.
     - ⚪ **Trắng / Xám:** Số tiền không đổi hoặc dòng đã chọn bỏ qua.
     - 🟡 **Vàng (Warning):** Có cảnh báo hoặc ngoại lệ cần người dùng xác nhận.
     - 🔴 **Đỏ nhạt (Exception):** Ngoại lệ chưa xử lý hoặc lỗi chặn cứng.
   - Hỗ trợ **Tìm kiếm không dấu**, **Lọc theo trạng thái**, **Sắp xếp theo cột**.

### Bước ③: Xử lý Ngoại lệ (nếu có) & Xuất kết quả (`Execute`)
1. Nếu có ngoại lệ cần quyết định (ví dụ: lệch tên `X4`, hoặc nhân viên kỳ trước nghỉ việc không có trong bảng lương kỳ này `X5`):
   - Chọn dòng ngoại lệ trên bảng Preview, chọn phương án xử lý ở thanh dưới cùng (ví dụ: *Chấp nhận ghi*, *Bỏ qua*, *Giữ nguyên số cũ*, *Đặt về 0*, hoặc *Xóa trắng ô tiền*) rồi bấm **`✔ Áp dụng xử lý`**.
2. Khi trạng thái chuyển sang **`🟢 AN TOÀN (Sẵn sàng ghi)`**, bấm **`💾 Execute (Xuất file mới)`**:
   - Công cụ tự động sao lưu file gốc vào `data/backup/`.
   - Tạo file ngân hàng mới tại `data/output/` (ví dụ: `4 MauChiHoLuongQuaTaiKhoanACB_KetQua.xls`).
   - Thực hiện **Hậu kiểm (Verifier)** toàn bộ ô trên file mới.
   - Tự động xuất **Báo cáo đối chiếu Excel (`.xlsx`)** gồm 4 sheet (`Tổng hợp`, `Chi tiết`, `Exception`, `Kiểm tra & Hậu kiểm`) vào `data/reports/`.
3. Bấm **`📂 Mở thư mục xuất`** để lấy file ngân hàng mới nộp cho ngân hàng.

---

## 5. Bảng Mã Ngoại lệ (Exception) & Cách Xử lý

| Mã | Ý nghĩa | Quy tắc xử lý trên phần mềm |
|---|---|---|
| **X1** | Có trong bảng lương nhưng không tìm thấy Số TK trong bất kỳ file ngân hàng nào (nhân viên mới / đổi NH) | Chỉ ghi vào Báo cáo Exception để kế toán nắm danh sách, không tự chèn dòng mới vào file ngân hàng (tránh vỡ công thức tổng). |
| **X2** | Dòng có Số thẻ nhưng trống Số tài khoản (nhận tiền mặt / chưa mở TK) | Tự động bỏ qua và liệt kê trong Báo cáo. |
| **X3** | Số tiền thực lãnh trống, bằng `0`, âm hoặc số thập phân lẻ | Không ghi dòng đó, đưa vào Báo cáo để kiểm tra lại. |
| **X4** | Trùng khớp Số tài khoản nhưng Họ tên giữa bảng lương và file ngân hàng khác nhau | Cảnh báo màu vàng; kế toán kiểm tra và chọn **Chấp nhận ghi (`ACCEPT`)** hoặc **Bỏ qua (`SKIP`)**. |
| **X5** | Có trong file ngân hàng kỳ trước nhưng không có trong bảng lương tháng này (nghỉ việc / thai sản) | Kế toán bắt buộc chọn 1 trong 3 phương án: **Đặt số tiền = 0 (`SET_ZERO`)**, **Xóa trắng ô tiền (`CLEAR`)**, hoặc **Giữ nguyên số cũ (`KEEP`)**. |
| **X6** | Tài khoản ngân hàng ở file Thông thường thiếu Mã ngân hàng (Citad 8 số) | Chặn ghi dòng lỗi và cảnh báo rõ vị trí dòng. |
| **X7** | Giá trị tiền nằm ở cột kế bên (nghi nhập lệch cột trong bảng lương) | Cảnh báo để kế toán kiểm tra lại file lương hoặc chọn Bỏ qua. |
| **X8** | Trùng Số thẻ hoặc trùng Số tài khoản (2 dòng cùng 1 TK) | **Chặn cứng (Blocking Error)** để tránh ghi đè sai tiền; cần sửa lại file nguồn/đích trước khi xuất. |

---

## 6. Xử lý Sự cố Thường gặp (Troubleshooting)

- **Thông báo *"File đang được mở trong Excel hoặc bị khóa"*:**
  Hãy đóng file ngân hàng hoặc file kết quả đang mở trong Microsoft Excel, sau đó bấm **`Execute`** lại.
- **File gốc được đặt thuộc tính *Read-Only (Chỉ đọc)*:**
  Công cụ hỗ trợ đọc bình thường các file đầu vào *Read-Only* và tự động cấp quyền ghi cho bản sao mới tại `data/output/` mà vẫn giữ nguyên trạng thái *Read-Only* của file gốc.
- **Tra cứu lịch sử chạy & lỗi chi tiết:**
  Mở thư mục `data/logs/` và xem file `chi_ho_luong_YYYYMMDD.log` của ngày tương ứng.
