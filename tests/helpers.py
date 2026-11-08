"""Hàm dựng file Excel giả lập cho kiểm thử (không dùng dữ liệu lương thật)."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable, Optional

import openpyxl

# Dòng dữ liệu mẫu: (Số thẻ, Tên VNI cột C, Tên không dấu W, Số TK X, Số tiền P)
Row = tuple[Any, Any, Any, Any, Any]

DEFAULT_ROWS: list[Row] = [
    (None, "ATO18: Tổ 1XM", None, None, None),                          # dòng 5: tên tổ
    ("16021804", "Nguyeãn Thaønh Coâng", "NGUYEN THANH CONG", "0587952606", 5615871),  # dòng 6
    ("22050301", "Phaïm Thò Lieân", "PHAM THI LIEN", 24349081, 5917028.0),          # dòng 7
    (None, "Tổng cộng", None, None, 11532899),                          # dòng 8: tổng
    ("26032301", "Löông Vaên Haûi", "LUONG VAN HAI", None, 4988450),             # dòng 9: thiếu TK
    ("4/7/2012", "Traàn Ñöùc Traân", "TRAN DUC TRAN", 24230311, 2439377),          # dòng 10
]


def build_source_xlsx(
    path: Path,
    rows: Optional[Iterable[Row]] = None,
    sheet_names: Iterable[str] = ("HLCK06.26 (2)",),
    amount_header: str = "Hồng Lâm\nChuyển  Khoản",
    insert_col_before_p: bool = False,
    extra_hidden_sheet: bool = False,
) -> Path:
    """Tạo file lương giả lập có cấu trúc giống file thật (tiêu đề dòng 3-4, dữ liệu từ dòng 5)."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    data = list(DEFAULT_ROWS if rows is None else rows)
    for name in sheet_names:
        ws = wb.create_sheet(name)
        ws["A1"] = "CTY TNHH SX TMDV HỒNG LÂM"
        ws["A2"] = "BẢNG LƯƠNG THÁNG 06/2026"
        ws["A3"], ws["B3"], ws["C3"] = "Stt", "Số thẻ", "Họ Và Tên"
        ws.merge_cells("B3:B4")
        ws.merge_cells("C3:C4")
        amount_col = "Q" if insert_col_before_p else "P"
        ws[f"{amount_col}3"] = amount_header
        if insert_col_before_p:
            ws["P3"] = "Cột mới chèn"
        ws.merge_cells(f"{amount_col}3:{amount_col}4")
        for i, (card, name_vni, name_plain, acc, amount) in enumerate(data, start=5):
            ws[f"B{i}"] = card
            ws[f"C{i}"] = name_vni
            ws[f"W{i}"] = name_plain
            ws[f"X{i}"] = acc
            ws[f"{amount_col}{i}"] = amount
    if extra_hidden_sheet:
        hidden = wb.create_sheet("An")
        hidden["A1"] = "x"
        hidden.sheet_state = "hidden"
    wb.save(path)
    return path


def build_source_xls(path: Path, rows: Optional[Iterable[Row]] = None, sheet_name: str = "HLCK06.26 (2)") -> Path:
    """Tạo file .xls giả lập (cần thư viện xlwt)."""
    import xlwt

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet(sheet_name)
    ws.write(0, 0, "CTY TNHH SX TMDV HỒNG LÂM")
    ws.write(2, 1, "Số thẻ")
    ws.write(2, 2, "Họ Và Tên")
    ws.write(2, 15, "Hồng Lâm Chuyển Khoản")
    text_style = xlwt.easyxf(num_format_str="@")
    for i, (card, name_vni, name_plain, acc, amount) in enumerate(DEFAULT_ROWS if rows is None else rows, start=4):
        for col, val in ((1, card), (2, name_vni), (22, name_plain), (23, acc), (15, amount)):
            if val is not None:
                ws.write(i, col, val, text_style if isinstance(val, str) else xlwt.Style.default_style)
    wb.save(str(path))
    return path


# =============================================================================
# Mẫu ngân hàng (file đích) giả lập
# =============================================================================

# Dòng mẫu ACB: (Tên NV cột B, Số TK cột C, Số tiền cột D)
AcbRow = tuple[Any, Any, Any]

DEFAULT_ACB_ROWS: list[AcbRow] = [
    ("NGUYEN THANH CONG", 24349081, 4500000),
    ("PHAM THI LIEN", 24230311, 5100000),
    ("TRAN DUC TRAN", 24112233, 3200000),
]

# Dòng mẫu Thông thường: (Tên NV cột B, Số tiền cột E, Mã NH cột F, Số TK cột G)
TtRow = tuple[Any, Any, Any, Any]

DEFAULT_TT_ROWS: list[TtRow] = [
    ("NGUYEN VAN AN", 6000000, "79203001", "0587952606"),
    ("LE THI BINH", 7200000, "79304001", "0071001234567"),
    ("HOANG MINH", 5400000, "79310001", "19034567890011"),
]


def build_acb_xls(
    path: Path,
    rows: Optional[Iterable[AcbRow]] = None,
    sheet_name: str = "CNBD",
    with_guide: bool = True,
    amount_header: str = "Số tiền",
) -> Path:
    """Tạo file mẫu ACB (.xls) giống cấu trúc 4 MauChiHoLuongQuaTaiKhoanACB.xls."""
    import xlwt

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet(sheet_name)
    ws.write(0, 0, "DANH SÁCH CHI HỘ LƯƠNG QUA TÀI KHOẢN ACB")
    ws.write(1, 0, "STT")
    ws.write(1, 1, "Tên NV")
    ws.write(1, 2, "Số tài khoản")
    ws.write(1, 3, amount_header)
    for idx, (name, acc, amt) in enumerate(DEFAULT_ACB_ROWS if rows is None else rows, start=1):
        r = idx + 1  # bắt đầu từ dòng 3 (index 2)
        ws.write(r, 0, idx)
        if name is not None:
            ws.write(r, 1, name)
        if acc is not None:
            ws.write(r, 2, acc)
        if amt is not None:
            ws.write(r, 3, amt)
    if with_guide:
        g = wb.add_sheet("Hướng dẫn nhập liệu")
        g.write(0, 0, "HƯỚNG DẪN NHẬP LIỆU")
        g.write(1, 0, "Không nhập dấu tiếng Việt, không thừa khoảng trắng.")
        g.write(2, 0, "Nội dung tối đa 210 ký tự, không gồm dấu &")
    wb.save(str(path))
    return path


def build_tt_xls(
    path: Path,
    rows: Optional[Iterable[TtRow]] = None,
    sheet_name: str = "Dán giá trị vào đây để TT",
    with_soan_lenh: bool = True,
    with_guide: bool = True,
) -> Path:
    """Tạo file mẫu Thông thường (.xls) giống 5 MauChiHoLuongThongThuong.xls."""
    import xlwt

    wb = xlwt.Workbook(encoding="utf-8")
    if with_soan_lenh:
        sl = wb.add_sheet("Soạn lệnh")
        sl.write(1, 1, "Tên NV")
        sl.write(1, 4, "Số tiền")
        sl.write(1, 5, "Ngân hàng")  # khác tiêu đề "Mã ngân hàng" của sheet dán
        sl.write(1, 6, "Số tài khoản")
    ws = wb.add_sheet(sheet_name)
    ws.write(1, 0, "STT")
    ws.write(1, 1, "Tên NV")
    ws.write(1, 4, "Số tiền")
    ws.write(1, 5, "Mã ngân hàng")
    ws.write(1, 6, "Số tài khoản")
    ws.write(1, 10, "Số thẻ")  # cột K trùng tên "Số thẻ" (thẻ ATM, không phải mã NV)
    text_style = xlwt.easyxf(num_format_str="@")
    for idx, (name, amt, bank, acc) in enumerate(DEFAULT_TT_ROWS if rows is None else rows, start=1):
        r = idx + 1
        ws.write(r, 0, idx)
        if name is not None:
            ws.write(r, 1, name)
        if amt is not None:
            ws.write(r, 4, amt)
        if bank is not None:
            ws.write(r, 5, bank, text_style if isinstance(bank, str) else xlwt.Style.default_style)
        if acc is not None:
            ws.write(r, 6, acc, text_style if isinstance(acc, str) else xlwt.Style.default_style)
    if with_guide:
        g = wb.add_sheet("Hướng dẫn nhập liệu")
        g.write(0, 0, "Không dấu, không khoảng trắng thừa.")
    wb.save(str(path))
    return path


def build_target_xlsx(
    path: Path,
    rows: Optional[Iterable[AcbRow]] = None,
    sheet_name: str = "CNBD",
    with_guide: bool = False,
) -> Path:
    """Tạo file mẫu ACB định dạng .xlsx (để kiểm thử công thức, merge cell, v.v.)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = sheet_name
    ws["A2"], ws["B2"], ws["C2"], ws["D2"] = "STT", "Tên NV", "Số tài khoản", "Số tiền"
    for idx, (name, acc, amt) in enumerate(DEFAULT_ACB_ROWS if rows is None else rows, start=1):
        r = idx + 2
        ws[f"A{r}"] = idx
        ws[f"B{r}"] = name
        ws[f"C{r}"] = acc
        ws[f"D{r}"] = amt
    if with_guide:
        g = wb.create_sheet("Hướng dẫn nhập liệu")
        g["A1"] = "Không nhập dấu tiếng Việt, tối đa 210 ký tự"
    wb.save(path)
    return path

