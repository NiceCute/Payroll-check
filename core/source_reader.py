"""
source_reader.py - Đọc file lương (file nguồn) theo cấu hình SourceProfile.

Luồng xử lý:
1. Chọn sheet (theo tên cố định / mẫu tên / người dùng chỉ định).
2. Kiểm tra tiêu đề các cột đã mapping -> lệch là DỪNG (tránh đọc nhầm cột tiền).
3. Duyệt từng dòng dữ liệu:
     - Cột "Số thẻ" trống  -> bỏ qua dòng.
     - Có "Số thẻ"         -> lấy Họ tên, Số tài khoản, Số tiền -> tạo SourceRecord.
4. Kiểm tra trùng Số thẻ / Số tài khoản trong toàn file.

Module CHỈ ĐỌC, không sửa file.
"""

from __future__ import annotations

import logging
from collections import defaultdict
from pathlib import Path
from typing import Any, Callable, Optional

from .excel_reader import (
    PathLike,
    SheetSelectionError,
    WorkbookReader,
    cell_display_text,
    cell_problem,
    check_column_headers,
    select_sheet_name,
)
from .models import (
    CellData,
    CellKind,
    ColumnRef,
    Issue,
    IssueCode,
    Severity,
    SheetData,
    SkippedRow,
    SourceProfile,
    SourceReadResult,
    SourceRecord,
)
from .normalize import (
    NormResult,
    clean_text,
    has_letters,
    looks_like_legacy_font,
    normalize_account,
    normalize_amount,
    normalize_card,
    normalize_name_key,
)

__all__ = [
    "DEFAULT_SOURCE_PROFILE",
    "SheetSelectionError",   # giữ để tương thích: được định nghĩa trong excel_reader
    "check_headers",
    "parse_sheet",
    "read_source",
    "select_sheet",
]

log = logging.getLogger(__name__)

# Hàm ghi nhận vấn đề của 1 dòng: (mã, mức độ, thông báo, cột, số thẻ)
AddIssueFn = Callable[[IssueCode, Severity, str, ColumnRef, Optional[str]], None]

# Kiểm tra "nội dung cột có hợp lý không" - phát hiện cột bị lệch khi file lương bị chèn/xóa cột
# (đặc biệt với cột KHÔNG có tiêu đề như Họ tên W, Số tài khoản X).
SUSPECT_MIN_RECORDS: int = 3      # Ít hơn số dòng này thì không đủ cơ sở thống kê
SUSPECT_BAD_RATIO: float = 0.5    # > 50% dòng có giá trị bất thường ở 1 cột -> nghi mapping sai, DỪNG


# =============================================================================
# CẤU HÌNH MẶC ĐỊNH CHO FILE "BẢNG LƯƠNG THÁNG xx.xxxx.xlsx"
# -----------------------------------------------------------------------------
# Đây là cấu hình dự phòng khi chưa có file mapping (JSON) do giao diện tạo ra.
# Nếu file lương thay đổi cấu trúc, CHỈ CẦN SỬA CÁC DÒNG CÓ ĐÁNH DẤU  <-- MAPPING
#
#   ColumnRef("P", "Hồng Lâm Chuyển Khoản")
#              ^    ^
#              |    +-- tiêu đề mong đợi tại cột đó (None nếu cột không có tiêu đề).
#              |        Tool sẽ DỪNG nếu tiêu đề thực tế khác -> chống đọc nhầm cột.
#              +------- chữ cái cột trong Excel.
# =============================================================================
DEFAULT_SOURCE_PROFILE = SourceProfile(
    name="Bảng lương Hồng Lâm - sheet HLCK",
    # Tên sheet cố định. Để None để tìm theo mẫu bên dưới.
    sheet_name=None,
    # Mẫu tên sheet: '*' = chuỗi bất kỳ. "HLCK*" khớp "HLCK05.26 (2)", "HLCK06.26"...
    # Nếu khớp nhiều sheet, tool sẽ yêu cầu người dùng chọn (không tự đoán).
    sheet_patterns=("HLCK*",),
    # Dòng tiêu đề (tiêu đề bị merge 2 dòng 3-4)
    header_rows=(3, 4),
    # Dòng dữ liệu đầu tiên
    data_start_row=5,
    # Dòng dữ liệu cuối: None = đọc đến hết sheet
    data_end_row=None,

    # ---------------- MAPPING CỘT ----------------
    card_no=ColumnRef("B", "Số thẻ", label="Số thẻ"),                              # <-- MAPPING: Số thẻ (mã NV)
    full_name=ColumnRef("W", None, label="Họ tên (không dấu)"),                     # <-- MAPPING: cột W không có tiêu đề
    account_no=ColumnRef("X", None, label="Số tài khoản"),                         # <-- MAPPING: cột X không có tiêu đề
    amount=ColumnRef("P", "Hồng Lâm Chuyển Khoản", label="Hồng Lâm Chuyển Khoản"),  # <-- MAPPING: số tiền chuyển khoản
)


# =============================================================================
# 1. Chọn sheet
# =============================================================================

def select_sheet(sheet_names: list[str], profile: SourceProfile, override: Optional[str] = None) -> str:
    """Xác định sheet cần đọc (xem excel_reader.select_sheet_name)."""
    return select_sheet_name(sheet_names, profile.sheet_name, profile.sheet_patterns, override)


# =============================================================================
# 2. Kiểm tra tiêu đề
# =============================================================================

def check_headers(sheet: SheetData, profile: SourceProfile) -> list[Issue]:
    """So tiêu đề thực tế với tiêu đề mong đợi của từng cột đã mapping."""
    return check_column_headers(sheet, profile.header_rows, profile.columns())


# =============================================================================
# 3. Đọc dữ liệu
# =============================================================================

def read_source(
    path: PathLike,
    profile: SourceProfile = DEFAULT_SOURCE_PROFILE,
    sheet_name: Optional[str] = None,
) -> SourceReadResult:
    """Đọc file lương.

    Raises:
        ConfigError         : cấu hình không hợp lệ.
        ExcelReadError      : không đọc được file.
        SheetSelectionError : không xác định được sheet (giao diện cần hỏi người dùng).

    Trả về SourceReadResult. Kiểm tra `result.is_usable` trước khi dùng.
    """
    profile.ensure_valid()
    with WorkbookReader(path) as wb:
        chosen = select_sheet(wb.sheet_names, profile, sheet_name)
        sheet = wb.read_sheet(chosen)
    log.info("Đọc file lương '%s', sheet '%s' (%d dòng)", Path(path).name, chosen, sheet.max_row)
    return parse_sheet(sheet, profile, file_path=str(path))


def parse_sheet(sheet: SheetData, profile: SourceProfile, file_path: str = "") -> SourceReadResult:
    """Phân tích 1 sheet đã đọc (tách riêng để dễ kiểm thử)."""
    result = SourceReadResult(file_path=file_path, sheet_name=sheet.name, profile_name=profile.name)

    header_issues = check_headers(sheet, profile)
    if header_issues:
        result.header_ok = False
        result.issues.extend(header_issues)
        return result  # Sai cấu trúc -> KHÔNG đọc dữ liệu để tránh dùng nhầm cột

    cols = profile.columns()
    mapped_indexes = [ref.index for ref in cols.values()]
    last_row = profile.data_end_row if profile.data_end_row is not None else sheet.max_row
    legacy_font_rows: list[int] = []

    for r in range(profile.data_start_row, last_row + 1):
        row_issues: list[Issue] = []

        def add(code: IssueCode, sev: Severity, msg: str, ref: ColumnRef, card: Optional[str] = None) -> None:
            row_issues.append(Issue(code, sev, msg, sheet=sheet.name, row=r, column=ref.letter, card_no=card))

        # ---------- Bước 1: Số thẻ ----------
        card_ref = profile.card_no
        card_cell = sheet.cell(r, card_ref.index)
        card_problem = cell_problem(card_cell)
        if card_problem:
            add(card_problem[0], Severity.ERROR, f"Ô Số thẻ: {card_problem[1]}", card_ref)
            result.issues.extend(row_issues)
            result.skipped.append(SkippedRow(r, "Ô Số thẻ bị lỗi", _snapshot(sheet, r, cols)))
            continue

        card_res = normalize_card(card_cell.value)
        if card_res.is_error:
            add(card_res.code or IssueCode.CARD_INVALID, Severity.ERROR, card_res.detail, card_ref)
            result.issues.extend(row_issues)
            result.skipped.append(SkippedRow(r, "Số thẻ không hợp lệ", _snapshot(sheet, r, cols)))
            continue
        if card_res.value is None:
            # Không có Số thẻ -> bỏ qua. Chỉ ghi nhận nếu dòng có dữ liệu thực sự ở cột khác
            # (vd dòng Tổng cộng) để phục vụ đối soát.
            if _has_meaningful_data(sheet, r, mapped_indexes):
                result.skipped.append(SkippedRow(r, "Không có Số thẻ", _snapshot(sheet, r, cols)))
            continue
        card_no: str = card_res.value
        if card_res.code:
            add(card_res.code, card_res.severity or Severity.WARNING, card_res.detail, card_ref, card_no)

        # ---------- Bước 2: Họ tên ----------
        name_ref = profile.full_name
        name_cell = sheet.cell(r, name_ref.index)
        full_name = ""
        problem = cell_problem(name_cell)
        if problem:
            add(problem[0], Severity.ERROR, f"Ô Họ tên: {problem[1]}", name_ref, card_no)
        else:
            full_name = clean_text(name_cell.value)
            if not full_name:
                add(IssueCode.NAME_MISSING, Severity.WARNING, "Thiếu Họ tên.", name_ref, card_no)
            elif looks_like_legacy_font(full_name):
                legacy_font_rows.append(r)

        # ---------- Bước 2: Số tài khoản ----------
        acc_ref = profile.account_no
        account_no = _read_value(sheet, r, acc_ref, normalize_account, "Số tài khoản", card_no, add)
        if account_no is None and not any(i.column == acc_ref.letter for i in row_issues):
            add(IssueCode.ACCOUNT_MISSING, Severity.WARNING,
                "Có Số thẻ nhưng không có Số tài khoản.", acc_ref, card_no)

        # ---------- Bước 2: Số tiền ----------
        amt_ref = profile.amount
        amount = _read_value(sheet, r, amt_ref, normalize_amount, "Số tiền", card_no, add)
        if amount is None and not any(i.column == amt_ref.letter for i in row_issues):
            add(IssueCode.AMOUNT_MISSING, Severity.ERROR, "Có Số thẻ nhưng không có Số tiền.", amt_ref, card_no)

        has_error = any(i.severity == Severity.ERROR for i in row_issues)
        result.issues.extend(row_issues)
        result.records.append(SourceRecord(
            row=r,
            card_no=card_no,
            full_name=full_name,
            name_key=normalize_name_key(full_name),
            account_no=account_no,
            amount=amount,
            has_error=has_error,
        ))

    if legacy_font_rows:
        result.issues.append(Issue(
            code=IssueCode.NAME_LEGACY_FONT,
            severity=Severity.WARNING,
            message=(
                f"{len(legacy_font_rows)} tên đang gõ bằng font cũ VNI/TCVN3 (ví dụ dòng {legacy_font_rows[0]}). "
                "Nên mapping sang cột tên không dấu để kiểm tra chéo chính xác."
            ),
            sheet=sheet.name,
            column=profile.full_name.letter,
        ))

    if not result.records:
        result.issues.append(Issue(
            code=IssueCode.NO_RECORDS,
            severity=Severity.ERROR,
            message=(
                f"Không tìm thấy dòng nào có Số thẻ (cột {profile.card_no.letter}) "
                f"từ dòng {profile.data_start_row}. Kiểm tra lại sheet hoặc mapping."
            ),
            sheet=sheet.name,
            column=profile.card_no.letter,
        ))

    result.issues.extend(_duplicate_issues(result.records, sheet.name, profile))
    result.issues.extend(_column_content_issues(result, sheet.name, profile))
    return result


# =============================================================================
# 4. Hàm phụ
# =============================================================================

def _read_value(
    sheet: SheetData,
    row: int,
    ref: ColumnRef,
    normalizer: Callable[[Any], NormResult[Any]],
    label: str,
    card_no: str,
    add: AddIssueFn,
) -> Any:
    """Đọc + chuẩn hóa 1 ô; ghi nhận vấn đề qua hàm add. Trả về giá trị hoặc None."""
    cell: CellData = sheet.cell(row, ref.index)
    problem = cell_problem(cell)
    if problem:
        add(problem[0], Severity.ERROR, f"Ô {label}: {problem[1]}", ref, card_no)
        return None
    res = normalizer(cell.value)
    if res.code is not None:
        add(res.code, res.severity or Severity.ERROR, res.detail, ref, card_no)
    return None if res.is_error else res.value


def _snapshot(sheet: SheetData, row: int, cols: dict[str, ColumnRef]) -> dict[str, str]:
    return {key: cell_display_text(sheet.cell(row, ref.index)) for key, ref in cols.items()}


def _has_meaningful_data(sheet: SheetData, row: int, col_indexes: list[int]) -> bool:
    """Dòng có ít nhất 1 ô khác trống và khác số 0 (số 0 ở dòng tên tổ không có ý nghĩa đối soát)."""
    for c in col_indexes:
        cell = sheet.cell(row, c)
        if cell.is_empty:
            continue
        if cell.kind == CellKind.NUMBER and cell.value == 0:
            continue
        return True
    return False


def _column_content_issues(result: SourceReadResult, sheet_name: str, profile: SourceProfile) -> list[Issue]:
    """Phát hiện cột bị map nhầm dựa trên nội dung (vì cột Họ tên / Số TK có thể không có tiêu đề).

    Nếu quá nửa số dòng có giá trị bất thường ở cùng 1 cột -> khả năng cao file lương đã bị
    chèn/xóa cột -> chặn, yêu cầu kiểm tra lại mapping (không tự đoán cột đúng).
    """
    records = result.records
    total = len(records)
    if total < SUSPECT_MIN_RECORDS:
        return []

    unusual_acc_rows = {
        i.row for i in result.issues if i.code in (IssueCode.ACCOUNT_UNUSUAL_LENGTH, IssueCode.ACCOUNT_INVALID)
    }
    bad_counts: dict[str, int] = {
        "full_name": sum(1 for r in records if not has_letters(r.full_name)),
        "account_no": sum(1 for r in records if r.account_no is None or r.row in unusual_acc_rows),
        "amount": sum(1 for r in records if r.amount is None),
    }
    descriptions = {
        "full_name": "trống hoặc không chứa chữ cái",
        "account_no": "trống, sai định dạng hoặc độ dài bất thường",
        "amount": "trống hoặc không hợp lệ",
    }
    issues: list[Issue] = []
    for key, bad in bad_counts.items():
        if bad / total > SUSPECT_BAD_RATIO:
            ref = profile.columns()[key]
            issues.append(Issue(
                IssueCode.COLUMN_CONTENT_SUSPECT, Severity.ERROR,
                f"{bad}/{total} dòng có {ref.display()} {descriptions[key]}. "
                "Có thể file lương đã bị chèn/xóa cột -> hãy kiểm tra lại mapping.",
                sheet=sheet_name, column=ref.letter,
            ))
    return issues


def _duplicate_issues(records: list[SourceRecord], sheet_name: str, profile: SourceProfile) -> list[Issue]:
    issues: list[Issue] = []

    by_card: dict[str, list[int]] = defaultdict(list)
    by_acc: dict[str, list[int]] = defaultdict(list)
    by_name: dict[str, list[int]] = defaultdict(list)
    for rec in records:
        by_card[rec.card_no].append(rec.row)
        if rec.account_no:
            by_acc[rec.account_no].append(rec.row)
        if rec.name_key:
            by_name[rec.name_key].append(rec.row)

    for card, rows in by_card.items():
        if len(rows) > 1:
            issues.append(Issue(
                IssueCode.DUPLICATE_CARD, Severity.ERROR,
                f"Số thẻ '{card}' bị trùng ở các dòng {rows}. Phải sửa file lương trước khi tiếp tục.",
                sheet=sheet_name, column=profile.card_no.letter, card_no=card,
            ))
    for acc, rows in by_acc.items():
        if len(rows) > 1:
            issues.append(Issue(
                IssueCode.DUPLICATE_ACCOUNT, Severity.ERROR,
                f"Số tài khoản '{acc}' bị trùng ở các dòng {rows}. Phải sửa file lương trước khi tiếp tục.",
                sheet=sheet_name, column=profile.account_no.letter,
            ))
    for name, rows in by_name.items():
        if len(rows) > 1:
            issues.append(Issue(
                IssueCode.DUPLICATE_NAME, Severity.INFO,
                f"Tên '{name}' xuất hiện ở các dòng {rows} (được phép - tool không dùng tên làm khóa).",
                sheet=sheet_name, column=profile.full_name.letter,
            ))
    return issues
