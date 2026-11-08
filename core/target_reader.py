"""
target_reader.py - Đọc file mẫu ngân hàng (file đích) theo cấu hình TargetTemplate.

Luồng xử lý:
1. Chọn sheet nhập liệu (tên cố định / mẫu tên / tự nhận theo tiêu đề nếu chỉ 1 sheet khớp).
2. Kiểm tra tiêu đề cột -> lệch là DỪNG (tránh ghi số tiền vào nhầm cột).
3. Đọc sheet "Hướng dẫn nhập liệu" (nếu có) -> gợi ý quy tắc.
4. Đọc từng dòng: Số tài khoản (khóa), Tên NV, Mã ngân hàng, Số tiền hiện có.
5. Kiểm tra trùng Số tài khoản, quy tắc nhập liệu, cột bị lệch.

Module CHỈ ĐỌC, không sửa file.
"""

from __future__ import annotations

import logging
import re
from collections import defaultdict
from pathlib import Path
from typing import Optional

from .excel_reader import (
    PathLike,
    SheetSelectionError,
    WorkbookReader,
    cell_display_text,
    cell_problem,
    check_column_headers,
    select_sheet_name,
    sheet_matches,
)
from .models import (
    CellKind,
    ColumnRef,
    GuideInfo,
    Issue,
    IssueCode,
    Severity,
    SheetData,
    TargetReadResult,
    TargetRow,
    TargetRules,
    TargetTemplate,
)
from .normalize import (
    clean_text,
    has_letters,
    normalize_account,
    normalize_amount,
    normalize_bank_code,
    normalize_name_key,
    remove_diacritics,
)

log = logging.getLogger(__name__)

__all__ = [
    "DEFAULT_TARGET_TEMPLATES",
    "TEMPLATE_ACB",
    "TEMPLATE_THONG_THUONG",
    "check_rules",
    "parse_guide",
    "parse_target_sheet",
    "read_target",
    "select_target_sheet",
]

SUSPECT_MIN_ROWS: int = 3
SUSPECT_BAD_RATIO: float = 0.5


# =============================================================================
# CẤU HÌNH MẶC ĐỊNH CHO CÁC MẪU NGÂN HÀNG HIỆN CÓ
# -----------------------------------------------------------------------------
# Dùng khi chưa có file mapping JSON. Sửa các dòng đánh dấu  <-- MAPPING  nếu mẫu thay đổi.
#   ColumnRef("D", "Số tiền")  =  cột D, tiêu đề mong đợi "Số tiền" (tool DỪNG nếu tiêu đề khác).
# =============================================================================

# Mẫu "4 MauChiHoLuongQuaTaiKhoanACB.xls" - chuyển trong hệ thống ACB
TEMPLATE_ACB = TargetTemplate(
    name="ACB - Chi hộ lương qua tài khoản ACB",
    description="Chuyển lương cho nhân viên có tài khoản ACB.",
    file_patterns=("*MauChiHoLuongQuaTaiKhoanACB*",),   # <-- MAPPING: nhận diện theo tên file
    sheet_name="CNBD",                                   # <-- MAPPING: sheet nhập liệu
    header_rows=(2,),                                    # <-- MAPPING: dòng tiêu đề
    data_start_row=3,                                    # <-- MAPPING: dòng dữ liệu đầu tiên
    account_no=ColumnRef("C", "Số tài khoản", label="Số tài khoản"),   # <-- MAPPING: khóa tìm dòng
    full_name=ColumnRef("B", "Tên NV", label="Tên NV"),                # <-- MAPPING: kiểm tra chéo
    amount=ColumnRef("D", "Số tiền", label="Số tiền"),                 # <-- MAPPING: cột GHI số tiền
    bank_code=None,                                      # mẫu ACB không có cột mã ngân hàng
    rules=TargetRules(no_diacritics=True, no_extra_spaces=True),
)

# Mẫu "5 MauChiHoLuongThongThuong.xls" - chuyển liên ngân hàng
TEMPLATE_THONG_THUONG = TargetTemplate(
    name="Thông thường - Chi hộ lương liên ngân hàng",
    description="Chuyển lương cho nhân viên có tài khoản ở ngân hàng khác (có Mã ngân hàng).",
    file_patterns=("*MauChiHoLuongThongThuong*",),        # <-- MAPPING
    sheet_patterns=("Dán giá trị*",),                     # <-- MAPPING: KHÔNG dùng sheet "Soạn lệnh" (bị khóa)
    header_rows=(2,),
    data_start_row=3,
    account_no=ColumnRef("G", "Số tài khoản", label="Số tài khoản"),   # <-- MAPPING
    full_name=ColumnRef("B", "Tên NV", label="Tên NV"),                # <-- MAPPING
    amount=ColumnRef("E", "Số tiền", label="Số tiền"),                 # <-- MAPPING: cột GHI số tiền
    bank_code=ColumnRef("F", "Mã ngân hàng", label="Mã ngân hàng"),    # <-- MAPPING: chỉ đọc
    rules=TargetRules(no_diacritics=True, no_extra_spaces=True),
)

DEFAULT_TARGET_TEMPLATES: tuple[TargetTemplate, ...] = (TEMPLATE_ACB, TEMPLATE_THONG_THUONG)


# =============================================================================
# 1. Chọn sheet nhập liệu
# =============================================================================

def select_target_sheet(
    wb: WorkbookReader,
    template: TargetTemplate,
    override: Optional[str] = None,
) -> tuple[str, Optional[Issue]]:
    """Chọn sheet nhập liệu. Trả về (tên sheet, thông báo nếu phải tự nhận theo tiêu đề).

    Nếu tên/mẫu không xác định được duy nhất 1 sheet: thử các sheet KHÔNG phải sheet hướng dẫn;
    chỉ tự chọn khi DUY NHẤT 1 sheet có tiêu đề khớp hoàn toàn mapping. Ngược lại báo lỗi để người dùng chọn.
    """
    try:
        return select_sheet_name(wb.sheet_names, template.sheet_name, template.sheet_patterns, override), None
    except SheetSelectionError as exc:
        if override is not None:
            raise
        pool = exc.candidates or wb.sheet_names
        candidates = [n for n in pool if not sheet_matches(n, template.guide_patterns)]
        passing = [
            n for n in candidates
            if not check_column_headers(wb.read_sheet(n), template.header_rows, template.columns())
        ]
        if len(passing) == 1:
            note = Issue(
                IssueCode.SHEET_AUTO_SELECTED, Severity.WARNING,
                f"Không tìm thấy sheet theo cấu hình; đã chọn sheet '{passing[0]}' vì tiêu đề cột khớp mapping.",
                sheet=passing[0],
            )
            return passing[0], note
        raise SheetSelectionError(exc.user_message, candidates=passing or candidates) from exc


# =============================================================================
# 2. Sheet "Hướng dẫn nhập liệu"
# =============================================================================

_RE_MAX_LEN = re.compile(r"toi da\s+(\d+)\s*ky tu")
_RE_FORBIDDEN = re.compile(r"khong (?:gom|chua|nhap|co)\s+(?:dau|ky tu)\s+([^\w\s])")


def _plain(text: str) -> str:
    return remove_diacritics(clean_text(text)).lower()


def parse_guide(sheet: SheetData) -> GuideInfo:
    """Đọc nội dung sheet hướng dẫn và GỢI Ý quy tắc (người dùng phải xác nhận trước khi áp dụng)."""
    lines: list[str] = []
    for r in range(1, sheet.max_row + 1):
        parts = [clean_text(cell_display_text(sheet.cell(r, c))) for c in range(1, sheet.max_col + 1)]
        line = " | ".join(p for p in parts if p)
        if line:
            lines.append(line)

    plain = " ".join(_plain(line) for line in lines)
    max_len_match = _RE_MAX_LEN.search(plain)
    max_len = int(max_len_match.group(1)) if max_len_match else None
    forbidden = "".join(sorted(set(_RE_FORBIDDEN.findall(plain))))
    suggested = TargetRules(
        no_diacritics="khong nhap dau tieng viet" in plain or "khong dau" in plain,
        no_extra_spaces="khoang trang" in plain,
        max_text_length=None,   # giới hạn trong hướng dẫn áp dụng cho "Nội dung giao dịch", không phải tên
        forbidden_chars=forbidden,
    )
    return GuideInfo(sheet_name=sheet.name, lines=tuple(lines), suggested_rules=suggested, max_length_found=max_len)


# =============================================================================
# 3. Kiểm tra quy tắc nhập liệu
# =============================================================================

def check_rules(raw_text: str, rules: TargetRules) -> list[str]:
    """Trả về danh sách vi phạm quy tắc của 1 giá trị văn bản (rỗng = hợp lệ)."""
    problems: list[str] = []
    if rules.no_diacritics and remove_diacritics(raw_text) != raw_text:
        problems.append("có dấu tiếng Việt")
    if rules.no_extra_spaces and raw_text != " ".join(raw_text.split()):
        problems.append("thừa khoảng trắng")
    if rules.max_text_length is not None and len(raw_text) > rules.max_text_length:
        problems.append(f"dài {len(raw_text)} ký tự (tối đa {rules.max_text_length})")
    bad = sorted({ch for ch in raw_text if ch in rules.forbidden_chars})
    if bad:
        problems.append(f"chứa ký tự không cho phép {''.join(bad)!r}")
    return problems


# =============================================================================
# 4. Đọc file đích
# =============================================================================

def read_target(
    path: PathLike,
    template: TargetTemplate,
    sheet_name: Optional[str] = None,
) -> TargetReadResult:
    """Đọc file mẫu ngân hàng.

    Raises:
        ConfigError         : cấu hình không hợp lệ.
        ExcelReadError      : không đọc được file.
        SheetSelectionError : không xác định được sheet (giao diện cần hỏi người dùng).
    """
    template.ensure_valid()
    with WorkbookReader(path) as wb:
        chosen, note = select_target_sheet(wb, template, sheet_name)
        sheet = wb.read_sheet(chosen)
        guide_names = [n for n in wb.sheet_names if n != chosen and sheet_matches(n, template.guide_patterns)]
        guide = parse_guide(wb.read_sheet(guide_names[0])) if guide_names else None
        file_format = wb.format.value
    log.info("Đọc file ngân hàng '%s', sheet '%s'", Path(path).name, chosen)

    result = parse_target_sheet(sheet, template, file_path=str(path), file_format=file_format)
    result.guide = guide
    if note is not None:
        result.issues.insert(0, note)
    if guide is not None:
        result.issues.insert(0, Issue(
            IssueCode.GUIDE_SHEET_FOUND, Severity.INFO,
            f"File có sheet hướng dẫn '{guide.sheet_name}' ({len(guide.lines)} dòng). "
            "Hãy xem và xác nhận quy tắc nhập liệu.",
            sheet=guide.sheet_name,
        ))
    return result


def parse_target_sheet(
    sheet: SheetData,
    template: TargetTemplate,
    file_path: str = "",
    file_format: str = "",
) -> TargetReadResult:
    """Phân tích sheet nhập liệu đã đọc (tách riêng để dễ kiểm thử)."""
    result = TargetReadResult(
        file_path=file_path, file_format=file_format, sheet_name=sheet.name, template_name=template.name
    )
    header_issues = check_column_headers(sheet, template.header_rows, template.columns())
    if header_issues:
        result.header_ok = False
        result.issues.extend(header_issues)
        return result

    cols = template.columns()
    col_indexes = [ref.index for ref in cols.values()]
    last_row = template.data_end_row if template.data_end_row is not None else sheet.max_row
    unusual_acc_rows: set[int] = set()

    for r in range(template.data_start_row, last_row + 1):
        if all(sheet.cell(r, c).is_empty for c in col_indexes):
            continue  # dòng trống

        row_issues: list[Issue] = []

        def add(code: IssueCode, sev: Severity, msg: str, ref: ColumnRef) -> None:
            row_issues.append(Issue(code, sev, msg, sheet=sheet.name, row=r, column=ref.letter))

        # ---------- Số tài khoản (khóa) ----------
        acc_ref = template.account_no
        acc_cell = sheet.cell(r, acc_ref.index)
        account_no: Optional[str] = None
        problem = cell_problem(acc_cell)
        if problem:
            add(problem[0], Severity.ERROR, f"Ô Số tài khoản: {problem[1]}", acc_ref)
        elif acc_cell.is_empty:
            add(IssueCode.TARGET_ROW_INCOMPLETE, Severity.WARNING,
                "Dòng có dữ liệu nhưng không có Số tài khoản -> tool sẽ không ghi vào dòng này.", acc_ref)
        else:
            res = normalize_account(acc_cell.value)
            if res.code is not None:
                add(res.code, res.severity or Severity.ERROR, res.detail, acc_ref)
                if res.code in (IssueCode.ACCOUNT_UNUSUAL_LENGTH, IssueCode.ACCOUNT_INVALID):
                    unusual_acc_rows.add(r)
            account_no = None if res.is_error else res.value

        # ---------- Tên NV ----------
        name_ref = template.full_name
        name_cell = sheet.cell(r, name_ref.index)
        raw_name = ""
        problem = cell_problem(name_cell)
        if problem:
            add(problem[0], Severity.WARNING, f"Ô Tên NV: {problem[1]}", name_ref)
        elif not name_cell.is_empty:
            raw_name = cell_display_text(name_cell)
            violations = check_rules(raw_name, template.rules)
            if violations:
                add(IssueCode.RULE_VIOLATION, Severity.WARNING,
                    f"Tên '{raw_name}' vi phạm hướng dẫn nhập liệu: {', '.join(violations)}.", name_ref)
        full_name = clean_text(raw_name)
        if not full_name and not problem:
            add(IssueCode.NAME_MISSING, Severity.WARNING, "Thiếu Tên NV.", name_ref)

        # ---------- Mã ngân hàng (tùy chọn) ----------
        bank_code: Optional[str] = None
        if template.bank_code is not None:
            bank_ref = template.bank_code
            bank_cell = sheet.cell(r, bank_ref.index)
            problem = cell_problem(bank_cell)
            if problem:
                add(problem[0], Severity.ERROR, f"Ô Mã ngân hàng: {problem[1]}", bank_ref)
            else:
                bres = normalize_bank_code(bank_cell.value)
                if bres.is_error:
                    add(bres.code or IssueCode.BANK_CODE_INVALID, Severity.ERROR, bres.detail, bank_ref)
                elif bres.value is None:
                    add(IssueCode.BANK_CODE_MISSING, Severity.ERROR, "Thiếu Mã ngân hàng.", bank_ref)
                bank_code = bres.value

        # ---------- Số tiền hiện có (sẽ bị ghi đè) ----------
        amt_ref = template.amount
        amt_cell = sheet.cell(r, amt_ref.index)
        current_amount: Optional[int] = None
        if amt_cell.is_formula:
            # Ghi đè sẽ làm mất công thức -> không cho ghi dòng này
            add(IssueCode.TARGET_AMOUNT_IS_FORMULA, Severity.ERROR,
                "Ô Số tiền đang chứa công thức; tool không ghi đè công thức.", amt_ref)
        elif amt_cell.kind == CellKind.ERROR:
            add(IssueCode.EXCEL_ERROR_VALUE, Severity.WARNING,
                f"Ô Số tiền hiện đang báo lỗi '{amt_cell.value}' (sẽ được ghi đè).", amt_ref)
        else:
            ares = normalize_amount(amt_cell.value)
            current_amount = ares.value if not ares.is_error else None
            if ares.is_error:
                add(ares.code or IssueCode.AMOUNT_NOT_NUMBER, Severity.WARNING,
                    f"Số tiền cũ không hợp lệ (sẽ được ghi đè): {ares.detail}", amt_ref)

        has_error = any(i.severity == Severity.ERROR for i in row_issues)
        result.issues.extend(row_issues)
        result.rows.append(TargetRow(
            row=r,
            account_no=account_no,
            full_name=full_name,
            name_key=normalize_name_key(full_name),
            bank_code=bank_code,
            current_amount=current_amount,
            account_stored_as_number=acc_cell.kind == CellKind.NUMBER,
            has_error=has_error,
        ))

    if not result.rows:
        result.issues.append(Issue(
            IssueCode.NO_RECORDS, Severity.ERROR,
            f"Không có dòng dữ liệu nào từ dòng {template.data_start_row}. "
            "V1 chỉ cập nhật dòng có sẵn -> kiểm tra lại file hoặc mapping.",
            sheet=sheet.name,
        ))

    result.issues.extend(_duplicate_issues(result.rows, sheet.name, template))
    result.issues.extend(_column_content_issues(result.rows, unusual_acc_rows, sheet.name, template))
    return result


# =============================================================================
# 5. Hàm phụ
# =============================================================================

def _duplicate_issues(rows: list[TargetRow], sheet_name: str, template: TargetTemplate) -> list[Issue]:
    by_acc: dict[str, list[int]] = defaultdict(list)
    for row in rows:
        if row.account_no:
            by_acc[row.account_no].append(row.row)
    return [
        Issue(
            IssueCode.DUPLICATE_ACCOUNT, Severity.ERROR,
            f"Số tài khoản '{acc}' xuất hiện nhiều lần ở các dòng {lines} trong file ngân hàng. "
            "Phải sửa file trước khi tiếp tục.",
            sheet=sheet_name, column=template.account_no.letter,
        )
        for acc, lines in by_acc.items() if len(lines) > 1
    ]


def _column_content_issues(
    rows: list[TargetRow], unusual_acc_rows: set[int], sheet_name: str, template: TargetTemplate
) -> list[Issue]:
    total = len(rows)
    if total < SUSPECT_MIN_ROWS:
        return []
    checks = {
        "account_no": (
            sum(1 for r in rows if r.account_no is None or r.row in unusual_acc_rows),
            "trống, sai định dạng hoặc độ dài bất thường",
        ),
        "full_name": (sum(1 for r in rows if not has_letters(r.full_name)), "trống hoặc không chứa chữ cái"),
    }
    issues: list[Issue] = []
    for key, (bad, desc) in checks.items():
        if bad / total > SUSPECT_BAD_RATIO:
            ref = template.columns()[key]
            issues.append(Issue(
                IssueCode.COLUMN_CONTENT_SUSPECT, Severity.ERROR,
                f"{bad}/{total} dòng có {ref.display()} {desc}. Có thể mapping sai cột -> hãy kiểm tra lại.",
                sheet=sheet_name, column=ref.letter,
            ))
    return issues
