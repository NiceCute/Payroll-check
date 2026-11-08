"""
models.py - Định nghĩa các đối tượng dữ liệu dùng chung cho toàn bộ tool.

Module này KHÔNG phụ thuộc module nào khác trong dự án (chỉ dùng thư viện chuẩn),
để mọi module khác có thể import an toàn.

Quy ước:
- Số dòng / số cột luôn đánh số giống Excel: bắt đầu từ 1.
- Định danh (Số thẻ, Số tài khoản, Mã ngân hàng) luôn là chuỗi (str), không bao giờ là số.
- Số tiền luôn là số nguyên (int), đơn vị VND.
"""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, time, timedelta
from enum import Enum
from typing import Any, Optional, Union


# =============================================================================
# 1. Lỗi dùng chung
# =============================================================================

class ToolError(Exception):
    """Lỗi nghiệp vụ có thông báo tiếng Việt dễ hiểu để hiển thị cho người dùng."""

    def __init__(self, user_message: str, *, detail: str = "") -> None:
        super().__init__(user_message)
        self.user_message: str = user_message
        self.detail: str = detail  # thông tin kỹ thuật, chỉ ghi vào log

    def __str__(self) -> str:
        return self.user_message if not self.detail else f"{self.user_message} ({self.detail})"


class ConfigError(ToolError):
    """Cấu hình mapping (profile/template) không hợp lệ."""


# =============================================================================
# 2. Mức độ & mã vấn đề
# =============================================================================

class Severity(str, Enum):
    ERROR = "ERROR"      # Chặn ghi file
    WARNING = "WARNING"  # Cần người dùng xem / xác nhận
    INFO = "INFO"        # Chỉ thông báo


class IssueCode(str, Enum):
    # --- Cấu trúc file / mapping ---
    HEADER_MISMATCH = "HEADER_MISMATCH"
    COLUMN_CONTENT_SUSPECT = "COLUMN_CONTENT_SUSPECT"
    NO_RECORDS = "NO_RECORDS"
    FORMULA_NOT_CALCULATED = "FORMULA_NOT_CALCULATED"
    EXCEL_ERROR_VALUE = "EXCEL_ERROR_VALUE"
    # --- Số thẻ ---
    CARD_STORED_AS_NUMBER = "CARD_STORED_AS_NUMBER"
    CARD_INVALID = "CARD_INVALID"
    DUPLICATE_CARD = "DUPLICATE_CARD"
    # --- Số tài khoản ---
    ACCOUNT_MISSING = "ACCOUNT_MISSING"
    ACCOUNT_INVALID = "ACCOUNT_INVALID"
    ACCOUNT_UNUSUAL_LENGTH = "ACCOUNT_UNUSUAL_LENGTH"
    DUPLICATE_ACCOUNT = "DUPLICATE_ACCOUNT"
    # --- Họ tên ---
    NAME_MISSING = "NAME_MISSING"
    NAME_LEGACY_FONT = "NAME_LEGACY_FONT"
    DUPLICATE_NAME = "DUPLICATE_NAME"
    # --- Số tiền ---
    AMOUNT_MISSING = "AMOUNT_MISSING"
    AMOUNT_NOT_NUMBER = "AMOUNT_NOT_NUMBER"
    AMOUNT_NOT_INTEGER = "AMOUNT_NOT_INTEGER"
    AMOUNT_NEGATIVE = "AMOUNT_NEGATIVE"
    AMOUNT_ZERO = "AMOUNT_ZERO"
    AMOUNT_UNUSUALLY_LARGE = "AMOUNT_UNUSUALLY_LARGE"
    # --- File đích (mẫu ngân hàng) ---
    BANK_CODE_MISSING = "BANK_CODE_MISSING"
    BANK_CODE_INVALID = "BANK_CODE_INVALID"
    TARGET_ROW_INCOMPLETE = "TARGET_ROW_INCOMPLETE"
    TARGET_AMOUNT_IS_FORMULA = "TARGET_AMOUNT_IS_FORMULA"
    RULE_VIOLATION = "RULE_VIOLATION"
    GUIDE_SHEET_FOUND = "GUIDE_SHEET_FOUND"
    SHEET_AUTO_SELECTED = "SHEET_AUTO_SELECTED"


# Các mã lỗi làm cả file không dùng được (chặn toàn bộ quá trình ghi)
BLOCKING_CODES: frozenset[IssueCode] = frozenset({
    IssueCode.HEADER_MISMATCH,
    IssueCode.COLUMN_CONTENT_SUSPECT,
    IssueCode.NO_RECORDS,
    IssueCode.DUPLICATE_CARD,
    IssueCode.DUPLICATE_ACCOUNT,
})


@dataclass(frozen=True)
class Issue:
    """Một vấn đề phát hiện được khi đọc / kiểm tra dữ liệu."""

    code: IssueCode
    severity: Severity
    message: str
    sheet: Optional[str] = None
    row: Optional[int] = None
    column: Optional[str] = None   # chữ cái cột, ví dụ "P"
    card_no: Optional[str] = None

    def location(self) -> str:
        parts: list[str] = []
        if self.sheet:
            parts.append(f"sheet '{self.sheet}'")
        if self.column and self.row:
            parts.append(f"ô {self.column}{self.row}")
        elif self.row:
            parts.append(f"dòng {self.row}")
        elif self.column:
            parts.append(f"cột {self.column}")
        return ", ".join(parts)

    def __str__(self) -> str:
        loc = self.location()
        return f"[{self.severity.value}] {self.message}" + (f" ({loc})" if loc else "")


# =============================================================================
# 3. Chuyển đổi chữ cái cột <-> số thứ tự cột
# =============================================================================

_COLUMN_RE = re.compile(r"^[A-Z]{1,3}$")
MAX_EXCEL_COLUMN = 16384  # XFD


def column_index(letter: str) -> int:
    """'A' -> 1, 'P' -> 16, 'AA' -> 27. Báo lỗi nếu không phải chữ cái cột hợp lệ."""
    if not isinstance(letter, str):
        raise ConfigError(f"Tên cột phải là chữ cái (ví dụ 'P'), nhận được: {letter!r}")
    text = letter.strip().upper()
    if not _COLUMN_RE.match(text):
        raise ConfigError(f"Tên cột không hợp lệ: {letter!r}. Ví dụ hợp lệ: A, P, AB")
    result = 0
    for ch in text:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    if result > MAX_EXCEL_COLUMN:
        raise ConfigError(f"Tên cột vượt quá giới hạn Excel (XFD): {letter!r}")
    return result


def column_letter(index: int) -> str:
    """1 -> 'A', 16 -> 'P', 27 -> 'AA'."""
    if not isinstance(index, int) or isinstance(index, bool) or not 1 <= index <= MAX_EXCEL_COLUMN:
        raise ConfigError(f"Số thứ tự cột không hợp lệ: {index!r}")
    letters = ""
    n = index
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


# =============================================================================
# 4. Dữ liệu ô / sheet (kết quả của excel_reader)
# =============================================================================

class CellKind(str, Enum):
    EMPTY = "EMPTY"
    TEXT = "TEXT"
    NUMBER = "NUMBER"
    BOOL = "BOOL"
    DATE = "DATE"
    ERROR = "ERROR"   # #N/A, #VALUE!, #REF! ...


@dataclass(frozen=True, slots=True)
class CellData:
    """Giá trị 1 ô sau khi đọc.

    is_formula = True nghĩa là ô chứa công thức (chỉ phát hiện được với .xlsx).
    Nếu is_formula = True và kind = EMPTY: công thức CHƯA được Excel tính/lưu giá trị.
    """

    value: Any
    kind: CellKind
    is_formula: bool = False

    @property
    def is_empty(self) -> bool:
        if self.kind == CellKind.EMPTY:
            return True
        return self.kind == CellKind.TEXT and str(self.value).strip() == ""

    @property
    def formula_not_calculated(self) -> bool:
        return self.is_formula and self.kind == CellKind.EMPTY


EMPTY_CELL = CellData(None, CellKind.EMPTY)


def classify_value(value: Any, *, is_error: bool = False) -> CellKind:
    """Xác định loại ô theo kiểu dữ liệu Python."""
    if value is None:
        return CellKind.EMPTY
    if is_error:
        return CellKind.ERROR
    if isinstance(value, bool):
        return CellKind.BOOL
    if isinstance(value, (int, float)):
        return CellKind.NUMBER
    if isinstance(value, (datetime, date, time, timedelta)):
        return CellKind.DATE
    return CellKind.TEXT


@dataclass(frozen=True)
class SheetInfo:
    name: str
    visible: bool


@dataclass
class SheetData:
    """Nội dung 1 sheet, lưu dạng thưa (chỉ ô có dữ liệu hoặc có công thức)."""

    info: SheetInfo
    cells: dict[tuple[int, int], CellData] = field(default_factory=dict)
    max_row: int = 0
    max_col: int = 0

    @property
    def name(self) -> str:
        return self.info.name

    def cell(self, row: int, col: int) -> CellData:
        """Lấy ô theo (dòng, cột), đánh số từ 1. Ô ngoài vùng dữ liệu trả về ô trống."""
        return self.cells.get((row, col), EMPTY_CELL)

    def cell_at(self, row: int, letter: str) -> CellData:
        return self.cell(row, column_index(letter))

    def row_is_empty(self, row: int, cols: Optional[list[int]] = None) -> bool:
        columns = cols if cols is not None else range(1, self.max_col + 1)
        return all(self.cell(row, c).is_empty for c in columns)


# =============================================================================
# 5. Cấu hình mapping file nguồn (file lương)
# =============================================================================

@dataclass(frozen=True)
class ColumnRef:
    """Tham chiếu tới 1 cột trong file Excel.

    letter          : chữ cái cột, ví dụ "P".
    expected_header : tiêu đề mong đợi tại cột đó (để phát hiện file bị chèn/xóa cột).
                      Để None nếu cột không có tiêu đề (ví dụ cột số tài khoản trong file lương).
    label           : tên hiển thị trên giao diện / báo cáo.
    """

    letter: str
    expected_header: Optional[str] = None
    label: str = ""

    def __post_init__(self) -> None:
        # Chuẩn hóa chữ cái + kiểm tra hợp lệ ngay khi tạo
        object.__setattr__(self, "letter", column_letter(column_index(self.letter)))

    @property
    def index(self) -> int:
        return column_index(self.letter)

    def display(self) -> str:
        name = self.label or self.expected_header or ""
        return f"{name} (cột {self.letter})" if name else f"cột {self.letter}"

    def to_dict(self) -> dict[str, Any]:
        return {"letter": self.letter, "expected_header": self.expected_header, "label": self.label}

    @staticmethod
    def from_dict(data: dict[str, Any]) -> "ColumnRef":
        if not isinstance(data, dict) or "letter" not in data:
            raise ConfigError(f"Cấu hình cột không hợp lệ: {data!r}")
        header = data.get("expected_header")
        return ColumnRef(
            letter=str(data["letter"]),
            expected_header=_opt_str(header),
            label=str(data.get("label") or ""),
        )


def _is_positive_int(value: Any) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value >= 1


def _validate_layout(
    *,
    name: str,
    sheet_name: Optional[str],
    sheet_patterns: tuple[str, ...],
    header_rows: tuple[int, ...],
    data_start_row: int,
    data_end_row: Optional[int],
    columns: dict[str, ColumnRef],
    labels: dict[str, str],
) -> list[str]:
    """Kiểm tra chung cho cấu hình file lương và mẫu ngân hàng."""
    problems: list[str] = []
    if not isinstance(name, str) or not name.strip():
        problems.append("Tên cấu hình không được để trống.")
    if not sheet_name and not sheet_patterns:
        problems.append("Phải khai báo tên sheet hoặc mẫu tên sheet.")
    if not header_rows:
        problems.append("Phải khai báo ít nhất 1 dòng tiêu đề.")
    valid_header_rows: list[int] = []
    for r in header_rows:
        if _is_positive_int(r):
            valid_header_rows.append(r)
        else:
            problems.append(f"Dòng tiêu đề không hợp lệ: {r!r}")
    if not _is_positive_int(data_start_row):
        problems.append(f"Dòng bắt đầu dữ liệu không hợp lệ: {data_start_row!r}")
    else:
        if valid_header_rows and data_start_row <= max(valid_header_rows):
            problems.append("Dòng bắt đầu dữ liệu phải nằm dưới các dòng tiêu đề.")
        if data_end_row is not None and (not _is_positive_int(data_end_row) or data_end_row < data_start_row):
            problems.append("Dòng kết thúc dữ liệu phải lớn hơn hoặc bằng dòng bắt đầu.")
    # Không cho 2 trường dùng chung 1 cột -> tránh nhầm số tiền với số tài khoản
    used: dict[str, str] = {}
    for key, ref in columns.items():
        if ref.letter in used:
            problems.append(
                f"Cột {ref.letter} đang được dùng cho cả '{labels.get(used[ref.letter], used[ref.letter])}' "
                f"và '{labels.get(key, key)}'."
            )
        used[ref.letter] = key
    return problems


def _opt_str(value: Any) -> Optional[str]:
    return str(value) if value not in (None, "") else None


def _str_tuple(value: Any) -> tuple[str, ...]:
    if value in (None, ""):
        return ()
    if isinstance(value, str):
        return (value,)
    if not isinstance(value, (list, tuple)):
        raise TypeError(f"Cần danh sách, nhận được {type(value).__name__}")
    return tuple(str(v) for v in value)


def _int_tuple(value: Any) -> tuple[int, ...]:
    if value in (None, ""):
        return ()
    if not isinstance(value, (list, tuple)):
        raise TypeError(f"Cần danh sách số dòng, nhận được {type(value).__name__}")
    return tuple(int(v) for v in value)


# Tên các trường bắt buộc trong file nguồn (dùng cho validate & giao diện)
SOURCE_FIELDS: dict[str, str] = {
    "card_no": "Số thẻ",
    "full_name": "Họ và tên",
    "account_no": "Số tài khoản",
    "amount": "Số tiền chuyển khoản",
}


@dataclass(frozen=True)
class SourceProfile:
    """Cấu hình cách đọc file lương.

    sheet_name     : tên sheet cố định. Nếu None -> tìm theo sheet_patterns.
    sheet_patterns : mẫu tên sheet, '*' = chuỗi bất kỳ, '?' = 1 ký tự. Không phân biệt hoa thường.
    header_rows    : các dòng chứa tiêu đề cột (có thể nhiều dòng do merge).
    data_start_row : dòng dữ liệu đầu tiên.
    data_end_row   : dòng dữ liệu cuối cùng; None = đọc tới hết sheet.
    """

    name: str
    card_no: ColumnRef
    full_name: ColumnRef
    account_no: ColumnRef
    amount: ColumnRef
    header_rows: tuple[int, ...] = (1,)
    data_start_row: int = 2
    data_end_row: Optional[int] = None
    sheet_name: Optional[str] = None
    sheet_patterns: tuple[str, ...] = ()

    def columns(self) -> dict[str, ColumnRef]:
        return {
            "card_no": self.card_no,
            "full_name": self.full_name,
            "account_no": self.account_no,
            "amount": self.amount,
        }

    def validate(self) -> list[str]:
        """Trả về danh sách lỗi cấu hình (rỗng = hợp lệ)."""
        return _validate_layout(
            name=self.name,
            sheet_name=self.sheet_name,
            sheet_patterns=self.sheet_patterns,
            header_rows=self.header_rows,
            data_start_row=self.data_start_row,
            data_end_row=self.data_end_row,
            columns=self.columns(),
            labels=SOURCE_FIELDS,
        )

    def ensure_valid(self) -> None:
        problems = self.validate()
        if problems:
            raise ConfigError("Cấu hình file lương không hợp lệ:\n- " + "\n- ".join(problems))

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "sheet_name": self.sheet_name,
            "sheet_patterns": list(self.sheet_patterns),
            "header_rows": list(self.header_rows),
            "data_start_row": self.data_start_row,
            "data_end_row": self.data_end_row,
            "columns": {k: v.to_dict() for k, v in self.columns().items()},
        }

    @staticmethod
    def from_dict(data: dict[str, Any]) -> "SourceProfile":
        try:
            cols = data["columns"]
            end_row = data.get("data_end_row")
            profile = SourceProfile(
                name=str(data["name"]),
                sheet_name=_opt_str(data.get("sheet_name")),
                sheet_patterns=_str_tuple(data.get("sheet_patterns")),
                header_rows=_int_tuple(data.get("header_rows")),
                data_start_row=int(data["data_start_row"]),
                # 0 không được coi là "không giới hạn" -> để validate() báo lỗi
                data_end_row=int(end_row) if end_row is not None else None,
                card_no=ColumnRef.from_dict(cols["card_no"]),
                full_name=ColumnRef.from_dict(cols["full_name"]),
                account_no=ColumnRef.from_dict(cols["account_no"]),
                amount=ColumnRef.from_dict(cols["amount"]),
            )
        except ConfigError:
            raise
        except (KeyError, TypeError, ValueError) as exc:
            raise ConfigError("File cấu hình file lương bị thiếu hoặc sai thông tin.", detail=repr(exc)) from exc
        profile.ensure_valid()
        return profile


# =============================================================================
# 6. Kết quả đọc file nguồn
# =============================================================================

@dataclass(frozen=True)
class SourceRecord:
    """1 nhân sự đọc từ file lương (dòng có Số thẻ)."""

    row: int                       # số dòng trong Excel
    card_no: str                   # Số thẻ - khóa định danh chính
    full_name: str                 # tên đã làm sạch khoảng trắng (giữ nguyên chữ)
    name_key: str                  # tên không dấu, viết hoa - CHỈ dùng để kiểm tra chéo
    account_no: Optional[str]      # None nếu trống / không hợp lệ
    amount: Optional[int]          # None nếu trống / không hợp lệ
    has_error: bool = False        # True nếu dòng có ít nhất 1 lỗi mức ERROR

    @property
    def is_transferable(self) -> bool:
        """Đủ điều kiện để đưa vào file chuyển khoản (chưa xét file đích).

        Số tiền phải > 0: ngân hàng không nhận lệnh chuyển 0 đồng.
        """
        return (
            not self.has_error
            and self.account_no is not None
            and self.amount is not None
            and self.amount > 0
        )


@dataclass(frozen=True)
class SkippedRow:
    """Dòng bị bỏ qua: không có Số thẻ (nhưng có dữ liệu ở cột khác) hoặc ô Số thẻ bị lỗi."""

    row: int
    reason: str
    snapshot: dict[str, Any]


@dataclass
class SourceReadResult:
    file_path: str
    sheet_name: str
    profile_name: str
    records: list[SourceRecord] = field(default_factory=list)
    skipped: list[SkippedRow] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    header_ok: bool = True

    @property
    def errors(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == Severity.ERROR]

    @property
    def warnings(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == Severity.WARNING]

    @property
    def blocking_issues(self) -> list[Issue]:
        """Lỗi mức FILE (sai mapping, trùng định danh) -> không được dùng kết quả để ghi file.

        Lỗi mức DÒNG (thiếu TK, số tiền sai...) không nằm ở đây: dòng đó bị đánh dấu
        has_error / không đủ điều kiện và sẽ hiển thị trong Exception List.
        """
        return [i for i in self.issues if i.code in BLOCKING_CODES]

    @property
    def is_usable(self) -> bool:
        return not self.blocking_issues

    @property
    def transferable(self) -> list[SourceRecord]:
        return [r for r in self.records if r.is_transferable]

    @property
    def total_transferable_amount(self) -> int:
        return sum(r.amount or 0 for r in self.transferable)

    def summary(self) -> dict[str, int]:
        return {
            "records": len(self.records),
            "transferable": len(self.transferable),
            "missing_account": sum(1 for r in self.records if r.account_no is None),
            "zero_amount": sum(1 for r in self.records if r.amount == 0),
            "rows_with_error": sum(1 for r in self.records if r.has_error),
            "skipped_rows": len(self.skipped),
            "errors": len(self.errors),
            "warnings": len(self.warnings),
            "total_transferable_amount": self.total_transferable_amount,
        }


# =============================================================================
# 7. Cấu hình mẫu ngân hàng (file đích)
# =============================================================================

TARGET_FIELDS: dict[str, str] = {
    "account_no": "Số tài khoản",
    "full_name": "Tên nhân viên",
    "amount": "Số tiền",
    "bank_code": "Mã ngân hàng",
}

DEFAULT_GUIDE_PATTERNS: tuple[str, ...] = ("Hướng dẫn*", "Huong dan*")


@dataclass(frozen=True)
class TargetRules:
    """Quy tắc nhập liệu của ngân hàng (lấy từ sheet "Hướng dẫn nhập liệu", người dùng xác nhận).

    Áp dụng cho cột Tên NV (cột văn bản duy nhất tool kiểm tra trong V1).
    Số tiền luôn phải là số nguyên > 0 (không cấu hình được, vì an toàn).
    """

    no_diacritics: bool = False        # không nhập dấu tiếng Việt
    no_extra_spaces: bool = False      # không thừa khoảng trắng (đầu/cuối/giữa)
    max_text_length: Optional[int] = None
    forbidden_chars: str = ""          # ví dụ "#"

    def validate(self) -> list[str]:
        if self.max_text_length is not None and not _is_positive_int(self.max_text_length):
            return [f"Độ dài tối đa không hợp lệ: {self.max_text_length!r}"]
        return []

    def to_dict(self) -> dict[str, Any]:
        return {
            "no_diacritics": self.no_diacritics,
            "no_extra_spaces": self.no_extra_spaces,
            "max_text_length": self.max_text_length,
            "forbidden_chars": self.forbidden_chars,
        }

    @staticmethod
    def from_dict(data: Any) -> "TargetRules":
        if data in (None, ""):
            return TargetRules()
        if not isinstance(data, dict):
            raise ConfigError(f"Cấu hình quy tắc nhập liệu không hợp lệ: {data!r}")
        max_len = data.get("max_text_length")
        return TargetRules(
            no_diacritics=bool(data.get("no_diacritics", False)),
            no_extra_spaces=bool(data.get("no_extra_spaces", False)),
            max_text_length=int(max_len) if max_len not in (None, "") else None,
            forbidden_chars=str(data.get("forbidden_chars") or ""),
        )


@dataclass(frozen=True)
class TargetTemplate:
    """Cấu hình 1 mẫu biểu ngân hàng.

    file_patterns  : mẫu tên file để tự nhận diện (vd "*MauChiHoLuongQuaTaiKhoanACB*").
    account_no     : cột khóa để tìm dòng (BẮT BUỘC có tiêu đề mong đợi).
    full_name      : cột tên - chỉ để kiểm tra chéo.
    amount         : cột sẽ được ghi số tiền (BẮT BUỘC có tiêu đề mong đợi).
    bank_code      : cột mã ngân hàng (tùy chọn, chỉ đọc).
    guide_patterns : mẫu tên sheet hướng dẫn nhập liệu.
    """

    name: str
    account_no: ColumnRef
    full_name: ColumnRef
    amount: ColumnRef
    bank_code: Optional[ColumnRef] = None
    header_rows: tuple[int, ...] = (1,)
    data_start_row: int = 2
    data_end_row: Optional[int] = None
    sheet_name: Optional[str] = None
    sheet_patterns: tuple[str, ...] = ()
    file_patterns: tuple[str, ...] = ()
    guide_patterns: tuple[str, ...] = DEFAULT_GUIDE_PATTERNS
    rules: TargetRules = field(default_factory=TargetRules)
    description: str = ""

    def columns(self) -> dict[str, ColumnRef]:
        cols = {"account_no": self.account_no, "full_name": self.full_name, "amount": self.amount}
        if self.bank_code is not None:
            cols["bank_code"] = self.bank_code
        return cols

    def validate(self) -> list[str]:
        problems = _validate_layout(
            name=self.name,
            sheet_name=self.sheet_name,
            sheet_patterns=self.sheet_patterns,
            header_rows=self.header_rows,
            data_start_row=self.data_start_row,
            data_end_row=self.data_end_row,
            columns=self.columns(),
            labels=TARGET_FIELDS,
        )
        # File đích luôn có tiêu đề -> bắt buộc khai báo để chống ghi nhầm cột
        for key in ("account_no", "amount"):
            if not self.columns()[key].expected_header:
                problems.append(f"Cột '{TARGET_FIELDS[key]}' bắt buộc phải khai báo tiêu đề mong đợi.")
        problems.extend(self.rules.validate())
        return problems

    def ensure_valid(self) -> None:
        problems = self.validate()
        if problems:
            raise ConfigError(f"Cấu hình mẫu '{self.name}' không hợp lệ:\n- " + "\n- ".join(problems))

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "description": self.description,
            "file_patterns": list(self.file_patterns),
            "sheet_name": self.sheet_name,
            "sheet_patterns": list(self.sheet_patterns),
            "guide_patterns": list(self.guide_patterns),
            "header_rows": list(self.header_rows),
            "data_start_row": self.data_start_row,
            "data_end_row": self.data_end_row,
            "columns": {k: v.to_dict() for k, v in self.columns().items()},
            "rules": self.rules.to_dict(),
        }

    @staticmethod
    def from_dict(data: dict[str, Any]) -> "TargetTemplate":
        try:
            cols = data["columns"]
            if not isinstance(cols, dict):
                raise TypeError("columns phải là đối tượng")
            end_row = data.get("data_end_row")
            guide = data.get("guide_patterns")
            template = TargetTemplate(
                name=str(data["name"]),
                description=str(data.get("description") or ""),
                file_patterns=_str_tuple(data.get("file_patterns")),
                sheet_name=_opt_str(data.get("sheet_name")),
                sheet_patterns=_str_tuple(data.get("sheet_patterns")),
                guide_patterns=_str_tuple(guide) if guide is not None else DEFAULT_GUIDE_PATTERNS,
                header_rows=_int_tuple(data.get("header_rows")),
                data_start_row=int(data["data_start_row"]),
                data_end_row=int(end_row) if end_row is not None else None,
                account_no=ColumnRef.from_dict(cols["account_no"]),
                full_name=ColumnRef.from_dict(cols["full_name"]),
                amount=ColumnRef.from_dict(cols["amount"]),
                bank_code=ColumnRef.from_dict(cols["bank_code"]) if cols.get("bank_code") else None,
                rules=TargetRules.from_dict(data.get("rules")),
            )
        except ConfigError:
            raise
        except (KeyError, TypeError, ValueError) as exc:
            raise ConfigError("File cấu hình mẫu ngân hàng bị thiếu hoặc sai thông tin.", detail=repr(exc)) from exc
        template.ensure_valid()
        return template


# =============================================================================
# 8. Kết quả đọc file đích
# =============================================================================

@dataclass(frozen=True)
class GuideInfo:
    """Nội dung sheet "Hướng dẫn nhập liệu" + quy tắc gợi ý (người dùng cần xác nhận)."""

    sheet_name: str
    lines: tuple[str, ...]
    suggested_rules: TargetRules
    max_length_found: Optional[int] = None


@dataclass(frozen=True)
class TargetRow:
    """1 dòng dữ liệu đang có trong file ngân hàng."""

    row: int
    account_no: Optional[str]
    full_name: str
    name_key: str
    bank_code: Optional[str]
    current_amount: Optional[int]          # số tiền đang có (tháng trước), None nếu trống/không hợp lệ
    account_stored_as_number: bool = False  # ô TK lưu dạng số -> có thể đã mất số 0 đầu
    has_error: bool = False                 # lỗi định danh / công thức -> KHÔNG được ghi vào dòng này

    @property
    def is_writable(self) -> bool:
        return not self.has_error and self.account_no is not None


@dataclass
class TargetReadResult:
    file_path: str
    file_format: str
    sheet_name: str
    template_name: str
    rows: list[TargetRow] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    guide: Optional[GuideInfo] = None
    header_ok: bool = True

    @property
    def errors(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == Severity.ERROR]

    @property
    def warnings(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == Severity.WARNING]

    @property
    def blocking_issues(self) -> list[Issue]:
        return [i for i in self.issues if i.code in BLOCKING_CODES]

    @property
    def is_usable(self) -> bool:
        return not self.blocking_issues

    def by_account(self) -> dict[str, TargetRow]:
        """{số TK: dòng}. Chỉ an toàn khi is_usable (không trùng TK)."""
        return {r.account_no: r for r in self.rows if r.account_no is not None}

    def summary(self) -> dict[str, int]:
        return {
            "rows": len(self.rows),
            "writable": sum(1 for r in self.rows if r.is_writable),
            "rows_with_error": sum(1 for r in self.rows if r.has_error),
            "missing_account": sum(1 for r in self.rows if r.account_no is None),
            "errors": len(self.errors),
            "warnings": len(self.warnings),
            "current_total_amount": sum(r.current_amount or 0 for r in self.rows),
        }


# Kiểu dùng chung đại diện cho cấu hình mapping nguồn hoặc đích
Template = Union[SourceProfile, TargetTemplate]


# =============================================================================
# 9. Đối chiếu (Reconciler), Kế hoạch ghi (WritePlan) & Ngoại lệ (ExceptionItem)
# =============================================================================

@dataclass(frozen=True)
class MasterRecord:
    """Bản ghi Master Dataset chuẩn (dùng cho V2; V1 tùy chọn để cắm sẵn kiến trúc).

    - card_no     : Số thẻ (khóa chính)
    - full_name   : Họ tên chuẩn
    - account_no  : Số tài khoản chuẩn
    - bank_code   : Mã ngân hàng (None nếu chuyển nội bộ ACB)
    - target_type : Tên mẫu / loại file đích tương ứng (tùy chọn)
    """

    card_no: str
    full_name: str
    account_no: str
    bank_code: Optional[str] = None
    target_type: Optional[str] = None


class ExceptionCode(str, Enum):
    """Mã ngoại lệ nghiệp vụ X1 -> X10 (Phase 1 - Mục A4)."""

    X1_NOT_IN_MASTER = "X1"             # Số thẻ chưa có trong Master (V2)
    X2_ACCOUNT_MISMATCH_MASTER = "X2"   # TK nguồn != TK Master (V2)
    X3_MISSING_ACCOUNT = "X3"           # Có Số thẻ nhưng không có Số TK hợp lệ
    X4_NOT_IN_TARGET = "X4"             # Có trong lương nhưng không có dòng trong file đích
    X5_TARGET_NOT_IN_SOURCE = "X5"      # Dòng trong file đích không có trong lương mới
    X6_TARGET_MISMATCH = "X6"           # Khớp TK nhưng lệch tên / lệch mã ngân hàng
    X7_INVALID_AMOUNT = "X7"            # Số tiền không hợp lệ (trống, <= 0, lẻ, lỗi)
    X8_DUPLICATE_KEY = "X8"             # Trùng Số thẻ hoặc trùng Số TK (chặn cứng)
    X9_PERIOD_CONFLICT = "X9"           # Tháng/năm mâu thuẫn (V2)
    X10_RULE_VIOLATION = "X10"          # Vi phạm hướng dẫn nhập liệu / lỗi trên dòng đích

    # Bí danh ngắn gọn
    X1 = "X1"
    X2 = "X2"
    X3 = "X3"
    X4 = "X4"
    X5 = "X5"
    X6 = "X6"
    X7 = "X7"
    X8 = "X8"
    X9 = "X9"
    X10 = "X10"


class ResolutionAction(str, Enum):
    """Lựa chọn xử lý của người dùng cho từng dòng ngoại lệ trên bảng Preview."""

    SKIP = "SKIP"                       # Bỏ qua không ghi dòng nguồn này (chỉ báo cáo)
    CONFIRM_WRITE = "CONFIRM_WRITE"     # Xác nhận vẫn ghi số tiền mới (dùng cho X6, X10)
    KEEP = "KEEP"                       # Giữ nguyên số tiền cũ trong file đích (dùng cho X5)
    SET_ZERO = "SET_ZERO"               # Đặt số tiền = 0 vào ô Số tiền đích (dùng cho X5, X7)
    CLEAR = "CLEAR"                     # Xóa trắng ô Số tiền đích (dùng cho X5, X7)
    OVERRIDE_AMOUNT = "OVERRIDE_AMOUNT" # Nhập tay số tiền nguyên > 0 trong Preview (dùng cho X7)


class MatchStatus(str, Enum):
    """Trạng thái hiển thị của 1 dòng trên bảng Preview (Mục 5 - V1)."""

    WILL_UPDATE = "WILL_UPDATE"   # 🟢 Sẽ cập nhật (giá trị mới khác giá trị cũ)
    UNCHANGED = "UNCHANGED"       # ⚪ Không đổi (giá trị mới trùng giá trị cũ hoặc chọn KEEP)
    PENDING = "PENDING"           # 🟡 Cần xác nhận (còn Exception chưa chọn xử lý)
    SKIPPED = "SKIPPED"           # ⚪ Đã bỏ qua (X3, X4, X7 hoặc người dùng chọn SKIP)
    ERROR = "ERROR"               # 🔴 Lỗi / Chặn cứng (X8 trùng khóa hoặc lỗi cấu trúc)


@dataclass
class ExceptionItem:
    """1 mục ngoại lệ trong Exception List cần người dùng xem / xử lý."""

    id: str
    code: ExceptionCode
    severity: Severity
    message: str
    card_no: Optional[str] = None
    account_no: Optional[str] = None
    source_name: str = ""
    target_name: str = ""
    source_row: Optional[int] = None
    target_file: Optional[str] = None
    target_sheet: Optional[str] = None
    target_row: Optional[int] = None
    old_amount: Optional[int] = None
    source_amount: Optional[int] = None
    allowed_actions: tuple[ResolutionAction, ...] = ()
    action: Optional[ResolutionAction] = None
    override_amount: Optional[int] = None

    @property
    def is_blocking(self) -> bool:
        """True nếu là lỗi chặn cứng (X8), không cho phép bỏ qua trên giao diện."""
        return self.code == ExceptionCode.X8_DUPLICATE_KEY or not self.allowed_actions

    @property
    def is_resolved(self) -> bool:
        """True nếu ngoại lệ đã có phương án xử lý hợp lệ."""
        if self.is_blocking or self.action is None:
            return False
        if self.action not in self.allowed_actions:
            return False
        if self.action == ResolutionAction.OVERRIDE_AMOUNT:
            return _is_positive_int(self.override_amount)
        return True


@dataclass
class MatchItem:
    """1 dòng kết quả đối chiếu hiển thị trên bảng Preview và Báo cáo."""

    card_no: Optional[str]
    account_no: Optional[str]
    source_name: str
    target_name: str
    bank_code: Optional[str]
    source_row: Optional[int]
    target_file: Optional[str]
    target_sheet: Optional[str]
    target_row: Optional[int]
    target_cell: Optional[str]
    old_amount: Optional[int]
    new_amount: Optional[int]
    source_amount: Optional[int]
    status: MatchStatus
    exception_ids: tuple[str, ...] = ()
    note: str = ""


@dataclass(frozen=True)
class CellUpdate:
    """1 ô cụ thể trong file đích được lên kế hoạch ghi (chỉ cột Số tiền theo R6)."""

    file_path: str
    sheet_name: str
    row: int
    column: str                 # chữ cái cột Số tiền (vd "D", "E")
    old_value: Optional[int]
    new_value: Optional[int]    # int >= 0, hoặc None nếu xóa trắng ô (CLEAR)
    card_no: Optional[str] = None
    account_no: Optional[str] = None
    full_name: str = ""
    reason: str = "MATCHED"

    def __post_init__(self) -> None:
        object.__setattr__(self, "column", column_letter(column_index(self.column)))
        if self.new_value is not None:
            if not isinstance(self.new_value, int) or isinstance(self.new_value, bool) or self.new_value < 0:
                raise ConfigError(f"Giá trị ghi vào ô {self.column}{self.row} không hợp lệ: {self.new_value!r}")

    @property
    def cell(self) -> str:
        return f"{self.column}{self.row}"

    @property
    def col_index(self) -> int:
        return column_index(self.column)

    @property
    def is_changed(self) -> bool:
        return self.old_value != self.new_value


@dataclass
class TargetWritePlan:
    """Kế hoạch ghi cho 1 file ngân hàng cụ thể."""

    file_path: str
    file_format: str
    sheet_name: str
    template_name: str
    amount_column: str
    updates: list[CellUpdate] = field(default_factory=list)
    total_target_rows: int = 0

    @property
    def changed_updates(self) -> list[CellUpdate]:
        return [u for u in self.updates if u.is_changed]

    @property
    def total_old_amount(self) -> int:
        return sum(u.old_value or 0 for u in self.updates)

    @property
    def total_new_amount(self) -> int:
        return sum(u.new_value or 0 for u in self.updates)


@dataclass
class WritePlan:
    """Kết quả đối chiếu tổng thể: danh sách ô cần ghi + bảng MatchItem + danh sách ExceptionItem."""

    source_file: str
    source_sheet: str
    target_plans: list[TargetWritePlan] = field(default_factory=list)
    updates: list[CellUpdate] = field(default_factory=list)
    matches: list[MatchItem] = field(default_factory=list)
    exceptions: list[ExceptionItem] = field(default_factory=list)
    blocking_issues: list[Issue] = field(default_factory=list)
    total_source_records: int = 0
    total_source_transferable_amount: int = 0

    @property
    def write_plan(self) -> "WritePlan":
        """Thuộc tính tương thích để truy cập .write_plan hoặc dùng trực tiếp đối tượng WritePlan."""
        return self

    @property
    def changed_updates(self) -> list[CellUpdate]:
        return [u for u in self.updates if u.is_changed]

    @property
    def blocking_exceptions(self) -> list[ExceptionItem]:
        return [e for e in self.exceptions if e.is_blocking]

    @property
    def unresolved_exceptions(self) -> list[ExceptionItem]:
        return [e for e in self.exceptions if not e.is_resolved]

    @property
    def has_blocking_errors(self) -> bool:
        return bool(self.blocking_issues or self.blocking_exceptions)

    @property
    def is_ready(self) -> bool:
        """True khi không có lỗi chặn cứng và toàn bộ Exception đều đã được xử lý."""
        return not self.has_blocking_errors and not self.unresolved_exceptions

    @property
    def total_new_amount(self) -> int:
        return sum(u.new_value or 0 for u in self.updates)

    @property
    def total_old_amount(self) -> int:
        return sum(u.old_value or 0 for u in self.updates)

    def for_file(self, file_path: Any) -> Optional[TargetWritePlan]:
        target_str = str(file_path)
        return next((tp for tp in self.target_plans if tp.file_path == target_str), None)

    def resolve(
        self,
        exception_id: str,
        action: ResolutionAction,
        *,
        override_amount: Optional[int] = None,
    ) -> "WritePlan":
        """Cập nhật lựa chọn xử lý cho 1 ExceptionItem và tính lại WritePlan."""
        from .reconciler import resolve_exception

        return resolve_exception(self, exception_id, action, override_amount=override_amount)

    def summary(self) -> dict[str, int]:
        return {
            "total_source_records": self.total_source_records,
            "total_source_transferable_amount": self.total_source_transferable_amount,
            "scheduled_updates": len(self.updates),
            "changed_cells": len(self.changed_updates),
            "will_update": sum(1 for m in self.matches if m.status == MatchStatus.WILL_UPDATE),
            "unchanged": sum(1 for m in self.matches if m.status == MatchStatus.UNCHANGED),
            "pending": sum(1 for m in self.matches if m.status == MatchStatus.PENDING),
            "skipped": sum(1 for m in self.matches if m.status == MatchStatus.SKIPPED),
            "error": sum(1 for m in self.matches if m.status == MatchStatus.ERROR),
            "exceptions_total": len(self.exceptions),
            "exceptions_unresolved": len(self.unresolved_exceptions),
            "exceptions_blocking": len(self.blocking_exceptions),
            "total_new_amount": self.total_new_amount,
        }


# Bí danh tương thích
ReconcileResult = WritePlan


@dataclass
class RunResult:
    """Kết quả hoàn tất 1 lần chạy toàn trình (dùng cho workflow / reporter)."""

    write_plan: WritePlan
    backup_files: list[str] = field(default_factory=list)
    output_files: list[str] = field(default_factory=list)
    report_file: Optional[str] = None
    verified: bool = False


# =============================================================================
# 10. Kiểm tra an toàn trước khi ghi (Safety Layer)
# =============================================================================

class SafetyError(ToolError):
    """Lỗi chốt chặn an toàn: không cho phép ghi file khi dữ liệu chưa an toàn."""


class SafetyCode(str, Enum):
    """Mã phát hiện của lớp kiểm tra an toàn (core.safety)."""

    AMBIGUOUS_MAPPING = "AMBIGUOUS_MAPPING"                     # Mapping mơ hồ / trùng ô / sai cột
    DUPLICATE_IDENTITY = "DUPLICATE_IDENTITY"                   # Nhiều người trùng định danh (Số thẻ / Số TK)
    MISSING_OR_INVALID_AMOUNT = "MISSING_OR_INVALID_AMOUNT"     # Thiếu số tiền hoặc số tiền không hợp lệ
    MISSING_ACCOUNT = "MISSING_ACCOUNT"                         # Thiếu Số tài khoản ở ô dự kiến ghi
    ABNORMAL_CHANGE = "ABNORMAL_CHANGE"                         # Thay đổi bất thường (đột biến, > 1 tỷ, đặt 0/xóa...)
    INCONSISTENT_DATA = "INCONSISTENT_DATA"                     # Dữ liệu không nhất quán giữa MatchItem và CellUpdate
    UNRESOLVED_EXCEPTION = "UNRESOLVED_EXCEPTION"               # Còn Exception chưa được người dùng xử lý
    ROW_BALANCE_MISMATCH = "ROW_BALANCE_MISMATCH"               # Lệch cân đối số dòng (K8)
    AMOUNT_BALANCE_MISMATCH = "AMOUNT_BALANCE_MISMATCH"         # Lệch cân đối tổng tiền (K8)
    OVERWRITE_ORIGINAL_FORBIDDEN = "OVERWRITE_ORIGINAL_FORBIDDEN"  # Vi phạm R8: ghi đè file gốc
    MISSING_BACKUP = "MISSING_BACKUP"                           # Chưa có bản sao lưu trước khi xuất
    FILE_LOCKED = "FILE_LOCKED"                                 # File đang mở trong Excel (K9)


@dataclass(frozen=True)
class SafetyFinding:
    """1 lỗi hoặc cảnh báo do Safety Layer phát hiện."""

    code: SafetyCode
    severity: Severity
    message: str
    file_path: Optional[str] = None
    sheet: Optional[str] = None
    row: Optional[int] = None
    column: Optional[str] = None
    card_no: Optional[str] = None
    account_no: Optional[str] = None

    def __str__(self) -> str:
        return self.message

    def __contains__(self, item: str) -> bool:
        return item in self.message or item in self.code.value


@dataclass
class ValidationResult:
    """Kết quả kiểm tra an toàn trước khi ghi file."""

    errors: list[SafetyFinding] = field(default_factory=list)
    warnings: list[SafetyFinding] = field(default_factory=list)
    row_balance_ok: bool = True
    amount_balance_ok: bool = True

    @property
    def is_valid(self) -> bool:
        return not self.errors

    @property
    def is_safe(self) -> bool:
        return not self.errors

    @property
    def ok(self) -> bool:
        return not self.errors

    @property
    def error_messages(self) -> list[str]:
        return [e.message for e in self.errors]

    @property
    def warning_messages(self) -> list[str]:
        return [w.message for w in self.warnings]

    def ensure_safe(self) -> None:
        """Ném SafetyError nếu có bất kỳ lỗi an toàn nào."""
        if self.errors:
            bullet_list = "\n- ".join(e.message for e in self.errors)
            raise SafetyError(
                f"Chốt chặn an toàn: phát hiện {len(self.errors)} lỗi, không được phép ghi file:\n- {bullet_list}"
            )

    def summary(self) -> dict[str, int]:
        return {
            "errors": len(self.errors),
            "warnings": len(self.warnings),
            "row_balance_ok": int(self.row_balance_ok),
            "amount_balance_ok": int(self.amount_balance_ok),
        }


# =============================================================================
# 11. Ghi file kết quả (Excel Writer Layer)
# =============================================================================

class ExcelWriteError(ToolError):
    """Lỗi khi tạo bản sao lưu hoặc ghi dữ liệu vào file Excel đích."""


# Kiểu đại diện cho cấu hình mapping và file đích truyền vào Excel Writer
TemplateProfile = Union[TargetTemplate, SourceProfile]
TargetWorkbook = Union[str, TargetReadResult]


@dataclass(frozen=True)
class TargetWriteOutput:
    """Kết quả ghi của 1 file ngân hàng cụ thể."""

    source_target_file: str     # Đường dẫn file ngân hàng gốc (đầu vào)
    backup_file: str            # Đường dẫn file backup đã tạo trong data/backup
    output_file: str            # Đường dẫn file kết quả mới đã tạo trong data/output
    sheet_name: str             # Tên sheet đã cập nhật
    file_format: str            # "xls" hoặc "xlsx"
    updated_cells: int          # Tổng số ô trong kế hoạch đã thực hiện ghi
    changed_cells: int          # Số ô có giá trị mới khác giá trị cũ
    total_new_amount: int       # Tổng số tiền mới trên các ô đã ghi của file này
    original_sha256: str = ""   # Mã băm SHA-256 của file gốc (đảm bảo gốc không bị sửa)


@dataclass
class WriteExecutionResult:
    """Kết quả thực thi của Excel Writer Layer cho toàn bộ WritePlan."""

    write_plan: WritePlan
    validation: ValidationResult
    target_outputs: list[TargetWriteOutput] = field(default_factory=list)
    source_backup_file: Optional[str] = None

    @property
    def backup_files(self) -> list[str]:
        files: list[str] = []
        if self.source_backup_file:
            files.append(self.source_backup_file)
        files.extend(t.backup_file for t in self.target_outputs if t.backup_file)
        return files

    @property
    def output_files(self) -> list[str]:
        return [t.output_file for t in self.target_outputs]

    def to_run_result(self, *, report_file: Optional[str] = None, verified: bool = False) -> RunResult:
        return RunResult(
            write_plan=self.write_plan,
            backup_files=self.backup_files,
            output_files=self.output_files,
            report_file=report_file,
            verified=verified,
        )


# =============================================================================
# 12. Hậu kiểm & Đối chiếu tính nhất quán (Verification Layer)
# =============================================================================

class VerificationError(ToolError):
    """Lỗi khi hậu kiểm phát hiện dữ liệu không nhất quán hoặc file xuất sai lệch."""


class VerificationCode(str, Enum):
    """Mã phát hiện của Verification Layer (core.verifier)."""

    VALIDATION_NOT_SAFE = "VALIDATION_NOT_SAFE"
    UNRESOLVED_EXCEPTION = "UNRESOLVED_EXCEPTION"
    ROW_DISCREPANCY = "ROW_DISCREPANCY"
    AMOUNT_DISCREPANCY = "AMOUNT_DISCREPANCY"
    SOURCE_RECONCILE_MISMATCH = "SOURCE_RECONCILE_MISMATCH"
    RECONCILE_PLAN_MISMATCH = "RECONCILE_PLAN_MISMATCH"
    MISSING_OUTPUT_FILE = "MISSING_OUTPUT_FILE"
    MISSING_BACKUP_FILE = "MISSING_BACKUP_FILE"
    ORIGINAL_FILE_MODIFIED = "ORIGINAL_FILE_MODIFIED"
    SHEET_STRUCTURE_CHANGED = "SHEET_STRUCTURE_CHANGED"
    POST_WRITE_VALUE_MISMATCH = "POST_WRITE_VALUE_MISMATCH"
    UNEXPECTED_CELL_MODIFIED = "UNEXPECTED_CELL_MODIFIED"


@dataclass(frozen=True)
class CellDiff:
    """Chênh lệch phát hiện tại 1 ô khi đối chiếu hậu kiểm file kết quả với file gốc."""

    file_path: str
    sheet_name: str
    row: int
    column: str
    expected_value: Any
    actual_value: Any
    is_planned_cell: bool
    message: str

    @property
    def cell(self) -> str:
        return f"{self.column}{self.row}"


@dataclass(frozen=True)
class VerificationFinding:
    """1 lỗi hoặc cảnh báo do Verification Layer phát hiện."""

    code: VerificationCode
    severity: Severity
    message: str
    file_path: Optional[str] = None
    sheet: Optional[str] = None
    row: Optional[int] = None
    column: Optional[str] = None

    def __str__(self) -> str:
        return self.message

    def __contains__(self, item: str) -> bool:
        return item in self.message or item in self.code.value


@dataclass
class VerificationResult:
    """Kết quả kiểm tra tính nhất quán (Source -> Reconcile -> WritePlan) & hậu kiểm diff file xuất."""

    # --- Chỉ số dòng ---
    total_source_records: int = 0
    total_matched_records: int = 0
    total_unmatched_source_records: int = 0
    total_unmatched_target_records: int = 0
    total_exceptions: int = 0
    resolved_exceptions: int = 0
    unresolved_exceptions: int = 0
    blocking_exceptions: int = 0
    total_written_records: int = 0
    total_source_written_records: int = 0
    total_zero_or_clear_records: int = 0
    total_skipped_records: int = 0

    # --- Chỉ số tiền (VND) ---
    total_source_amount: int = 0
    total_source_transferable_amount: int = 0
    total_planned_write_amount: int = 0
    total_skipped_transferable_amount: int = 0
    total_override_delta: int = 0

    # --- Chênh lệch ---
    row_discrepancy: int = 0
    amount_discrepancy: int = 0

    # --- Trạng thái nhất quán từng chặng ---
    source_to_reconcile_ok: bool = True
    reconcile_to_plan_ok: bool = True
    post_write_diff_ok: bool = True

    # --- Chi tiết phát hiện ---
    errors: list[VerificationFinding] = field(default_factory=list)
    warnings: list[VerificationFinding] = field(default_factory=list)
    cell_diffs: list[CellDiff] = field(default_factory=list)
    verified_files: list[str] = field(default_factory=list)

    @property
    def verified(self) -> bool:
        return (
            not self.errors
            and self.row_discrepancy == 0
            and self.amount_discrepancy == 0
            and self.source_to_reconcile_ok
            and self.reconcile_to_plan_ok
            and self.post_write_diff_ok
        )

    @property
    def is_verified(self) -> bool:
        return self.verified

    @property
    def is_valid(self) -> bool:
        return self.verified

    @property
    def ok(self) -> bool:
        return self.verified

    @property
    def error_messages(self) -> list[str]:
        return [e.message for e in self.errors]

    @property
    def warning_messages(self) -> list[str]:
        return [w.message for w in self.warnings]

    def ensure_verified(self) -> None:
        """Ném VerificationError nếu phát hiện bất kỳ sai lệch nào."""
        if not self.verified:
            bullet_list = "\n- ".join(e.message for e in self.errors) or (
                f"row_discrepancy={self.row_discrepancy}, amount_discrepancy={self.amount_discrepancy}"
            )
            raise VerificationError(
                f"Hậu kiểm không đạt ({len(self.errors)} lỗi, {len(self.cell_diffs)} ô lệch):\n- {bullet_list}"
            )

    def summary(self) -> dict[str, int]:
        return {
            "total_source_records": self.total_source_records,
            "total_matched_records": self.total_matched_records,
            "total_unmatched_source_records": self.total_unmatched_source_records,
            "total_unmatched_target_records": self.total_unmatched_target_records,
            "total_exceptions": self.total_exceptions,
            "resolved_exceptions": self.resolved_exceptions,
            "unresolved_exceptions": self.unresolved_exceptions,
            "blocking_exceptions": self.blocking_exceptions,
            "total_written_records": self.total_written_records,
            "total_source_written_records": self.total_source_written_records,
            "total_zero_or_clear_records": self.total_zero_or_clear_records,
            "total_skipped_records": self.total_skipped_records,
            "total_source_amount": self.total_source_amount,
            "total_source_transferable_amount": self.total_source_transferable_amount,
            "total_planned_write_amount": self.total_planned_write_amount,
            "total_skipped_transferable_amount": self.total_skipped_transferable_amount,
            "total_override_delta": self.total_override_delta,
            "row_discrepancy": self.row_discrepancy,
            "amount_discrepancy": self.amount_discrepancy,
            "errors": len(self.errors),
            "warnings": len(self.warnings),
            "cell_diffs": len(self.cell_diffs),
            "verified": int(self.verified),
        }


# =============================================================================
# 13. Điều phối quy trình (Workflow Layer)
# =============================================================================

class WorkflowError(ToolError):
    """Lỗi khi điều phối quy trình xử lý."""


class WorkflowStage(str, Enum):
    """Trạng thái / bước hiện tại của quy trình điều phối."""

    IDLE = "IDLE"
    TEMPLATES_LOADED = "TEMPLATES_LOADED"
    SOURCE_READ = "SOURCE_READ"
    TARGETS_READ = "TARGETS_READ"
    RECONCILED = "RECONCILED"
    AWAITING_RESOLUTION = "AWAITING_RESOLUTION"
    SAFETY_CHECKED = "SAFETY_CHECKED"
    SAFETY_FAILED = "SAFETY_FAILED"
    PRE_VERIFIED = "PRE_VERIFIED"
    VERIFICATION_FAILED = "VERIFICATION_FAILED"
    WRITTEN = "WRITTEN"
    COMPLETED = "COMPLETED"
    FAILED = "FAILED"


WorkflowState = WorkflowStage


@dataclass
class RuntimeConfig:
    """Cấu hình thời gian chạy cho 1 phiên điều phối Workflow."""

    templates_dir: str = "templates"
    output_dir: str = "data/output"
    backup_dir: str = "data/backup"
    output_suffix: str = "_KetQua"
    source_sheet: Optional[str] = None
    source_profile_name: Optional[str] = None
    source_profile: Optional[SourceProfile] = None
    target_templates: Optional[Any] = None
    output_paths: Optional[Any] = None
    resolutions: dict[str, Any] = field(default_factory=dict)
    master_records: Optional[dict[str, MasterRecord]] = None
    auto_create_default_templates: bool = True
    overwrite_output: bool = False
    backup_source_file: bool = True
    writer_engine: str = "auto"
    dry_run: bool = False
    cleanup_output_on_verify_failure: bool = True
    raise_on_failure: bool = False


WorkflowConfig = RuntimeConfig


@dataclass
class WorkflowResult:
    """Kết quả tổng thể của 1 phiên chạy Workflow (bao gồm trạng thái từng bước và RunResult)."""

    stage: WorkflowStage = WorkflowStage.IDLE
    source_result: Optional[SourceReadResult] = None
    source_profile: Optional[SourceProfile] = None
    target_results: list[TargetReadResult] = field(default_factory=list)
    target_templates: list[TargetTemplate] = field(default_factory=list)
    unmatched_target_files: list[str] = field(default_factory=list)
    write_plan: Optional[WritePlan] = None
    exceptions: list[ExceptionItem] = field(default_factory=list)
    validation_result: Optional[ValidationResult] = None
    pre_verification_result: Optional[VerificationResult] = None
    write_execution_result: Optional[WriteExecutionResult] = None
    verification_result: Optional[VerificationResult] = None
    run_result: Optional[RunResult] = None
    stopped_reason: Optional[str] = None
    rolled_back_files: list[str] = field(default_factory=list)

    @property
    def is_success(self) -> bool:
        return (
            self.stage == WorkflowStage.COMPLETED
            and self.verification_result is not None
            and self.verification_result.is_verified
            and (self.validation_result is None or self.validation_result.is_safe)
        )

    @property
    def ok(self) -> bool:
        return self.is_success

    @property
    def succeeded(self) -> bool:
        return self.is_success

    @property
    def verified(self) -> bool:
        if self.verification_result is not None:
            return self.verification_result.is_verified
        if self.run_result is not None:
            return self.run_result.verified
        return False

    @property
    def can_proceed_to_write(self) -> bool:
        return (
            self.write_plan is not None
            and self.validation_result is not None
            and self.validation_result.is_safe
            and (self.pre_verification_result is None or self.pre_verification_result.is_verified)
        )

    @property
    def unresolved_exceptions(self) -> list[ExceptionItem]:
        return [e for e in self.exceptions if not e.is_resolved]

    @property
    def blocking_exceptions(self) -> list[ExceptionItem]:
        return [e for e in self.exceptions if e.is_blocking]

    @property
    def has_unresolved_exceptions(self) -> bool:
        return any(not e.is_resolved or e.is_blocking for e in self.exceptions)

    @property
    def output_files(self) -> list[str]:
        if self.run_result is not None:
            return list(self.run_result.output_files)
        if self.write_execution_result is not None:
            return self.write_execution_result.output_files
        return []

    @property
    def backup_files(self) -> list[str]:
        if self.run_result is not None:
            return list(self.run_result.backup_files)
        if self.write_execution_result is not None:
            return self.write_execution_result.backup_files
        return []

    @property
    def errors(self) -> list[str]:
        msgs: list[str] = []
        if self.stopped_reason:
            msgs.append(self.stopped_reason)
        if self.validation_result is not None:
            for e in self.validation_result.error_messages:
                if e not in msgs:
                    msgs.append(e)
        if self.pre_verification_result is not None:
            for e in self.pre_verification_result.error_messages:
                if e not in msgs:
                    msgs.append(e)
        if self.verification_result is not None:
            for e in self.verification_result.error_messages:
                if e not in msgs:
                    msgs.append(e)
        return msgs

    @property
    def warnings(self) -> list[str]:
        msgs: list[str] = []
        if self.validation_result is not None:
            for w in self.validation_result.warning_messages:
                if w not in msgs:
                    msgs.append(w)
        if self.pre_verification_result is not None:
            for w in self.pre_verification_result.warning_messages:
                if w not in msgs:
                    msgs.append(w)
        if self.verification_result is not None:
            for w in self.verification_result.warning_messages:
                if w not in msgs:
                    msgs.append(w)
        return msgs

    def to_run_result(self) -> RunResult:
        if self.run_result is not None:
            return self.run_result
        if self.write_plan is None:
            raise WorkflowError("Chưa có WritePlan để tạo RunResult.")
        return RunResult(
            write_plan=self.write_plan,
            backup_files=self.backup_files,
            output_files=self.output_files,
            verified=self.verified,
        )

    def summary(self) -> dict[str, Any]:
        return {
            "stage": self.stage.value,
            "is_success": self.is_success,
            "verified": self.verified,
            "total_source_records": self.write_plan.total_source_records if self.write_plan else 0,
            "total_updates": len(self.write_plan.updates) if self.write_plan else 0,
            "exceptions_total": len(self.exceptions),
            "exceptions_unresolved": len(self.unresolved_exceptions),
            "output_files": len(self.output_files),
            "backup_files": len(self.backup_files),
            "errors": len(self.errors),
            "warnings": len(self.warnings),
        }


# =============================================================================
# 14. Báo cáo kết quả (Reporting Layer)
# =============================================================================

class ReportError(ToolError):
    """Lỗi khi tổng hợp dữ liệu báo cáo hoặc xuất file báo cáo Excel (.xlsx)."""


@dataclass(frozen=True)
class TargetFileSummaryRow:
    """Dòng tổng hợp theo từng file ngân hàng đích trong Summary Report."""

    target_file: str
    sheet_name: str
    template_name: str
    amount_column: str
    total_target_rows: int
    scheduled_updates: int
    changed_cells: int
    total_old_amount: int
    total_new_amount: int
    delta_amount: int
    output_file: str = ""
    backup_file: str = ""


@dataclass
class SummaryReportData:
    """Dữ liệu báo cáo tổng hợp (Summary Report)."""

    created_at: str = ""
    stage: str = "IDLE"
    is_safe: bool = False
    is_verified: bool = False
    source_file: str = ""
    source_sheet: str = ""
    total_source_records: int = 0
    total_matched_records: int = 0
    total_unmatched_source_records: int = 0
    total_unmatched_target_records: int = 0
    total_written_records: int = 0
    total_source_written_records: int = 0
    total_zero_or_clear_records: int = 0
    total_skipped_records: int = 0
    total_exceptions: int = 0
    resolved_exceptions: int = 0
    unresolved_exceptions: int = 0
    blocking_exceptions: int = 0
    total_source_amount: int = 0
    total_source_transferable_amount: int = 0
    total_old_target_amount: int = 0
    total_planned_write_amount: int = 0
    total_skipped_transferable_amount: int = 0
    total_override_delta: int = 0
    row_discrepancy: int = 0
    amount_discrepancy: int = 0
    safety_errors_count: int = 0
    safety_warnings_count: int = 0
    verification_errors_count: int = 0
    verification_warnings_count: int = 0
    cell_diffs_count: int = 0
    output_files: list[str] = field(default_factory=list)
    backup_files: list[str] = field(default_factory=list)
    target_summaries: list[TargetFileSummaryRow] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class DetailReportRow:
    """1 dòng chi tiết đối chiếu & cập nhật trong báo cáo."""

    stt: int
    card_no: str
    account_no: str
    source_name: str
    target_name: str
    bank_code: str
    source_row: Optional[int]
    target_file: str
    target_sheet: str
    target_row: Optional[int]
    target_cell: str
    old_amount: Optional[int]
    source_amount: Optional[int]
    new_amount: Optional[int]
    delta_amount: Optional[int]
    status: str
    exception_ids: str
    note: str


@dataclass(frozen=True)
class ExceptionReportRow:
    """1 dòng trong Exception Report."""

    stt: int
    exception_id: str
    code: str
    severity: str
    card_no: str
    account_no: str
    source_name: str
    target_name: str
    source_row: Optional[int]
    target_file: str
    target_sheet: str
    target_row: Optional[int]
    old_amount: Optional[int]
    source_amount: Optional[int]
    action: str
    override_amount: Optional[int]
    is_resolved: bool
    is_blocking: bool
    message: str


@dataclass(frozen=True)
class ValidationReportRow:
    """1 dòng trong Validation & Verification Report (Safety, Verifier, CellDiff)."""

    stt: int
    layer: str
    code: str
    severity: str
    file_path: str
    sheet: str
    row: Optional[int]
    column: str
    card_no: str
    account_no: str
    expected_value: Any = None
    actual_value: Any = None
    message: str = ""


@dataclass
class ReportResult:
    """Kết quả đầu ra của Reporting Layer (chứa dữ liệu các báo cáo và đường dẫn file .xlsx)."""

    summary_report: SummaryReportData = field(default_factory=SummaryReportData)
    detail_rows: list[DetailReportRow] = field(default_factory=list)
    exception_report: list[ExceptionReportRow] = field(default_factory=list)
    validation_report: list[ValidationReportRow] = field(default_factory=list)
    report_file: Optional[str] = None
    created_at: str = ""

    @property
    def summary(self) -> SummaryReportData:
        return self.summary_report

    @property
    def details(self) -> list[DetailReportRow]:
        return self.detail_rows

    @property
    def exceptions(self) -> list[ExceptionReportRow]:
        return self.exception_report

    @property
    def validations(self) -> list[ValidationReportRow]:
        return self.validation_report

    def to_dict(self) -> dict[str, Any]:
        return {
            "created_at": self.created_at,
            "report_file": self.report_file,
            "summary_report": self.summary_report.to_dict(),
            "detail_rows": [asdict(r) for r in self.detail_rows],
            "exception_report": [asdict(r) for r in self.exception_report],
            "validation_report": [asdict(r) for r in self.validation_report],
        }






