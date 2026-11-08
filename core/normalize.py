"""
normalize.py - Chuẩn hóa và kiểm tra giá trị đọc từ Excel.

Nguyên tắc an toàn kế toán:
- KHÔNG đoán. Giá trị mơ hồ -> trả về lỗi, không tự sửa.
- KHÔNG làm tròn số tiền. Số tiền có phần lẻ -> lỗi.
- Định danh (Số thẻ, Số tài khoản) luôn là chuỗi, giữ nguyên số 0 ở đầu.
- KHÔNG tự xóa khoảng trắng / ký tự ở GIỮA số tài khoản (chỉ cắt 2 đầu).
"""

from __future__ import annotations

import math
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Any, Generic, Optional, TypeVar

from .models import IssueCode, Severity

T = TypeVar("T")

# ---------------------------------------------------------------------------
# Ngưỡng kiểm tra (có thể chỉnh nếu nghiệp vụ thay đổi)
# ---------------------------------------------------------------------------
ACCOUNT_MIN_LEN: int = 6            # Số TK ngắn hơn -> cảnh báo (không chặn)
ACCOUNT_MAX_LEN: int = 19           # Số TK dài hơn -> cảnh báo (không chặn)
AMOUNT_WARNING_THRESHOLD: int = 1_000_000_000  # > 1 tỷ / người / lần -> cảnh báo
FLOAT_SAFE_INTEGER_DIGITS: int = 15  # Số lưu dạng số (float) chỉ chính xác tới 15 chữ số
_INTEGER_TOLERANCE: float = 1e-6     # Sai số dấu phẩy động chấp nhận được khi Excel lưu số nguyên

# Ký tự "vô hình" thường dính vào khi copy từ web / phần mềm khác
_INVISIBLE_CHARS = {
    "\u00a0": " ",   # non-breaking space
    "\u2007": " ",
    "\u202f": " ",
    "\u200b": "",    # zero-width space
    "\u200c": "",
    "\u200d": "",
    "\ufeff": "",    # BOM
}
_WHITESPACE_RE = re.compile(r"\s+")
_DIGITS_RE = re.compile(r"^\d+$")
_SIGNED_INT_RE = re.compile(r"^[+-]?\d+$")

# Ký tự Latin-1 KHÔNG có trong tiếng Việt Unicode nhưng xuất hiện khi văn bản gõ bằng
# font VNI / TCVN3 được đọc ra (ví dụ "Nguyeãn Thaønh Coâng").
_LEGACY_FONT_MARKERS = set("øöñïäëåæüûÿîçÑØÖÄËÏÜÛÎÇÅÆ")


@dataclass(frozen=True)
class NormResult(Generic[T]):
    """Kết quả chuẩn hóa 1 giá trị.

    value    : giá trị đã chuẩn hóa (None nếu trống hoặc không dùng được).
    code     : mã vấn đề nếu có (None = không có vấn đề).
    severity : mức độ vấn đề.
    detail   : mô tả tiếng Việt.
    """

    value: Optional[T]
    code: Optional[IssueCode] = None
    severity: Optional[Severity] = None
    detail: str = ""

    @property
    def ok(self) -> bool:
        return self.code is None

    @property
    def is_error(self) -> bool:
        return self.severity == Severity.ERROR

    @property
    def is_empty(self) -> bool:
        return self.value is None and self.code is None


def _ok(value: Optional[T]) -> NormResult[T]:
    return NormResult(value=value)


def _error(code: IssueCode, detail: str, value: Optional[T] = None) -> NormResult[T]:
    return NormResult(value=value, code=code, severity=Severity.ERROR, detail=detail)


def _warning(code: IssueCode, detail: str, value: Optional[T]) -> NormResult[T]:
    return NormResult(value=value, code=code, severity=Severity.WARNING, detail=detail)


# =============================================================================
# 1. Văn bản chung
# =============================================================================

def clean_text(value: Any) -> str:
    """Làm sạch văn bản: chuẩn Unicode NFC, bỏ ký tự vô hình, gộp khoảng trắng, cắt 2 đầu.

    Không đổi hoa/thường, không bỏ dấu.
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = unicodedata.normalize("NFC", str(value))
    for bad, repl in _INVISIBLE_CHARS.items():
        text = text.replace(bad, repl)
    return _WHITESPACE_RE.sub(" ", text).strip()


def strip_edges(value: Any) -> str:
    """Chỉ bỏ ký tự vô hình và khoảng trắng ở 2 đầu; GIỮ NGUYÊN phần giữa."""
    if value is None:
        return ""
    text = unicodedata.normalize("NFC", str(value))
    for bad in ("\u200b", "\u200c", "\u200d", "\ufeff"):
        text = text.replace(bad, "")
    return text.strip()  # str.strip() đã bao gồm NBSP và các khoảng trắng Unicode


def remove_diacritics(text: str) -> str:
    """Bỏ dấu tiếng Việt: 'Nguyễn Đức' -> 'Nguyen Duc'."""
    decomposed = unicodedata.normalize("NFD", text)
    stripped = "".join(ch for ch in decomposed if unicodedata.category(ch) != "Mn")
    return unicodedata.normalize("NFC", stripped.replace("đ", "d").replace("Đ", "D"))


def normalize_header(value: Any) -> str:
    """Chuẩn hóa tiêu đề cột để so sánh: gộp khoảng trắng/xuống dòng, không phân biệt hoa thường.

    GIỮ dấu tiếng Việt để tránh nhầm các tiêu đề gần giống nhau.
    """
    return clean_text(value).casefold()


def normalize_name_key(value: Any) -> str:
    """Khóa so sánh tên: bỏ dấu, viết hoa, gộp khoảng trắng.

    CHỈ dùng để kiểm tra chéo; KHÔNG BAO GIỜ dùng làm khóa định danh.
    """
    return remove_diacritics(clean_text(value)).upper()


def looks_like_legacy_font(text: str) -> bool:
    """True nếu văn bản có dấu hiệu gõ bằng font VNI/TCVN3 (không phải Unicode)."""
    return any(ch in _LEGACY_FONT_MARKERS for ch in text)


def has_letters(text: str) -> bool:
    """True nếu có ít nhất 1 chữ cái (dùng phát hiện cột tên bị map nhầm sang cột số)."""
    return any(ch.isalpha() for ch in text)


def _is_date_like(value: Any) -> bool:
    return isinstance(value, (datetime, date, time, timedelta))


# =============================================================================
# 2. Số thẻ (mã nhân viên)
# =============================================================================

def normalize_card(value: Any) -> NormResult[str]:
    """Chuẩn hóa Số thẻ.

    - Trống -> NormResult(None) (dòng sẽ bị bỏ qua).
    - Chuỗi -> cắt khoảng trắng 2 đầu, giữ nguyên nội dung (kể cả số 0 đầu).
    - Lưu dạng SỐ -> chấp nhận nhưng CẢNH BÁO vì có thể đã mất số 0 ở đầu.
    - Lưu dạng NGÀY / đúng-sai -> LỖI (Excel đã tự đổi kiểu, định danh bị hỏng).
    """
    if value is None:
        return _ok(None)
    if isinstance(value, bool) or _is_date_like(value):
        return _error(IssueCode.CARD_INVALID, f"Số thẻ đang bị Excel lưu dạng ngày/đúng-sai: {value!r}")
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or not value.is_integer()):
            return _error(IssueCode.CARD_INVALID, f"Số thẻ không phải số nguyên: {value!r}")
        if value < 0:
            return _error(IssueCode.CARD_INVALID, f"Số thẻ là số âm: {value!r}")
        text = str(int(value))
        return _warning(
            IssueCode.CARD_STORED_AS_NUMBER,
            f"Số thẻ '{text}' đang lưu dạng số trong Excel, có thể đã mất số 0 ở đầu.",
            text,
        )
    text = strip_edges(value)
    if not text:
        return _ok(None)
    return _ok(text)


# =============================================================================
# 3. Số tài khoản
# =============================================================================

def normalize_account(value: Any) -> NormResult[str]:
    """Chuẩn hóa Số tài khoản (chỉ gồm chữ số).

    - Trống -> NormResult(None).
    - Chuỗi -> cắt 2 đầu; phần còn lại PHẢI toàn chữ số, nếu không -> LỖI (không tự xóa ký tự).
    - Số nguyên -> chuyển sang chuỗi.
    - Số thực có phần lẻ hoặc dài hơn 15 chữ số -> LỖI (đã mất độ chính xác).
    - Độ dài bất thường -> CẢNH BÁO.
    """
    if value is None:
        return _ok(None)
    if isinstance(value, bool) or _is_date_like(value):
        return _error(IssueCode.ACCOUNT_INVALID, f"Số tài khoản bị lưu dạng ngày/đúng-sai: {value!r}")

    if isinstance(value, int):
        if value < 0:
            return _error(IssueCode.ACCOUNT_INVALID, f"Số tài khoản là số âm: {value}")
        text = str(value)
    elif isinstance(value, float):
        if math.isnan(value) or math.isinf(value) or value < 0 or not value.is_integer():
            return _error(IssueCode.ACCOUNT_INVALID, f"Số tài khoản không phải số nguyên: {value!r}")
        text = str(int(value))
        if len(text) > FLOAT_SAFE_INTEGER_DIGITS:
            return _error(
                IssueCode.ACCOUNT_INVALID,
                f"Số tài khoản dài {len(text)} chữ số nhưng lưu dạng số -> có thể đã sai các chữ số cuối. "
                "Hãy định dạng ô là Text và nhập lại.",
            )
    else:
        text = strip_edges(value)
        if not text:
            return _ok(None)
        if not _DIGITS_RE.match(text):
            return _error(
                IssueCode.ACCOUNT_INVALID,
                f"Số tài khoản '{text}' chứa ký tự không phải chữ số (khoảng trắng, dấu chấm, chữ...).",
            )

    if not ACCOUNT_MIN_LEN <= len(text) <= ACCOUNT_MAX_LEN:
        return _warning(
            IssueCode.ACCOUNT_UNUSUAL_LENGTH,
            f"Số tài khoản '{text}' có {len(text)} chữ số (bất thường, cần kiểm tra).",
            text,
        )
    return _ok(text)


# =============================================================================
# 4. Số tiền
# =============================================================================

def normalize_amount(value: Any) -> NormResult[int]:
    """Chuẩn hóa Số tiền (VND, số nguyên).

    - Trống -> NormResult(None).
    - Số có phần lẻ -> LỖI (không tự làm tròn).
    - Chuỗi chỉ chấp nhận dạng toàn chữ số (vd "5917028"). Dạng "5.917.028" hoặc "5,917,028"
      bị coi là LỖI vì không xác định chắc chắn dấu phân cách.
    - Âm -> LỖI. Bằng 0 -> CẢNH BÁO. Lớn bất thường -> CẢNH BÁO.
    """
    if value is None:
        return _ok(None)
    if isinstance(value, bool) or _is_date_like(value):
        return _error(IssueCode.AMOUNT_NOT_NUMBER, f"Số tiền không phải dạng số: {value!r}")

    amount: int
    if isinstance(value, int):
        amount = value
    elif isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return _error(IssueCode.AMOUNT_NOT_NUMBER, f"Số tiền không hợp lệ: {value!r}")
        nearest = round(value)
        if abs(value - nearest) > _INTEGER_TOLERANCE:
            return _error(
                IssueCode.AMOUNT_NOT_INTEGER,
                f"Số tiền {value!r} có phần thập phân. Ngân hàng yêu cầu số nguyên - "
                "hãy làm tròn trong file lương (tool không tự làm tròn).",
            )
        amount = int(nearest)
    else:
        text = clean_text(value)
        if not text:
            return _ok(None)
        if not _SIGNED_INT_RE.match(text):
            return _error(
                IssueCode.AMOUNT_NOT_NUMBER,
                f"Số tiền '{text}' là văn bản, không phải số. Hãy nhập dạng số trong Excel.",
            )
        amount = int(text)

    if amount < 0:
        return _error(IssueCode.AMOUNT_NEGATIVE, f"Số tiền âm: {amount:,}", amount)
    if amount == 0:
        return _warning(IssueCode.AMOUNT_ZERO, "Số tiền bằng 0.", amount)
    if amount > AMOUNT_WARNING_THRESHOLD:
        return _warning(IssueCode.AMOUNT_UNUSUALLY_LARGE, f"Số tiền lớn bất thường: {amount:,}", amount)
    return _ok(amount)


# =============================================================================
# 5. Mã ngân hàng
# =============================================================================

_BANK_CODE_RE = re.compile(r"^[A-Za-z0-9]+$")


def normalize_bank_code(value: Any) -> NormResult[str]:
    """Chuẩn hóa Mã ngân hàng (định danh: chỉ chữ và số, không tự sửa).

    - Trống -> NormResult(None).
    - Số nguyên (vd 6588.0) -> "6588".
    - Chuỗi -> cắt 2 đầu; còn ký tự khác chữ/số -> LỖI.
    """
    if value is None:
        return _ok(None)
    if isinstance(value, bool) or _is_date_like(value):
        return _error(IssueCode.BANK_CODE_INVALID, f"Mã ngân hàng bị lưu dạng ngày/đúng-sai: {value!r}")
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value) or not value.is_integer()):
            return _error(IssueCode.BANK_CODE_INVALID, f"Mã ngân hàng không phải số nguyên: {value!r}")
        if value < 0:
            return _error(IssueCode.BANK_CODE_INVALID, f"Mã ngân hàng là số âm: {value!r}")
        return _ok(str(int(value)))
    text = strip_edges(value)
    if not text:
        return _ok(None)
    if not _BANK_CODE_RE.match(text):
        return _error(IssueCode.BANK_CODE_INVALID, f"Mã ngân hàng '{text}' chứa ký tự không hợp lệ.")
    return _ok(text)
