"""
excel_reader.py - Đọc file Excel (.xlsx / .xlsm / .xls) thành dữ liệu thống nhất (SheetData).

- Module này CHỈ ĐỌC, không bao giờ ghi/sửa file.
- Nhận dạng định dạng theo NỘI DUNG file (không tin phần đuôi tên file).
- .xlsx: đọc giá trị đã tính (cache) + phát hiện ô công thức chưa được tính.
- .xls : đọc bằng xlrd (giá trị công thức luôn là kết quả đã lưu).
"""

from __future__ import annotations

import fnmatch
import logging
import os
import warnings
from enum import Enum
from pathlib import Path
from types import TracebackType
from typing import Any, Optional, Union

from .models import (
    CellData,
    CellKind,
    ColumnRef,
    Issue,
    IssueCode,
    Severity,
    SheetData,
    SheetInfo,
    ToolError,
    classify_value,
    column_letter,
)
from .normalize import normalize_header

log = logging.getLogger(__name__)

PathLike = Union[str, os.PathLike[str]]

_OLE2_SIGNATURE = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"   # .xls (Excel 97-2003)
_ZIP_SIGNATURE = b"PK\x03\x04"                          # .xlsx / .xlsm
_MIN_EXCEL_SIZE = 512


class ExcelReadError(ToolError):
    """Không đọc được file Excel."""


class FileFormat(str, Enum):
    XLSX = "xlsx"
    XLS = "xls"


class _XlrdLogWriter:
    """Đối tượng giống file, nhận thông báo của xlrd và chuyển vào logging."""

    def write(self, text: str) -> int:
        text = text.strip()
        if text:
            log.debug("xlrd: %s", text)
        return len(text)

    def flush(self) -> None:
        return None


# =============================================================================
# 1. Nhận dạng định dạng file
# =============================================================================

def detect_format(path: PathLike) -> FileFormat:
    """Xác định định dạng thật của file dựa vào các byte đầu tiên."""
    p = Path(path)
    if not p.exists():
        raise ExcelReadError(f"Không tìm thấy file: {p}")
    if not p.is_file():
        raise ExcelReadError(f"Đường dẫn không phải là file: {p}")
    if p.name.startswith("~$"):
        raise ExcelReadError(
            f"'{p.name}' là file tạm do Excel tạo ra khi đang mở file, không phải file dữ liệu. "
            "Hãy chọn file gốc (không có '~$' ở đầu tên)."
        )
    if p.suffix.lower() == ".xlsb":
        raise ExcelReadError("Không hỗ trợ định dạng .xlsb. Hãy mở bằng Excel và lưu lại dạng .xlsx.")

    try:
        size = p.stat().st_size
        with p.open("rb") as fh:
            head = fh.read(8)
    except PermissionError as exc:
        raise ExcelReadError(
            f"Không có quyền đọc file '{p.name}'. File có thể đang bị khóa bởi chương trình khác.",
            detail=repr(exc),
        ) from exc
    except OSError as exc:
        raise ExcelReadError(f"Không mở được file '{p.name}'.", detail=repr(exc)) from exc

    if size < _MIN_EXCEL_SIZE:
        raise ExcelReadError(f"File '{p.name}' quá nhỏ ({size} byte) - không phải file Excel hợp lệ.")
    if head.startswith(_OLE2_SIGNATURE):
        return FileFormat.XLS
    if head.startswith(_ZIP_SIGNATURE):
        return FileFormat.XLSX
    if head.lstrip().startswith(b"<"):
        raise ExcelReadError(
            f"File '{p.name}' thực chất là trang HTML/XML được đổi đuôi Excel. "
            "Hãy mở bằng Excel rồi 'Save As' thành Excel Workbook (.xlsx) hoặc Excel 97-2003 (.xls)."
        )
    raise ExcelReadError(f"File '{p.name}' không phải định dạng Excel được hỗ trợ (.xlsx, .xlsm, .xls).")


# =============================================================================
# 2. Đọc workbook
# =============================================================================

class WorkbookReader:
    """Mở 1 workbook để đọc. Dùng với 'with' để luôn giải phóng file.

        with WorkbookReader("luong.xlsx") as wb:
            print(wb.sheet_names)
            sheet = wb.read_sheet("HLCK05.26 (2)")
    """

    def __init__(self, path: PathLike) -> None:
        self.path: Path = Path(path)
        self.format: FileFormat = detect_format(self.path)
        self._cache: dict[str, SheetData] = {}
        self._values_wb: Any = None      # openpyxl workbook (giá trị) hoặc xlrd Book
        self._formulas_wb: Any = None    # openpyxl workbook (công thức) - mở khi cần
        self._infos: list[SheetInfo] = []
        self._open()

    # ----- vòng đời ---------------------------------------------------------
    def __enter__(self) -> "WorkbookReader":
        return self

    def __exit__(
        self,
        exc_type: Optional[type[BaseException]],
        exc: Optional[BaseException],
        tb: Optional[TracebackType],
    ) -> None:
        self.close()

    def close(self) -> None:
        for wb in (self._values_wb, self._formulas_wb):
            if wb is None:
                continue
            try:
                if self.format == FileFormat.XLSX:
                    wb.close()
                else:
                    wb.release_resources()
            except Exception:  # pragma: no cover - đóng file không được phép làm hỏng luồng chính
                log.debug("Lỗi khi đóng workbook %s", self.path, exc_info=True)
        self._values_wb = None
        self._formulas_wb = None

    # ----- thông tin sheet -------------------------------------------------
    @property
    def sheet_names(self) -> list[str]:
        return [i.name for i in self._infos]

    @property
    def sheet_infos(self) -> list[SheetInfo]:
        return list(self._infos)

    def sheet_info(self, name: str) -> SheetInfo:
        for info in self._infos:
            if info.name == name:
                return info
        raise ExcelReadError(
            f"Không tìm thấy sheet '{name}' trong file '{self.path.name}'.\n"
            f"Các sheet hiện có: {', '.join(repr(n) for n in self.sheet_names)}"
        )

    def read_sheet(self, name: str) -> SheetData:
        """Đọc toàn bộ nội dung 1 sheet (có cache)."""
        if name in self._cache:
            return self._cache[name]
        info = self.sheet_info(name)
        if self._values_wb is None:
            raise ExcelReadError("Workbook đã được đóng.")
        try:
            if self.format == FileFormat.XLSX:
                data = self._read_xlsx_sheet(info)
            else:
                data = self._read_xls_sheet(info)
        except ExcelReadError:
            raise
        except Exception as exc:
            raise ExcelReadError(
                f"Lỗi khi đọc sheet '{name}' trong file '{self.path.name}'.", detail=repr(exc)
            ) from exc
        self._cache[name] = data
        return data

    # ----- mở file ---------------------------------------------------------
    def _open(self) -> None:
        try:
            if self.format == FileFormat.XLSX:
                self._open_xlsx()
            else:
                self._open_xls()
        except ExcelReadError:
            self.close()
            raise
        except PermissionError as exc:
            self.close()
            raise ExcelReadError(
                f"Không có quyền đọc file '{self.path.name}' (có thể đang bị khóa).", detail=repr(exc)
            ) from exc
        except Exception as exc:
            self.close()
            raise ExcelReadError(
                f"Không mở được file '{self.path.name}'. File có thể bị hỏng hoặc đang đặt mật khẩu.",
                detail=repr(exc),
            ) from exc

    def _open_xlsx(self) -> None:
        import openpyxl

        with warnings.catch_warnings():
            # openpyxl cảnh báo về các tính năng nó không đọc (data validation mở rộng...) - bỏ qua khi chỉ đọc
            warnings.simplefilter("ignore")
            self._values_wb = openpyxl.load_workbook(
                self.path, read_only=True, data_only=True, keep_links=False
            )
        self._infos = [
            SheetInfo(name=ws.title, visible=(getattr(ws, "sheet_state", "visible") == "visible"))
            for ws in self._values_wb.worksheets
        ]

    def _open_xls(self) -> None:
        import xlrd

        try:
            # logfile: chuyển cảnh báo của xlrd vào log. Mặc định xlrd in ra sys.stdout,
            # mà sys.stdout = None khi chạy bản .exe không có cửa sổ console -> có thể gây lỗi.
            book = xlrd.open_workbook(str(self.path), on_demand=False, logfile=_XlrdLogWriter())
        except xlrd.XLRDError as exc:
            msg = str(exc).lower()
            if "encrypt" in msg:
                raise ExcelReadError(
                    f"File '{self.path.name}' đang đặt mật khẩu. Hãy gỡ mật khẩu rồi thử lại."
                ) from exc
            raise ExcelReadError(
                f"File '{self.path.name}' bị hỏng hoặc là file Excel có mật khẩu.", detail=repr(exc)
            ) from exc
        self._values_wb = book
        infos: list[SheetInfo] = []
        for idx in range(book.nsheets):
            sh = book.sheet_by_index(idx)
            infos.append(SheetInfo(name=sh.name, visible=(sh.visibility == 0)))
        self._infos = infos

    # ----- đọc .xlsx -------------------------------------------------------
    def _formula_workbook(self) -> Any:
        if self._formulas_wb is None:
            import openpyxl

            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                self._formulas_wb = openpyxl.load_workbook(
                    self.path, read_only=True, data_only=False, keep_links=False
                )
        return self._formulas_wb

    def _read_xlsx_sheet(self, info: SheetInfo) -> SheetData:
        data = SheetData(info=info)

        # Lượt 1: giá trị đã tính
        ws = self._values_wb[info.name]
        # Một số phần mềm ghi sai vùng dữ liệu (dimension) -> bỏ qua để đọc đủ mọi dòng
        ws.reset_dimensions()
        for row in ws.iter_rows():
            for cell in row:
                value = getattr(cell, "value", None)
                if value is None:
                    continue
                r, c = cell.row, cell.column
                is_error = getattr(cell, "data_type", None) == "e"
                if isinstance(value, str) and value == "" and not is_error:
                    continue
                data.cells[(r, c)] = CellData(value, classify_value(value, is_error=is_error))
                data.max_row = max(data.max_row, r)
                data.max_col = max(data.max_col, c)

        # Lượt 2: đánh dấu ô công thức (để phát hiện công thức chưa được Excel tính)
        fws = self._formula_workbook()[info.name]
        fws.reset_dimensions()
        for row in fws.iter_rows():
            for cell in row:
                if getattr(cell, "data_type", None) != "f":
                    continue
                r, c = cell.row, cell.column
                current = data.cells.get((r, c))
                if current is None:
                    data.cells[(r, c)] = CellData(None, CellKind.EMPTY, is_formula=True)
                else:
                    data.cells[(r, c)] = CellData(current.value, current.kind, is_formula=True)
                data.max_row = max(data.max_row, r)
                data.max_col = max(data.max_col, c)
        return data

    # ----- đọc .xls --------------------------------------------------------
    def _read_xls_sheet(self, info: SheetInfo) -> SheetData:
        import xlrd

        book = self._values_wb
        sh = book.sheet_by_name(info.name)
        data = SheetData(info=info)
        for r in range(sh.nrows):
            for c in range(sh.ncols):
                ctype = sh.cell_type(r, c)
                if ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    continue
                raw = sh.cell_value(r, c)
                cell = _convert_xls_cell(ctype, raw, book.datemode)
                if cell is None:
                    continue
                data.cells[(r + 1, c + 1)] = cell
                data.max_row = max(data.max_row, r + 1)
                data.max_col = max(data.max_col, c + 1)
        return data


def _convert_xls_cell(ctype: int, raw: Any, datemode: int) -> Optional[CellData]:
    import xlrd

    if ctype == xlrd.XL_CELL_TEXT:
        if raw == "":
            return None
        return CellData(raw, CellKind.TEXT)
    if ctype == xlrd.XL_CELL_NUMBER:
        return CellData(raw, CellKind.NUMBER)
    if ctype == xlrd.XL_CELL_BOOLEAN:
        return CellData(bool(raw), CellKind.BOOL)
    if ctype == xlrd.XL_CELL_ERROR:
        return CellData(xlrd.error_text_from_code.get(raw, f"#ERR{raw}"), CellKind.ERROR)
    if ctype == xlrd.XL_CELL_DATE:
        try:
            return CellData(xlrd.xldate.xldate_as_datetime(raw, datemode), CellKind.DATE)
        except Exception:  # ngày không hợp lệ -> giữ số gốc nhưng vẫn đánh dấu là ngày
            return CellData(raw, CellKind.DATE)
    return CellData(raw, classify_value(raw))


# =============================================================================
# 3. Hàm tiện ích
# =============================================================================

def list_sheets(path: PathLike) -> list[SheetInfo]:
    """Liệt kê các sheet trong file."""
    with WorkbookReader(path) as wb:
        return wb.sheet_infos


def read_sheet(path: PathLike, sheet_name: str) -> SheetData:
    """Đọc nhanh 1 sheet."""
    with WorkbookReader(path) as wb:
        return wb.read_sheet(sheet_name)


def cell_display_text(cell: CellData) -> str:
    """Chuỗi hiển thị của ô (dùng cho tiêu đề, preview)."""
    if cell.kind == CellKind.EMPTY or cell.value is None:
        return ""
    value = cell.value
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value)


def header_texts(sheet: SheetData, header_rows: tuple[int, ...] | list[int], col: int) -> list[str]:
    """Các tiêu đề (không rỗng) của 1 cột trong các dòng tiêu đề, theo thứ tự dòng."""
    texts: list[str] = []
    for r in header_rows:
        text = cell_display_text(sheet.cell(r, col)).strip()
        if text:
            texts.append(text)
    return texts


def column_headers(sheet: SheetData, header_rows: tuple[int, ...] | list[int]) -> dict[int, str]:
    """{số cột: tiêu đề ghép từ các dòng tiêu đề}. Dùng để hiển thị danh sách cột cho người dùng chọn."""
    result: dict[int, str] = {}
    for c in range(1, sheet.max_col + 1):
        texts = header_texts(sheet, header_rows, c)
        if texts:
            result[c] = " / ".join(" ".join(t.split()) for t in texts)
    return result


def find_header_columns(sheet: SheetData, header_rows: tuple[int, ...] | list[int], header: str) -> list[int]:
    """Tìm các cột có tiêu đề khớp (không phân biệt hoa thường/khoảng trắng/xuống dòng)."""
    wanted = normalize_header(header)
    if not wanted:
        return []
    found: list[int] = []
    for c in range(1, sheet.max_col + 1):
        if any(normalize_header(t) == wanted for t in header_texts(sheet, header_rows, c)):
            found.append(c)
    return found


# =============================================================================
# 4. Tiện ích dùng chung cho các bộ đọc theo cấu hình (file lương, file ngân hàng)
# =============================================================================

class SheetSelectionError(ToolError):
    """Không xác định được duy nhất 1 sheet -> giao diện cần cho người dùng chọn."""

    def __init__(self, user_message: str, candidates: list[str]) -> None:
        super().__init__(user_message)
        self.candidates: list[str] = candidates


def _norm_sheet_name(name: str) -> str:
    return name.strip().casefold()


def sheet_matches(name: str, patterns: tuple[str, ...] | list[str]) -> bool:
    """Tên sheet khớp 1 trong các mẫu ('*', '?'), không phân biệt hoa thường/khoảng trắng 2 đầu."""
    return any(fnmatch.fnmatchcase(_norm_sheet_name(name), _norm_sheet_name(p)) for p in patterns)


def select_sheet_name(
    sheet_names: list[str],
    exact: Optional[str],
    patterns: tuple[str, ...] | list[str],
    override: Optional[str] = None,
) -> str:
    """Xác định sheet cần đọc.

    Thứ tự ưu tiên:
      1. override (người dùng chọn trên giao diện) - phải tồn tại chính xác.
      2. exact - khớp chính xác, hoặc khớp duy nhất khi bỏ qua hoa/thường + khoảng trắng 2 đầu.
      3. patterns - phải khớp DUY NHẤT 1 sheet (không tự đoán khi có nhiều sheet khớp).
    """
    if override is not None:
        if override in sheet_names:
            return override
        raise SheetSelectionError(f"Không tìm thấy sheet '{override}'.", candidates=list(sheet_names))

    if exact:
        if exact in sheet_names:
            return exact
        loose = [n for n in sheet_names if _norm_sheet_name(n) == _norm_sheet_name(exact)]
        if len(loose) == 1:
            return loose[0]
        if not patterns:
            raise SheetSelectionError(
                f"Không tìm thấy sheet '{exact}'. Hãy chọn sheet cần đọc.", candidates=list(sheet_names)
            )

    matches = [n for n in sheet_names if sheet_matches(n, patterns)]
    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise SheetSelectionError(
            f"Không có sheet nào khớp mẫu {list(patterns)}. Hãy chọn sheet cần đọc.",
            candidates=list(sheet_names),
        )
    raise SheetSelectionError(
        f"Có {len(matches)} sheet khớp mẫu {list(patterns)}: "
        f"{', '.join(repr(m) for m in matches)}. Hãy chọn đúng sheet cần đọc.",
        candidates=matches,
    )


def check_column_headers(
    sheet: SheetData,
    header_rows: tuple[int, ...] | list[int],
    columns: dict[str, ColumnRef],
) -> list[Issue]:
    """So tiêu đề thực tế với tiêu đề mong đợi của từng cột đã mapping. Lệch -> lỗi HEADER_MISMATCH."""
    issues: list[Issue] = []
    for field_key, ref in columns.items():
        if not ref.expected_header:
            continue  # cột không có tiêu đề -> không kiểm tra được bằng tiêu đề
        actual = header_texts(sheet, header_rows, ref.index)
        wanted = normalize_header(ref.expected_header)
        if any(normalize_header(t) == wanted for t in actual):
            continue
        found_at = find_header_columns(sheet, header_rows, ref.expected_header)
        hint = (
            f" Tiêu đề này hiện nằm ở cột {', '.join(column_letter(c) for c in found_at)} "
            "-> cần cập nhật lại mapping."
            if found_at else " Không tìm thấy tiêu đề này ở cột nào khác."
        )
        actual_txt = " / ".join(" ".join(t.split()) for t in actual) or "(trống)"
        issues.append(Issue(
            code=IssueCode.HEADER_MISMATCH,
            severity=Severity.ERROR,
            message=(
                f"Cột {ref.letter} ({ref.label or field_key}): mong đợi tiêu đề '{ref.expected_header}' "
                f"nhưng thực tế là '{actual_txt}'.{hint}"
            ),
            sheet=sheet.name,
            column=ref.letter,
        ))
    return issues


def cell_problem(cell: CellData) -> Optional[tuple[IssueCode, str]]:
    """Lỗi ở mức ô Excel (trước khi chuẩn hóa giá trị): công thức chưa tính, ô báo lỗi #N/A..."""
    if cell.formula_not_calculated:
        return (
            IssueCode.FORMULA_NOT_CALCULATED,
            "công thức chưa được Excel tính. Hãy mở file bằng Excel, bấm Lưu (Ctrl+S) rồi chạy lại.",
        )
    if cell.kind == CellKind.ERROR:
        return IssueCode.EXCEL_ERROR_VALUE, f"ô đang báo lỗi Excel '{cell.value}'."
    return None
