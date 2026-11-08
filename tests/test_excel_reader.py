import importlib.util
import shutil
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

import openpyxl

from core.excel_reader import (
    ExcelReadError,
    FileFormat,
    WorkbookReader,
    column_headers,
    detect_format,
    find_header_columns,
    list_sheets,
    read_sheet,
)
from core.models import CellKind, ConfigError, Issue, IssueCode, Severity, ToolError, column_index, column_letter
from tests.helpers import build_source_xls, build_source_xlsx

HAS_XLWT = importlib.util.find_spec("xlwt") is not None


class TempDirTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="chl_test_"))

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)


class TestColumnLetters(unittest.TestCase):
    def test_roundtrip(self) -> None:
        for letter, idx in (("A", 1), ("P", 16), ("X", 24), ("Z", 26), ("AA", 27), ("XFD", 16384)):
            self.assertEqual(column_index(letter), idx)
            self.assertEqual(column_letter(idx), letter)
        self.assertEqual(column_index(" p "), 16)

    def test_invalid(self) -> None:
        for bad in ("", "1", "A1", "XFE", "ABCD"):
            with self.assertRaises(ConfigError):
                column_index(bad)
        with self.assertRaises(ConfigError):
            column_letter(0)


class TestDetectFormat(TempDirTest):
    def test_xlsx_and_renamed(self) -> None:
        p = build_source_xlsx(self.tmp / "a.xlsx")
        self.assertEqual(detect_format(p), FileFormat.XLSX)
        renamed = self.tmp / "a_renamed.xls"
        shutil.copy(p, renamed)
        self.assertEqual(detect_format(renamed), FileFormat.XLSX)  # theo nội dung, không theo đuôi

    @unittest.skipUnless(HAS_XLWT, "cần xlwt để tạo file .xls")
    def test_xls(self) -> None:
        p = build_source_xls(self.tmp / "b.xls")
        self.assertEqual(detect_format(p), FileFormat.XLS)

    def test_errors(self) -> None:
        with self.assertRaises(ExcelReadError):
            detect_format(self.tmp / "khong_co.xlsx")
        html = self.tmp / "web.xls"
        html.write_text("<html><body>" + "x" * 1000 + "</body></html>", encoding="utf-8")
        with self.assertRaisesRegex(ExcelReadError, "HTML"):
            detect_format(html)
        tiny = self.tmp / "tiny.xlsx"
        tiny.write_bytes(b"PK\x03\x04abc")
        with self.assertRaises(ExcelReadError):
            detect_format(tiny)
        lock = self.tmp / "~$luong.xlsx"
        lock.write_bytes(b"x" * 2000)
        with self.assertRaisesRegex(ExcelReadError, "file tạm"):
            detect_format(lock)

    def test_corrupt_zip(self) -> None:
        bad = self.tmp / "hong.xlsx"
        bad.write_bytes(b"PK\x03\x04" + b"\x00" * 2000)
        with self.assertRaises(ExcelReadError):
            WorkbookReader(bad)


class TestReadXlsx(TempDirTest):
    def _make(self) -> Path:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = "Tiêu đề\n2 dòng"
        ws["B1"] = "Số tiền"
        ws["A2"] = "0123"
        ws["B2"] = 1000
        ws["C2"] = 1.5
        ws["D2"] = datetime(2026, 5, 31)
        ws["E2"] = True
        ws["F2"] = "=B2*2"     # openpyxl không tính công thức -> không có giá trị cache
        ws["G2"] = "#N/A"      # openpyxl lưu thành ô lỗi
        hidden = wb.create_sheet("An")
        hidden.sheet_state = "hidden"
        hidden["A1"] = "x"
        p = self.tmp / "kinds.xlsx"
        wb.save(p)
        return p

    def test_sheets_and_visibility(self) -> None:
        infos = list_sheets(self._make())
        self.assertEqual([i.name for i in infos], ["Data", "An"])
        self.assertTrue(infos[0].visible)
        self.assertFalse(infos[1].visible)

    def test_cell_kinds(self) -> None:
        s = read_sheet(self._make(), "Data")
        self.assertEqual(s.cell_at(2, "A").value, "0123")
        self.assertEqual(s.cell_at(2, "A").kind, CellKind.TEXT)
        self.assertEqual(s.cell_at(2, "B").kind, CellKind.NUMBER)
        self.assertEqual(s.cell_at(2, "D").kind, CellKind.DATE)
        self.assertEqual(s.cell_at(2, "E").kind, CellKind.BOOL)
        self.assertEqual(s.cell_at(2, "G").kind, CellKind.ERROR)
        f = s.cell_at(2, "F")
        self.assertTrue(f.is_formula)
        self.assertTrue(f.formula_not_calculated)
        self.assertTrue(s.cell(100, 100).is_empty)
        self.assertEqual((s.max_row, s.max_col), (2, 7))

    def test_sheet_not_found(self) -> None:
        with self.assertRaisesRegex(ExcelReadError, "Các sheet hiện có"):
            read_sheet(self._make(), "KhongCo")

    def test_header_helpers(self) -> None:
        p = build_source_xlsx(self.tmp / "src.xlsx")
        s = read_sheet(p, "HLCK06.26 (2)")
        headers = column_headers(s, (3, 4))
        self.assertEqual(headers[2], "Số thẻ")
        self.assertEqual(headers[16], "Hồng Lâm Chuyển Khoản")
        self.assertEqual(find_header_columns(s, (3, 4), "hồng lâm chuyển khoản"), [16])
        self.assertEqual(find_header_columns(s, (3, 4), "Không có"), [])


@unittest.skipUnless(HAS_XLWT, "cần xlwt để tạo file .xls")
class TestReadXls(TempDirTest):
    def test_read(self) -> None:
        p = build_source_xls(self.tmp / "src.xls")
        with WorkbookReader(p) as wb:
            self.assertEqual(wb.format, FileFormat.XLS)
            s = wb.read_sheet("HLCK06.26 (2)")
        self.assertEqual(s.cell_at(6, "X").value, "0587952606")  # text giữ số 0 đầu
        self.assertEqual(s.cell_at(6, "P").kind, CellKind.NUMBER)
        self.assertEqual(s.cell_at(3, "B").value, "Số thẻ")

    def test_kinds_date_bool(self) -> None:
        import xlwt

        wb = xlwt.Workbook()
        ws = wb.add_sheet("S")
        ws.write(0, 0, datetime(2026, 5, 31), xlwt.easyxf(num_format_str="DD/MM/YYYY"))
        ws.write(0, 1, True)
        ws.write(0, 2, "")
        p = self.tmp / "k.xls"
        wb.save(str(p))
        s = read_sheet(p, "S")
        self.assertEqual(s.cell(1, 1).kind, CellKind.DATE)
        self.assertEqual(s.cell(1, 1).value, datetime(2026, 5, 31))
        self.assertEqual(s.cell(1, 2).kind, CellKind.BOOL)
        self.assertTrue(s.cell(1, 3).is_empty)

    def test_works_without_stdout(self) -> None:
        """Bản .exe không có console: sys.stdout = None -> xlrd không được ghi ra stdout."""
        p = build_source_xls(self.tmp / "src.xls")
        saved = sys.stdout
        sys.stdout = None
        try:
            with WorkbookReader(p) as wb:
                wb.read_sheet(wb.sheet_names[0])
        finally:
            sys.stdout = saved


class TestModelsMisc(unittest.TestCase):
    def test_issue_str(self) -> None:
        i = Issue(IssueCode.ACCOUNT_MISSING, Severity.WARNING, "Thiếu TK", sheet="S", row=5, column="X")
        self.assertEqual(str(i), "[WARNING] Thiếu TK (sheet 'S', ô X5)")
        self.assertEqual(Issue(IssueCode.NO_RECORDS, Severity.ERROR, "m", column="B").location(), "cột B")
        self.assertEqual(Issue(IssueCode.NO_RECORDS, Severity.ERROR, "m", row=3).location(), "dòng 3")

    def test_tool_error_str(self) -> None:
        self.assertEqual(str(ToolError("Lỗi")), "Lỗi")
        self.assertEqual(str(ToolError("Lỗi", detail="x")), "Lỗi (x)")

    def test_row_is_empty(self) -> None:
        p = Path(tempfile.mkdtemp()) / "r.xlsx"
        try:
            build_source_xlsx(p)
            s = read_sheet(p, "HLCK06.26 (2)")
            self.assertTrue(s.row_is_empty(5, [column_index("P"), column_index("X")]))
            self.assertFalse(s.row_is_empty(6))
            self.assertTrue(s.row_is_empty(999))
        finally:
            shutil.rmtree(p.parent, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
