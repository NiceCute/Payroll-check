"""Kiểm thử cho core.excel_writer (Backup, tạo file mới, giữ nguyên format/công thức/merge/width/height)."""

from __future__ import annotations

import hashlib
import os
import stat
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import Border, Font, PatternFill, Side

from core.excel_reader import WorkbookReader
from core.excel_writer import (
    ExcelWriteError,
    ExcelWriter,
    backup_file,
    build_output_path,
    execute_write_plan,
    is_excel_com_available,
    write_target_file,
)
from core.models import (
    CellUpdate,
    ResolutionAction,
    SafetyCode,
    SafetyError,
    SafetyFinding,
    Severity,
    SourceReadResult,
    SourceRecord,
    TargetWritePlan,
    ValidationResult,
)
from core.normalize import normalize_name_key
from core.reconciler import reconcile
from core.safety import is_file_locked, validate_write_plan
from core.source_reader import read_source
from core.target_reader import TEMPLATE_ACB, TEMPLATE_THONG_THUONG, read_target
from tests.helpers import build_acb_xls, build_tt_xls


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _src_rec(
    row: int,
    card_no: str,
    full_name: str,
    account_no: Optional[str],
    amount: Optional[int],
) -> SourceRecord:
    return SourceRecord(
        row=row,
        card_no=card_no,
        full_name=full_name,
        name_key=normalize_name_key(full_name),
        account_no=account_no,
        amount=amount,
    )


def _build_rich_target_xlsx(path: Path, *, protect_cnbd: bool = False) -> Path:
    """Tạo file .xlsx có đầy đủ format, border, màu nền, merge cell, công thức, độ rộng cột, chiều cao dòng."""
    wb = openpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "CNBD"

    # Merge cell tiêu đề A1:D1
    ws["A1"] = "DANH SÁCH CHI HỘ LƯƠNG ACB"
    ws.merge_cells("A1:D1")

    # Dòng tiêu đề 2
    ws["A2"], ws["B2"], ws["C2"], ws["D2"], ws["E2"] = (
        "STT",
        "Tên NV",
        "Số tài khoản",
        "Số tiền",
        "Công thức kiểm tra",
    )

    # Độ rộng cột & chiều cao dòng
    ws.column_dimensions["B"].width = 28.5
    ws.column_dimensions["C"].width = 18.0
    ws.column_dimensions["D"].width = 16.0
    ws.row_dimensions[1].height = 30.0
    ws.row_dimensions[3].height = 22.5
    ws.row_dimensions[4].height = 22.5

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill = PatternFill(fill_type="solid", fgColor="FFFFCC")
    font = Font(name="Arial", size=11, bold=True, color="003366")

    data = [
        (1, "NGUYEN THANH CONG", "24349081", 4_500_000),
        (2, "PHAM THI LIEN", "24230311", 5_100_000),
        (3, "TRAN DUC TRAN", "24112233", 3_200_000),
    ]
    for idx, name, acc, amt in data:
        r = idx + 2
        ws[f"A{r}"] = idx
        ws[f"B{r}"] = name
        ws[f"C{r}"] = acc
        ws[f"C{r}"].number_format = "@"
        ws[f"D{r}"] = amt
        ws[f"D{r}"].number_format = "#,##0"
        ws[f"D{r}"].border = border
        ws[f"D{r}"].fill = fill
        ws[f"D{r}"].font = font
        ws[f"E{r}"] = f"=D{r}*1"

    # Dòng tổng có công thức ở cột E (không nằm trong vùng dòng nhân viên)
    ws["E6"] = "=SUM(D3:D5)"

    if protect_cnbd:
        ws.protection.sheet = True

    # Sheet ẩn và sheet khóa bảo vệ giống mẫu thực tế
    hidden = wb.create_sheet("DSid")
    hidden["A1"] = "MA_NH"
    hidden.sheet_state = "hidden"

    locked = wb.create_sheet("Soạn lệnh")
    locked["A1"] = "Sheet khóa"
    locked.protection.sheet = True

    guide = wb.create_sheet("Hướng dẫn nhập liệu")
    guide["A1"] = "Không nhập dấu tiếng Việt, không thừa khoảng trắng."

    wb.save(path)
    return path


class TestExcelWriterXlsxPreservation(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _run_xlsx_write_test(self, engine: str) -> None:
        target_path = _build_rich_target_xlsx(self.tmp / "Mau_ACB.xlsx")
        orig_hash = _sha256(target_path)

        tgt_read = read_target(target_path, TEMPLATE_ACB)
        src_read = SourceReadResult(
            file_path=str(self.tmp / "BangLuong.xlsx"),
            sheet_name="HLCK06.26",
            profile_name="Bảng lương",
            records=[
                _src_rec(5, "0001", "NGUYEN THANH CONG", "24349081", 6_100_000),
                _src_rec(6, "0002", "PHAM THI LIEN", "24230311", 7_250_000),
                _src_rec(7, "0003", "TRAN DUC TRAN", "24112233", 3_200_000),
            ],
        )
        plan = reconcile(src_read, (tgt_read, TEMPLATE_ACB))
        val = validate_write_plan(plan)

        writer = ExcelWriter(
            output_dir=self.tmp / "output",
            backup_dir=self.tmp / "backup",
            engine=engine,  # type: ignore[arg-type]
        )
        res = writer.write(plan, val, TEMPLATE_ACB, tgt_read)

        # 1. File gốc hoàn toàn không bị thay đổi (SHA-256 giống hệt)
        self.assertEqual(_sha256(target_path), orig_hash)

        # 2. Bản backup được tạo và khớp SHA-256 với bản gốc
        self.assertEqual(len(res.target_outputs), 1)
        out_info = res.target_outputs[0]
        bak_path = Path(out_info.backup_file)
        out_path = Path(out_info.output_file)
        self.assertTrue(bak_path.exists())
        self.assertEqual(_sha256(bak_path), orig_hash)
        self.assertTrue(out_path.exists())
        self.assertNotEqual(out_path.resolve(), target_path.resolve())

        # 3. Kiểm tra nội dung, format, border, màu sắc, merge cell, công thức, width, height
        wb = openpyxl.load_workbook(out_path, data_only=False)
        try:
            self.assertEqual(wb.sheetnames, ["CNBD", "DSid", "Soạn lệnh", "Hướng dẫn nhập liệu"])
            self.assertEqual(wb["DSid"].sheet_state, "hidden")
            self.assertTrue(wb["Soạn lệnh"].protection.sheet)

            ws = wb["CNBD"]
            # Giá trị mới tại D3, D4, D5
            self.assertEqual(ws["D3"].value, 6_100_000)
            self.assertEqual(ws["D4"].value, 7_250_000)
            self.assertEqual(ws["D5"].value, 3_200_000)

            # Các cột định danh B, C giữ nguyên 100%
            self.assertEqual(ws["B3"].value, "NGUYEN THANH CONG")
            self.assertEqual(ws["C3"].value, "24349081")
            self.assertEqual(ws["C3"].number_format, "@")

            # Merge cell A1:D1 giữ nguyên
            merged_strs = {str(r) for r in ws.merged_cells.ranges}
            self.assertIn("A1:D1", merged_strs)

            # Công thức ở E3:E6 giữ nguyên
            self.assertEqual(ws["E3"].value, "=D3*1")
            self.assertEqual(ws["E6"].value, "=SUM(D3:D5)")

            # Format, border, font, fill của D3 giữ nguyên
            self.assertEqual(ws["D3"].number_format, "#,##0")
            self.assertEqual(ws["D3"].border.left.style, "thin")
            self.assertEqual(ws["D3"].font.name, "Arial")
            self.assertTrue(ws["D3"].font.bold)
            self.assertIn("FFFFCC", (ws["D3"].fill.fgColor.rgb or "").upper())

            # Chiều rộng cột & chiều cao dòng giữ nguyên
            self.assertAlmostEqual(ws.column_dimensions["B"].width or 0.0, 28.5, delta=0.2)
            self.assertAlmostEqual(ws.column_dimensions["D"].width or 0.0, 16.0, delta=0.2)
            self.assertAlmostEqual(ws.row_dimensions[3].height or 0.0, 22.5, delta=0.2)
        finally:
            wb.close()

    def test_write_xlsx_via_openpyxl_engine(self) -> None:
        self._run_xlsx_write_test("openpyxl")

    def test_write_xlsx_via_com_engine(self) -> None:
        self._run_xlsx_write_test("com")


class TestExcelWriterXlsViaCom(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_write_xls_acb_and_tt_preserves_sheets_and_updates_only_amount(self) -> None:
        acb_path = build_acb_xls(self.tmp / "4 MauChiHoLuongQuaTaiKhoanACB.xls")
        tt_path = build_tt_xls(self.tmp / "5 MauChiHoLuongThongThuong.xls")
        acb_hash = _sha256(acb_path)
        tt_hash = _sha256(tt_path)

        r_acb = read_target(acb_path, TEMPLATE_ACB)
        r_tt = read_target(tt_path, TEMPLATE_THONG_THUONG)

        src = SourceReadResult(
            file_path=str(self.tmp / "BangLuong.xlsx"),
            sheet_name="HLCK06.26",
            profile_name="Bảng lương",
            records=[
                # 3 nhân viên ACB
                _src_rec(5, "0001", "NGUYEN THANH CONG", "24349081", 8_111_222),
                _src_rec(6, "0002", "PHAM THI LIEN", "24230311", 9_333_444),
                _src_rec(7, "0003", "TRAN DUC TRAN", "24112233", 4_555_666),
                # 3 nhân viên Thông thường (có số 0 đầu ở TK)
                _src_rec(8, "0004", "NGUYEN VAN AN", "0587952606", 6_777_888),
                _src_rec(9, "0005", "LE THI BINH", "0071001234567", 7_888_999),
                _src_rec(10, "0006", "HOANG MINH", "19034567890011", 5_400_000),
            ],
        )

        plan = reconcile(src, [(r_acb, TEMPLATE_ACB), (r_tt, TEMPLATE_THONG_THUONG)])
        val = validate_write_plan(plan)
        exec_res = execute_write_plan(
            plan,
            val,
            [TEMPLATE_ACB, TEMPLATE_THONG_THUONG],
            [r_acb, r_tt],
            output_dir=self.tmp / "output",
            backup_dir=self.tmp / "backup",
        )

        # File gốc không thay đổi
        self.assertEqual(_sha256(acb_path), acb_hash)
        self.assertEqual(_sha256(tt_path), tt_hash)
        self.assertEqual(len(exec_res.target_outputs), 2)

        # Đọc lại 2 file .xls kết quả
        out_acb = Path(exec_res.target_outputs[0].output_file)
        out_tt = Path(exec_res.target_outputs[1].output_file)

        r_acb_after = read_target(out_acb, TEMPLATE_ACB)
        self.assertEqual(
            [r.current_amount for r in r_acb_after.rows],
            [8_111_222, 9_333_444, 4_555_666],
        )
        self.assertIsNotNone(r_acb_after.guide)

        r_tt_after = read_target(out_tt, TEMPLATE_THONG_THUONG)
        self.assertEqual(
            [r.current_amount for r in r_tt_after.rows],
            [6_777_888, 7_888_999, 5_400_000],
        )
        # Số TK có số 0 đầu và Mã ngân hàng được giữ nguyên 100%
        self.assertEqual(r_tt_after.rows[0].account_no, "0587952606")
        self.assertEqual(r_tt_after.rows[1].account_no, "0071001234567")
        self.assertEqual(r_tt_after.rows[0].bank_code, "79203001")

    def test_write_set_zero_and_clear_for_x5(self) -> None:
        target_path = _build_rich_target_xlsx(self.tmp / "Mau_ACB.xlsx")
        tgt_read = read_target(target_path, TEMPLATE_ACB)
        src_read = SourceReadResult(
            file_path=str(self.tmp / "BangLuong.xlsx"),
            sheet_name="HLCK06.26",
            profile_name="Bảng lương",
            records=[
                _src_rec(5, "0001", "NGUYEN THANH CONG", "24349081", 6_000_000),
            ],
        )
        plan = reconcile(src_read, (tgt_read, TEMPLATE_ACB))
        # Dòng 4 chọn SET_ZERO, dòng 5 chọn CLEAR
        x5_list = [e for e in plan.exceptions if e.code.value == "X5"]
        self.assertEqual(len(x5_list), 2)
        plan.resolve(x5_list[0].id, ResolutionAction.SET_ZERO)
        plan.resolve(x5_list[1].id, ResolutionAction.CLEAR)

        res = execute_write_plan(
            plan,
            output_dir=self.tmp / "output",
            backup_dir=self.tmp / "backup",
            engine="openpyxl",
        )
        out_path = Path(res.target_outputs[0].output_file)
        wb = openpyxl.load_workbook(out_path, data_only=True)
        try:
            ws = wb["CNBD"]
            self.assertEqual(ws["D3"].value, 6_000_000)
            self.assertEqual(ws["D4"].value, 0)
            self.assertIsNone(ws["D5"].value)
        finally:
            wb.close()


class TestExcelWriterSafetyAndErrors(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_refuses_to_overwrite_original_file_r8(self) -> None:
        target_path = _build_rich_target_xlsx(self.tmp / "Mau_ACB.xlsx")
        tgt_read = read_target(target_path, TEMPLATE_ACB)
        src_read = SourceReadResult(
            file_path=str(self.tmp / "BangLuong.xlsx"),
            sheet_name="HLCK",
            profile_name="Bảng lương",
            records=[
                _src_rec(5, "0001", "NGUYEN THANH CONG", "24349081", 6_000_000),
                _src_rec(6, "0002", "PHAM THI LIEN", "24230311", 5_100_000),
                _src_rec(7, "0003", "TRAN DUC TRAN", "24112233", 3_200_000),
            ],
        )
        plan = reconcile(src_read, (tgt_read, TEMPLATE_ACB))
        with self.assertRaises(SafetyError):
            execute_write_plan(
                plan,
                output_paths=target_path,
                backup_dir=self.tmp / "backup",
                engine="openpyxl",
            )

    def test_refuses_to_write_when_validation_failed(self) -> None:
        target_path = _build_rich_target_xlsx(self.tmp / "Mau_ACB.xlsx")
        tgt_read = read_target(target_path, TEMPLATE_ACB)
        src_read = SourceReadResult(
            file_path=str(self.tmp / "BangLuong.xlsx"),
            sheet_name="HLCK",
            profile_name="Bảng lương",
            records=[_src_rec(5, "0001", "NGUYEN THANH CONG", "24349081", 6_000_000)],
        )
        # Còn 2 dòng X5 chưa xử lý -> plan chưa an toàn
        plan = reconcile(src_read, (tgt_read, TEMPLATE_ACB))
        bad_val = ValidationResult(
            errors=[SafetyFinding(SafetyCode.UNRESOLVED_EXCEPTION, Severity.ERROR, "Chưa xử lý X5")]
        )
        with self.assertRaises(SafetyError):
            execute_write_plan(plan, bad_val, engine="openpyxl")

    def test_refuses_protected_sheet_and_cleans_up_output(self) -> None:
        target_path = _build_rich_target_xlsx(self.tmp / "Protected.xlsx", protect_cnbd=True)
        tp = TargetWritePlan(
            file_path=str(target_path),
            file_format="xlsx",
            sheet_name="CNBD",
            template_name=TEMPLATE_ACB.name,
            amount_column="D",
            updates=[CellUpdate(str(target_path), "CNBD", 3, "D", 4_500_000, 6_000_000, "0001", "24349081")],
        )
        out_path = self.tmp / "output" / "Protected_KetQua.xlsx"
        with self.assertRaises(ExcelWriteError):
            write_target_file(
                tp,
                target_path,
                out_path,
                backup_dir=self.tmp / "backup",
                engine="openpyxl",
            )
        self.assertFalse(out_path.exists())

    def test_refuses_formula_or_merged_cell_and_cleans_up_output(self) -> None:
        target_path = _build_rich_target_xlsx(self.tmp / "FormulaTarget.xlsx")
        wb = openpyxl.load_workbook(target_path)
        wb["CNBD"]["D3"] = "=1000+2000"
        wb.save(target_path)
        wb.close()

        tp = TargetWritePlan(
            file_path=str(target_path),
            file_format="xlsx",
            sheet_name="CNBD",
            amount_column="D",
            template_name=TEMPLATE_ACB.name,
            updates=[CellUpdate(str(target_path), "CNBD", 3, "D", 3000, 6_000_000, "0001", "24349081")],
        )
        out_path = self.tmp / "output" / "FormulaTarget_KetQua.xlsx"
        with self.assertRaises(ExcelWriteError):
            write_target_file(
                tp,
                target_path,
                out_path,
                backup_dir=self.tmp / "backup",
                engine="openpyxl",
            )
        self.assertFalse(out_path.exists())

    def test_build_output_path_and_backup_helpers(self) -> None:
        src = _build_rich_target_xlsx(self.tmp / "Mau_ACB.xlsx")
        # Nếu output_dir trùng thư mục gốc và suffix="" -> vẫn tự gắn hậu tố để không trùng gốc
        out = build_output_path(src, self.tmp, suffix="")
        self.assertNotEqual(out.resolve(), src.resolve())

        bak = backup_file(src, self.tmp / "backup")
        self.assertTrue(bak.exists())
        self.assertEqual(_sha256(bak), _sha256(src))

    def test_template_row_bound_check_prevents_writing_into_header(self) -> None:
        target_path = _build_rich_target_xlsx(self.tmp / "Mau_ACB.xlsx")
        tp = TargetWritePlan(
            file_path=str(target_path),
            file_format="xlsx",
            sheet_name="CNBD",
            template_name=TEMPLATE_ACB.name,
            amount_column="D",
            updates=[CellUpdate(str(target_path), "CNBD", 2, "D", 0, 5_000_000, "0001", "24349081")],
        )
        out_path = self.tmp / "output" / "BadRow.xlsx"
        with self.assertRaises(SafetyError):
            write_target_file(
                tp,
                target_path,
                out_path,
                backup_dir=self.tmp / "backup",
                template=replace(TEMPLATE_ACB, data_start_row=3),
                engine="openpyxl",
            )

    def test_uat15_readonly_input_files(self) -> None:
        """UAT-15: File đầu vào có thuộc tính Read-Only vẫn đọc/backup/xuất file mới thành công."""
        target_xlsx = _build_rich_target_xlsx(self.tmp / "ReadOnly_ACB.xlsx")
        target_xls = build_acb_xls(self.tmp / "ReadOnly_ACB.xls")
        orig_xlsx_hash = _sha256(target_xlsx)
        orig_xls_hash = _sha256(target_xls)

        # Đặt cờ Read-Only cho cả 2 file đầu vào
        os.chmod(target_xlsx, stat.S_IREAD)
        os.chmod(target_xls, stat.S_IREAD)
        try:
            self.assertFalse(bool(target_xlsx.stat().st_mode & stat.S_IWRITE))
            self.assertFalse(bool(target_xls.stat().st_mode & stat.S_IWRITE))

            # is_file_locked không báo nhầm file Read-Only là đang bị khóa bởi Excel
            self.assertFalse(is_file_locked(target_xlsx))
            self.assertFalse(is_file_locked(target_xls))

            src_read = SourceReadResult(
                file_path=str(self.tmp / "BangLuong.xlsx"),
                sheet_name="HLCK06.26",
                profile_name="Bảng lương",
                records=[
                    _src_rec(5, "0001", "NGUYEN THANH CONG", "24349081", 6_111_000),
                    _src_rec(6, "0002", "PHAM THI LIEN", "24230311", 5_222_000),
                    _src_rec(7, "0003", "TRAN DUC TRAN", "24112233", 3_333_000),
                ],
            )

            # 1. Kiểm thử với file .xlsx Read-Only qua openpyxl
            tgt_read_xlsx = read_target(target_xlsx, TEMPLATE_ACB)
            plan_xlsx = reconcile(src_read, (tgt_read_xlsx, TEMPLATE_ACB))
            val_xlsx = validate_write_plan(plan_xlsx, check_disk=True)
            self.assertTrue(val_xlsx.is_safe)

            res_xlsx = execute_write_plan(
                plan_xlsx,
                val_xlsx,
                TEMPLATE_ACB,
                tgt_read_xlsx,
                output_dir=self.tmp / "output",
                backup_dir=self.tmp / "backup",
                engine="openpyxl",
            )
            out_xlsx = Path(res_xlsx.target_outputs[0].output_file)
            self.assertTrue(out_xlsx.exists())
            self.assertTrue(bool(out_xlsx.stat().st_mode & stat.S_IWRITE))
            self.assertEqual(_sha256(target_xlsx), orig_xlsx_hash)
            self.assertFalse(bool(target_xlsx.stat().st_mode & stat.S_IWRITE))

            # 2. Kiểm thử với file .xls Read-Only qua Excel COM
            if is_excel_com_available():
                tgt_read_xls = read_target(target_xls, TEMPLATE_ACB)
                plan_xls = reconcile(src_read, (tgt_read_xls, TEMPLATE_ACB))
                val_xls = validate_write_plan(plan_xls, check_disk=True)
                self.assertTrue(val_xls.is_safe)

                res_xls = execute_write_plan(
                    plan_xls,
                    val_xls,
                    TEMPLATE_ACB,
                    tgt_read_xls,
                    output_dir=self.tmp / "output",
                    backup_dir=self.tmp / "backup",
                )
                out_xls = Path(res_xls.target_outputs[0].output_file)
                self.assertTrue(out_xls.exists())
                self.assertTrue(bool(out_xls.stat().st_mode & stat.S_IWRITE))
                self.assertEqual(_sha256(target_xls), orig_xls_hash)
                self.assertFalse(bool(target_xls.stat().st_mode & stat.S_IWRITE))
                after_xls = read_target(out_xls, TEMPLATE_ACB)
                self.assertEqual(
                    [r.current_amount for r in after_xls.rows],
                    [6_111_000, 5_222_000, 3_333_000],
                )
        finally:
            os.chmod(target_xlsx, stat.S_IWRITE | stat.S_IREAD)
            os.chmod(target_xls, stat.S_IWRITE | stat.S_IREAD)


@unittest.skipUnless(
    os.environ.get("CHL_REAL_SOURCE")
    and os.environ.get("CHL_REAL_ACB")
    and os.environ.get("CHL_REAL_TT"),
    "Đặt CHL_REAL_SOURCE, CHL_REAL_ACB, CHL_REAL_TT để kiểm thử ghi trên bộ file tháng 05/2026 thật.",
)
class TestExcelWriterRealMay2026Files(unittest.TestCase):
    def test_write_real_may_2026_xls_files_end_to_end(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            src_path = Path(os.environ["CHL_REAL_SOURCE"])
            acb_path = Path(os.environ["CHL_REAL_ACB"])
            tt_path = Path(os.environ["CHL_REAL_TT"])

            acb_orig_hash = _sha256(acb_path)
            tt_orig_hash = _sha256(tt_path)

            src = read_source(src_path, sheet_name="HLCK05.26 (2)")
            r_acb = read_target(acb_path, TEMPLATE_ACB)
            r_tt = read_target(tt_path, TEMPLATE_THONG_THUONG)

            plan = reconcile(src, [(r_acb, TEMPLATE_ACB), (r_tt, TEMPLATE_THONG_THUONG)])
            val = validate_write_plan(plan)

            exec_res = execute_write_plan(
                plan,
                val,
                [TEMPLATE_ACB, TEMPLATE_THONG_THUONG],
                [r_acb, r_tt],
                output_dir=tmp / "output",
                backup_dir=tmp / "backup",
            )

            # File gốc hoàn toàn nguyên vẹn
            self.assertEqual(_sha256(acb_path), acb_orig_hash)
            self.assertEqual(_sha256(tt_path), tt_orig_hash)
            self.assertEqual(len(exec_res.target_outputs), 2)
            self.assertIsNotNone(exec_res.source_backup_file)

            out_acb = Path(exec_res.target_outputs[0].output_file)
            out_tt = Path(exec_res.target_outputs[1].output_file)
            self.assertTrue(out_acb.exists())
            self.assertTrue(out_tt.exists())

            # Đọc lại 2 file .xls đã xuất và kiểm tra số dòng + tổng tiền + danh sách sheet
            r_acb_out = read_target(out_acb, TEMPLATE_ACB)
            r_tt_out = read_target(out_tt, TEMPLATE_THONG_THUONG)
            self.assertEqual(len(r_acb_out.rows), 79)
            self.assertEqual(r_acb_out.summary()["current_total_amount"], 405_476_927)
            self.assertEqual(len(r_tt_out.rows), 7)
            self.assertEqual(r_tt_out.summary()["current_total_amount"], 39_661_674)

            with WorkbookReader(out_tt) as wb_tt:
                self.assertEqual(
                    wb_tt.sheet_names,
                    ["Dán giá trị vào đây để TT", "Soạn lệnh", "Hướng dẫn nhập liệu", "DSid", "dstinhNH"],
                )


if __name__ == "__main__":
    unittest.main()
