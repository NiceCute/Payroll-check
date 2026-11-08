"""
test_preview_table.py - Unit test cho GUI Module 2 (gui/preview_table.py).
"""

from __future__ import annotations

import shutil
import tempfile
import tkinter as tk
import unittest
from pathlib import Path

from core.models import ExceptionCode, ResolutionAction, ValidationResult, WritePlan
from core.reconciler import reconcile, resolve_exception
from core.safety import validate_write_plan
from core.source_reader import read_source
from core.target_reader import TEMPLATE_ACB, TEMPLATE_THONG_THUONG, read_target
from gui.preview_table import (
    PreviewColumn,
    PreviewFilter,
    PreviewTable,
    TAG_EXCEPTION,
    TAG_UPDATE,
    TAG_WARNING,
    ValidationFilter,
)
from tests.helpers import build_source_xlsx, build_target_xlsx, build_tt_xls


class TestPreviewTable(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="chl_test_preview_table_"))
        self.root = tk.Tk()
        self.root.withdraw()
        self.table = PreviewTable(self.root)

    def tearDown(self) -> None:
        try:
            self.table.destroy()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def _build_plan_with_exceptions_and_warnings(self) -> tuple[WritePlan, ValidationResult]:
        src_path = build_source_xlsx(
            self.tmp_dir / "bang_luong.xlsx",
            rows=[
                ("HL001", "Nguyễn Văn An", "NGUYEN VAN AN", "0587952606", 6_500_000),
                ("HL002", "Trần Thị Bình", "TRAN THI BINH", "24349081", 8_200_000),
                ("HL003", "Lê Hoàng Minh", "LE HOANG MINH", None, 4_900_000),  # X3_MISSING_ACCOUNT (Warning, đã SKIP)
            ],
        )
        acb_path = build_target_xlsx(
            self.tmp_dir / "acb.xlsx",
            rows=[
                ("PHAM TEN KHAC", "24349081", 5_100_000),  # X6_TARGET_MISMATCH (chưa xử lý -> Exception)
                ("NGUOI NGHI VIEC", "999888777", 3_200_000),  # X5_TARGET_NOT_IN_SOURCE (chưa xử lý -> Exception)
            ],
            with_guide=True,
        )
        tt_path = build_tt_xls(
            self.tmp_dir / "tt.xls",
            rows=[
                ("NGUYEN VAN AN", 6_000_000, "79203001", "0587952606"),
            ],
        )

        src_res = read_source(src_path)
        acb_res = read_target(acb_path, TEMPLATE_ACB)
        tt_res = read_target(tt_path, TEMPLATE_THONG_THUONG)
        plan = reconcile(src_res, [(acb_res, TEMPLATE_ACB), (tt_res, TEMPLATE_THONG_THUONG)])
        val_res = validate_write_plan(plan)
        return plan, val_res

    def test_columns_display_and_highlight_exception_and_warning(self) -> None:
        plan, val_res = self._build_plan_with_exceptions_and_warnings()
        rows = self.table.load_data(plan, validation_result=val_res)

        # Có 4 dòng: 3 dòng nguồn (HL001, HL002, HL003) + 1 dòng đích dư X5
        self.assertEqual(len(rows), 4)

        # Dòng 0 (HL001 - Thông thường): khớp hoàn toàn, có Mã NH 79203001, tag_update
        r0 = rows[0]
        self.assertEqual(r0.card_no, "HL001")
        self.assertEqual(r0.account_no, "0587952606")
        self.assertEqual(r0.bank_code, "79203001")
        self.assertEqual(r0.display_amount, 6_500_000)
        self.assertEqual(r0.validation_status, "OK")
        self.assertEqual(r0.row_tag, TAG_UPDATE)
        self.assertEqual(self.table.get_row_tags(0), (TAG_UPDATE,))

        # Dòng 1 (HL002 - X6 lệch tên chưa xử lý): Highlight Exception (TAG_EXCEPTION)
        r1 = rows[1]
        self.assertEqual(r1.card_no, "HL002")
        self.assertTrue(r1.has_exception)
        self.assertTrue(r1.is_unresolved)
        self.assertEqual(r1.validation_status, "ERROR")
        self.assertEqual(r1.row_tag, TAG_EXCEPTION)
        self.assertEqual(self.table.get_row_tags(1), (TAG_EXCEPTION,))

        # Dòng 2 (HL003 - X3 thiếu TK, đã tự động SKIP, mức WARNING): Highlight Warning (TAG_WARNING)
        r2 = rows[2]
        self.assertEqual(r2.card_no, "HL003")
        self.assertTrue(r2.has_warning)
        self.assertEqual(r2.validation_status, "WARNING")
        self.assertEqual(r2.row_tag, TAG_WARNING)
        self.assertEqual(self.table.get_row_tags(2), (TAG_WARNING,))

        # Dòng 3 (X5 dòng đích dư chưa xử lý): Highlight Exception (TAG_EXCEPTION)
        r3 = rows[3]
        self.assertEqual(r3.account_no, "999888777")
        self.assertEqual(r3.row_tag, TAG_EXCEPTION)
        self.assertEqual(self.table.get_row_tags(3), (TAG_EXCEPTION,))

    def test_filter_search_and_sort(self) -> None:
        plan, val_res = self._build_plan_with_exceptions_and_warnings()
        self.table.load_data(plan, validation_result=val_res)

        # 1. Search không dấu theo tên "nguyen van an"
        res_search = self.table.set_search_query("nguyen van an")
        self.assertEqual(len(res_search), 1)
        self.assertEqual(res_search[0].card_no, "HL001")

        # Search theo mã ngân hàng "79203001"
        res_bank = self.table.set_search_query("79203001")
        self.assertEqual(len(res_bank), 1)
        self.assertEqual(res_bank[0].account_no, "0587952606")
        self.table.clear_search()
        self.assertEqual(len(self.table.visible_rows()), 4)

        # 2. Filter theo trạng thái Match & Exception
        self.assertEqual(len(self.table.set_filter(PreviewFilter.WILL_UPDATE)), 1)
        self.assertEqual(len(self.table.set_filter(PreviewFilter.EXCEPTIONS_ONLY)), 3)
        self.assertEqual(len(self.table.set_filter(PreviewFilter.UNRESOLVED_ONLY)), 2)
        self.assertEqual(len(self.table.set_filter(PreviewFilter.WARNINGS_ONLY)), 3)

        # 3. Filter theo Validation status
        self.table.reset_filters()
        self.assertEqual(len(self.table.set_validation_filter(ValidationFilter.OK)), 1)
        self.assertEqual(len(self.table.set_validation_filter(ValidationFilter.WARNING)), 1)
        self.assertEqual(len(self.table.set_validation_filter(ValidationFilter.ERROR)), 2)
        self.table.reset_filters()

        # 4. Sort theo Số tiền tăng dần và giảm dần
        asc_rows = self.table.sort_by(PreviewColumn.AMOUNT, descending=False)
        self.assertEqual(
            [r.display_amount for r in asc_rows],
            [3_200_000, 4_900_000, 6_500_000, 8_200_000],
        )
        desc_rows = self.table.sort_by(PreviewColumn.AMOUNT, descending=True)
        self.assertEqual(
            [r.display_amount for r in desc_rows],
            [8_200_000, 6_500_000, 4_900_000, 3_200_000],
        )

        # Sort theo Họ và tên
        name_rows = self.table.sort_by(PreviewColumn.FULL_NAME, descending=False)
        self.assertEqual(name_rows[0].card_no, "HL003")  # Lê Hoàng Minh

    def test_after_resolving_exceptions_tags_update_to_warning(self) -> None:
        plan, _ = self._build_plan_with_exceptions_and_warnings()
        for exc in plan.exceptions:
            if exc.code == ExceptionCode.X6_TARGET_MISMATCH:
                resolve_exception(plan, exc.id, ResolutionAction.CONFIRM_WRITE)
            elif exc.code == ExceptionCode.X5_TARGET_NOT_IN_SOURCE:
                resolve_exception(plan, exc.id, ResolutionAction.SET_ZERO)

        val_res = validate_write_plan(plan)
        rows = self.table.load_data(plan, validation_result=val_res)
        # Khi đã xử lý hết ngoại lệ -> không còn dòng nào mang TAG_EXCEPTION
        self.assertEqual([r for r in rows if r.row_tag == TAG_EXCEPTION], [])
        self.assertEqual(rows[1].row_tag, TAG_WARNING)


if __name__ == "__main__":
    unittest.main()
