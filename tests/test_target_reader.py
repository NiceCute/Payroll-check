"""Kiểm thử cho core.target_reader và các model liên quan (TargetTemplate, TargetRules, TargetReadResult)."""

from __future__ import annotations

import os
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path

import openpyxl

from core.excel_reader import SheetSelectionError
from core.models import (
    CellData,
    CellKind,
    ColumnRef,
    ConfigError,
    IssueCode,
    Severity,
    SheetData,
    SheetInfo,
    TargetRules,
    TargetTemplate,
)
from core.target_reader import (
    TEMPLATE_ACB,
    TEMPLATE_THONG_THUONG,
    check_rules,
    parse_guide,
    parse_target_sheet,
    read_target,
)
from tests.helpers import build_acb_xls, build_target_xlsx, build_tt_xls


class TestTargetTemplateModel(unittest.TestCase):
    def test_roundtrip_dict(self) -> None:
        for tpl in (TEMPLATE_ACB, TEMPLATE_THONG_THUONG):
            restored = TargetTemplate.from_dict(tpl.to_dict())
            self.assertEqual(restored, tpl)

    def test_validation_rejects_duplicate_columns(self) -> None:
        bad = replace(TEMPLATE_ACB, amount=ColumnRef("C", "Số tiền"))  # trùng cột C với account_no
        with self.assertRaises(ConfigError):
            bad.ensure_valid()

    def test_validation_rejects_missing_expected_header(self) -> None:
        bad = replace(TEMPLATE_ACB, amount=ColumnRef("D", ""))  # thiếu expected_header
        with self.assertRaises(ConfigError):
            bad.ensure_valid()

    def test_validation_rejects_bad_row_order(self) -> None:
        bad = replace(TEMPLATE_ACB, header_rows=(3,), data_start_row=3)
        with self.assertRaises(ConfigError):
            bad.ensure_valid()

    def test_data_end_row_zero_rejected(self) -> None:
        d = TEMPLATE_ACB.to_dict()
        d["data_end_row"] = 0
        with self.assertRaises(ConfigError):
            TargetTemplate.from_dict(d)


class TestGuideAndRules(unittest.TestCase):
    def test_check_rules(self) -> None:
        rules = TargetRules(
            no_diacritics=True,
            no_extra_spaces=True,
            max_text_length=10,
            forbidden_chars="&@#",
        )
        self.assertEqual(check_rules("NGUYEN AN", rules), [])
        v1 = check_rules("NGUYỄN  AN & B", rules)
        self.assertEqual(len(v1), 4)  # có dấu, thừa khoảng trắng, quá dài, chứa &

    def test_parse_guide_extracts_suggestions(self) -> None:
        sheet = SheetData(
            info=SheetInfo("Hướng dẫn nhập liệu", True),
            cells={
                (1, 1): CellData("Không nhập dấu tiếng Việt, không khoảng trắng thừa.", CellKind.TEXT),
                (2, 1): CellData("Nội dung tối đa 210 ký tự, không gồm dấu &", CellKind.TEXT),
            },
            max_row=2,
            max_col=1,
        )
        info = parse_guide(sheet)
        self.assertEqual(info.sheet_name, "Hướng dẫn nhập liệu")
        self.assertEqual(len(info.lines), 2)
        self.assertTrue(info.suggested_rules.no_diacritics)
        self.assertTrue(info.suggested_rules.no_extra_spaces)
        self.assertIsNone(info.suggested_rules.max_text_length)
        self.assertEqual(info.max_length_found, 210)
        self.assertIn("&", info.suggested_rules.forbidden_chars)


class TestReadTargetSynthetic(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_read_acb_xls_happy_path(self) -> None:
        p = build_acb_xls(self.tmp / "4 MauChiHoLuongQuaTaiKhoanACB.xls")
        res = read_target(p, TEMPLATE_ACB)
        self.assertTrue(res.header_ok)
        self.assertTrue(res.is_usable)
        self.assertEqual(len(res.rows), 3)
        self.assertEqual(res.rows[0].account_no, "24349081")
        self.assertTrue(res.rows[0].account_stored_as_number)
        self.assertEqual(res.rows[0].current_amount, 4500000)
        self.assertIsNotNone(res.guide)
        self.assertTrue(any(i.code == IssueCode.GUIDE_SHEET_FOUND for i in res.issues))
        self.assertEqual(res.summary()["current_total_amount"], 12800000)

    def test_read_tt_xls_happy_path_preserves_leading_zero_and_bank_code(self) -> None:
        p = build_tt_xls(self.tmp / "5 MauChiHoLuongThongThuong.xls")
        res = read_target(p, TEMPLATE_THONG_THUONG)
        self.assertTrue(res.header_ok)
        self.assertTrue(res.is_usable)
        self.assertEqual(res.sheet_name, "Dán giá trị vào đây để TT")
        self.assertEqual(len(res.rows), 3)
        self.assertEqual(res.rows[0].account_no, "0587952606")
        self.assertFalse(res.rows[0].account_stored_as_number)
        self.assertEqual(res.rows[0].bank_code, "79203001")

    def test_header_mismatch_blocks(self) -> None:
        p = build_acb_xls(self.tmp / "acb.xls", amount_header="Cột khác")
        res = read_target(p, TEMPLATE_ACB)
        self.assertFalse(res.header_ok)
        self.assertFalse(res.is_usable)
        self.assertTrue(any(i.code == IssueCode.HEADER_MISMATCH for i in res.issues))
        self.assertEqual(res.rows, [])

    def test_auto_selects_sheet_by_headers_when_name_differs(self) -> None:
        p = build_acb_xls(self.tmp / "acb_other_sheet.xls", sheet_name="SheetKhac")
        res = read_target(p, TEMPLATE_ACB)
        self.assertEqual(res.sheet_name, "SheetKhac")
        self.assertTrue(any(i.code == IssueCode.SHEET_AUTO_SELECTED for i in res.issues))

    def test_ambiguous_sheets_raise_selection_error(self) -> None:
        p = self.tmp / "two_sheets.xlsx"
        wb = openpyxl.Workbook()
        ws1 = wb.active
        assert ws1 is not None
        ws1.title = "S1"
        ws2 = wb.create_sheet("S2")
        for ws in (ws1, ws2):
            ws["B2"], ws["C2"], ws["D2"] = "Tên NV", "Số tài khoản", "Số tiền"
            ws["B3"], ws["C3"], ws["D3"] = "NGUYEN AN", 24349081, 1000000
        wb.save(p)
        with self.assertRaises(SheetSelectionError) as ctx:
            read_target(p, TEMPLATE_ACB)
        self.assertEqual(ctx.exception.candidates, ["S1", "S2"])

    def test_duplicate_account_is_blocking_error(self) -> None:
        rows = [
            ("NGUYEN AN", 24349081, 1000000),
            ("LE BINH", 24349081, 2000000),
        ]
        p = build_acb_xls(self.tmp / "dup.xls", rows=rows)
        res = read_target(p, TEMPLATE_ACB)
        self.assertFalse(res.is_usable)
        self.assertTrue(any(i.code == IssueCode.DUPLICATE_ACCOUNT for i in res.issues))

    def test_formula_in_target_amount_marks_row_error(self) -> None:
        p = build_target_xlsx(self.tmp / "formula.xlsx")
        wb = openpyxl.load_workbook(p)
        wb["CNBD"]["D3"] = "=1000+2000"
        wb.save(p)
        res = read_target(p, TEMPLATE_ACB)
        self.assertTrue(res.rows[0].has_error)
        self.assertFalse(res.rows[0].is_writable)
        self.assertTrue(any(i.code == IssueCode.TARGET_AMOUNT_IS_FORMULA for i in res.issues))

    def test_incomplete_row_without_account_warns_not_writable(self) -> None:
        rows = [
            ("NGUYEN AN", 24349081, 1000000),
            ("GHI CHU TONG", None, 1000000),
        ]
        p = build_acb_xls(self.tmp / "inc.xls", rows=rows)
        res = read_target(p, TEMPLATE_ACB)
        self.assertTrue(res.is_usable)
        self.assertEqual(len(res.rows), 2)
        self.assertEqual(res.summary()["writable"], 1)
        self.assertTrue(any(i.code == IssueCode.TARGET_ROW_INCOMPLETE for i in res.issues))

    def test_rule_violation_warns(self) -> None:
        rows = [("NGUYỄN VĂN AN", 24349081, 1000000)]
        p = build_acb_xls(self.tmp / "rule.xls", rows=rows)
        res = read_target(p, TEMPLATE_ACB)
        self.assertTrue(any(i.code == IssueCode.RULE_VIOLATION for i in res.issues))

    def test_missing_bank_code_in_tt_is_row_error(self) -> None:
        rows = [("NGUYEN VAN AN", 6000000, None, "0587952606")]
        p = build_tt_xls(self.tmp / "tt_missing_bank.xls", rows=rows)
        res = read_target(p, TEMPLATE_THONG_THUONG)
        self.assertTrue(any(i.code == IssueCode.BANK_CODE_MISSING for i in res.issues))
        self.assertTrue(res.rows[0].has_error)
        self.assertFalse(res.rows[0].is_writable)

    def test_suspect_column_content_blocks(self) -> None:
        rows = [
            ("111", "abc", 1000),
            ("222", "def", 2000),
            ("333", "ghi", 3000),
        ]
        p = build_acb_xls(self.tmp / "suspect.xls", rows=rows)
        res = read_target(p, TEMPLATE_ACB)
        self.assertFalse(res.is_usable)
        suspect = [i for i in res.issues if i.code == IssueCode.COLUMN_CONTENT_SUSPECT]
        self.assertEqual(len(suspect), 2)

    def test_no_records_is_blocking_error(self) -> None:
        p = build_acb_xls(self.tmp / "empty.xls", rows=[])
        res = read_target(p, TEMPLATE_ACB)
        self.assertFalse(res.is_usable)
        self.assertTrue(any(i.code == IssueCode.NO_RECORDS for i in res.issues))

    def test_excel_error_in_account_column_marks_row_error(self) -> None:
        sheet = SheetData(
            info=SheetInfo("CNBD", True),
            cells={
                (2, 1): CellData("STT", CellKind.TEXT),
                (2, 2): CellData("Tên NV", CellKind.TEXT),
                (2, 3): CellData("Số tài khoản", CellKind.TEXT),
                (2, 4): CellData("Số tiền", CellKind.TEXT),
                (3, 1): CellData(1, CellKind.NUMBER),
                (3, 2): CellData("NGUYEN AN", CellKind.TEXT),
                (3, 3): CellData("#REF!", CellKind.ERROR),
                (3, 4): CellData(1000000, CellKind.NUMBER),
            },
            max_row=3,
            max_col=4,
        )
        res = parse_target_sheet(sheet, TEMPLATE_ACB)
        self.assertTrue(res.rows[0].has_error)
        self.assertTrue(any(i.code == IssueCode.EXCEL_ERROR_VALUE and i.severity == Severity.ERROR for i in res.issues))


@unittest.skipUnless(
    os.environ.get("CHL_REAL_ACB") and os.environ.get("CHL_REAL_TT"),
    "Đặt CHL_REAL_ACB và CHL_REAL_TT để chạy kiểm thử trên 2 file mẫu ngân hàng thật.",
)
class TestReadTargetRealFiles(unittest.TestCase):
    def test_real_acb_and_tt_match_source_total(self) -> None:
        acb_path = Path(os.environ["CHL_REAL_ACB"])
        tt_path = Path(os.environ["CHL_REAL_TT"])
        r_acb = read_target(acb_path, TEMPLATE_ACB)
        r_tt = read_target(tt_path, TEMPLATE_THONG_THUONG)

        self.assertTrue(r_acb.header_ok)
        self.assertTrue(r_acb.is_usable)
        self.assertEqual(len(r_acb.rows), 79)
        self.assertEqual(r_acb.summary()["writable"], 79)

        self.assertTrue(r_tt.header_ok)
        self.assertTrue(r_tt.is_usable)
        self.assertEqual(len(r_tt.rows), 7)
        self.assertEqual(r_tt.summary()["writable"], 7)
        self.assertIn("0587952606", r_tt.by_account())


if __name__ == "__main__":
    unittest.main()
