import hashlib
import importlib.util
import os
import shutil
import tempfile
import unittest
from dataclasses import replace
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

import openpyxl

from core.models import ColumnRef, ConfigError, IssueCode, Severity, SourceProfile, SourceReadResult
from core.source_reader import DEFAULT_SOURCE_PROFILE, SheetSelectionError, read_source, select_sheet
from tests.helpers import build_source_xls, build_source_xlsx

HAS_XLWT = importlib.util.find_spec("xlwt") is not None


def codes(result: SourceReadResult) -> list[IssueCode]:
    return [i.code for i in result.issues]


class Base(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="chl_src_"))

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)


class TestHappyPath(Base):
    def test_read_default_structure(self) -> None:
        p = build_source_xlsx(self.tmp / "luong.xlsx", extra_hidden_sheet=True)
        res = read_source(p)

        self.assertTrue(res.is_usable, [str(i) for i in res.issues])
        self.assertEqual(res.sheet_name, "HLCK06.26 (2)")
        self.assertEqual([r.card_no for r in res.records], ["16021804", "22050301", "26032301", "4/7/2012"])

        first = res.records[0]
        self.assertEqual(first.account_no, "0587952606")      # giữ số 0 đầu
        self.assertEqual(first.amount, 5615871)
        self.assertEqual(first.full_name, "NGUYEN THANH CONG")
        self.assertEqual(res.records[1].account_no, "24349081")  # số -> chuỗi
        self.assertEqual(res.records[1].amount, 5917028)          # 5917028.0 -> int

        # Người thiếu số TK: vẫn là record, có cảnh báo, KHÔNG đủ điều kiện chuyển khoản
        no_acc = res.records[2]
        self.assertIsNone(no_acc.account_no)
        self.assertFalse(no_acc.is_transferable)
        self.assertIn(IssueCode.ACCOUNT_MISSING, codes(res))

        self.assertEqual(len(res.transferable), 3)
        self.assertEqual(res.total_transferable_amount, 5615871 + 5917028 + 2439377)

        # Dòng tổng cộng (không có Số thẻ nhưng có số tiền) được ghi nhận là bỏ qua
        self.assertEqual([s.row for s in res.skipped], [8])
        # Dòng tên tổ không có dữ liệu ở cột mapping -> không cần ghi nhận
        self.assertNotIn(5, [s.row for s in res.skipped])

    @unittest.skipUnless(HAS_XLWT, "cần xlwt để tạo file .xls")
    def test_read_xls(self) -> None:
        p = build_source_xls(self.tmp / "luong.xls")
        res = read_source(p)
        self.assertTrue(res.is_usable, [str(i) for i in res.issues])
        self.assertEqual(len(res.records), 4)
        self.assertEqual(res.records[0].account_no, "0587952606")


class TestSheetSelection(Base):
    def test_ambiguous_requires_user_choice(self) -> None:
        p = build_source_xlsx(self.tmp / "luong.xlsx", sheet_names=("HLCK06.26 (2)", "HLCK06.26"))
        with self.assertRaises(SheetSelectionError) as ctx:
            read_source(p)
        self.assertEqual(sorted(ctx.exception.candidates), ["HLCK06.26", "HLCK06.26 (2)"])
        res = read_source(p, sheet_name="HLCK06.26 (2)")
        self.assertEqual(res.sheet_name, "HLCK06.26 (2)")

    def test_no_match(self) -> None:
        p = build_source_xlsx(self.tmp / "luong.xlsx", sheet_names=("Sheet1",))
        with self.assertRaises(SheetSelectionError) as ctx:
            read_source(p)
        self.assertEqual(ctx.exception.candidates, ["Sheet1"])

    def test_select_sheet_rules(self) -> None:
        names = ["05.2026", "HLCK05.26 (2)", "HLCK05.26", "TM05.26 "]
        prof = replace(DEFAULT_SOURCE_PROFILE, sheet_name="tm05.26", sheet_patterns=())
        self.assertEqual(select_sheet(names, prof), "TM05.26 ")      # khớp lỏng duy nhất
        prof2 = replace(DEFAULT_SOURCE_PROFILE, sheet_patterns=("HLCK* (2)",))
        self.assertEqual(select_sheet(names, prof2), "HLCK05.26 (2)")
        with self.assertRaises(SheetSelectionError):
            select_sheet(names, DEFAULT_SOURCE_PROFILE, override="Không có")


class TestSafetyChecks(Base):
    def test_header_mismatch_stops_reading(self) -> None:
        p = build_source_xlsx(self.tmp / "luong.xlsx", insert_col_before_p=True)
        res = read_source(p)
        self.assertFalse(res.header_ok)
        self.assertFalse(res.is_usable)
        self.assertEqual(res.records, [])
        msg = res.blocking_issues[0].message
        self.assertIn("cột Q", msg)

    def test_duplicates(self) -> None:
        rows = [
            ("A1", None, "NGUYEN VAN A", "111111", 100),
            ("A1", None, "NGUYEN VAN B", "222222", 200),    # trùng Số thẻ
            ("A3", None, "NGUYEN VAN A", "333333", 300),    # trùng tên -> chỉ INFO
            ("A4", None, "TRAN VAN D", "333333", 400),      # trùng TK
        ]
        res = read_source(build_source_xlsx(self.tmp / "dup.xlsx", rows=rows))
        blocking = [i.code for i in res.blocking_issues]
        self.assertIn(IssueCode.DUPLICATE_CARD, blocking)
        self.assertIn(IssueCode.DUPLICATE_ACCOUNT, blocking)
        dup_name = [i for i in res.issues if i.code == IssueCode.DUPLICATE_NAME]
        self.assertEqual(len(dup_name), 1)
        self.assertEqual(dup_name[0].severity, Severity.INFO)

    def test_name_duplicate_alone_is_not_blocking(self) -> None:
        rows = [("A1", None, "NGUYEN VAN A", "111111", 100), ("A2", None, "NGUYEN VAN A", "222222", 200)]
        res = read_source(build_source_xlsx(self.tmp / "n.xlsx", rows=rows))
        self.assertTrue(res.is_usable)

    def test_bad_amounts_mark_row_error(self) -> None:
        good = [(f"G{i}", None, f"NGUOI {i}", f"99999{i}0", 1000 + i) for i in range(6)]
        rows = good + [
            ("A1", None, "A", "111111", 5917028.11),   # lẻ
            ("A2", None, "B", "222222", "5.917.028"),  # text
            ("A3", None, "C", "333333", -5),           # âm
            ("A4", None, "D", "444444", None),         # trống
            ("A5", None, "E", "555555", "#N/A"),       # lỗi Excel
        ]
        res = read_source(build_source_xlsx(self.tmp / "amt.xlsx", rows=rows))
        self.assertTrue(res.is_usable, [str(i) for i in res.blocking_issues])  # lỗi mức dòng không chặn cả file
        bad = [r for r in res.records if r.card_no.startswith("A")]
        self.assertTrue(all(r.has_error and r.amount is None for r in bad))
        c = codes(res)
        for code in (IssueCode.AMOUNT_NOT_INTEGER, IssueCode.AMOUNT_NOT_NUMBER,
                     IssueCode.AMOUNT_NEGATIVE, IssueCode.AMOUNT_MISSING, IssueCode.EXCEL_ERROR_VALUE):
            self.assertIn(code, c)
        self.assertEqual(len(res.transferable), 6)

    def test_zero_amount_not_transferable(self) -> None:
        rows = [("A1", None, "A", "111111", 0), ("A2", None, "B", "222222", 100)]
        res = read_source(build_source_xlsx(self.tmp / "zero.xlsx", rows=rows))
        self.assertFalse(res.records[0].is_transferable)
        self.assertIn(IssueCode.AMOUNT_ZERO, codes(res))
        self.assertEqual(res.summary()["zero_amount"], 1)
        self.assertEqual(res.total_transferable_amount, 100)

    def test_inserted_column_before_unlabeled_columns_is_blocked(self) -> None:
        """Chèn cột giữa P và W -> W/X (không có tiêu đề) bị lệch -> phải chặn, không đọc sai lặng lẽ."""
        p = build_source_xlsx(self.tmp / "shift.xlsx")
        wb = openpyxl.load_workbook(p)
        wb.active.insert_cols(18)
        wb.save(p)
        res = read_source(p)
        self.assertFalse(res.is_usable)
        self.assertIn(IssueCode.COLUMN_CONTENT_SUSPECT, [i.code for i in res.blocking_issues])

    def test_account_column_pointing_to_small_numbers_is_blocked(self) -> None:
        rows = [(f"C{i}", None, f"NGUOI {i}", 500 + i, 1000) for i in range(5)]  # "số TK" chỉ 3 chữ số
        res = read_source(build_source_xlsx(self.tmp / "acc.xlsx", rows=rows))
        suspect = [i for i in res.blocking_issues if i.code == IssueCode.COLUMN_CONTENT_SUSPECT]
        self.assertEqual([i.column for i in suspect], ["X"])

    def test_card_cell_error_and_name_cell_error(self) -> None:
        rows = [("#REF!", None, "A", "111111", 100), ("A2", None, "#N/A", "222222", 200)]
        res = read_source(build_source_xlsx(self.tmp / "err.xlsx", rows=rows))
        self.assertEqual([s.row for s in res.skipped], [5])
        self.assertEqual(res.skipped[0].reason, "Ô Số thẻ bị lỗi")
        self.assertEqual(len(res.records), 1)
        self.assertTrue(res.records[0].has_error)          # ô tên lỗi Excel -> dòng lỗi
        self.assertFalse(res.records[0].is_transferable)

    def test_data_end_row(self) -> None:
        res = read_source(build_source_xlsx(self.tmp / "end.xlsx"), replace(DEFAULT_SOURCE_PROFILE, data_end_row=7))
        self.assertEqual([r.row for r in res.records], [6, 7])

    def test_reader_does_not_modify_file(self) -> None:
        p = build_source_xlsx(self.tmp / "ro.xlsx")
        before = hashlib.sha256(p.read_bytes()).hexdigest()
        read_source(p)
        self.assertEqual(hashlib.sha256(p.read_bytes()).hexdigest(), before)

    def test_formula_not_calculated(self) -> None:
        p = build_source_xlsx(self.tmp / "f.xlsx", rows=[("A1", None, "A", "111111", "=1+1")])
        res = read_source(p)
        self.assertIn(IssueCode.FORMULA_NOT_CALCULATED, codes(res))
        self.assertTrue(res.records[0].has_error)

    def test_card_number_and_date(self) -> None:
        p = self.tmp / "card.xlsx"
        build_source_xlsx(p, rows=[(2101003, None, "A", "111111", 100), ("X", None, "B", "222222", 200)])
        wb = openpyxl.load_workbook(p)
        wb["HLCK06.26 (2)"]["B6"] = datetime(2012, 4, 7)   # Excel tự đổi Số thẻ thành ngày
        wb.save(p)
        res = read_source(p)
        self.assertEqual(res.records[0].card_no, "2101003")
        self.assertIn(IssueCode.CARD_STORED_AS_NUMBER, codes(res))
        self.assertIn(IssueCode.CARD_INVALID, codes(res))
        self.assertEqual([s.row for s in res.skipped], [6])
        self.assertEqual(len(res.records), 1)

    def test_legacy_font_warning(self) -> None:
        prof = replace(DEFAULT_SOURCE_PROFILE, full_name=ColumnRef("C", "Họ Và Tên", "Họ tên"))
        res = read_source(build_source_xlsx(self.tmp / "vni.xlsx"), prof)
        legacy = [i for i in res.issues if i.code == IssueCode.NAME_LEGACY_FONT]
        self.assertEqual(len(legacy), 1)  # gộp 1 cảnh báo cho cả file

    def test_no_records(self) -> None:
        res = read_source(build_source_xlsx(self.tmp / "empty.xlsx", rows=[(None, "Tổng", None, None, 0)]))
        self.assertIn(IssueCode.NO_RECORDS, [i.code for i in res.blocking_issues])


class TestProfile(unittest.TestCase):
    def test_same_column_twice_rejected(self) -> None:
        bad = replace(DEFAULT_SOURCE_PROFILE, account_no=ColumnRef("P"))
        self.assertTrue(any("Cột P" in p for p in bad.validate()))
        with self.assertRaises(ConfigError):
            bad.ensure_valid()

    def test_data_row_must_be_below_header(self) -> None:
        bad = replace(DEFAULT_SOURCE_PROFILE, data_start_row=4)
        with self.assertRaises(ConfigError):
            bad.ensure_valid()

    def test_invalid_letter(self) -> None:
        with self.assertRaises(ConfigError):
            ColumnRef("P1")

    def test_dict_roundtrip(self) -> None:
        data = DEFAULT_SOURCE_PROFILE.to_dict()
        self.assertEqual(SourceProfile.from_dict(data), DEFAULT_SOURCE_PROFILE)
        del data["columns"]["amount"]
        with self.assertRaises(ConfigError):
            SourceProfile.from_dict(data)

    def test_from_dict_coerces_types(self) -> None:
        data = DEFAULT_SOURCE_PROFILE.to_dict()
        data["sheet_name"] = 123
        data["columns"]["card_no"]["expected_header"] = 5
        data["columns"]["card_no"]["letter"] = "b"
        data["header_rows"] = ["3", "4"]
        prof = SourceProfile.from_dict(data)
        self.assertEqual(prof.sheet_name, "123")
        self.assertEqual(prof.card_no.expected_header, "5")
        self.assertEqual(prof.card_no.letter, "B")
        self.assertEqual(prof.header_rows, (3, 4))

    def test_from_dict_rejects_bad_values(self) -> None:
        mutations: list[Callable[[dict[str, Any]], object]] = [
            lambda d: d.update(data_end_row=0),         # 0 không được hiểu là "không giới hạn"
            lambda d: d.update(header_rows=["x"]),
            lambda d: d.update(columns=[]),
            lambda d: d["columns"]["amount"].update(letter="P1"),
            lambda d: d.update(sheet_name=None, sheet_patterns=[]),
        ]
        for mutate in mutations:
            data = DEFAULT_SOURCE_PROFILE.to_dict()
            mutate(data)
            with self.assertRaises(ConfigError):
                SourceProfile.from_dict(data)

    def test_validate_mixed_header_types_no_crash(self) -> None:
        bad = replace(DEFAULT_SOURCE_PROFILE, header_rows=("x", 3))  # type: ignore[arg-type]
        problems = bad.validate()
        self.assertTrue(any("Dòng tiêu đề không hợp lệ" in p for p in problems))

    def test_column_ref_display(self) -> None:
        self.assertEqual(ColumnRef("x", None, "Số tài khoản").display(), "Số tài khoản (cột X)")
        self.assertEqual(ColumnRef("Q").display(), "cột Q")


# -----------------------------------------------------------------------------
# Kiểm thử với file lương THẬT (không đưa file thật vào dự án vì là dữ liệu cá nhân).
# Chạy:  set CHL_REAL_SOURCE=C:\duong\dan\BANG LUONG THANG 05.2026.xlsx
# -----------------------------------------------------------------------------
REAL_SOURCE = os.environ.get("CHL_REAL_SOURCE", "")


@unittest.skipUnless(REAL_SOURCE and Path(REAL_SOURCE).exists(), "chưa đặt biến CHL_REAL_SOURCE")
class TestRealFile(unittest.TestCase):
    def test_may_2026(self) -> None:
        res = read_source(REAL_SOURCE, sheet_name="HLCK05.26 (2)")
        self.assertTrue(res.is_usable, [str(i) for i in res.blocking_issues])
        self.assertEqual(len(res.records), 90)
        self.assertEqual(len(res.transferable), 86)
        self.assertEqual(sum(1 for r in res.records if r.account_no is None), 4)
        self.assertEqual(res.total_transferable_amount, 445_138_601)  # khớp dòng Tổng cộng P125
        acc = {r.card_no: r.account_no for r in res.records}
        self.assertEqual(acc["16021804"], "0587952606")


if __name__ == "__main__":
    unittest.main()
