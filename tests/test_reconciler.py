"""Kiểm thử cho core.reconciler (quy tắc R1 -> R8, ngoại lệ X1 -> X10, WritePlan & ExceptionItem)."""

from __future__ import annotations

import os
import unittest
from pathlib import Path
from typing import Optional

from core.models import (
    CellUpdate,
    ColumnRef,
    ConfigError,
    ExceptionCode,
    Issue,
    IssueCode,
    MasterRecord,
    MatchStatus,
    ResolutionAction,
    Severity,
    SkippedRow,
    SourceReadResult,
    SourceRecord,
    TargetReadResult,
    TargetRow,
    TargetTemplate,
)
from core.normalize import normalize_name_key
from core.reconciler import (
    ReconcileError,
    TargetInput,
    apply_resolutions,
    reconcile,
    resolve_exception,
)
from core.source_reader import read_source
from core.target_reader import TEMPLATE_ACB, TEMPLATE_THONG_THUONG, read_target


def _src_rec(
    row: int,
    card_no: str,
    full_name: str,
    account_no: Optional[str],
    amount: Optional[int],
    *,
    has_error: bool = False,
) -> SourceRecord:
    return SourceRecord(
        row=row,
        card_no=card_no,
        full_name=full_name,
        name_key=normalize_name_key(full_name),
        account_no=account_no,
        amount=amount,
        has_error=has_error,
    )


def _tgt_row(
    row: int,
    account_no: Optional[str],
    full_name: str,
    current_amount: Optional[int] = 1_000_000,
    *,
    bank_code: Optional[str] = None,
    account_stored_as_number: bool = False,
    has_error: bool = False,
) -> TargetRow:
    return TargetRow(
        row=row,
        account_no=account_no,
        full_name=full_name,
        name_key=normalize_name_key(full_name),
        bank_code=bank_code,
        current_amount=current_amount,
        account_stored_as_number=account_stored_as_number,
        has_error=has_error,
    )


def _make_source(
    records: list[SourceRecord],
    *,
    skipped: Optional[list[SkippedRow]] = None,
    issues: Optional[list[Issue]] = None,
) -> SourceReadResult:
    return SourceReadResult(
        file_path="BANG_LUONG.xlsx",
        sheet_name="HLCK06.26 (2)",
        profile_name="Bảng lương",
        records=records,
        skipped=skipped or [],
        issues=issues or [],
    )


def _make_target(
    rows: list[TargetRow],
    *,
    file_path: str = "ACB.xls",
    sheet_name: str = "CNBD",
    template_name: str = "ACB - Chi hộ lương qua tài khoản ACB",
    issues: Optional[list[Issue]] = None,
) -> TargetReadResult:
    return TargetReadResult(
        file_path=file_path,
        file_format="xls",
        sheet_name=sheet_name,
        template_name=template_name,
        rows=rows,
        issues=issues or [],
    )


class TestReconcilerCoreRules(unittest.TestCase):
    def test_r1_r2_r3_matches_strictly_by_account_never_by_name(self) -> None:
        """2 nhân sự trùng hoàn toàn Họ và tên nhưng khác Số thẻ & Số TK -> khớp đúng từng dòng."""
        src = _make_source(
            [
                _src_rec(5, "0001", "NGUYEN VAN AN", "24340001", 5_500_000),
                _src_rec(6, "0002", "NGUYEN VAN AN", "24340002", 6_800_000),
            ],
            skipped=[SkippedRow(4, "Dòng tên tổ (không có Số thẻ)", {})],
        )
        tgt = _make_target([
            _tgt_row(3, "24340002", "NGUYEN VAN AN", 4_000_000),
            _tgt_row(4, "24340001", "NGUYEN VAN AN", 5_500_000),
        ])

        plan = reconcile(src, (tgt, TEMPLATE_ACB))
        self.assertTrue(plan.is_ready)
        self.assertEqual(len(plan.exceptions), 0)
        self.assertEqual(len(plan.updates), 2)

        by_row = {u.row: u for u in plan.updates}
        # Dòng 3 trong file NH là TK 24340002 (Số thẻ 0002) -> 6,800,000 (WILL_UPDATE)
        self.assertEqual(by_row[3].card_no, "0002")
        self.assertEqual(by_row[3].account_no, "24340002")
        self.assertEqual(by_row[3].cell, "D3")
        self.assertEqual(by_row[3].old_value, 4_000_000)
        self.assertEqual(by_row[3].new_value, 6_800_000)
        self.assertTrue(by_row[3].is_changed)

        # Dòng 4 trong file NH là TK 24340001 (Số thẻ 0001) -> 5,500,000 (UNCHANGED)
        self.assertEqual(by_row[4].card_no, "0001")
        self.assertEqual(by_row[4].new_value, 5_500_000)
        self.assertFalse(by_row[4].is_changed)
        self.assertEqual(len(plan.changed_updates), 1)
        self.assertEqual(plan.summary()["will_update"], 1)
        self.assertEqual(plan.summary()["unchanged"], 1)

    def test_never_matches_by_name_when_account_differs(self) -> None:
        """Cùng tên 'NGUYEN VAN AN' nhưng khác Số TK -> tuyệt đối không tự ghép theo tên (sinh X4 + X5)."""
        src = _make_source([_src_rec(5, "0001", "NGUYEN VAN AN", "11111111", 5_000_000)])
        tgt = _make_target([_tgt_row(3, "99999999", "NGUYEN VAN AN", 4_000_000)])

        plan = reconcile(src, (tgt, TEMPLATE_ACB))
        codes = [e.code for e in plan.exceptions]
        self.assertIn(ExceptionCode.X4_NOT_IN_TARGET, codes)
        self.assertIn(ExceptionCode.X5_TARGET_NOT_IN_SOURCE, codes)
        self.assertEqual(plan.updates, [])
        self.assertFalse(plan.is_ready)  # X5 bắt buộc người dùng chọn xử lý

    def test_r4_r6_preserves_leading_zero_and_only_writes_amount_column(self) -> None:
        src = _make_source([_src_rec(5, "0007", "LE THI BINH", "0587952606", 7_200_000)])
        tgt = _make_target(
            [_tgt_row(3, "0587952606", "LE THI BINH", 6_000_000, bank_code="79203001")],
            file_path="ThongThuong.xls",
            sheet_name="Dán giá trị vào đây để TT",
            template_name=TEMPLATE_THONG_THUONG.name,
        )
        plan = reconcile(src, (tgt, TEMPLATE_THONG_THUONG))
        self.assertTrue(plan.is_ready)
        self.assertEqual(len(plan.updates), 1)
        u = plan.updates[0]
        self.assertEqual(u.cell, "E3")
        self.assertEqual(u.column, "E")
        self.assertEqual(u.account_no, "0587952606")
        self.assertEqual(u.card_no, "0007")
        self.assertEqual(u.new_value, 7_200_000)


class TestReconcilerExceptionsV1(unittest.TestCase):
    def test_x3_missing_account_is_report_only_skipped(self) -> None:
        src = _make_source(
            [_src_rec(138, "26032301", "LUONG VAN HAI", None, 4_988_450)],
            issues=[
                Issue(
                    IssueCode.ACCOUNT_MISSING,
                    Severity.WARNING,
                    "Dòng có Số thẻ '26032301' nhưng thiếu Số tài khoản.",
                    row=138,
                )
            ],
        )
        tgt = _make_target([_tgt_row(3, "24349081", "NGUYEN AN", 1_000_000)])
        plan = reconcile(src, tgt)

        x3_list = [e for e in plan.exceptions if e.code == ExceptionCode.X3_MISSING_ACCOUNT]
        self.assertEqual(len(x3_list), 1)
        self.assertTrue(x3_list[0].is_resolved)
        self.assertEqual(x3_list[0].action, ResolutionAction.SKIP)
        self.assertEqual(plan.matches[0].status, MatchStatus.SKIPPED)

    def test_x4_not_in_target_and_leading_zero_hint(self) -> None:
        """Số TK nguồn có số 0 đầu ('0587952606'), đích lưu số mất số 0 ('587952606') -> báo X4 kèm cảnh báo, không tự ghép."""
        src = _make_source([_src_rec(5, "0001", "NGUYEN AN", "0587952606", 5_000_000)])
        tgt = _make_target([
            _tgt_row(3, "587952606", "NGUYEN AN", 4_000_000, account_stored_as_number=True)
        ])
        plan = reconcile(src, tgt)
        x4 = next(e for e in plan.exceptions if e.code == ExceptionCode.X4_NOT_IN_TARGET)
        self.assertIn("mất số 0 ở đầu", x4.message)
        self.assertEqual(x4.action, ResolutionAction.SKIP)
        self.assertEqual(plan.updates, [])

    def test_x5_target_not_in_source_requires_choice_keep_zero_or_clear(self) -> None:
        src = _make_source([_src_rec(5, "0001", "NGUYEN AN", "24340001", 5_000_000)])
        tgt = _make_target([
            _tgt_row(3, "24340001", "NGUYEN AN", 5_000_000),
            _tgt_row(4, "24349999", "NHAN VIEN NGHI VIEC", 4_500_000),
        ])
        plan = reconcile(src, tgt)
        self.assertFalse(plan.is_ready)
        x5 = next(e for e in plan.exceptions if e.code == ExceptionCode.X5_TARGET_NOT_IN_SOURCE)
        self.assertFalse(x5.is_resolved)

        # 1. Chọn KEEP -> giữ nguyên, không tạo CellUpdate cho dòng 4
        plan.resolve(x5.id, ResolutionAction.KEEP)
        self.assertTrue(plan.is_ready)
        self.assertEqual([u.row for u in plan.updates], [3])

        # 2. Chọn SET_ZERO -> tạo CellUpdate new_value=0 cho dòng 4
        plan.resolve(x5.id, ResolutionAction.SET_ZERO)
        self.assertTrue(plan.is_ready)
        by_row = {u.row: u for u in plan.updates}
        self.assertEqual(by_row[4].new_value, 0)
        self.assertEqual(by_row[4].reason, "X5_SET_ZERO")

        # 3. Chọn CLEAR -> tạo CellUpdate new_value=None cho dòng 4
        plan.resolve(x5.id, ResolutionAction.CLEAR)
        self.assertTrue(plan.is_ready)
        by_row = {u.row: u for u in plan.updates}
        self.assertIsNone(by_row[4].new_value)
        self.assertEqual(by_row[4].reason, "X5_CLEAR")

    def test_x6_name_mismatch_requires_confirmation_or_skip(self) -> None:
        src = _make_source([_src_rec(5, "0001", "NGUYEN VAN AN", "24340001", 6_000_000)])
        tgt = _make_target([_tgt_row(3, "24340001", "TRAN THI KHAC", 5_000_000)])

        plan = reconcile(src, tgt)
        self.assertFalse(plan.is_ready)
        self.assertEqual(plan.updates, [])
        x6 = next(e for e in plan.exceptions if e.code == ExceptionCode.X6_TARGET_MISMATCH)
        self.assertFalse(x6.is_resolved)
        self.assertEqual(plan.matches[0].status, MatchStatus.PENDING)

        # Người dùng chọn SKIP -> không ghi
        plan.resolve(x6.id, ResolutionAction.SKIP)
        self.assertTrue(plan.is_ready)
        self.assertEqual(plan.updates, [])
        self.assertEqual(plan.matches[0].status, MatchStatus.SKIPPED)

        # Người dùng chọn CONFIRM_WRITE -> ghi số tiền mới
        plan.resolve(x6.id, ResolutionAction.CONFIRM_WRITE)
        self.assertTrue(plan.is_ready)
        self.assertEqual(len(plan.updates), 1)
        self.assertEqual(plan.updates[0].new_value, 6_000_000)
        self.assertEqual(plan.matches[0].status, MatchStatus.WILL_UPDATE)

    def test_x7_invalid_amount_skipped_by_default_and_supports_override(self) -> None:
        src = _make_source(
            [_src_rec(5, "0001", "NGUYEN AN", "24340001", None, has_error=True)],
            issues=[
                Issue(
                    IssueCode.AMOUNT_NOT_INTEGER,
                    Severity.ERROR,
                    "Số tiền 5000000.5 có phần thập phân.",
                    row=5,
                )
            ],
        )
        tgt = _make_target([_tgt_row(3, "24340001", "NGUYEN AN", 4_000_000)])
        plan = reconcile(src, tgt)

        # Mặc định X7 được SKIP (không ghi dòng lỗi số tiền) và không sinh X5 giả cho dòng 3
        self.assertEqual(len(plan.exceptions), 1)
        x7 = plan.exceptions[0]
        self.assertEqual(x7.code, ExceptionCode.X7_INVALID_AMOUNT)
        self.assertEqual(x7.action, ResolutionAction.SKIP)
        self.assertEqual(plan.updates, [])

        # Sửa tay số tiền không hợp lệ (<= 0) phải bị từ chối
        with self.assertRaises(ReconcileError):
            resolve_exception(plan, x7.id, ResolutionAction.OVERRIDE_AMOUNT, override_amount=0)

        # Sửa tay số tiền nguyên > 0 hợp lệ -> được đưa vào WritePlan
        resolve_exception(plan, x7.id, ResolutionAction.OVERRIDE_AMOUNT, override_amount=5_000_001)
        self.assertTrue(plan.is_ready)
        self.assertEqual(len(plan.updates), 1)
        self.assertEqual(plan.updates[0].new_value, 5_000_001)
        self.assertEqual(plan.updates[0].reason, "OVERRIDE_AMOUNT")

    def test_x8_duplicate_card_or_account_hard_blocks(self) -> None:
        # 1. Trùng Số thẻ trong nguồn
        src_dup_card = _make_source([
            _src_rec(5, "0001", "NGUYEN A", "24340001", 5_000_000),
            _src_rec(6, "0001", "NGUYEN B", "24340002", 6_000_000),
        ])
        tgt = _make_target([
            _tgt_row(3, "24340001", "NGUYEN A", 5_000_000),
            _tgt_row(4, "24340002", "NGUYEN B", 6_000_000),
        ])
        plan1 = reconcile(src_dup_card, tgt)
        self.assertFalse(plan1.is_ready)
        self.assertTrue(plan1.has_blocking_errors)
        self.assertEqual(plan1.updates, [])
        x8 = plan1.exceptions[0]
        self.assertEqual(x8.code, ExceptionCode.X8_DUPLICATE_KEY)
        self.assertTrue(x8.is_blocking)
        with self.assertRaises(ReconcileError):
            plan1.resolve(x8.id, ResolutionAction.SKIP)

        # 2. Trùng Số TK giữa 2 file ngân hàng khác nhau trong cùng 1 lần chạy
        src_ok = _make_source([_src_rec(5, "0001", "NGUYEN A", "24340001", 5_000_000)])
        tgt_1 = _make_target([_tgt_row(3, "24340001", "NGUYEN A", 5_000_000)], file_path="ACB_1.xls")
        tgt_2 = _make_target([_tgt_row(7, "24340001", "NGUYEN A", 5_000_000)], file_path="ACB_2.xls")
        plan2 = reconcile(src_ok, [tgt_1, tgt_2])
        self.assertFalse(plan2.is_ready)
        self.assertTrue(plan2.has_blocking_errors)
        self.assertEqual(plan2.updates, [])
        self.assertTrue(any(e.code == ExceptionCode.X8_DUPLICATE_KEY for e in plan2.exceptions))

    def test_x10_target_formula_error_never_allows_overwrite(self) -> None:
        src = _make_source([_src_rec(5, "0001", "NGUYEN A", "24340001", 5_000_000)])
        tgt = _make_target(
            [_tgt_row(3, "24340001", "NGUYEN A", None, has_error=True)],
            issues=[
                Issue(
                    IssueCode.TARGET_AMOUNT_IS_FORMULA,
                    Severity.ERROR,
                    "Ô Số tiền đang chứa công thức; tool không ghi đè công thức.",
                    row=3,
                    column="D",
                )
            ],
        )
        plan = reconcile(src, tgt)
        self.assertFalse(plan.is_ready)
        x10 = next(e for e in plan.exceptions if e.code == ExceptionCode.X10_RULE_VIOLATION)
        self.assertEqual(x10.allowed_actions, (ResolutionAction.SKIP,))
        with self.assertRaises(ReconcileError):
            plan.resolve(x10.id, ResolutionAction.CONFIRM_WRITE)
        plan.resolve(x10.id, ResolutionAction.SKIP)
        self.assertTrue(plan.is_ready)
        self.assertEqual(plan.updates, [])

    def test_x10_rule_violation_warning_allows_confirm_or_skip(self) -> None:
        src = _make_source([_src_rec(5, "0001", "NGUYEN VAN AN", "24340001", 5_000_000)])
        tgt = _make_target(
            [_tgt_row(3, "24340001", "NGUYỄN VĂN AN", 4_000_000)],
            issues=[
                Issue(
                    IssueCode.RULE_VIOLATION,
                    Severity.WARNING,
                    "Tên 'NGUYỄN VĂN AN' vi phạm hướng dẫn nhập liệu: có dấu tiếng Việt.",
                    row=3,
                    column="B",
                )
            ],
        )
        plan = reconcile(src, tgt)
        self.assertFalse(plan.is_ready)
        x10 = next(e for e in plan.exceptions if e.code == ExceptionCode.X10_RULE_VIOLATION)
        plan.resolve(x10.id, ResolutionAction.CONFIRM_WRITE)
        self.assertTrue(plan.is_ready)
        self.assertEqual(len(plan.updates), 1)
        self.assertEqual(plan.updates[0].new_value, 5_000_000)

    def test_v2_master_dataset_hook(self) -> None:
        src = _make_source([
            _src_rec(5, "0001", "NGUYEN AN", "99999999", 5_000_000),  # lệch TK so với Master
            _src_rec(6, "0002", "LE BINH", "24340002", 6_000_000),    # chưa có trong Master
        ])
        tgt = _make_target([
            _tgt_row(3, "24340001", "NGUYEN AN", 4_000_000),
            _tgt_row(4, "24340002", "LE BINH", 6_000_000),
        ])
        master = {
            "0001": MasterRecord(
                card_no="0001",
                full_name="NGUYEN AN",
                account_no="24340001",
                target_type=TEMPLATE_ACB.name,
            )
        }
        plan = reconcile(src, tgt, master=master)
        codes = {e.code for e in plan.exceptions}
        self.assertIn(ExceptionCode.X1_NOT_IN_MASTER, codes)
        self.assertIn(ExceptionCode.X2_ACCOUNT_MISMATCH_MASTER, codes)
        self.assertFalse(plan.is_ready)

        # Xác nhận cả X1 và X2 -> dùng TK chuẩn của Master cho 0001 và ghi cả 2 dòng
        resolutions = {e.id: ResolutionAction.CONFIRM_WRITE for e in plan.exceptions}
        apply_resolutions(plan, resolutions)
        self.assertTrue(plan.is_ready)
        self.assertEqual(len(plan.updates), 2)

    def test_cell_update_and_target_input_validation(self) -> None:
        with self.assertRaises(ConfigError):
            CellUpdate("a.xls", "S", 3, "D", 100, -1)
        bad_tpl = TargetTemplate(
            name="Bad",
            account_no=ColumnRef("C", "Số TK"),
            full_name=ColumnRef("B", "Tên"),
            amount=ColumnRef("C", "Số tiền"),
            sheet_name="CNBD",
        )
        tgt = _make_target([])
        with self.assertRaises(ConfigError):
            TargetInput(tgt, bad_tpl).resolved_amount_column()


@unittest.skipUnless(
    os.environ.get("CHL_REAL_SOURCE")
    and os.environ.get("CHL_REAL_ACB")
    and os.environ.get("CHL_REAL_TT"),
    "Đặt CHL_REAL_SOURCE, CHL_REAL_ACB, CHL_REAL_TT để chạy đối chiếu trên bộ file tháng 05/2026 thật.",
)
class TestReconcileRealMay2026Files(unittest.TestCase):
    def test_reconcile_real_source_with_acb_and_tt(self) -> None:
        src = read_source(Path(os.environ["CHL_REAL_SOURCE"]), sheet_name="HLCK05.26 (2)")
        r_acb = read_target(Path(os.environ["CHL_REAL_ACB"]), TEMPLATE_ACB)
        r_tt = read_target(Path(os.environ["CHL_REAL_TT"]), TEMPLATE_THONG_THUONG)

        plan = reconcile(
            src,
            [
                (r_acb, TEMPLATE_ACB),
                (r_tt, TEMPLATE_THONG_THUONG),
            ],
        )

        # 86 nhân sự có TK khớp 100% (79 ACB + 7 Thông thường), 4 nhân sự cuối bảng thiếu TK -> X3
        self.assertTrue(plan.is_ready)
        self.assertFalse(plan.has_blocking_errors)
        self.assertEqual(len(plan.updates), 86)
        self.assertEqual(plan.total_new_amount, 445_138_601)

        acb_plan = plan.for_file(r_acb.file_path)
        tt_plan = plan.for_file(r_tt.file_path)
        assert acb_plan is not None and tt_plan is not None
        self.assertEqual(len(acb_plan.updates), 79)
        self.assertEqual(acb_plan.total_new_amount, 405_476_927)
        self.assertTrue(all(u.column == "D" for u in acb_plan.updates))

        self.assertEqual(len(tt_plan.updates), 7)
        self.assertEqual(tt_plan.total_new_amount, 39_661_674)
        self.assertTrue(all(u.column == "E" for u in tt_plan.updates))

        # Chỉ có đúng 4 Exception X3 (4 dòng có Số thẻ nhưng không có Số TK ở cuối bảng lương)
        self.assertEqual(len(plan.exceptions), 4)
        self.assertTrue(all(e.code == ExceptionCode.X3_MISSING_ACCOUNT for e in plan.exceptions))
        self.assertEqual(plan.unresolved_exceptions, [])


if __name__ == "__main__":
    unittest.main()
