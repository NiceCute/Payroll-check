"""Kiểm thử cho core.safety (Safety Layer trước khi ghi file)."""

from __future__ import annotations

import os
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path
from typing import Optional

from core.models import (
    CellUpdate,
    Issue,
    IssueCode,
    MatchItem,
    MatchStatus,
    ResolutionAction,
    RunResult,
    SafetyCode,
    SafetyError,
    Severity,
    SourceReadResult,
    SourceRecord,
    TargetReadResult,
    TargetRow,
    TargetWritePlan,
    WritePlan,
)
from core.normalize import normalize_name_key
from core.reconciler import reconcile
from core.safety import (
    check_file_locks,
    ensure_safe_to_write,
    is_file_locked,
    validate_safety,
    validate_write_plan,
)
from core.source_reader import read_source
from core.target_reader import TEMPLATE_ACB, TEMPLATE_THONG_THUONG, read_target


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


def _tgt_row(
    row: int,
    account_no: Optional[str],
    full_name: str,
    current_amount: Optional[int] = 5_000_000,
) -> TargetRow:
    return TargetRow(
        row=row,
        account_no=account_no,
        full_name=full_name,
        name_key=normalize_name_key(full_name),
        bank_code=None,
        current_amount=current_amount,
    )


def _build_valid_plan() -> WritePlan:
    src = SourceReadResult(
        file_path="BANG_LUONG.xlsx",
        sheet_name="HLCK06.26 (2)",
        profile_name="Bảng lương",
        records=[
            _src_rec(5, "0001", "NGUYEN VAN AN", "24340001", 5_500_000),
            _src_rec(6, "0002", "LE THI BINH", "24340002", 6_200_000),
        ],
    )
    tgt = TargetReadResult(
        file_path="ACB.xls",
        file_format="xls",
        sheet_name="CNBD",
        template_name=TEMPLATE_ACB.name,
        rows=[
            _tgt_row(3, "24340001", "NGUYEN VAN AN", 5_000_000),
            _tgt_row(4, "24340002", "LE THI BINH", 6_000_000),
        ],
    )
    return reconcile(src, (tgt, TEMPLATE_ACB))


class TestSafetyHappyPath(unittest.TestCase):
    def test_valid_plan_is_safe(self) -> None:
        plan = _build_valid_plan()
        res = validate_write_plan(plan)
        self.assertTrue(res.is_valid)
        self.assertTrue(res.is_safe)
        self.assertTrue(res.ok)
        self.assertEqual(res.errors, [])
        self.assertEqual(res.warnings, [])
        self.assertTrue(res.row_balance_ok)
        self.assertTrue(res.amount_balance_ok)
        ensure_safe_to_write(plan)


class TestSafetyDetections(unittest.TestCase):
    def test_detects_ambiguous_mapping_and_duplicate_cell_coords(self) -> None:
        plan = _build_valid_plan()
        # 1. Có blocking_issues từ bước đọc (HEADER_MISMATCH)
        plan.blocking_issues.append(
            Issue(IssueCode.HEADER_MISMATCH, Severity.ERROR, "Cột D tiêu đề không khớp", sheet="CNBD", column="D")
        )
        # 2. Hai lệnh ghi đè vào cùng 1 ô D3 và 1 lệnh sai cột E
        dup_cell = CellUpdate("ACB.xls", "CNBD", 3, "D", 5_000_000, 5_500_000, "0009", "24349999")
        wrong_col = CellUpdate("ACB.xls", "CNBD", 5, "E", 5_000_000, 5_500_000, "0010", "24348888")
        plan.target_plans[0].updates.extend([dup_cell, wrong_col])
        plan.updates.extend([dup_cell, wrong_col])

        res = validate_write_plan(plan)
        self.assertFalse(res.is_safe)
        codes = [e.code for e in res.errors]
        self.assertIn(SafetyCode.AMBIGUOUS_MAPPING, codes)
        with self.assertRaises(SafetyError):
            res.ensure_safe()

    def test_detects_duplicate_identities_x8_and_update_duplicates(self) -> None:
        # 1. Trùng Số thẻ trong nguồn (X8)
        src = SourceReadResult(
            file_path="BANG_LUONG.xlsx",
            sheet_name="HLCK",
            profile_name="Bảng lương",
            records=[
                _src_rec(5, "0001", "NGUYEN AN", "24340001", 5_000_000),
                _src_rec(6, "0001", "LE BINH", "24340002", 6_000_000),
            ],
        )
        tgt = TargetReadResult(
            file_path="ACB.xls",
            file_format="xls",
            sheet_name="CNBD",
            template_name=TEMPLATE_ACB.name,
            rows=[
                _tgt_row(3, "24340001", "NGUYEN AN", 5_000_000),
                _tgt_row(4, "24340002", "LE BINH", 6_000_000),
            ],
        )
        plan = reconcile(src, tgt)
        res = validate_write_plan(plan)
        self.assertFalse(res.is_safe)
        self.assertTrue(any(e.code == SafetyCode.DUPLICATE_IDENTITY for e in res.errors))

        # 2. Can thiệp thủ công làm 2 CellUpdate trùng Số thẻ / Số TK
        plan2 = _build_valid_plan()
        u0 = plan2.updates[0]
        u1_dup = replace(plan2.updates[1], card_no=u0.card_no, account_no=u0.account_no)
        plan2.updates[1] = u1_dup
        plan2.target_plans[0].updates[1] = u1_dup
        res2 = validate_write_plan(plan2)
        self.assertFalse(res2.is_safe)
        self.assertTrue(any(e.code == SafetyCode.DUPLICATE_IDENTITY for e in res2.errors))

    def test_detects_unresolved_exceptions_until_resolved(self) -> None:
        src = SourceReadResult(
            file_path="BANG_LUONG.xlsx",
            sheet_name="HLCK",
            profile_name="Bảng lương",
            records=[_src_rec(5, "0001", "NGUYEN VAN AN", "24340001", 5_500_000)],
        )
        tgt = TargetReadResult(
            file_path="ACB.xls",
            file_format="xls",
            sheet_name="CNBD",
            template_name=TEMPLATE_ACB.name,
            rows=[
                _tgt_row(3, "24340001", "TEN KHAC HAN", 5_000_000),     # X6
                _tgt_row(4, "24349999", "NGHI VIEC", 4_000_000),        # X5
            ],
        )
        plan = reconcile(src, tgt)
        res1 = validate_write_plan(plan)
        self.assertFalse(res1.is_safe)
        self.assertTrue(any(e.code == SafetyCode.UNRESOLVED_EXCEPTION for e in res1.errors))

        # Sau khi người dùng xử lý hết X6 (CONFIRM_WRITE) và X5 (SET_ZERO) -> hợp lệ, kèm cảnh báo
        for exc in plan.exceptions:
            if exc.code.value == "X6":
                plan.resolve(exc.id, ResolutionAction.CONFIRM_WRITE)
            elif exc.code.value == "X5":
                plan.resolve(exc.id, ResolutionAction.SET_ZERO)

        res2 = validate_write_plan(plan)
        self.assertTrue(res2.is_safe)
        self.assertTrue(any(w.code == SafetyCode.ABNORMAL_CHANGE for w in res2.warnings))

    def test_detects_missing_account_and_missing_amount_in_updates(self) -> None:
        plan = _build_valid_plan()
        # 1. Sửa CellUpdate làm mất account_no và làm new_value = 0 dù reason = "MATCHED"
        bad_u = replace(plan.updates[0], account_no="", new_value=0, reason="MATCHED")
        plan.updates[0] = bad_u
        plan.target_plans[0].updates[0] = bad_u

        res = validate_write_plan(plan)
        self.assertFalse(res.is_safe)
        codes = {e.code for e in res.errors}
        self.assertIn(SafetyCode.MISSING_ACCOUNT, codes)
        self.assertIn(SafetyCode.MISSING_OR_INVALID_AMOUNT, codes)

    def test_detects_inconsistent_data_and_balance_mismatch_k8(self) -> None:
        # 1. Lệch số tiền giữa CellUpdate và MatchItem / tổng tiền nguồn (K8)
        plan1 = _build_valid_plan()
        tampered = replace(plan1.updates[0], new_value=5_500_001)  # lệch 1 đồng
        plan1.updates[0] = tampered
        plan1.target_plans[0].updates[0] = tampered

        res1 = validate_write_plan(plan1)
        self.assertFalse(res1.is_safe)
        self.assertFalse(res1.amount_balance_ok)
        self.assertTrue(any(e.code == SafetyCode.AMOUNT_BALANCE_MISMATCH for e in res1.errors))

        # 2. Lệch cân đối số dòng (K8): total_source_records không khớp
        plan2 = _build_valid_plan()
        plan2.total_source_records = 99
        res2 = validate_write_plan(plan2)
        self.assertFalse(res2.is_safe)
        self.assertFalse(res2.row_balance_ok)
        self.assertTrue(any(e.code == SafetyCode.ROW_BALANCE_MISMATCH for e in res2.errors))

        # 3. Orphan CellUpdate không gắn với MatchItem nào
        plan3 = _build_valid_plan()
        orphan = CellUpdate("ACB.xls", "CNBD", 99, "D", 1_000, 2_000, "9999", "24347777")
        plan3.updates.append(orphan)
        plan3.target_plans[0].updates.append(orphan)
        res3 = validate_write_plan(plan3)
        self.assertFalse(res3.is_safe)
        self.assertTrue(any(e.code == SafetyCode.INCONSISTENT_DATA for e in res3.errors))

    def test_detects_abnormal_changes_warnings_and_empty_updates_error(self) -> None:
        # 1. Số tiền > 1 tỷ và tăng đột biến > 3 lần -> cảnh báo (WARNING)
        src = SourceReadResult(
            file_path="BANG_LUONG.xlsx",
            sheet_name="HLCK",
            profile_name="Bảng lương",
            records=[_src_rec(5, "0001", "NGUYEN AN", "24340001", 1_500_000_000)],
        )
        tgt = TargetReadResult(
            file_path="ACB.xls",
            file_format="xls",
            sheet_name="CNBD",
            template_name=TEMPLATE_ACB.name,
            rows=[_tgt_row(3, "24340001", "NGUYEN AN", 10_000_000)],
        )
        plan = reconcile(src, tgt)
        res = validate_write_plan(plan)
        self.assertTrue(res.is_safe)
        self.assertGreaterEqual(len(res.warnings), 2)
        self.assertTrue(all(w.code == SafetyCode.ABNORMAL_CHANGE for w in res.warnings))

        # 2. Giảm đột biến (<= 20% kỳ trước) -> cảnh báo
        src_drop = SourceReadResult(
            file_path="BANG_LUONG.xlsx",
            sheet_name="HLCK",
            profile_name="Bảng lương",
            records=[_src_rec(5, "0001", "NGUYEN AN", "24340001", 1_000_000)],
        )
        tgt_drop = TargetReadResult(
            file_path="ACB.xls",
            file_format="xls",
            sheet_name="CNBD",
            template_name=TEMPLATE_ACB.name,
            rows=[_tgt_row(3, "24340001", "NGUYEN AN", 10_000_000)],
        )
        plan_drop = reconcile(src_drop, tgt_drop)
        res_drop = validate_write_plan(plan_drop)
        self.assertTrue(res_drop.is_safe)
        self.assertTrue(any("giảm đột biến" in w.message for w in res_drop.warnings))

        # 3. Kế hoạch rỗng (0 updates) -> ERROR chặn ghi
        empty_plan = WritePlan(
            source_file="BANG_LUONG.xlsx",
            source_sheet="HLCK",
            target_plans=[TargetWritePlan("ACB.xls", "xls", "CNBD", "ACB", "D")],
            matches=[
                MatchItem(
                    card_no="0001", account_no=None, source_name="A", target_name="",
                    bank_code=None, source_row=5, target_file=None, target_sheet=None,
                    target_row=None, target_cell=None, old_amount=None, new_amount=None,
                    source_amount=5_000_000, status=MatchStatus.SKIPPED,
                )
            ],
            total_source_records=1,
            total_source_transferable_amount=0,
        )
        res_empty = validate_write_plan(empty_plan)
        self.assertFalse(res_empty.is_safe)
        self.assertTrue(any(e.code == SafetyCode.ABNORMAL_CHANGE for e in res_empty.errors))

    def test_run_result_forbids_overwriting_original_file_r8(self) -> None:
        plan = _build_valid_plan()
        # Đường dẫn output trùng với file gốc ACB.xls -> chặn cứng theo R8
        bad_run = RunResult(
            write_plan=plan,
            backup_files=[],
            output_files=["ACB.xls"],
        )
        res = validate_safety(bad_run)
        self.assertFalse(res.is_safe)
        codes = {e.code for e in res.errors}
        self.assertIn(SafetyCode.OVERWRITE_ORIGINAL_FORBIDDEN, codes)
        self.assertTrue(any(w.code == SafetyCode.MISSING_BACKUP for w in res.warnings))

        # Đường dẫn output sang file mới + có backup -> an toàn
        good_run = RunResult(
            write_plan=plan,
            backup_files=["data/backup/ACB_202606.xls"],
            output_files=["data/output/ACB_Thang06_2026.xls"],
        )
        res_ok = validate_safety(good_run)
        self.assertTrue(res_ok.is_safe)

    def test_file_lock_detection_k9(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "ACB.xls"
            p.write_bytes(b"dummy")
            self.assertFalse(is_file_locked(p))

            # Giả lập file tạm ~$ACB.xls do Excel tạo khi đang mở file
            lock_p = Path(tmp) / "~$ACB.xls"
            lock_p.write_bytes(b"lock")
            self.assertTrue(is_file_locked(p))
            findings = check_file_locks([p])
            self.assertEqual(len(findings), 1)
            self.assertEqual(findings[0].code, SafetyCode.FILE_LOCKED)


@unittest.skipUnless(
    os.environ.get("CHL_REAL_SOURCE")
    and os.environ.get("CHL_REAL_ACB")
    and os.environ.get("CHL_REAL_TT"),
    "Đặt CHL_REAL_SOURCE, CHL_REAL_ACB, CHL_REAL_TT để kiểm tra Safety Layer trên bộ file thật.",
)
class TestSafetyRealMay2026Files(unittest.TestCase):
    def test_real_may_2026_plan_passes_all_safety_gates(self) -> None:
        src = read_source(Path(os.environ["CHL_REAL_SOURCE"]), sheet_name="HLCK05.26 (2)")
        r_acb = read_target(Path(os.environ["CHL_REAL_ACB"]), TEMPLATE_ACB)
        r_tt = read_target(Path(os.environ["CHL_REAL_TT"]), TEMPLATE_THONG_THUONG)

        plan = reconcile(src, [(r_acb, TEMPLATE_ACB), (r_tt, TEMPLATE_THONG_THUONG)])
        val = validate_write_plan(plan, check_disk=True)

        self.assertTrue(val.is_safe)
        self.assertEqual(val.errors, [])
        self.assertTrue(val.row_balance_ok)
        self.assertTrue(val.amount_balance_ok)


if __name__ == "__main__":
    unittest.main()
