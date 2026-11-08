"""
test_verifier.py - Unit test cho Module 6 (core/verifier.py).
"""

from __future__ import annotations

import os
import shutil
import tempfile
import unittest
from pathlib import Path

import openpyxl

from core.excel_writer import execute_write_plan, is_excel_com_available
from core.models import (
    CellUpdate,
    ExceptionCode,
    ResolutionAction,
    RunResult,
    VerificationCode,
    VerificationError,
)
from core.reconciler import reconcile, resolve_exception
from core.safety import validate_write_plan
from core.source_reader import read_source
from core.target_reader import TEMPLATE_ACB, TEMPLATE_THONG_THUONG, read_target
from core.verifier import (
    Verifier,
    ensure_verified,
    verify,
    verify_output_file_diff,
)
from tests.helpers import build_source_xlsx, build_target_xlsx


class TestVerifier(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="chl_test_verifier_"))

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def _build_clean_files(self) -> tuple[Path, Path]:
        src_path = build_source_xlsx(
            self.tmp_dir / "bang_luong.xlsx",
            rows=[
                ("HL001", "NGUYEN VAN A", "NGUYEN VAN A", "111222333", 6_500_000),
                ("HL002", "TRAN THI B", "TRAN THI B", "444555666", 7_200_000),
                ("HL003", "LE VAN C", "LE VAN C", "0587952606", 8_100_000),
            ],
        )
        acb_path = build_target_xlsx(
            self.tmp_dir / "acb.xlsx",
            rows=[
                ("NGUYEN VAN A", "111222333", 5_000_000),
                ("TRAN THI B", "444555666", 7_000_000),
                ("LE VAN C", "0587952606", 8_000_000),
            ],
            with_guide=True,
        )
        return src_path, acb_path

    def test_clean_chain_and_post_write_diff_xlsx(self) -> None:
        src_path, acb_path = self._build_clean_files()
        src_res = read_source(src_path)
        acb_res = read_target(acb_path, TEMPLATE_ACB)

        plan = reconcile(src_res, [(acb_res, TEMPLATE_ACB)])
        val_res = validate_write_plan(plan)
        exec_res = execute_write_plan(
            plan,
            val_res,
            [TEMPLATE_ACB],
            [acb_res],
            output_dir=self.tmp_dir / "out",
            backup_dir=self.tmp_dir / "bak",
            engine="openpyxl",
        )
        run_res = RunResult(
            write_plan=plan,
            output_files=exec_res.output_files,
            backup_files=exec_res.backup_files,
        )

        vres = verify(
            run_result=run_res,
            validation_result=val_res,
            write_execution_result=exec_res,
            source=src_res,
        )
        self.assertTrue(vres.is_verified)
        self.assertTrue(vres.source_to_reconcile_ok)
        self.assertTrue(vres.reconcile_to_plan_ok)
        self.assertTrue(vres.post_write_diff_ok)
        self.assertEqual(vres.total_source_records, 3)
        self.assertEqual(vres.total_matched_records, 3)
        self.assertEqual(vres.total_exceptions, 0)
        self.assertEqual(vres.total_written_records, 3)
        self.assertEqual(vres.total_source_written_records, 3)
        self.assertEqual(vres.total_source_amount, 21_800_000)
        self.assertEqual(vres.total_source_transferable_amount, 21_800_000)
        self.assertEqual(vres.total_planned_write_amount, 21_800_000)
        self.assertEqual(vres.row_discrepancy, 0)
        self.assertEqual(vres.amount_discrepancy, 0)
        self.assertEqual(len(vres.cell_diffs), 0)
        self.assertEqual(len(vres.verified_files), 1)

        # Kiểm tra thông qua class Verifier và ensure_verified
        v_obj = Verifier()
        self.assertTrue(v_obj.verify(run_res, val_res, exec_res, source=src_res).is_verified)
        self.assertTrue(ensure_verified(run_res, val_res, exec_res, source=src_res).is_verified)

    def test_resolved_exceptions_skip_zero_override_balance(self) -> None:
        src_path = build_source_xlsx(
            self.tmp_dir / "src_exc.xlsx",
            rows=[
                ("HL001", "NGUYEN VAN A", "NGUYEN VAN A", "111222333", 6_500_000),
                ("HL002", "TRAN THI B", "TRAN THI B", "444555666", 0),  # X7_INVALID_AMOUNT
                ("HL003", "LE VAN C", "LE VAN C", None, 5_000_000),  # X3_MISSING_ACCOUNT (tự động SKIP)
            ],
        )
        acb_path = build_target_xlsx(
            self.tmp_dir / "acb_exc.xlsx",
            rows=[
                ("PHAM VAN KHAC", "111222333", 6_000_000),  # X6_TARGET_MISMATCH (lệch tên)
                ("TRAN THI B", "444555666", 7_000_000),
                ("NGUOI NGHI VIEC", "999888777", 4_500_000),  # X5_TARGET_NOT_IN_SOURCE
            ],
        )
        src_res = read_source(src_path)
        acb_res = read_target(acb_path, TEMPLATE_ACB)
        plan = reconcile(src_res, [(acb_res, TEMPLATE_ACB)])

        # Khi chưa xử lý X6, X7, X5 -> verify báo lỗi UNRESOLVED_EXCEPTION
        v_unresolved = verify(plan, source=src_res, verify_files_on_disk=False)
        self.assertFalse(v_unresolved.is_verified)
        self.assertIn(
            VerificationCode.UNRESOLVED_EXCEPTION,
            {e.code for e in v_unresolved.errors},
        )

        # Xử lý: X6 -> CONFIRM_WRITE, X7 -> OVERRIDE_AMOUNT (8_000_000), X5 -> SET_ZERO
        for exc in plan.exceptions:
            if exc.code == ExceptionCode.X6_TARGET_MISMATCH:
                resolve_exception(plan, exc.id, ResolutionAction.CONFIRM_WRITE)
            elif exc.code == ExceptionCode.X7_INVALID_AMOUNT:
                resolve_exception(plan, exc.id, ResolutionAction.OVERRIDE_AMOUNT, override_amount=8_000_000)
            elif exc.code == ExceptionCode.X5_TARGET_NOT_IN_SOURCE:
                resolve_exception(plan, exc.id, ResolutionAction.SET_ZERO)

        val_res = validate_write_plan(plan)
        self.assertTrue(val_res.is_safe)

        vres = verify(plan, val_res, source=src_res, verify_files_on_disk=False)
        self.assertTrue(vres.is_verified)
        self.assertEqual(vres.total_source_records, 3)
        self.assertEqual(vres.total_matched_records, 2)
        self.assertEqual(vres.total_exceptions, 4)
        self.assertEqual(vres.resolved_exceptions, 4)
        self.assertEqual(vres.unresolved_exceptions, 0)
        self.assertEqual(vres.total_written_records, 3)  # 2 nguồn + 1 dòng X5=0
        self.assertEqual(vres.total_source_written_records, 2)
        self.assertEqual(vres.total_zero_or_clear_records, 1)
        self.assertEqual(vres.total_skipped_records, 1)
        self.assertEqual(vres.total_source_amount, 11_500_000)
        self.assertEqual(vres.total_source_transferable_amount, 6_500_000)
        self.assertEqual(vres.total_override_delta, 8_000_000)
        self.assertEqual(vres.total_planned_write_amount, 14_500_000)
        self.assertEqual(vres.row_discrepancy, 0)
        self.assertEqual(vres.amount_discrepancy, 0)

    def test_detects_discrepancies_and_chain_tampering(self) -> None:
        src_path, acb_path = self._build_clean_files()
        src_res = read_source(src_path)
        acb_res = read_target(acb_path, TEMPLATE_ACB)
        plan = reconcile(src_res, [(acb_res, TEMPLATE_ACB)])

        # Cố tình làm lệch total_source_records và giá trị 1 CellUpdate
        plan.total_source_records = 5
        old_u = plan.updates[0]
        tampered_u = CellUpdate(
            file_path=old_u.file_path,
            sheet_name=old_u.sheet_name,
            row=old_u.row,
            column=old_u.column,
            old_value=old_u.old_value,
            new_value=99_999_999,
            card_no=old_u.card_no,
            account_no=old_u.account_no,
            full_name=old_u.full_name,
            reason=old_u.reason,
        )
        plan.updates[0] = tampered_u
        plan.target_plans[0].updates[0] = tampered_u

        vres = verify(plan, source=src_res, verify_files_on_disk=False)
        self.assertFalse(vres.is_verified)
        self.assertFalse(vres.source_to_reconcile_ok)
        self.assertFalse(vres.reconcile_to_plan_ok)
        self.assertNotEqual(vres.row_discrepancy, 0)
        self.assertNotEqual(vres.amount_discrepancy, 0)

        err_codes = {e.code for e in vres.errors}
        self.assertIn(VerificationCode.SOURCE_RECONCILE_MISMATCH, err_codes)
        self.assertIn(VerificationCode.RECONCILE_PLAN_MISMATCH, err_codes)
        self.assertIn(VerificationCode.ROW_DISCREPANCY, err_codes)
        self.assertIn(VerificationCode.AMOUNT_DISCREPANCY, err_codes)

        with self.assertRaises(VerificationError):
            ensure_verified(plan, source=src_res, verify_files_on_disk=False)

    def test_detects_post_write_planned_and_unplanned_cell_tampering(self) -> None:
        src_path, acb_path = self._build_clean_files()
        src_res = read_source(src_path)
        acb_res = read_target(acb_path, TEMPLATE_ACB)

        plan = reconcile(src_res, [(acb_res, TEMPLATE_ACB)])
        val_res = validate_write_plan(plan)
        exec_res = execute_write_plan(
            plan,
            val_res,
            [TEMPLATE_ACB],
            [acb_res],
            output_dir=self.tmp_dir / "out",
            backup_dir=self.tmp_dir / "bak",
            engine="openpyxl",
        )

        # Can thiệp trái phép vào 1 ô kế hoạch (D3) + 1 ô ngoài kế hoạch (B3)
        acb_out_path = Path(exec_res.target_outputs[0].output_file)
        wb = openpyxl.load_workbook(acb_out_path)
        ws = wb["CNBD"]
        ws["D3"].value = 123_456  # sai lệch ô trong kế hoạch
        ws["B3"].value = "TEN BI SUA TRAI PHEP"  # sửa ô ngoài kế hoạch
        wb.save(acb_out_path)
        wb.close()

        vres = verify(exec_res, val_res, source=src_res)
        self.assertFalse(vres.is_verified)
        self.assertFalse(vres.post_write_diff_ok)
        err_codes = {e.code for e in vres.errors}
        self.assertIn(VerificationCode.POST_WRITE_VALUE_MISMATCH, err_codes)
        self.assertIn(VerificationCode.UNEXPECTED_CELL_MODIFIED, err_codes)

        planned_diffs = [d for d in vres.cell_diffs if d.is_planned_cell]
        unplanned_diffs = [d for d in vres.cell_diffs if not d.is_planned_cell]
        self.assertEqual(len(planned_diffs), 1)
        self.assertEqual(planned_diffs[0].cell, "D3")
        self.assertEqual(planned_diffs[0].expected_value, 6_500_000)
        self.assertEqual(planned_diffs[0].actual_value, 123_456)

        self.assertEqual(len(unplanned_diffs), 1)
        self.assertEqual(unplanned_diffs[0].cell, "B3")

    def test_detects_original_file_modified_and_missing_files(self) -> None:
        src_path, acb_path = self._build_clean_files()
        src_res = read_source(src_path)
        acb_res = read_target(acb_path, TEMPLATE_ACB)

        plan = reconcile(src_res, [(acb_res, TEMPLATE_ACB)])
        val_res = validate_write_plan(plan)
        exec_res = execute_write_plan(
            plan,
            val_res,
            [TEMPLATE_ACB],
            [acb_res],
            output_dir=self.tmp_dir / "out",
            backup_dir=self.tmp_dir / "bak",
            engine="openpyxl",
        )

        # Sửa file gốc sau khi ghi -> SHA256 thay đổi & xóa file backup của file đích
        with acb_path.open("ab") as fh:
            fh.write(b" ")
        Path(exec_res.target_outputs[0].backup_file).unlink()

        vres = verify(exec_res, val_res, source=src_res)
        self.assertFalse(vres.is_verified)
        err_codes = {e.code for e in vres.errors}
        self.assertIn(VerificationCode.ORIGINAL_FILE_MODIFIED, err_codes)
        self.assertIn(VerificationCode.MISSING_BACKUP_FILE, err_codes)

    def test_real_may_2026_end_to_end_verification(self) -> None:
        real_src = os.environ.get("CHL_REAL_SOURCE")
        real_acb = os.environ.get("CHL_REAL_ACB")
        real_tt = os.environ.get("CHL_REAL_TT")
        if not (real_src and real_acb and real_tt):
            self.skipTest("Bỏ qua test dữ liệu thật (chưa đặt biến môi trường CHL_REAL_*).")
        if not (Path(real_src).exists() and Path(real_acb).exists() and Path(real_tt).exists()):
            self.skipTest("Không tìm thấy bộ file thật tháng 05/2026.")
        if not is_excel_com_available():
            self.skipTest("Cần Excel COM để ghi file .xls thật.")

        src_res = read_source(real_src, sheet_name="HLCK05.26 (2)")
        acb_res = read_target(real_acb, TEMPLATE_ACB)
        tt_res = read_target(real_tt, TEMPLATE_THONG_THUONG)

        plan = reconcile(src_res, [(acb_res, TEMPLATE_ACB), (tt_res, TEMPLATE_THONG_THUONG)])
        val_res = validate_write_plan(plan)
        exec_res = execute_write_plan(
            plan,
            val_res,
            [TEMPLATE_ACB, TEMPLATE_THONG_THUONG],
            [acb_res, tt_res],
            output_dir=self.tmp_dir / "real_out",
            backup_dir=self.tmp_dir / "real_bak",
        )

        vres = verify(exec_res, val_res, source=src_res)
        self.assertTrue(vres.is_verified)
        self.assertEqual(vres.total_source_records, 90)
        self.assertEqual(vres.total_matched_records, 86)
        self.assertEqual(vres.total_exceptions, 4)
        self.assertEqual(vres.resolved_exceptions, 4)
        self.assertEqual(vres.total_written_records, 86)
        self.assertEqual(vres.total_skipped_records, 4)
        self.assertEqual(vres.total_source_transferable_amount, 445_138_601)
        self.assertEqual(vres.total_planned_write_amount, 445_138_601)
        self.assertEqual(vres.row_discrepancy, 0)
        self.assertEqual(vres.amount_discrepancy, 0)
        self.assertEqual(len(vres.cell_diffs), 0)
        self.assertEqual(len(vres.verified_files), 2)

        # Kiểm tra trực tiếp hàm verify_output_file_diff
        for tp, out_meta in zip(plan.target_plans, exec_res.target_outputs):
            diffs = verify_output_file_diff(
                original_path=out_meta.source_target_file,
                output_path=out_meta.output_file,
                target_plan=tp,
            )
            self.assertEqual(diffs, [])


if __name__ == "__main__":
    unittest.main()
