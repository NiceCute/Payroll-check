"""
test_workflow.py - Unit test cho Module 7 (core/workflow.py).
"""

from __future__ import annotations

import os
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core.excel_writer import is_excel_com_available
from core.models import (
    ExceptionCode,
    ResolutionAction,
    RuntimeConfig,
    SafetyError,
    VerificationCode,
    VerificationFinding,
    VerificationResult,
    Severity,
    WorkflowStage,
)
from core.workflow import (
    Workflow,
    execute_workflow,
    prepare_workflow,
    run_workflow,
)
from tests.helpers import build_source_xlsx, build_target_xlsx


class TestWorkflow(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="chl_test_workflow_"))
        self.cfg = RuntimeConfig(
            templates_dir=str(self.tmp_dir / "templates"),
            output_dir=str(self.tmp_dir / "output"),
            backup_dir=str(self.tmp_dir / "backup"),
            writer_engine="openpyxl",
        )

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def _build_clean_pair(self) -> tuple[Path, Path]:
        src_path = build_source_xlsx(
            self.tmp_dir / "BẢNG LƯƠNG THÁNG 06.2026.xlsx",
            rows=[
                ("HL001", "NGUYEN VAN A", "NGUYEN VAN A", "111222333", 6_500_000),
                ("HL002", "TRAN THI B", "TRAN THI B", "444555666", 7_200_000),
                ("HL003", "LE VAN C", "LE VAN C", "0587952606", 8_100_000),
            ],
        )
        acb_path = build_target_xlsx(
            self.tmp_dir / "4 MauChiHoLuongQuaTaiKhoanACB.xlsx",
            rows=[
                ("NGUYEN VAN A", "111222333", 5_000_000),
                ("TRAN THI B", "444555666", 7_000_000),
                ("LE VAN C", "0587952606", 8_000_000),
            ],
            with_guide=True,
        )
        return src_path, acb_path

    def test_full_pipeline_clean_xlsx(self) -> None:
        src_path, acb_path = self._build_clean_pair()

        res = run_workflow(src_path, acb_path, self.cfg)
        self.assertTrue(res.is_success)
        self.assertEqual(res.stage, WorkflowStage.COMPLETED)
        self.assertTrue(res.verified)
        self.assertIsNotNone(res.run_result)
        assert res.run_result is not None
        self.assertTrue(res.run_result.verified)
        self.assertEqual(len(res.output_files), 1)
        self.assertEqual(len(res.backup_files), 2)
        self.assertTrue(Path(res.output_files[0]).exists())
        self.assertEqual(res.to_run_result(), res.run_result)

        assert res.verification_result is not None
        self.assertEqual(res.verification_result.total_source_records, 3)
        self.assertEqual(res.verification_result.total_written_records, 3)
        self.assertEqual(res.verification_result.total_planned_write_amount, 21_800_000)
        self.assertEqual(res.verification_result.row_discrepancy, 0)
        self.assertEqual(res.verification_result.amount_discrepancy, 0)

    def test_interactive_exception_flow_and_safety_stop(self) -> None:
        src_path = build_source_xlsx(
            self.tmp_dir / "BẢNG LƯƠNG THÁNG 06.2026.xlsx",
            rows=[
                ("HL001", "NGUYEN VAN A", "NGUYEN VAN A", "111222333", 6_500_000),
                ("HL002", "TRAN THI B", "TRAN THI B", "444555666", 7_200_000),
            ],
        )
        acb_path = build_target_xlsx(
            self.tmp_dir / "4 MauChiHoLuongQuaTaiKhoanACB.xlsx",
            rows=[
                ("TEN KHAC", "111222333", 5_000_000),        # X6_TARGET_MISMATCH
                ("TRAN THI B", "444555666", 7_000_000),
                ("NGUOI NGHI VIEC", "999888777", 4_000_000),  # X5_TARGET_NOT_IN_SOURCE
            ],
            with_guide=True,
        )

        wf = Workflow(self.cfg)
        prep = wf.prepare(src_path, acb_path)
        self.assertEqual(prep.stage, WorkflowStage.AWAITING_RESOLUTION)
        self.assertTrue(prep.has_unresolved_exceptions)
        self.assertEqual(len(prep.unresolved_exceptions), 2)
        self.assertFalse(prep.can_proceed_to_write)

        # Gọi execute() khi còn ngoại lệ chưa xử lý -> Safety chặn ngay, không ghi file
        stopped = wf.execute()
        self.assertEqual(stopped.stage, WorkflowStage.SAFETY_FAILED)
        self.assertFalse(stopped.is_success)
        self.assertEqual(len(stopped.output_files), 0)

        # Kiểm tra raise_on_failure=True ném SafetyError
        wf.config.raise_on_failure = True
        with self.assertRaises(SafetyError):
            wf.execute()
        wf.config.raise_on_failure = False

        # Xử lý từng ngoại lệ qua Workflow.resolve_exception
        for exc in list(wf.exceptions):
            if exc.code == ExceptionCode.X6_TARGET_MISMATCH:
                wf.resolve_exception(exc.id, ResolutionAction.CONFIRM_WRITE)
            elif exc.code == ExceptionCode.X5_TARGET_NOT_IN_SOURCE:
                wf.resolve_exception(exc.id, ResolutionAction.SET_ZERO)

        self.assertFalse(wf.result.has_unresolved_exceptions)
        self.assertTrue(wf.result.can_proceed_to_write)
        self.assertEqual(wf.stage, WorkflowStage.PRE_VERIFIED)

        # Thực thi xuất file sau khi đã xử lý hết ngoại lệ
        final_res = execute_workflow(wf)
        self.assertTrue(final_res.is_success)
        self.assertEqual(final_res.stage, WorkflowStage.COMPLETED)
        self.assertEqual(len(final_res.output_files), 1)
        assert final_res.verification_result is not None
        self.assertEqual(final_res.verification_result.total_written_records, 3)
        self.assertEqual(final_res.verification_result.total_planned_write_amount, 13_700_000)

    def test_rollback_when_post_write_verifier_fails(self) -> None:
        src_path, acb_path = self._build_clean_pair()
        wf = Workflow(self.cfg)
        wf.prepare(src_path, acb_path)

        real_pre_ver = wf.verify_plan()
        bad_post_ver = VerificationResult(
            post_write_diff_ok=False,
            errors=[
                VerificationFinding(
                    code=VerificationCode.POST_WRITE_VALUE_MISMATCH,
                    severity=Severity.ERROR,
                    message="Giả lập lỗi hậu kiểm sau khi ghi file.",
                )
            ],
        )

        # Lần gọi verify thứ 1 (pre-write) trả kết quả hợp lệ, lần gọi thứ 2 (post-write) trả lỗi
        with patch("core.workflow.verify", side_effect=[real_pre_ver, bad_post_ver]):
            res = wf.execute()

        self.assertFalse(res.is_success)
        self.assertEqual(res.stage, WorkflowStage.VERIFICATION_FAILED)
        self.assertIsNotNone(res.run_result)
        assert res.run_result is not None
        self.assertFalse(res.run_result.verified)
        # File output mới phải bị thu hồi (xóa) và ghi nhận vào rolled_back_files
        self.assertEqual(len(res.rolled_back_files), 1)
        self.assertFalse(Path(res.rolled_back_files[0]).exists())
        self.assertEqual(res.output_files, [])
        # File backup vẫn được giữ nguyên
        self.assertEqual(len(res.backup_files), 2)
        for bak in res.backup_files:
            self.assertTrue(Path(bak).exists())

    def test_unmatched_target_file_stops_in_prepare(self) -> None:
        import openpyxl

        src_path, _ = self._build_clean_pair()
        unknown_target = self.tmp_dir / "FileKhongTheoMau.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "SheetLa"
        ws["A1"], ws["B1"] = "CotKhongQuen", "GiaTriKhac"
        ws["A2"], ws["B2"] = "X", 100
        wb.save(unknown_target)
        wb.close()

        res = prepare_workflow(src_path, unknown_target, self.cfg)
        self.assertEqual(res.stage, WorkflowStage.FAILED)
        self.assertFalse(res.is_success)
        self.assertEqual(len(res.unmatched_target_files), 1)

    def test_dry_run_mode_does_not_write_files(self) -> None:
        src_path, acb_path = self._build_clean_pair()
        self.cfg.dry_run = True
        res = run_workflow(src_path, acb_path, self.cfg)
        self.assertEqual(res.stage, WorkflowStage.PRE_VERIFIED)
        self.assertEqual(res.output_files, [])
        self.assertIsNotNone(res.run_result)
        assert res.run_result is not None
        self.assertTrue(res.run_result.verified)

    def test_real_may_2026_end_to_end_workflow(self) -> None:
        real_src = os.environ.get("CHL_REAL_SOURCE")
        real_acb = os.environ.get("CHL_REAL_ACB")
        real_tt = os.environ.get("CHL_REAL_TT")
        if not (real_src and real_acb and real_tt):
            self.skipTest("Bỏ qua test dữ liệu thật (chưa đặt biến môi trường CHL_REAL_*).")
        if not (Path(real_src).exists() and Path(real_acb).exists() and Path(real_tt).exists()):
            self.skipTest("Không tìm thấy bộ file thật tháng 05/2026.")
        if not is_excel_com_available():
            self.skipTest("Cần Excel COM để ghi file .xls thật.")

        src_copy = self.tmp_dir / "BẢNG LƯƠNG THÁNG 05.2026.xlsx"
        acb_copy = self.tmp_dir / "4 MauChiHoLuongQuaTaiKhoanACB.xls"
        tt_copy = self.tmp_dir / "5 MauChiHoLuongThongThuong.xls"
        shutil.copy2(real_src, src_copy)
        shutil.copy2(real_acb, acb_copy)
        shutil.copy2(real_tt, tt_copy)

        cfg = RuntimeConfig(
            templates_dir=str(self.tmp_dir / "templates"),
            output_dir=str(self.tmp_dir / "real_out"),
            backup_dir=str(self.tmp_dir / "real_bak"),
            source_sheet="HLCK05.26 (2)",
            writer_engine="auto",
        )
        res = run_workflow(src_copy, [acb_copy, tt_copy], cfg)
        self.assertTrue(res.is_success)
        self.assertEqual(res.stage, WorkflowStage.COMPLETED)
        self.assertTrue(res.verified)
        self.assertEqual(len(res.output_files), 2)
        self.assertEqual(len(res.backup_files), 3)
        assert res.verification_result is not None
        self.assertEqual(res.verification_result.total_source_records, 90)
        self.assertEqual(res.verification_result.total_matched_records, 86)
        self.assertEqual(res.verification_result.total_written_records, 86)
        self.assertEqual(res.verification_result.total_planned_write_amount, 445_138_601)
        self.assertEqual(res.verification_result.row_discrepancy, 0)
        self.assertEqual(res.verification_result.amount_discrepancy, 0)


if __name__ == "__main__":
    unittest.main()
