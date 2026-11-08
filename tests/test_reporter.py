"""
test_reporter.py - Unit test cho Module 8 (core/reporter.py).
"""

from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

import openpyxl

from core.models import (
    CellDiff,
    RuntimeConfig,
    Severity,
    VerificationCode,
    VerificationFinding,
    VerificationResult,
)
from core.reporter import (
    Reporter,
    SHEET_DETAIL,
    SHEET_EXCEPTION,
    SHEET_SUMMARY,
    SHEET_VALIDATION,
    build_report_data,
    generate_report,
)
from core.workflow import prepare_workflow, run_workflow
from tests.helpers import build_source_xlsx, build_target_xlsx


class TestReporter(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="chl_test_reporter_"))
        self.cfg = RuntimeConfig(
            templates_dir=str(self.tmp_dir / "templates"),
            output_dir=str(self.tmp_dir / "output"),
            backup_dir=str(self.tmp_dir / "backup"),
            writer_engine="openpyxl",
        )
        self.report_dir = self.tmp_dir / "reports"

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def _build_clean_pair(self) -> tuple[Path, Path]:
        src_path = build_source_xlsx(
            self.tmp_dir / "BẢNG LƯƠNG THÁNG 06.2026.xlsx",
            rows=[
                ("00123", "NGUYEN VAN A", "NGUYEN VAN A", "0587952606", 6_500_000),
                ("00124", "TRAN THI B", "TRAN THI B", "444555666", 7_200_000),
                ("00125", "LE VAN C", "LE VAN C", "111222333", 8_100_000),
            ],
        )
        acb_path = build_target_xlsx(
            self.tmp_dir / "4 MauChiHoLuongQuaTaiKhoanACB.xlsx",
            rows=[
                ("NGUYEN VAN A", "0587952606", 5_000_000),
                ("TRAN THI B", "444555666", 7_000_000),
                ("LE VAN C", "111222333", 8_000_000),
            ],
            with_guide=True,
        )
        return src_path, acb_path

    def test_generate_report_from_workflow_and_run_result_xlsx(self) -> None:
        src_path, acb_path = self._build_clean_pair()
        wf_res = run_workflow(src_path, acb_path, self.cfg)

        rep = generate_report(
            run_result=wf_res.run_result,
            workflow_result=wf_res,
            verification_result=wf_res.verification_result,
            validation_result=wf_res.validation_result,
            report_dir=self.report_dir,
        )

        self.assertIsNotNone(rep.report_file)
        assert rep.report_file is not None
        report_path = Path(rep.report_file)
        self.assertTrue(report_path.exists())
        assert wf_res.run_result is not None
        self.assertEqual(wf_res.run_result.report_file, rep.report_file)

        # Kiểm tra SummaryReportData
        s = rep.summary_report
        self.assertTrue(s.is_safe)
        self.assertTrue(s.is_verified)
        self.assertEqual(s.total_source_records, 3)
        self.assertEqual(s.total_matched_records, 3)
        self.assertEqual(s.total_written_records, 3)
        self.assertEqual(s.total_source_transferable_amount, 21_800_000)
        self.assertEqual(s.total_old_target_amount, 20_000_000)
        self.assertEqual(s.total_planned_write_amount, 21_800_000)
        self.assertEqual(s.row_discrepancy, 0)
        self.assertEqual(s.amount_discrepancy, 0)
        self.assertEqual(len(s.target_summaries), 1)
        self.assertEqual(s.target_summaries[0].delta_amount, 1_800_000)

        # Kiểm tra DetailReportRow
        self.assertEqual(len(rep.detail_rows), 3)
        self.assertEqual(rep.detail_rows[0].card_no, "00123")
        self.assertEqual(rep.detail_rows[0].account_no, "0587952606")

        # Đọc lại file Excel báo cáo (.xlsx) để kiểm tra 4 sheet và định dạng Text '@'
        wb = openpyxl.load_workbook(report_path)
        self.assertEqual(
            wb.sheetnames,
            [SHEET_SUMMARY, SHEET_DETAIL, SHEET_EXCEPTION, SHEET_VALIDATION],
        )
        ws_detail = wb[SHEET_DETAIL]
        self.assertEqual(ws_detail["B2"].value, "00123")
        self.assertEqual(ws_detail["B2"].number_format, "@")
        self.assertEqual(ws_detail["C2"].value, "0587952606")
        self.assertEqual(ws_detail["C2"].number_format, "@")
        self.assertEqual(ws_detail["N2"].value, 6_500_000)
        wb.close()

    def test_exception_and_validation_reports_populated(self) -> None:
        src_path = build_source_xlsx(
            self.tmp_dir / "BẢNG LƯƠNG THÁNG 06.2026.xlsx",
            rows=[
                ("001", "NGUYEN VAN A", "NGUYEN VAN A", "0587952606", 6_500_000),
                ("002", "LE VAN C", "LE VAN C", None, 5_000_000),  # X3_MISSING_ACCOUNT
            ],
        )
        acb_path = build_target_xlsx(
            self.tmp_dir / "4 MauChiHoLuongQuaTaiKhoanACB.xlsx",
            rows=[
                ("TEN KHAC", "0587952606", 5_000_000),        # X6_TARGET_MISMATCH
                ("NGUOI NGHI VIEC", "999888777", 4_000_000),  # X5_TARGET_NOT_IN_SOURCE
            ],
            with_guide=True,
        )

        wf_res = prepare_workflow(src_path, acb_path, self.cfg)
        custom_ver = VerificationResult(
            post_write_diff_ok=False,
            errors=[
                VerificationFinding(
                    code=VerificationCode.POST_WRITE_VALUE_MISMATCH,
                    severity=Severity.ERROR,
                    message="Ô D3 sai lệch hậu kiểm.",
                    file_path=str(acb_path),
                    sheet="CNBD",
                    row=3,
                    column="D",
                )
            ],
            cell_diffs=[
                CellDiff(
                    file_path=str(acb_path),
                    sheet_name="CNBD",
                    row=3,
                    column="D",
                    expected_value=6_500_000,
                    actual_value=5_000_000,
                    is_planned_cell=True,
                    message="Lệch giá trị ô D3",
                )
            ],
        )

        reporter = Reporter(report_dir=self.report_dir)
        rep = reporter.generate(
            workflow_result=wf_res,
            verification_result=custom_ver,
            validation_result=wf_res.validation_result,
        )

        self.assertEqual(len(rep.exception_report), 3)
        exc_codes = {e.code for e in rep.exception_report}
        self.assertEqual(exc_codes, {"X3", "X5", "X6"})

        layers = {v.layer for v in rep.validation_report}
        self.assertIn("SAFETY", layers)
        self.assertIn("VERIFIER", layers)
        self.assertIn("CELL_DIFF", layers)

        # Kiểm tra to_dict() và dữ liệu không ghi file
        rep_mem = build_report_data(workflow_result=wf_res)
        self.assertIsNone(rep_mem.report_file)
        d = rep_mem.to_dict()
        self.assertEqual(len(d["exception_report"]), 3)


if __name__ == "__main__":
    unittest.main()
