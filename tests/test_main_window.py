"""
test_main_window.py - Unit test cho GUI Module 1 (gui/main_window.py).
"""

from __future__ import annotations

import shutil
import tempfile
import tkinter as tk
import unittest
from pathlib import Path

from core.models import ExceptionCode, ResolutionAction, RuntimeConfig, WorkflowStage
from core.reporter import Reporter
from gui.main_window import AUTO_TEMPLATE_OPTION, MainWindow
from tests.helpers import build_source_xlsx, build_target_xlsx


class TestMainWindow(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="chl_test_gui_main_"))
        self.root = tk.Tk()
        self.root.withdraw()

        self.cfg = RuntimeConfig(
            templates_dir=str(self.tmp_dir / "templates"),
            output_dir=str(self.tmp_dir / "output"),
            backup_dir=str(self.tmp_dir / "backup"),
            writer_engine="openpyxl",
        )
        self.reporter = Reporter(report_dir=self.tmp_dir / "reports")
        self.win = MainWindow(
            self.root,
            config=self.cfg,
            reporter=self.reporter,
            interactive=False,
        )

    def tearDown(self) -> None:
        try:
            self.win.destroy()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_file_selection_template_selection_analyze_preview_execute(self) -> None:
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

        # 1. Chọn file nguồn và file đích thông qua browse callback
        self.win.browse_source_file(dialog_fn=lambda **_: str(src_path))
        self.win.browse_target_files(dialog_fn=lambda **_: [str(acb_path)])
        self.assertEqual(self.win.source_file, str(src_path))
        self.assertEqual(self.win.target_files, [str(acb_path)])

        # 2. Kiểm tra danh sách Template đã nạp
        self.assertIn(AUTO_TEMPLATE_OPTION, list(self.win.cmb_source_template["values"]))
        self.assertIn(AUTO_TEMPLATE_OPTION, list(self.win.cmb_target_template["values"]))
        self.win.select_target_template(AUTO_TEMPLATE_OPTION)

        # 3. Nhấn nút Analyze
        wf_res = self.win.analyze()
        self.assertEqual(wf_res.stage, WorkflowStage.PRE_VERIFIED)
        stats = self.win.get_statistics()
        self.assertEqual(stats.total_source_records, 3)
        self.assertEqual(stats.total_matched_records, 3)
        self.assertEqual(stats.total_written_records, 3)
        self.assertEqual(stats.total_exceptions, 0)
        self.assertEqual(stats.total_planned_write_amount, 21_800_000)
        self.assertIn("3 dòng nguồn", self.win.stat_records_var.get())
        self.assertIn("0 ngoại lệ", self.win.stat_exceptions_var.get())
        self.assertIn("21,800,000 đ", self.win.stat_amount_var.get())

        # 4. Nhấn nút Preview
        prev_res = self.win.preview()
        self.assertIsNotNone(prev_res.write_plan)
        summary_text = self.win.txt_summary_view.get("1.0", tk.END)
        self.assertIn("Bảng đối chiếu chi tiết", summary_text)

        # 5. Nhấn nút Execute
        exec_res = self.win.execute()
        self.assertTrue(exec_res.is_success)
        self.assertEqual(exec_res.stage, WorkflowStage.COMPLETED)
        self.assertEqual(len(exec_res.output_files), 1)
        self.assertTrue(Path(exec_res.output_files[0]).exists())
        self.assertIsNotNone(self.win.last_report)
        assert self.win.last_report is not None
        assert self.win.last_report.report_file is not None
        self.assertTrue(Path(self.win.last_report.report_file).exists())

    def test_exception_stats_and_resolution_from_main_window(self) -> None:
        src_path = build_source_xlsx(
            self.tmp_dir / "BẢNG LƯƠNG THÁNG 06.2026.xlsx",
            rows=[
                ("HL001", "NGUYEN VAN A", "NGUYEN VAN A", "111222333", 6_500_000),
            ],
        )
        acb_path = build_target_xlsx(
            self.tmp_dir / "4 MauChiHoLuongQuaTaiKhoanACB.xlsx",
            rows=[
                ("TEN KHAC", "111222333", 5_000_000),        # X6_TARGET_MISMATCH
                ("NGUOI NGHI VIEC", "999888777", 4_000_000),  # X5_TARGET_NOT_IN_SOURCE
            ],
            with_guide=True,
        )

        self.win.set_source_file(src_path)
        self.win.set_target_files([acb_path])

        wf_res = self.win.analyze()
        self.assertEqual(wf_res.stage, WorkflowStage.AWAITING_RESOLUTION)
        stats = self.win.get_statistics()
        self.assertEqual(stats.total_exceptions, 2)
        self.assertEqual(stats.unresolved_exceptions, 2)

        for exc in list(wf_res.exceptions):
            if exc.code == ExceptionCode.X6_TARGET_MISMATCH:
                self.win.resolve_exception(exc.id, ResolutionAction.CONFIRM_WRITE)
            elif exc.code == ExceptionCode.X5_TARGET_NOT_IN_SOURCE:
                self.win.resolve_exception(exc.id, ResolutionAction.SET_ZERO)

        stats_after = self.win.get_statistics()
        self.assertEqual(stats_after.unresolved_exceptions, 0)
        self.assertEqual(stats_after.resolved_exceptions, 2)
        self.assertTrue(stats_after.is_safe)

        final_res = self.win.execute()
        self.assertTrue(final_res.is_success)


if __name__ == "__main__":
    unittest.main()
