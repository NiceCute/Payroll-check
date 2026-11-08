"""
test_mapping_dialog.py - Unit test cho GUI Module 3 (gui/mapping_dialog.py).
"""

from __future__ import annotations

import shutil
import tempfile
import tkinter as tk
import unittest
from pathlib import Path

from core.models import ConfigError
from core.template_store import TemplateStore
from gui.mapping_dialog import MappingDialog
from tests.helpers import build_source_xlsx, build_target_xlsx


class TestMappingDialog(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp_dir = Path(tempfile.mkdtemp(prefix="chl_test_mapping_dlg_"))
        self.templates_dir = self.tmp_dir / "templates"
        self.store = TemplateStore(self.templates_dir)
        self.root = tk.Tk()
        self.root.withdraw()

        self.dlg = MappingDialog(
            self.root,
            template_store=self.store,
            interactive=False,
        )
        self.dlg.withdraw()

    def tearDown(self) -> None:
        try:
            self.dlg.destroy()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_inspect_source_and_target_columns_and_links(self) -> None:
        src_path = build_source_xlsx(self.tmp_dir / "bang_luong.xlsx")
        tgt_path = build_target_xlsx(self.tmp_dir / "acb.xlsx", with_guide=True)

        src_opts = self.dlg.inspect_source_excel(src_path)
        tgt_opts = self.dlg.inspect_target_excel(tgt_path)

        src_by_letter = {o.letter: o.header for o in src_opts}
        self.assertEqual(src_by_letter["B"], "Số thẻ")
        self.assertEqual(src_by_letter["C"], "Họ Và Tên")
        self.assertIn("Chuyển Khoản", src_by_letter["P"])

        tgt_by_letter = {o.letter: o.header for o in tgt_opts}
        self.assertEqual(tgt_by_letter["B"], "Tên NV")
        self.assertEqual(tgt_by_letter["C"], "Số tài khoản")
        self.assertEqual(tgt_by_letter["D"], "Số tiền")

        links = self.dlg.get_source_target_links()
        self.assertEqual(len(links), 5)
        self.assertTrue(all(link.status == "OK" for link in links))

    def test_create_edit_save_reload_and_delete_mapping_json(self) -> None:
        # 1. Tạo mới cấu hình nguồn và lưu vào templates/*.json
        self.dlg.new_source_mapping()
        self.dlg.source_name_var.set("Bảng lương Nhà máy 2")
        self.dlg.source_sheet_patterns_var.set("NM2*")
        self.dlg.set_source_field_mapping("card_no", "B", "Số thẻ")
        self.dlg.set_source_field_mapping("full_name", "C", "Họ Và Tên")
        self.dlg.set_source_field_mapping("account_no", "X", "")
        self.dlg.set_source_field_mapping("amount", "P", "Thực lãnh CK")

        src_json = self.dlg.save_source_mapping()
        self.assertTrue(src_json.exists())
        src_names, _ = self.dlg.reload_saved_mappings()
        self.assertIn("Bảng lương Nhà máy 2", src_names)

        # 2. Tạo mới mẫu ngân hàng đích và lưu vào templates/*.json
        self.dlg.new_target_mapping()
        self.dlg.target_name_var.set("Mẫu Vietcombank")
        self.dlg.target_file_patterns_var.set("*VCB*")
        self.dlg.target_sheet_var.set("VCB_TT")
        self.dlg.set_target_field_mapping("account_no", "C", "Số tài khoản")
        self.dlg.set_target_field_mapping("full_name", "B", "Tên nhân viên")
        self.dlg.set_target_field_mapping("amount", "D", "Số tiền")
        self.dlg.set_target_field_mapping("bank_code", "E", "Mã NH")

        tgt_json = self.dlg.save_target_mapping()
        self.assertTrue(tgt_json.exists())
        _, tgt_names = self.dlg.reload_saved_mappings()
        self.assertIn("Mẫu Vietcombank", tgt_names)

        # 3. Sửa mẫu ngân hàng vừa tạo (đổi cột amount sang F) và lưu lại
        self.dlg.load_target_mapping("Mẫu Vietcombank")
        self.dlg.set_target_field_mapping("amount", "F", "Số tiền chi")
        self.dlg.save_target_mapping()

        reloaded_t = self.dlg.load_target_mapping("Mẫu Vietcombank")
        self.assertEqual(reloaded_t.amount.letter, "F")
        self.assertEqual(reloaded_t.amount.expected_header, "Số tiền chi")

        # 4. Xóa cấu hình vừa tạo
        bak_src = self.dlg.delete_source_mapping("Bảng lương Nhà máy 2")
        bak_tgt = self.dlg.delete_target_mapping("Mẫu Vietcombank")
        self.assertTrue(bak_src.exists())
        self.assertTrue(bak_tgt.exists())
        src_after, tgt_after = self.dlg.reload_saved_mappings()
        self.assertNotIn("Bảng lương Nhà máy 2", src_after)
        self.assertNotIn("Mẫu Vietcombank", tgt_after)

    def test_validate_warns_on_missing_and_duplicate_columns(self) -> None:
        # 1. Thiếu cột bắt buộc trên mẫu đích
        self.dlg.new_target_mapping()
        self.dlg.target_name_var.set("Mẫu lỗi thiếu cột")
        self.dlg.target_sheet_var.set("CNBD")
        self.dlg.set_target_field_mapping("account_no", "C", "Số tài khoản")
        self.dlg.set_target_field_mapping("full_name", "B", "Tên NV")
        # Để trống amount

        val_missing = self.dlg.validate_target_mapping()
        self.assertFalse(val_missing.is_valid)
        self.assertTrue(any("Thiếu mapping bắt buộc" in e for e in val_missing.errors))
        with self.assertRaises(ConfigError):
            self.dlg.save_target_mapping()

        # 2. Trùng cột (account_no và amount cùng trỏ vào cột C)
        self.dlg.set_target_field_mapping("amount", "C", "Số tiền")
        val_dup = self.dlg.validate_target_mapping()
        self.assertFalse(val_dup.is_valid)
        self.assertTrue(any("bị trùng" in e for e in val_dup.errors))
        links = self.dlg.refresh_mapping_preview()
        self.assertTrue(any(l.status == "DUPLICATE" for l in links))
        self.assertIn("CẢNH BÁO", self.dlg.validation_message_var.get())


if __name__ == "__main__":
    unittest.main()
