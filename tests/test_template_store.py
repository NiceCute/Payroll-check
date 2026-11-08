"""Kiểm thử cho core.template_store."""

from __future__ import annotations

import json
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path

from core.models import ColumnRef
from core.source_reader import DEFAULT_SOURCE_PROFILE
from core.target_reader import TEMPLATE_ACB, TEMPLATE_THONG_THUONG
from core.template_store import (
    BACKUP_DIR_NAME,
    MAX_BACKUPS_PER_FILE,
    TemplateStore,
    TemplateStoreError,
    slugify,
)
from tests.helpers import build_acb_xls, build_tt_xls


class TestTemplateStore(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name) / "templates"
        self.store = TemplateStore(self.root)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_slugify(self) -> None:
        self.assertEqual(
            slugify("ACB - Chi hộ lương qua tài khoản ACB"),
            "acb_chi_ho_luong_qua_tai_khoan_acb",
        )
        self.assertEqual(slugify("   !!!   "), "mau")

    def test_ensure_defaults_is_idempotent(self) -> None:
        created = self.store.ensure_defaults()
        self.assertEqual(len(created), 3)
        self.assertEqual(len(self.store.list_sources()), 1)
        self.assertEqual(len(self.store.list_targets()), 2)
        # Gọi lần 2 không tạo thêm
        self.assertEqual(self.store.ensure_defaults(), [])

    def test_save_get_overwrite_backup_and_delete(self) -> None:
        path = self.store.save_target(TEMPLATE_ACB)
        self.assertTrue(path.exists())
        loaded = self.store.get_target(TEMPLATE_ACB.name)
        self.assertEqual(loaded, TEMPLATE_ACB)

        # Không cho ghi đè nếu overwrite=False
        with self.assertRaises(TemplateStoreError):
            self.store.save_target(TEMPLATE_ACB, overwrite=False)

        # Cho phép ghi đè và sinh bản backup
        updated = replace(TEMPLATE_ACB, description="Đã cập nhật")
        self.store.save_target(updated, overwrite=True)
        self.assertEqual(self.store.get_target(TEMPLATE_ACB.name).description, "Đã cập nhật")
        backups = list((self.root / BACKUP_DIR_NAME).glob("*.json"))
        self.assertEqual(len(backups), 1)

        # Xóa cấu hình -> vẫn còn bản sao lưu
        del_backup = self.store.delete("target", TEMPLATE_ACB.name)
        self.assertTrue(del_backup.exists())
        self.assertEqual(self.store.list_targets(), [])

    def test_rename_in_place_via_previous_name(self) -> None:
        p1 = self.store.save_target(TEMPLATE_ACB)
        renamed = replace(TEMPLATE_ACB, name="Mẫu ACB Mới")
        p2 = self.store.save_target(renamed, previous_name=TEMPLATE_ACB.name)
        self.assertEqual(p1, p2)
        self.assertEqual([s.name for s in self.store.list_targets()], ["Mẫu ACB Mới"])

    def test_rename_collision_with_overwrite_removes_old_file(self) -> None:
        p_acb = self.store.save_target(TEMPLATE_ACB)
        p_tt = self.store.save_target(TEMPLATE_THONG_THUONG)
        # Đổi tên Thông thường sang trùng tên ACB với overwrite=True
        merged = replace(TEMPLATE_THONG_THUONG, name=TEMPLATE_ACB.name)
        saved = self.store.save_target(merged, overwrite=True, previous_name=TEMPLATE_THONG_THUONG.name)
        self.assertEqual(saved, p_acb)
        self.assertFalse(p_tt.exists())
        self.assertEqual(len(self.store.list_targets()), 1)

    def test_corrupt_and_unsupported_json_isolated_in_load_errors(self) -> None:
        self.store.save_source(DEFAULT_SOURCE_PROFILE)
        (self.root / "bad_syntax.json").write_text("{not valid json", encoding="utf-8")
        (self.root / "bad_version.json").write_text(
            json.dumps({"type": "target", "version": 99, "data": {}}), encoding="utf-8"
        )
        (self.root / "bad_kind.json").write_text(
            json.dumps({"type": "unknown", "version": 1, "data": {}}), encoding="utf-8"
        )
        sources = self.store.list_sources()
        self.assertEqual(len(sources), 1)
        self.assertEqual(len(self.store.load_errors), 3)

    def test_duplicate_names_in_two_files_quarantined(self) -> None:
        self.store.save_target(TEMPLATE_ACB)
        dup_path = self.root / "target_duplicate_manual.json"
        dup_path.write_text(
            json.dumps({"type": "target", "version": 1, "data": TEMPLATE_ACB.to_dict()}),
            encoding="utf-8",
        )
        targets = self.store.list_targets()
        self.assertEqual(targets, [])
        self.assertEqual(len(self.store.load_errors), 1)
        self.assertIn("nhiều file cùng tên", self.store.load_errors[0].message)

    def test_backup_rotation_keeps_at_most_max(self) -> None:
        self.store.save_target(TEMPLATE_ACB)
        for i in range(MAX_BACKUPS_PER_FILE + 4):
            self.store.save_target(replace(TEMPLATE_ACB, description=f"v{i}"), overwrite=True)
        backups = list((self.root / BACKUP_DIR_NAME).glob("*.json"))
        self.assertLessEqual(len(backups), MAX_BACKUPS_PER_FILE)

    def test_match_targets_identifies_acb_and_tt(self) -> None:
        self.store.ensure_defaults()
        work = Path(self._tmp.name)
        acb_file = build_acb_xls(work / "4_MauChiHoLuongQuaTaiKhoanACB-Xuong_L2.xls")
        tt_file = build_tt_xls(work / "5_MauChiHoLuongThongThuong-Xuong_L2.xls")

        m_acb = self.store.match_targets(acb_file)
        self.assertEqual(len(m_acb), 1)
        self.assertEqual(m_acb[0].template.name, TEMPLATE_ACB.name)
        self.assertTrue(m_acb[0].by_file_name)
        self.assertTrue(m_acb[0].by_headers)

        m_tt = self.store.match_targets(tt_file)
        self.assertEqual(len(m_tt), 1)
        self.assertEqual(m_tt[0].template.name, TEMPLATE_THONG_THUONG.name)
        self.assertTrue(m_tt[0].by_file_name)

        # File đổi tên bất kỳ nhưng cấu trúc cột là ACB -> vẫn nhận diện được theo tiêu đề cột
        renamed = build_acb_xls(work / "danh_sach_thang_6.xls")
        m_renamed = self.store.match_targets(renamed)
        self.assertEqual(len(m_renamed), 1)
        self.assertEqual(m_renamed[0].template.name, TEMPLATE_ACB.name)
        self.assertFalse(m_renamed[0].by_file_name)
        self.assertTrue(m_renamed[0].by_headers)

    def test_save_rejects_invalid_template(self) -> None:
        bad = replace(TEMPLATE_ACB, amount=ColumnRef("C", "Số tiền"))
        with self.assertRaises( Exception ):
            self.store.save_target(bad)


if __name__ == "__main__":
    unittest.main()
