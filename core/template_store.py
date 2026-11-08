"""
template_store.py - Lưu / đọc / sửa cấu hình mapping (JSON) cho file lương và các mẫu ngân hàng.

Thư mục templates/:
    source_<ten>.json   {"type": "source", "version": 1, "data": {...SourceProfile...}}
    target_<ten>.json   {"type": "target", "version": 1, "data": {...TargetTemplate...}}

An toàn:
- Ghi nguyên tử (file tạm -> thay thế) để không bao giờ để lại file JSON hỏng giữa chừng.
- Trước khi ghi đè / xóa: lưu bản cũ vào templates/_backup/ kèm thời gian.
- File JSON hỏng không làm sập chương trình: được liệt kê trong `load_errors`.
"""

from __future__ import annotations

import fnmatch
import json
import logging
import os
import re
import shutil
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Generic, Literal, Optional, TypeVar, Union

from .excel_reader import PathLike, SheetSelectionError, WorkbookReader, check_column_headers
from .models import ConfigError, SourceProfile, TargetTemplate, ToolError
from .normalize import remove_diacritics
from .source_reader import DEFAULT_SOURCE_PROFILE
from .target_reader import DEFAULT_TARGET_TEMPLATES, select_target_sheet

log = logging.getLogger(__name__)

SCHEMA_VERSION: int = 1
BACKUP_DIR_NAME: str = "_backup"
MAX_BACKUPS_PER_FILE: int = 20

Kind = Literal["source", "target"]
_PREFIX: dict[Kind, str] = {"source": "source_", "target": "target_"}
T = TypeVar("T", SourceProfile, TargetTemplate)


class TemplateStoreError(ToolError):
    """Lỗi khi lưu / đọc cấu hình mapping."""


@dataclass(frozen=True)
class StoredItem(Generic[T]):
    """1 cấu hình đã lưu kèm đường dẫn file."""

    path: Path
    item: T

    @property
    def name(self) -> str:
        return self.item.name


@dataclass
class LoadError:
    path: Path
    message: str


@dataclass
class TemplateMatch:
    """Kết quả nhận diện mẫu cho 1 file ngân hàng."""

    template: TargetTemplate
    by_file_name: bool          # khớp theo tên file
    by_headers: bool            # tiêu đề cột khớp mapping
    sheet_name: Optional[str] = None


@dataclass
class _Loaded:
    sources: list[StoredItem[SourceProfile]] = field(default_factory=list)
    targets: list[StoredItem[TargetTemplate]] = field(default_factory=list)
    errors: list[LoadError] = field(default_factory=list)


def slugify(name: str) -> str:
    """Tên cấu hình -> tên file an toàn: bỏ dấu, chữ thường, chỉ a-z0-9_."""
    text = remove_diacritics(name).lower()
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return (text or "mau")[:60]


class TemplateStore:
    """Kho cấu hình mapping trên đĩa."""

    def __init__(self, root: PathLike) -> None:
        self.root: Path = Path(root)
        self.load_errors: list[LoadError] = []

    # ------------------------------------------------------------------ đọc
    def _scan(self) -> _Loaded:
        loaded = _Loaded()
        if not self.root.exists():
            return loaded
        for path in sorted(self.root.glob("*.json")):
            try:
                kind, item = _read_file(path)
            except TemplateStoreError as exc:
                loaded.errors.append(LoadError(path, exc.user_message))
                log.warning("Bỏ qua file cấu hình lỗi %s: %s", path, exc)
                continue
            if isinstance(item, SourceProfile):
                loaded.sources.append(StoredItem(path, item))
            else:
                loaded.targets.append(StoredItem(path, item))
        _check_duplicate_names(loaded)
        self.load_errors = loaded.errors
        return loaded

    def list_sources(self) -> list[StoredItem[SourceProfile]]:
        return self._scan().sources

    def list_targets(self) -> list[StoredItem[TargetTemplate]]:
        return self._scan().targets

    def get_source(self, name: str) -> SourceProfile:
        return _find(self.list_sources(), name, "cấu hình file lương").item

    def get_target(self, name: str) -> TargetTemplate:
        return _find(self.list_targets(), name, "mẫu ngân hàng").item

    # ------------------------------------------------------------------ ghi
    def save_source(self, profile: SourceProfile, *, overwrite: bool = False,
                    previous_name: Optional[str] = None) -> Path:
        profile.ensure_valid()
        return self._save("source", profile.name, profile.to_dict(), overwrite, previous_name)

    def save_target(self, template: TargetTemplate, *, overwrite: bool = False,
                    previous_name: Optional[str] = None) -> Path:
        template.ensure_valid()
        return self._save("target", template.name, template.to_dict(), overwrite, previous_name)

    def delete(self, kind: Kind, name: str) -> Path:
        items: list[Any] = self.list_sources() if kind == "source" else self.list_targets()
        stored = _find(items, name, "cấu hình")
        backup = self._backup(stored.path)
        stored.path.unlink()
        log.info("Đã xóa cấu hình '%s' (bản sao: %s)", name, backup)
        return backup

    def _save(self, kind: Kind, name: str, data: dict[str, Any], overwrite: bool,
              previous_name: Optional[str]) -> Path:
        """Lưu cấu hình. previous_name: tên cũ khi người dùng đổi tên lúc sửa."""
        items: list[StoredItem[Any]] = list(self.list_sources()) if kind == "source" else list(self.list_targets())
        key = name.strip().casefold()
        same_name = next((s for s in items if s.name.strip().casefold() == key), None)
        old = next((s for s in items if previous_name is not None and s.name == previous_name), None)
        is_self = same_name is not None and old is not None and same_name.path == old.path

        if same_name is not None and not overwrite and not is_self:
            raise TemplateStoreError(f"Đã có cấu hình tên '{name}'. Hãy đặt tên khác hoặc chọn ghi đè.")

        if same_name is not None:
            path: Path = same_name.path
        elif old is not None:
            path = old.path             # đổi tên -> giữ file cũ, cập nhật nội dung
        else:
            path = self._new_path(kind, name)

        payload = {"type": kind, "version": SCHEMA_VERSION, "saved_at": _now_iso(), "data": data}
        self.root.mkdir(parents=True, exist_ok=True)
        if path.exists():
            self._backup(path)
        _atomic_write_json(path, payload)
        # Kiểm tra lại: đọc file vừa ghi phải ra đúng nội dung
        _, reread = _read_file(path)
        if reread.to_dict() != data:
            raise TemplateStoreError(f"Lưu cấu hình '{name}' không thành công (kiểm tra lại không khớp).")
        # Đổi tên trùng sang 1 cấu hình khác (đã chọn ghi đè) -> bỏ file của tên cũ để không còn 2 bản
        if old is not None and old.path != path and old.path.exists():
            self._backup(old.path)
            old.path.unlink()
        log.info("Đã lưu cấu hình '%s' -> %s", name, path)
        return path

    def _new_path(self, kind: Kind, name: str) -> Path:
        base = _PREFIX[kind] + slugify(name)
        path = self.root / f"{base}.json"
        i = 2
        while path.exists():
            path = self.root / f"{base}_{i}.json"
            i += 1
        return path

    def _backup(self, path: Path) -> Path:
        backup_dir = self.root / BACKUP_DIR_NAME
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        target = backup_dir / f"{path.stem}.{stamp}.json"
        try:
            shutil.copy2(path, target)
        except OSError as exc:
            raise TemplateStoreError(f"Không tạo được bản sao lưu cho '{path.name}'.", detail=repr(exc)) from exc
        old = sorted(backup_dir.glob(f"{path.stem}.*.json"))
        for extra in old[:-MAX_BACKUPS_PER_FILE]:
            try:
                extra.unlink()
            except OSError:  # pragma: no cover - dọn dẹp không được làm hỏng luồng chính
                log.debug("Không xóa được bản sao cũ %s", extra)
        return target

    # -------------------------------------------------------------- mặc định
    def ensure_defaults(self) -> list[Path]:
        """Tạo cấu hình mặc định nếu thư mục chưa có cấu hình nào (không ghi đè cấu hình người dùng)."""
        created: list[Path] = []
        loaded = self._scan()
        if not loaded.sources:
            created.append(self.save_source(DEFAULT_SOURCE_PROFILE))
        if not loaded.targets:
            for template in DEFAULT_TARGET_TEMPLATES:
                created.append(self.save_target(template))
        return created

    # ------------------------------------------------------------- nhận diện
    def match_targets(self, file_path: PathLike) -> list[TemplateMatch]:
        """Tìm mẫu phù hợp với 1 file ngân hàng.

        Chỉ trả về mẫu có tiêu đề cột khớp mapping. Sắp xếp: khớp cả tên file lên trước.
        Nhiều kết quả -> giao diện cho người dùng chọn (không tự đoán).
        """
        return match_templates(file_path, [s.item for s in self.list_targets()])


# =============================================================================
# Hàm cấp module
# =============================================================================

def match_templates(file_path: PathLike, templates: list[TargetTemplate]) -> list[TemplateMatch]:
    file_name = Path(file_path).name.casefold()
    results: list[TemplateMatch] = []
    with WorkbookReader(file_path) as wb:
        for template in templates:
            by_name = any(fnmatch.fnmatchcase(file_name, p.casefold()) for p in template.file_patterns)
            try:
                sheet, _ = select_target_sheet(wb, template)
            except SheetSelectionError:
                continue
            if check_column_headers(wb.read_sheet(sheet), template.header_rows, template.columns()):
                continue
            results.append(TemplateMatch(template, by_file_name=by_name, by_headers=True, sheet_name=sheet))
    results.sort(key=lambda m: (not m.by_file_name, m.template.name))
    return results


def _read_file(path: Path) -> tuple[Kind, Union[SourceProfile, TargetTemplate]]:
    try:
        text = path.read_text(encoding="utf-8-sig")
    except OSError as exc:
        raise TemplateStoreError(f"Không đọc được file cấu hình '{path.name}'.", detail=repr(exc)) from exc
    try:
        payload = json.loads(text)
    except json.JSONDecodeError as exc:
        raise TemplateStoreError(
            f"File cấu hình '{path.name}' bị hỏng (sai định dạng JSON, dòng {exc.lineno})."
        ) from exc
    if not isinstance(payload, dict):
        raise TemplateStoreError(f"File cấu hình '{path.name}' không đúng cấu trúc.")
    version = payload.get("version")
    if not isinstance(version, int) or version > SCHEMA_VERSION or version < 1:
        raise TemplateStoreError(
            f"File cấu hình '{path.name}' có phiên bản {version!r} không được hỗ trợ "
            f"(phiên bản hỗ trợ: {SCHEMA_VERSION})."
        )
    kind = payload.get("type")
    data = payload.get("data")
    if not isinstance(data, dict):
        raise TemplateStoreError(f"File cấu hình '{path.name}' thiếu phần 'data'.")
    try:
        if kind == "source":
            return "source", SourceProfile.from_dict(data)
        if kind == "target":
            return "target", TargetTemplate.from_dict(data)
    except ConfigError as exc:
        raise TemplateStoreError(f"File '{path.name}': {exc.user_message}", detail=exc.detail) from exc
    raise TemplateStoreError(f"File cấu hình '{path.name}' có loại không hợp lệ: {kind!r}.")


def _atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    fd, tmp_name = tempfile.mkstemp(prefix=f".{path.stem}.", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=2)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp_name, path)
    except OSError as exc:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass
        raise TemplateStoreError(f"Không ghi được file cấu hình '{path.name}'.", detail=repr(exc)) from exc


def _find(items: list[StoredItem[T]], name: str, label: str) -> StoredItem[T]:
    for s in items:
        if s.name == name:
            return s
    raise TemplateStoreError(f"Không tìm thấy {label} '{name}'.")


def _check_duplicate_names(loaded: _Loaded) -> None:
    """2 file cùng tên cấu hình -> không biết dùng file nào -> loại cả 2, báo lỗi."""
    for attr in ("sources", "targets"):
        items: list[Any] = getattr(loaded, attr)
        seen: dict[str, list[Any]] = {}
        for s in items:
            seen.setdefault(s.name.strip().casefold(), []).append(s)
        dup_keys = {k for k, v in seen.items() if len(v) > 1}
        if not dup_keys:
            continue
        for key in dup_keys:
            files = ", ".join(s.path.name for s in seen[key])
            loaded.errors.append(LoadError(
                seen[key][0].path, f"Có nhiều file cùng tên cấu hình '{seen[key][0].name}': {files}. Hãy xóa bớt."
            ))
        setattr(loaded, attr, [s for s in items if s.name.strip().casefold() not in dup_keys])


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")
