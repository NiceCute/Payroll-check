"""
excel_writer.py - Ghi kết quả đối chiếu ra file ngân hàng mới, giữ nguyên 100% định dạng.

Tuân thủ kiến trúc V1 & quy tắc an toàn kế toán:
1. Kiểm tra chốt an toàn (ValidationResult / Safety Layer) trước khi thực hiện.
2. Tự động sao lưu (backup) file gốc vào thư mục backup kèm dấu thời gian (R8).
3. Tuyệt đối KHÔNG ghi đè lên file gốc: copy nguyên bản file gốc sang file output mới,
   kiểm tra SHA-256 của file gốc trước và sau khi chạy (R8).
4. Chỉ cập nhật đúng các ô trong WritePlan (R6).
5. Không phá khóa sheet bị bảo vệ (K10), không ghi đè ô chứa công thức hoặc ô gộp (K5).
6. Giữ nguyên 100% layout, format, border, màu sắc, merge cell, công thức, chiều rộng cột,
   chiều cao dòng, sheet ẩn/khóa cho cả .xls và .xlsx thông qua Excel COM (K5, K9).

Module CHỈ THỰC HIỆN GHI FILE (không xử lý GUI, không xử lý workflow, không tạo báo cáo).
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import logging
import os
import shutil
import stat
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Literal, Mapping, Optional, Sequence, Union

from .excel_reader import FileFormat, PathLike, detect_format
from .models import (
    CellUpdate,
    ConfigError,
    ExcelWriteError,
    SafetyError,
    SourceProfile,
    TargetReadResult,
    TargetTemplate,
    TargetWorkbook,
    TargetWriteOutput,
    TargetWritePlan,
    TemplateProfile,
    ValidationResult,
    WriteExecutionResult,
    WritePlan,
)
from .safety import check_file_locks, validate_write_plan

log = logging.getLogger(__name__)

WriterEngine = Literal["auto", "com", "win32com", "powershell_com", "openpyxl"]

DEFAULT_BACKUP_DIR: str = "data/backup"
DEFAULT_OUTPUT_DIR: str = "data/output"
DEFAULT_OUTPUT_SUFFIX: str = "_KetQua"
COM_TIMEOUT_SECONDS: int = 120

__all__ = [
    "DEFAULT_BACKUP_DIR",
    "DEFAULT_OUTPUT_DIR",
    "DEFAULT_OUTPUT_SUFFIX",
    "ExcelWriteError",
    "ExcelWriter",
    "TargetWorkbook",
    "TargetWriteOutput",
    "TemplateProfile",
    "WriteExecutionResult",
    "WriterEngine",
    "backup_file",
    "build_output_path",
    "execute_write_plan",
    "is_excel_com_available",
    "write_excel",
    "write_target_file",
]


# =============================================================================
# 1. Giao diện chính (Function & Class API)
# =============================================================================

class ExcelWriter:
    """Bộ ghi file kết quả ngân hàng theo WritePlan."""

    def __init__(
        self,
        *,
        output_dir: PathLike = DEFAULT_OUTPUT_DIR,
        backup_dir: PathLike = DEFAULT_BACKUP_DIR,
        output_suffix: str = DEFAULT_OUTPUT_SUFFIX,
        overwrite_output: bool = False,
        backup_source_file: bool = True,
        engine: WriterEngine = "auto",
    ) -> None:
        self.output_dir = Path(output_dir)
        self.backup_dir = Path(backup_dir)
        self.output_suffix = output_suffix
        self.overwrite_output = overwrite_output
        self.backup_source_file = backup_source_file
        self.engine: WriterEngine = engine

    def write(
        self,
        write_plan: WritePlan,
        validation_result: Optional[ValidationResult] = None,
        template_profile: Optional[
            Union[TemplateProfile, Sequence[TargetTemplate], Mapping[str, TargetTemplate]]
        ] = None,
        target_workbook: Optional[
            Union[TargetWorkbook, Sequence[TargetWorkbook], Mapping[str, PathLike]]
        ] = None,
        *,
        output_paths: Optional[Union[PathLike, Sequence[PathLike], Mapping[str, PathLike]]] = None,
    ) -> WriteExecutionResult:
        return execute_write_plan(
            write_plan=write_plan,
            validation_result=validation_result,
            template_profile=template_profile,
            target_workbook=target_workbook,
            output_dir=self.output_dir,
            backup_dir=self.backup_dir,
            output_paths=output_paths,
            output_suffix=self.output_suffix,
            overwrite_output=self.overwrite_output,
            backup_source_file=self.backup_source_file,
            engine=self.engine,
        )


def execute_write_plan(
    write_plan: WritePlan,
    validation_result: Optional[ValidationResult] = None,
    template_profile: Optional[
        Union[TemplateProfile, Sequence[TargetTemplate], Mapping[str, TargetTemplate]]
    ] = None,
    target_workbook: Optional[
        Union[TargetWorkbook, Sequence[TargetWorkbook], Mapping[str, PathLike]]
    ] = None,
    *,
    output_dir: PathLike = DEFAULT_OUTPUT_DIR,
    backup_dir: PathLike = DEFAULT_BACKUP_DIR,
    output_paths: Optional[Union[PathLike, Sequence[PathLike], Mapping[str, PathLike]]] = None,
    output_suffix: str = DEFAULT_OUTPUT_SUFFIX,
    overwrite_output: bool = False,
    backup_source_file: bool = True,
    engine: WriterEngine = "auto",
) -> WriteExecutionResult:
    """Thực thi `WritePlan`: kiểm tra an toàn -> backup file gốc -> copy sang file mới -> cập nhật đúng các ô.

    Args:
        write_plan        : Kế hoạch ghi từ `reconciler`.
        validation_result : Kết quả kiểm tra từ `safety` (nếu None sẽ tự động kiểm tra).
        template_profile  : Cấu hình mẫu ngân hàng (để kiểm tra chéo cột Số tiền & phạm vi dòng).
        target_workbook   : Đường dẫn hoặc `TargetReadResult` của file ngân hàng đầu vào (nếu cần ánh xạ lại).
        output_dir        : Thư mục chứa file kết quả mới.
        backup_dir        : Thư mục chứa bản sao lưu tự động của file đầu vào.
        output_paths      : Đường dẫn file output cụ thể (tùy chọn).
        output_suffix     : Hậu tố đặt tên file kết quả mới (mặc định `_KetQua`).
        overwrite_output  : Cho phép ghi đè lên file output cũ (KHÔNG BAO GIỜ ghi đè file gốc).
        backup_source_file: Có sao lưu cả file bảng lương nguồn (nếu tồn tại trên đĩa) hay không.
        engine            : Cơ chế ghi (`"auto"`, `"com"`, `"win32com"`, `"powershell_com"`, `"openpyxl"`).

    Returns:
        `WriteExecutionResult` chứa đường dẫn các bản backup, các file output mới và thống kê.
    """
    # 1. Nếu người gọi truyền ValidationResult chưa đạt -> chặn ngay
    if validation_result is not None and not validation_result.is_safe:
        validation_result.ensure_safe()

    # 2. Luôn chạy kiểm tra an toàn đầy đủ trên WritePlan hiện tại
    fresh_validation = validate_write_plan(write_plan, check_disk=False)
    if not fresh_validation.is_safe:
        fresh_validation.ensure_safe()
    effective_validation = validation_result if validation_result is not None else fresh_validation

    # 3. Chuẩn hóa ánh xạ TemplateProfile & TargetWorkbook & OutputPaths
    tpl_map = _resolve_template_map(write_plan, template_profile)
    workbook_map = _resolve_target_workbook_map(write_plan, target_workbook)
    out_map = _resolve_output_path_map(
        write_plan=write_plan,
        workbook_map=workbook_map,
        output_dir=Path(output_dir),
        output_paths=output_paths,
        output_suffix=output_suffix,
        overwrite_output=overwrite_output,
    )

    # 4. Kiểm tra chéo với TemplateProfile (nếu có) & kiểm tra file gốc không trùng file output (R8)
    all_original_paths: set[str] = set()
    if write_plan.source_file:
        all_original_paths.add(_norm_path_key(write_plan.source_file))
    for tp in write_plan.target_plans:
        orig_target = workbook_map[tp.file_path]
        all_original_paths.add(_norm_path_key(str(orig_target)))
        _verify_against_template(tp, tpl_map.get(tp.file_path))

    seen_outputs: set[str] = set()
    for tp in write_plan.target_plans:
        out_path = out_map[tp.file_path]
        norm_out = _norm_path_key(str(out_path))
        if norm_out in all_original_paths:
            raise SafetyError(
                f"Vi phạm quy tắc an toàn R8: đường dẫn file kết quả '{out_path}' "
                "trùng với đường dẫn file gốc đầu vào."
            )
        if norm_out in seen_outputs:
            raise ExcelWriteError(
                f"Đường dẫn file kết quả '{out_path}' bị dùng chung cho nhiều file ngân hàng."
            )
        seen_outputs.add(norm_out)

    # 5. Kiểm tra các file đích trên đĩa có tồn tại và không bị mở khóa trong Excel (K9)
    for tp in write_plan.target_plans:
        src_target = workbook_map[tp.file_path]
        if not src_target.exists() or not src_target.is_file():
            raise ExcelWriteError(f"Không tìm thấy file ngân hàng đầu vào: '{src_target}'.")
        lock_findings = check_file_locks([src_target])
        if lock_findings:
            raise ExcelWriteError(lock_findings[0].message)
        out_target = out_map[tp.file_path]
        if out_target.exists():
            out_locks = check_file_locks([out_target])
            if out_locks:
                raise ExcelWriteError(out_locks[0].message)

    # 6. Sao lưu file bảng lương nguồn (nếu có trên đĩa)
    backup_root = Path(backup_dir)
    source_backup: Optional[str] = None
    if backup_source_file and write_plan.source_file:
        src_salary_path = Path(write_plan.source_file)
        if src_salary_path.exists() and src_salary_path.is_file():
            source_backup = str(backup_file(src_salary_path, backup_root))

    # 7. Thực hiện backup -> copy -> cập nhật ô cho từng file đích
    exec_result = WriteExecutionResult(
        write_plan=write_plan,
        validation=effective_validation,
        source_backup_file=source_backup,
    )
    created_outputs: list[Path] = []
    try:
        for tp in write_plan.target_plans:
            if not tp.updates:
                continue
            target_out = write_target_file(
                target_plan=tp,
                source_target_path=workbook_map[tp.file_path],
                output_path=out_map[tp.file_path],
                backup_dir=backup_root,
                template=tpl_map.get(tp.file_path),
                overwrite_output=overwrite_output,
                engine=engine,
            )
            created_outputs.append(Path(target_out.output_file))
            exec_result.target_outputs.append(target_out)
    except Exception:
        # Nếu bất kỳ file nào trong lô bị lỗi, dọn dẹp các file output vừa tạo để tránh trạng thái dở dang
        for p in created_outputs:
            _safe_unlink(p)
        raise

    return exec_result


# Bí danh hàm ngắn gọn
write_excel = execute_write_plan


# =============================================================================
# 2. Ghi 1 file ngân hàng đơn lẻ
# =============================================================================

def write_target_file(
    target_plan: TargetWritePlan,
    source_target_path: PathLike,
    output_path: PathLike,
    *,
    backup_dir: PathLike = DEFAULT_BACKUP_DIR,
    template: Optional[TargetTemplate] = None,
    overwrite_output: bool = False,
    engine: WriterEngine = "auto",
) -> TargetWriteOutput:
    """Sao lưu `source_target_path`, copy sang `output_path` và chỉ ghi các ô trong `target_plan.updates`."""
    src_path = Path(source_target_path)
    out_path = Path(output_path)

    if not src_path.exists() or not src_path.is_file():
        raise ExcelWriteError(f"Không tìm thấy file ngân hàng gốc: '{src_path}'.")

    if _norm_path_key(str(src_path)) == _norm_path_key(str(out_path)):
        raise SafetyError(
            f"Vi phạm quy tắc R8: không được phép ghi đè trực tiếp lên file gốc '{src_path}'."
        )

    if out_path.exists() and not overwrite_output:
        raise ExcelWriteError(
            f"File kết quả '{out_path}' đã tồn tại. Hãy chọn tên khác hoặc bật cho phép ghi đè file kết quả."
        )

    _verify_against_template(target_plan, template)

    # Xác thực định dạng thật của file đầu vào (.xls hay .xlsx)
    actual_format = detect_format(src_path)

    # Nếu đuôi file output khác đuôi file gốc -> chặn để tránh hỏng định dạng Excel
    if out_path.suffix.lower() != src_path.suffix.lower():
        raise ExcelWriteError(
            f"Đuôi file kết quả ('{out_path.suffix}') phải giữ nguyên giống file gốc ('{src_path.suffix}')."
        )

    # Tính SHA-256 của file gốc trước khi thao tác
    orig_size = src_path.stat().st_size
    orig_hash = _sha256_file(src_path)

    # Bước 1: Tạo bản sao lưu tự động của file gốc vào backup_dir
    bak_path = backup_file(src_path, backup_dir, expected_sha256=orig_hash)

    # Bước 2: Copy nguyên bản file gốc sang đường dẫn output mới (và gỡ cờ Read-Only nếu file gốc chỉ đọc)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        if out_path.exists():
            _make_writable(out_path)
        shutil.copy2(src_path, out_path)
        _make_writable(out_path)
    except OSError as exc:
        raise ExcelWriteError(
            f"Không thể tạo file kết quả mới tại '{out_path}'.", detail=repr(exc)
        ) from exc

    # Bước 3: Mở bản copy tại out_path và chỉ cập nhật các ô trong target_plan.updates
    try:
        _apply_updates_to_workbook(
            workbook_path=out_path,
            file_format=actual_format,
            sheet_name=target_plan.sheet_name,
            updates=target_plan.updates,
            engine=engine,
        )
    except Exception:
        _safe_unlink(out_path)
        raise

    # Bước 4: Xác nhận file gốc hoàn toàn không bị thay đổi dù chỉ 1 byte
    if src_path.stat().st_size != orig_size or _sha256_file(src_path) != orig_hash:
        _safe_unlink(out_path)
        raise SafetyError(
            f"Phát hiện file gốc '{src_path.name}' bị thay đổi ngoài ý muốn! Đã hủy file kết quả."
        )

    if not out_path.exists() or out_path.stat().st_size == 0:
        raise ExcelWriteError(f"File kết quả '{out_path}' không được tạo đúng cách.")

    log.info(
        "Đã xuất file ngân hàng '%s' (%d ô cập nhật, backup: '%s')",
        out_path,
        len(target_plan.updates),
        bak_path,
    )
    return TargetWriteOutput(
        source_target_file=str(src_path),
        backup_file=str(bak_path),
        output_file=str(out_path),
        sheet_name=target_plan.sheet_name,
        file_format=actual_format.value,
        updated_cells=len(target_plan.updates),
        changed_cells=len(target_plan.changed_updates),
        total_new_amount=target_plan.total_new_amount,
        original_sha256=orig_hash,
    )


# =============================================================================
# 3. Sao lưu (Backup) & Đặt tên File Output (R8)
# =============================================================================

def backup_file(
    source_path: PathLike,
    backup_dir: PathLike = DEFAULT_BACKUP_DIR,
    *,
    expected_sha256: Optional[str] = None,
) -> Path:
    """Tạo bản sao lưu nguyên vẹn của `source_path` trong `backup_dir` kèm dấu thời gian."""
    src = Path(source_path)
    if not src.exists() or not src.is_file():
        raise ExcelWriteError(f"Không thể sao lưu vì file không tồn tại: '{src}'.")

    bdir = Path(backup_dir)
    bdir.mkdir(parents=True, exist_ok=True)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    candidate = bdir / f"{src.stem}_{stamp}{src.suffix}"
    seq = 2
    while candidate.exists():
        candidate = bdir / f"{src.stem}_{stamp}_{seq}{src.suffix}"
        seq += 1

    try:
        shutil.copy2(src, candidate)
        _make_writable(candidate)
    except OSError as exc:
        raise ExcelWriteError(
            f"Không tạo được bản sao lưu cho '{src.name}' tại '{bdir}'.", detail=repr(exc)
        ) from exc

    src_hash = expected_sha256 or _sha256_file(src)
    bak_hash = _sha256_file(candidate)
    if bak_hash != src_hash or candidate.stat().st_size != src.stat().st_size:
        _safe_unlink(candidate)
        raise ExcelWriteError(
            f"Kiểm tra toàn vẹn bản sao lưu '{candidate.name}' thất bại (SHA-256 không khớp)."
        )

    log.info("Đã tạo bản sao lưu an toàn: '%s' -> '%s'", src.name, candidate)
    return candidate


def build_output_path(
    original_path: PathLike,
    output_dir: PathLike = DEFAULT_OUTPUT_DIR,
    *,
    suffix: str = DEFAULT_OUTPUT_SUFFIX,
    avoid_collision: bool = True,
) -> Path:
    """Sinh đường dẫn file output mới (không bao giờ trùng với `original_path`)."""
    orig = Path(original_path)
    out_dir = Path(output_dir)
    clean_suffix = suffix if suffix else DEFAULT_OUTPUT_SUFFIX
    base_name = f"{orig.stem}{clean_suffix}{orig.suffix}"
    candidate = out_dir / base_name

    # Nếu người dùng trỏ output_dir trùng thư mục gốc và suffix rỗng -> bắt buộc thêm hậu tố
    if _norm_path_key(str(candidate)) == _norm_path_key(str(orig)):
        candidate = out_dir / f"{orig.stem}{DEFAULT_OUTPUT_SUFFIX}{orig.suffix}"

    if avoid_collision and candidate.exists():
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        candidate = out_dir / f"{orig.stem}{clean_suffix}_{stamp}{orig.suffix}"
        seq = 2
        while candidate.exists():
            candidate = out_dir / f"{orig.stem}{clean_suffix}_{stamp}_{seq}{orig.suffix}"
            seq += 1

    return candidate


# =============================================================================
# 4. Các Engine ghi dữ liệu ô (Win32COM / PowerShell COM / OpenPyXL)
# =============================================================================

def _has_win32com() -> bool:
    return (
        importlib.util.find_spec("pythoncom") is not None
        and importlib.util.find_spec("win32com.client") is not None
    )


def is_excel_com_available() -> bool:
    """Kiểm tra máy hiện tại có dùng được Microsoft Excel COM hay không."""
    if os.name != "nt":
        return False
    if _has_win32com():
        return True
    try:
        proc = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                "$xl = New-Object -ComObject Excel.Application; $xl.Quit(); "
                "[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)",
            ],
            capture_output=True,
            timeout=30,
            check=False,
        )
        return proc.returncode == 0
    except OSError:
        return False


def _apply_updates_to_workbook(
    *,
    workbook_path: Path,
    file_format: FileFormat,
    sheet_name: str,
    updates: Sequence[CellUpdate],
    engine: WriterEngine,
) -> None:
    if not updates:
        return

    chosen = _select_engine(file_format, engine)
    if chosen == "win32com":
        _write_via_win32com(workbook_path, sheet_name, updates)
    elif chosen == "powershell_com":
        _write_via_powershell_com(workbook_path, sheet_name, updates)
    elif chosen == "openpyxl":
        _write_via_openpyxl(workbook_path, file_format, sheet_name, updates)
    else:
        raise ConfigError(f"Engine ghi Excel không hợp lệ: {engine!r}")


def _select_engine(file_format: FileFormat, engine: WriterEngine) -> str:
    if engine == "win32com":
        if not _has_win32com():
            raise ExcelWriteError("Không tìm thấy thư viện pywin32 (win32com) trên máy.")
        return "win32com"

    if engine == "powershell_com":
        if os.name != "nt":
            raise ExcelWriteError("Excel COM chỉ hỗ trợ trên hệ điều hành Windows.")
        return "powershell_com"

    if engine == "openpyxl":
        if file_format != FileFormat.XLSX:
            raise ExcelWriteError(
                "Định dạng .xls bắt buộc phải ghi bằng Microsoft Excel (COM) để giữ nguyên 100% "
                "định dạng, công thức và sheet ẩn/khóa (K5, K13)."
            )
        return "openpyxl"

    if engine in ("auto", "com"):
        if _has_win32com():
            return "win32com"
        if os.name == "nt":
            return "powershell_com"
        if engine == "auto" and file_format == FileFormat.XLSX:
            return "openpyxl"
        raise ExcelWriteError(
            "Không thể khởi tạo Microsoft Excel (COM). Hãy cài đặt Microsoft Excel trên Windows (K13)."
        )

    raise ConfigError(f"Engine ghi Excel không được hỗ trợ: {engine!r}")


def _write_via_win32com(
    workbook_path: Path,
    sheet_name: str,
    updates: Sequence[CellUpdate],
) -> None:
    """Ghi trực tiếp bằng Excel COM thông qua pywin32 (khi có cài pywin32)."""
    import pythoncom
    import win32com.client

    abs_path = str(workbook_path.resolve())
    pythoncom.CoInitialize()
    xl = None
    wb = None
    try:
        try:
            xl = win32com.client.DispatchEx("Excel.Application")
        except Exception as exc:
            raise ExcelWriteError(
                "Không khởi động được Microsoft Excel (COM). Hãy kiểm tra Excel đã cài đặt chưa.",
                detail=repr(exc),
            ) from exc

        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        xl.ScreenUpdating = False
        try:
            xl.AskToUpdateLinks = False
            xl.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception:
            pass

        try:
            wb = xl.Workbooks.Open(abs_path, UpdateLinks=0, ReadOnly=False)
        except Exception as exc:
            raise ExcelWriteError(
                f"Excel không mở được file '{workbook_path.name}'.", detail=repr(exc)
            ) from exc

        try:
            ws = wb.Worksheets(sheet_name)
        except Exception as exc:
            raise ExcelWriteError(
                f"Không tìm thấy sheet '{sheet_name}' trong file '{workbook_path.name}'.",
                detail=repr(exc),
            ) from exc

        if bool(ws.ProtectContents):
            raise ExcelWriteError(
                f"Sheet '{sheet_name}' trong '{workbook_path.name}' đang bị khóa bảo vệ (Protect Sheet). "
                "Tool không tự phá khóa (K10)."
            )

        # Kiểm tra trước toàn bộ các ô sẽ ghi (không ghi đè công thức hoặc ô gộp)
        for u in updates:
            cell = ws.Cells(u.row, u.col_index)
            if bool(cell.MergeCells):
                raise ExcelWriteError(
                    f"Ô {u.cell} tại sheet '{sheet_name}' đang bị gộp (Merge Cells); "
                    "không thể ghi để tránh làm hỏng cấu trúc biểu mẫu."
                )
            if bool(cell.HasFormula):
                raise ExcelWriteError(
                    f"Ô {u.cell} tại sheet '{sheet_name}' đang chứa công thức; "
                    "tool tuyệt đối không ghi đè công thức (R6, K5)."
                )

        # Chỉ cập nhật Value2 của đúng các ô trong WritePlan
        for u in updates:
            ws.Cells(u.row, u.col_index).Value2 = u.new_value

        wb.Save()
        wb.Close(SaveChanges=False)
        wb = None
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


_POWERSHELL_COM_SCRIPT = r"""
param(
    [Parameter(Mandatory=$true)][string]$PayloadPath,
    [Parameter(Mandatory=$true)][string]$ResultPath
)
$ErrorActionPreference = "Stop"

function Write-Result([bool]$ok, [string]$errorMsg, [int]$updatedCount) {
    $obj = @{
        ok = $ok
        error = $errorMsg
        updated = $updatedCount
    }
    $json = $obj | ConvertTo-Json -Compress
    [System.IO.File]::WriteAllText($ResultPath, $json, [System.Text.Encoding]::UTF8)
}

$xl = $null
$wb = $null
$ws = $null
try {
    $raw = [System.IO.File]::ReadAllText($PayloadPath, [System.Text.Encoding]::UTF8)
    $payload = $raw | ConvertFrom-Json

    try {
        $xl = New-Object -ComObject Excel.Application
    } catch {
        Write-Result $false ("Không khởi động được Microsoft Excel (COM): " + $_.Exception.Message) 0
        exit 0
    }

    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    $xl.ScreenUpdating = $false
    try {
        $xl.AskToUpdateLinks = $false
        $xl.AutomationSecurity = 3
    } catch {}

    try {
        $wb = $xl.Workbooks.Open([string]$payload.file_path, 0, $false)
    } catch {
        Write-Result $false ("Excel không mở được file: " + $_.Exception.Message) 0
        exit 0
    }

    try {
        $ws = $wb.Worksheets.Item([string]$payload.sheet_name)
    } catch {
        Write-Result $false ("Không tìm thấy sheet '" + [string]$payload.sheet_name + "' trong file.") 0
        exit 0
    }

    if ([bool]$ws.ProtectContents) {
        Write-Result $false ("Sheet '" + [string]$payload.sheet_name + "' đang bị khóa bảo vệ (Protect Sheet). Tool không tự phá khóa (K10).") 0
        exit 0
    }

    foreach ($u in $payload.updates) {
        $r = [int]$u.row
        $c = [int]$u.col
        $cellName = [string]$u.cell
        $cell = $ws.Cells.Item($r, $c)
        if ([bool]$cell.MergeCells) {
            Write-Result $false ("Ô " + $cellName + " tại sheet '" + [string]$payload.sheet_name + "' đang bị gộp (Merge Cells); không thể ghi.") 0
            exit 0
        }
        if ([bool]$cell.HasFormula) {
            Write-Result $false ("Ô " + $cellName + " tại sheet '" + [string]$payload.sheet_name + "' đang chứa công thức; tool không ghi đè công thức (R6, K5).") 0
            exit 0
        }
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)
    }

    $count = 0
    foreach ($u in $payload.updates) {
        $r = [int]$u.row
        $c = [int]$u.col
        $cell = $ws.Cells.Item($r, $c)
        if ($null -eq $u.value) {
            $cell.Value2 = $null
        } else {
            $cell.Value2 = [double]([int64]$u.value)
        }
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)
        $count += 1
    }

    $wb.Save()
    $wb.Close($false)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb)
    $wb = $null

    Write-Result $true "" $count
} catch {
    Write-Result $false ("Lỗi khi ghi qua Excel COM: " + $_.Exception.Message) 0
} finally {
    if ($null -ne $ws) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) } catch {}
    }
    if ($null -ne $wb) {
        try { $wb.Close($false) } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) } catch {}
    }
    if ($null -ne $xl) {
        try { $xl.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
"""


def _write_via_powershell_com(
    workbook_path: Path,
    sheet_name: str,
    updates: Sequence[CellUpdate],
) -> None:
    """Điều khiển trực tiếp Microsoft Excel COM qua Windows PowerShell (hỗ trợ cả .xls và .xlsx)."""
    abs_path = str(workbook_path.resolve())
    payload = {
        "file_path": abs_path,
        "sheet_name": sheet_name,
        "updates": [
            {
                "row": u.row,
                "col": u.col_index,
                "cell": u.cell,
                "value": u.new_value,
            }
            for u in updates
        ],
    }

    with tempfile.TemporaryDirectory(prefix="chl_com_") as tmp_dir:
        tmp = Path(tmp_dir)
        script_path = tmp / "run_excel_com.ps1"
        payload_path = tmp / "payload.json"
        result_path = tmp / "result.json"

        script_path.write_text(_POWERSHELL_COM_SCRIPT, encoding="utf-8-sig")
        payload_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

        try:
            proc = subprocess.run(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(script_path),
                    "-PayloadPath",
                    str(payload_path),
                    "-ResultPath",
                    str(result_path),
                ],
                capture_output=True,
                timeout=COM_TIMEOUT_SECONDS,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            raise ExcelWriteError(
                f"Microsoft Excel (COM) phản hồi quá thời gian ({COM_TIMEOUT_SECONDS}s) "
                f"khi ghi file '{workbook_path.name}'."
            ) from exc
        except OSError as exc:
            raise ExcelWriteError(
                "Không thể khởi chạy trình điều khiển Excel COM trên Windows.",
                detail=repr(exc),
            ) from exc

        if not result_path.exists():
            stderr_text = proc.stderr.decode("utf-8", errors="replace").strip()
            raise ExcelWriteError(
                f"Microsoft Excel (COM) kết thúc bất thường khi ghi '{workbook_path.name}'.",
                detail=stderr_text or f"returncode={proc.returncode}",
            )

        res_data = json.loads(result_path.read_text(encoding="utf-8-sig"))
        if not res_data.get("ok"):
            err_msg = str(res_data.get("error") or "Lỗi không xác định từ Excel COM.")
            raise ExcelWriteError(err_msg)


def _write_via_openpyxl(
    workbook_path: Path,
    file_format: FileFormat,
    sheet_name: str,
    updates: Sequence[CellUpdate],
) -> None:
    """Ghi trực tiếp vào ô của file .xlsx bằng openpyxl (chỉ dùng cho .xlsx khi chỉ định)."""
    import openpyxl
    from openpyxl.cell.cell import MergedCell

    if file_format != FileFormat.XLSX:
        raise ExcelWriteError("openpyxl chỉ hỗ trợ định dạng .xlsx.")

    wb = openpyxl.load_workbook(workbook_path)
    try:
        if sheet_name not in wb.sheetnames:
            raise ExcelWriteError(
                f"Không tìm thấy sheet '{sheet_name}' trong file '{workbook_path.name}'."
            )
        ws = wb[sheet_name]
        if ws.protection.sheet:
            raise ExcelWriteError(
                f"Sheet '{sheet_name}' trong '{workbook_path.name}' đang bị khóa bảo vệ (Protect Sheet). "
                "Tool không tự phá khóa (K10)."
            )

        merged_ranges = list(ws.merged_cells.ranges)
        for u in updates:
            cell = ws.cell(row=u.row, column=u.col_index)
            if isinstance(cell, MergedCell) or any(
                mr.min_row <= u.row <= mr.max_row and mr.min_col <= u.col_index <= mr.max_col
                for mr in merged_ranges
            ):
                raise ExcelWriteError(
                    f"Ô {u.cell} tại sheet '{sheet_name}' đang bị gộp (Merge Cells); "
                    "không thể ghi để tránh làm hỏng cấu trúc biểu mẫu."
                )
            if cell.data_type == "f" or (isinstance(cell.value, str) and cell.value.startswith("=")):
                raise ExcelWriteError(
                    f"Ô {u.cell} tại sheet '{sheet_name}' đang chứa công thức; "
                    "tool tuyệt đối không ghi đè công thức (R6, K5)."
                )

        for u in updates:
            ws.cell(row=u.row, column=u.col_index).value = u.new_value

        wb.save(workbook_path)
    finally:
        wb.close()


# =============================================================================
# 5. Hàm nội bộ kiểm tra & chuẩn hóa tham số
# =============================================================================

def _resolve_template_map(
    plan: WritePlan,
    template_profile: Optional[
        Union[TemplateProfile, Sequence[TargetTemplate], Mapping[str, TargetTemplate]]
    ],
) -> dict[str, TargetTemplate]:
    result: dict[str, TargetTemplate] = {}
    if template_profile is None or isinstance(template_profile, SourceProfile):
        return result

    if isinstance(template_profile, TargetTemplate):
        for tp in plan.target_plans:
            if tp.template_name == template_profile.name or len(plan.target_plans) == 1:
                result[tp.file_path] = template_profile
        return result

    if isinstance(template_profile, Mapping):
        for tp in plan.target_plans:
            tpl = template_profile.get(tp.file_path) or template_profile.get(tp.template_name)
            if isinstance(tpl, TargetTemplate):
                result[tp.file_path] = tpl
        return result

    by_name = {t.name: t for t in template_profile if isinstance(t, TargetTemplate)}
    for tp in plan.target_plans:
        if tp.template_name in by_name:
            result[tp.file_path] = by_name[tp.template_name]
    return result


def _resolve_target_workbook_map(
    plan: WritePlan,
    target_workbook: Optional[
        Union[TargetWorkbook, Sequence[TargetWorkbook], Mapping[str, PathLike]]
    ],
) -> dict[str, Path]:
    result: dict[str, Path] = {tp.file_path: Path(tp.file_path) for tp in plan.target_plans}
    if target_workbook is None:
        return result

    if isinstance(target_workbook, TargetReadResult):
        if len(plan.target_plans) == 1:
            result[plan.target_plans[0].file_path] = Path(target_workbook.file_path)
        else:
            result[target_workbook.file_path] = Path(target_workbook.file_path)
        return result

    if isinstance(target_workbook, (str, os.PathLike)):
        if len(plan.target_plans) != 1:
            raise ConfigError(
                "Khi WritePlan có nhiều file đích, cần truyền target_workbook theo danh sách hoặc dict."
            )
        result[plan.target_plans[0].file_path] = Path(target_workbook)
        return result

    if isinstance(target_workbook, Mapping):
        for k, v in target_workbook.items():
            result[k] = Path(v)
        return result

    seq = list(target_workbook)
    if len(seq) != len(plan.target_plans):
        raise ConfigError(
            f"Số lượng target_workbook ({len(seq)}) không khớp với số file đích trong WritePlan "
            f"({len(plan.target_plans)})."
        )
    for tp, item in zip(plan.target_plans, seq):
        if isinstance(item, TargetReadResult):
            result[tp.file_path] = Path(item.file_path)
        else:
            result[tp.file_path] = Path(item)
    return result


def _resolve_output_path_map(
    *,
    write_plan: WritePlan,
    workbook_map: Mapping[str, Path],
    output_dir: Path,
    output_paths: Optional[Union[PathLike, Sequence[PathLike], Mapping[str, PathLike]]],
    output_suffix: str,
    overwrite_output: bool,
) -> dict[str, Path]:
    result: dict[str, Path] = {}
    if output_paths is None:
        for tp in write_plan.target_plans:
            result[tp.file_path] = build_output_path(
                workbook_map[tp.file_path],
                output_dir,
                suffix=output_suffix,
                avoid_collision=not overwrite_output,
            )
        return result

    if isinstance(output_paths, (str, os.PathLike)):
        if len(write_plan.target_plans) != 1:
            raise ConfigError(
                "Khi WritePlan có nhiều file đích, cần truyền output_paths dạng dict hoặc danh sách."
            )
        result[write_plan.target_plans[0].file_path] = Path(output_paths)
        return result

    if isinstance(output_paths, Mapping):
        for tp in write_plan.target_plans:
            custom = output_paths.get(tp.file_path) or output_paths.get(str(workbook_map[tp.file_path]))
            if custom is not None:
                result[tp.file_path] = Path(custom)
            else:
                result[tp.file_path] = build_output_path(
                    workbook_map[tp.file_path],
                    output_dir,
                    suffix=output_suffix,
                    avoid_collision=not overwrite_output,
                )
        return result

    seq = list(output_paths)
    if len(seq) != len(write_plan.target_plans):
        raise ConfigError(
            f"Số lượng output_paths ({len(seq)}) không khớp với số file đích ({len(write_plan.target_plans)})."
        )
    for tp, p in zip(write_plan.target_plans, seq):
        result[tp.file_path] = Path(p)
    return result


def _verify_against_template(
    target_plan: TargetWritePlan,
    template: Optional[TargetTemplate],
) -> None:
    if template is None:
        return
    template.ensure_valid()
    expected_col = template.amount.letter
    if target_plan.amount_column != expected_col:
        raise SafetyError(
            f"Cột ghi số tiền của '{Path(target_plan.file_path).name}' ({target_plan.amount_column}) "
            f"không khớp với cấu hình mẫu '{template.name}' ({expected_col})."
        )
    for u in target_plan.updates:
        if u.column != expected_col:
            raise SafetyError(
                f"Ô {u.cell} ghi vào cột '{u.column}' khác với cột Số tiền '{expected_col}' "
                f"trong mẫu '{template.name}'."
            )
        if u.row < template.data_start_row:
            raise SafetyError(
                f"Ô {u.cell} nằm ở dòng {u.row}, phía trên dòng bắt đầu dữ liệu "
                f"({template.data_start_row}) của mẫu '{template.name}' (nguy cơ ghi đè tiêu đề)."
            )
        if template.data_end_row is not None and u.row > template.data_end_row:
            raise SafetyError(
                f"Ô {u.cell} nằm ở dòng {u.row}, vượt quá dòng kết thúc dữ liệu "
                f"({template.data_end_row}) của mẫu '{template.name}'."
            )


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _make_writable(path: Path) -> None:
    """Đảm bảo bản copy (output/backup) có quyền ghi ngay cả khi file gốc bật cờ Read-Only."""
    try:
        if path.exists():
            os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
    except OSError:
        pass


def _safe_unlink(path: Path) -> None:
    try:
        if path.exists():
            _make_writable(path)
            path.unlink()
    except OSError:
        log.warning("Không xóa được file tạm/lỗi: %s", path)


def _norm_path_key(path_str: str) -> str:
    try:
        return os.path.normcase(str(Path(path_str).resolve()))
    except OSError:
        return os.path.normcase(os.path.abspath(path_str))
