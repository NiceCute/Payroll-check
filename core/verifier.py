"""
verifier.py - Lớp hậu kiểm (Verification Layer): kiểm tra tính nhất quán Source -> Reconcile -> WritePlan
và đối chiếu từng ô giữa file kết quả mới với file ngân hàng gốc.

Kiểm tra:
1. Tổng record nguồn (`total_source_records`)
2. Tổng record matched (`total_matched_records`)
3. Tổng exception (`total_exceptions`, `resolved_exceptions`, `unresolved_exceptions`, `blocking_exceptions`)
4. Tổng record ghi (`total_written_records`, `total_source_written_records`, `total_zero_or_clear_records`)
5. Tổng tiền nguồn (`total_source_amount`, `total_source_transferable_amount`)
6. Tổng tiền chuẩn bị ghi (`total_planned_write_amount`, `total_skipped_transferable_amount`, `total_override_delta`)
7. Chênh lệch dữ liệu (`row_discrepancy`, `amount_discrepancy`, `cell_diffs`)
8. Tính nhất quán xuyên suốt: `Source -> Reconcile -> WritePlan` (+ hậu kiểm diff từng ô trên file output).

Module THUẦN KIỂM TRA: không ghi file, không tạo report, không xử lý GUI.
"""

from __future__ import annotations

import hashlib
import math
from collections import defaultdict
from pathlib import Path
from typing import Callable, Optional, Union

from .excel_reader import ExcelReadError, PathLike, WorkbookReader
from .models import (
    CellData,
    CellDiff,
    CellKind,
    CellUpdate,
    MatchItem,
    MatchStatus,
    RunResult,
    Severity,
    SourceReadResult,
    TargetWriteOutput,
    TargetWritePlan,
    ValidationResult,
    VerificationCode,
    VerificationError,
    VerificationFinding,
    VerificationResult,
    WriteExecutionResult,
    WritePlan,
    column_letter,
)
from .normalize import clean_text, normalize_amount

__all__ = [
    "CellDiff",
    "VerificationCode",
    "VerificationError",
    "VerificationFinding",
    "VerificationResult",
    "Verifier",
    "ensure_verified",
    "verify",
    "verify_execution",
    "verify_output_file_diff",
    "verify_plan_consistency",
]

_ZERO_OR_CLEAR_REASONS: frozenset[str] = frozenset({
    "X5_SET_ZERO",
    "X5_CLEAR",
    "X7_SET_ZERO",
    "X7_CLEAR",
})
_CLEAR_REASONS: frozenset[str] = frozenset({"X5_CLEAR", "X7_CLEAR"})
_ZERO_REASONS: frozenset[str] = frozenset({"X5_SET_ZERO", "X7_SET_ZERO"})


# =============================================================================
# 1. Giao diện chính (Class & Function API)
# =============================================================================

class Verifier:
    """Bộ kiểm tra tính nhất quán dữ liệu và hậu kiểm file xuất."""

    def __init__(self, *, verify_files_on_disk: bool = True) -> None:
        self.verify_files_on_disk = verify_files_on_disk

    def verify(
        self,
        run_result: Optional[Union[RunResult, WriteExecutionResult, WritePlan]] = None,
        validation_result: Optional[ValidationResult] = None,
        write_execution_result: Optional[WriteExecutionResult] = None,
        *,
        write_plan: Optional[WritePlan] = None,
        source: Optional[SourceReadResult] = None,
    ) -> VerificationResult:
        return verify(
            run_result=run_result,
            validation_result=validation_result,
            write_execution_result=write_execution_result,
            write_plan=write_plan,
            source=source,
            verify_files_on_disk=self.verify_files_on_disk,
        )


def verify(
    run_result: Optional[Union[RunResult, WriteExecutionResult, WritePlan]] = None,
    validation_result: Optional[ValidationResult] = None,
    write_execution_result: Optional[WriteExecutionResult] = None,
    *,
    write_plan: Optional[WritePlan] = None,
    source: Optional[SourceReadResult] = None,
    verify_files_on_disk: bool = True,
) -> VerificationResult:
    """Thực hiện kiểm tra toàn diện:
    - Tính nhất quán `Source -> Reconcile -> WritePlan`
    - Các chỉ số tổng hợp dòng, tiền và độ lệch (`row_discrepancy`, `amount_discrepancy`)
    - Trạng thái `ValidationResult` và `ExceptionItem`
    - Hậu kiểm từng ô giữa file gốc và file output mới (nếu có thông tin file xuất).
    """
    eff_plan, eff_val, eff_exec, eff_run = _extract_inputs(
        run_result=run_result,
        validation_result=validation_result,
        write_execution_result=write_execution_result,
        write_plan=write_plan,
    )

    res = VerificationResult()

    def add_error(
        code: VerificationCode,
        message: str,
        *,
        file_path: Optional[str] = None,
        sheet: Optional[str] = None,
        row: Optional[int] = None,
        column: Optional[str] = None,
    ) -> None:
        res.errors.append(
            VerificationFinding(
                code=code,
                severity=Severity.ERROR,
                message=message,
                file_path=file_path,
                sheet=sheet,
                row=row,
                column=column,
            )
        )

    def add_warning(
        code: VerificationCode,
        message: str,
        *,
        file_path: Optional[str] = None,
        sheet: Optional[str] = None,
        row: Optional[int] = None,
        column: Optional[str] = None,
    ) -> None:
        res.warnings.append(
            VerificationFinding(
                code=code,
                severity=Severity.WARNING,
                message=message,
                file_path=file_path,
                sheet=sheet,
                row=row,
                column=column,
            )
        )

    # 1. Kiểm tra ValidationResult (nếu có)
    if eff_val is not None:
        if not eff_val.is_safe:
            for s_err in eff_val.errors:
                add_error(
                    VerificationCode.VALIDATION_NOT_SAFE,
                    f"Chốt an toàn (ValidationResult) chưa đạt [{s_err.code.value}]: {s_err.message}",
                    file_path=s_err.file_path,
                    sheet=s_err.sheet,
                    row=s_err.row,
                    column=s_err.column,
                )

    # 2. Kiểm tra tính nhất quán Source -> Reconcile -> WritePlan & tính các chỉ số tổng hợp
    _verify_chain_and_metrics(
        plan=eff_plan,
        source=source,
        res=res,
        add_error=add_error,
        add_warning=add_warning,
    )

    # 3. Hậu kiểm file kết quả trên đĩa (nếu có WriteExecutionResult hoặc RunResult có output_files)
    if verify_files_on_disk and (eff_exec is not None or (eff_run is not None and eff_run.output_files)):
        _verify_disk_outputs(
            plan=eff_plan,
            exec_result=eff_exec,
            run_result=eff_run,
            res=res,
            add_error=add_error,
        )

    return res


# Bí danh tương thích
verify_execution = verify
verify_plan_consistency = verify


def ensure_verified(
    run_result: Optional[Union[RunResult, WriteExecutionResult, WritePlan]] = None,
    validation_result: Optional[ValidationResult] = None,
    write_execution_result: Optional[WriteExecutionResult] = None,
    *,
    write_plan: Optional[WritePlan] = None,
    source: Optional[SourceReadResult] = None,
    verify_files_on_disk: bool = True,
) -> VerificationResult:
    """Thực hiện hậu kiểm và ném `VerificationError` nếu phát hiện bất kỳ sai lệch nào."""
    res = verify(
        run_result=run_result,
        validation_result=validation_result,
        write_execution_result=write_execution_result,
        write_plan=write_plan,
        source=source,
        verify_files_on_disk=verify_files_on_disk,
    )
    res.ensure_verified()
    return res


# =============================================================================
# 2. Kiểm tra chuỗi Source -> Reconcile -> WritePlan & Tính chỉ số
# =============================================================================

def _verify_chain_and_metrics(
    *,
    plan: WritePlan,
    source: Optional[SourceReadResult],
    res: VerificationResult,
    add_error: Callable[..., None],
    add_warning: Callable[..., None],
) -> None:
    source_matches = [m for m in plan.matches if m.source_row is not None]
    target_only_matches = [m for m in plan.matches if m.source_row is None]

    # --- Chỉ số Exception ---
    res.total_exceptions = len(plan.exceptions)
    res.resolved_exceptions = sum(1 for e in plan.exceptions if e.is_resolved)
    res.unresolved_exceptions = sum(1 for e in plan.exceptions if not e.is_resolved)
    res.blocking_exceptions = sum(1 for e in plan.exceptions if e.is_blocking)

    for exc in plan.exceptions:
        if exc.is_blocking or not exc.is_resolved:
            add_error(
                VerificationCode.UNRESOLVED_EXCEPTION,
                f"Ngoại lệ [{exc.id} - {exc.code.value}] chưa được xử lý hợp lệ: {exc.message}",
                file_path=exc.target_file,
                sheet=exc.target_sheet,
                row=exc.target_row or exc.source_row,
            )

    for issue in plan.blocking_issues:
        add_error(
            VerificationCode.SOURCE_RECONCILE_MISMATCH,
            f"Lỗi cấu trúc mức file [{issue.code.value}]: {issue.message}",
            sheet=issue.sheet,
            row=issue.row,
            column=issue.column,
        )

    # --- Chặng 1: Source -> Reconcile ---
    res.total_source_records = plan.total_source_records
    res.total_matched_records = sum(1 for m in source_matches if m.target_row is not None)
    res.total_unmatched_source_records = sum(1 for m in source_matches if m.target_row is None)
    res.total_unmatched_target_records = sum(
        1 for m in target_only_matches if m.account_no is not None
    )

    if source is not None:
        if len(source.records) != plan.total_source_records:
            res.source_to_reconcile_ok = False
            add_error(
                VerificationCode.SOURCE_RECONCILE_MISMATCH,
                f"Tổng record nguồn trong SourceReadResult ({len(source.records)}) "
                f"không khớp với WritePlan.total_source_records ({plan.total_source_records}).",
            )
        if source.total_transferable_amount != plan.total_source_transferable_amount:
            res.source_to_reconcile_ok = False
            add_error(
                VerificationCode.SOURCE_RECONCILE_MISMATCH,
                f"Tổng tiền đủ điều kiện chuyển khoản trong SourceReadResult "
                f"({source.total_transferable_amount:,} đ) không khớp với WritePlan "
                f"({plan.total_source_transferable_amount:,} đ).",
            )
        match_by_src_row: dict[int, list[MatchItem]] = defaultdict(list)
        for m in source_matches:
            if m.source_row is not None:
                match_by_src_row[m.source_row].append(m)

        for srec in source.records:
            m_hits = match_by_src_row.get(srec.row, [])
            if len(m_hits) != 1:
                res.source_to_reconcile_ok = False
                add_error(
                    VerificationCode.SOURCE_RECONCILE_MISMATCH,
                    f"Dòng lương {srec.row} (Số thẻ '{srec.card_no}') xuất hiện {len(m_hits)} lần "
                    "trong danh sách đối chiếu (phải đúng 1 lần).",
                    row=srec.row,
                )
                continue
            m_hit = m_hits[0]
            if m_hit.card_no != srec.card_no or m_hit.source_amount != srec.amount:
                res.source_to_reconcile_ok = False
                add_error(
                    VerificationCode.SOURCE_RECONCILE_MISMATCH,
                    f"Dòng lương {srec.row} không nhất quán giữa SourceRecord "
                    f"(Số thẻ '{srec.card_no}', tiền={srec.amount}) và MatchItem "
                    f"(Số thẻ '{m_hit.card_no}', tiền={m_hit.source_amount}).",
                    row=srec.row,
                )

    if len(source_matches) != plan.total_source_records:
        res.source_to_reconcile_ok = False
        add_error(
            VerificationCode.SOURCE_RECONCILE_MISMATCH,
            f"Số dòng nguồn trong MatchItem ({len(source_matches)}) khác với "
            f"total_source_records ({plan.total_source_records}).",
        )

    if res.total_matched_records + res.total_unmatched_source_records != len(source_matches):
        res.source_to_reconcile_ok = False
        add_error(
            VerificationCode.SOURCE_RECONCILE_MISMATCH,
            "Tổng dòng matched + unmatched nguồn không bằng tổng dòng nguồn.",
        )

    # --- Chặng 2: Reconcile -> WritePlan ---
    flattened_tp_updates = [u for tp in plan.target_plans for u in tp.updates]
    if len(flattened_tp_updates) != len(plan.updates) or set(flattened_tp_updates) != set(plan.updates):
        res.reconcile_to_plan_ok = False
        add_error(
            VerificationCode.RECONCILE_PLAN_MISMATCH,
            f"Danh sách `plan.updates` ({len(plan.updates)} ô) không khớp với tổng hợp "
            f"từ `plan.target_plans` ({len(flattened_tp_updates)} ô).",
        )

    update_by_coord: dict[tuple[str, str, int], CellUpdate] = {}
    for u in plan.updates:
        key = (u.file_path, u.sheet_name, u.row)
        if key in update_by_coord:
            res.reconcile_to_plan_ok = False
            add_error(
                VerificationCode.RECONCILE_PLAN_MISMATCH,
                f"Ô {u.cell} tại '{Path(u.file_path).name}' bị trùng lặp trong WritePlan.updates.",
                file_path=u.file_path,
                sheet=u.sheet_name,
                row=u.row,
                column=u.column,
            )
        update_by_coord[key] = u

    res.total_written_records = len(plan.updates)
    res.total_source_written_records = sum(
        1 for u in plan.updates if u.card_no is not None and u.reason not in _ZERO_OR_CLEAR_REASONS
    )
    res.total_zero_or_clear_records = sum(
        1 for u in plan.updates if u.reason in _ZERO_OR_CLEAR_REASONS
    )
    res.total_skipped_records = sum(1 for m in source_matches if m.status == MatchStatus.SKIPPED)

    res.total_source_amount = sum(
        m.source_amount
        for m in source_matches
        if isinstance(m.source_amount, int) and m.source_amount > 0
    )
    res.total_source_transferable_amount = plan.total_source_transferable_amount
    res.total_planned_write_amount = sum(u.new_value or 0 for u in plan.updates)

    # Đối chiếu từng MatchItem với CellUpdate
    matched_update_coords: set[tuple[str, str, int]] = set()
    skipped_transferable_amount = 0
    override_delta = 0
    zero_or_clear_from_source_rows = 0

    for m in source_matches:
        is_transferable_candidate = (
            m.account_no is not None
            and isinstance(m.source_amount, int)
            and m.source_amount > 0
        )

        if m.status == MatchStatus.SKIPPED:
            if is_transferable_candidate and m.source_amount is not None:
                skipped_transferable_amount += m.source_amount
            if m.target_file and m.target_sheet and m.target_row is not None:
                coord = (m.target_file, m.target_sheet, m.target_row)
                if coord in update_by_coord:
                    res.reconcile_to_plan_ok = False
                    add_error(
                        VerificationCode.RECONCILE_PLAN_MISMATCH,
                        f"Dòng nguồn {m.source_row} (Số thẻ '{m.card_no}') ở trạng thái SKIPPED "
                        f"nhưng vẫn có lệnh ghi tại ô {update_by_coord[coord].cell}.",
                        file_path=m.target_file,
                        sheet=m.target_sheet,
                        row=m.target_row,
                    )
            continue

        if m.status in (MatchStatus.PENDING, MatchStatus.ERROR):
            res.reconcile_to_plan_ok = False
            add_error(
                VerificationCode.RECONCILE_PLAN_MISMATCH,
                f"Dòng nguồn {m.source_row} (Số thẻ '{m.card_no}') vẫn đang ở trạng thái {m.status.value}.",
                row=m.source_row,
            )
            continue

        # WILL_UPDATE hoặc UNCHANGED
        if not m.target_file or not m.target_sheet or m.target_row is None:
            res.reconcile_to_plan_ok = False
            add_error(
                VerificationCode.RECONCILE_PLAN_MISMATCH,
                f"Dòng nguồn {m.source_row} (Số thẻ '{m.card_no}') ở trạng thái {m.status.value} "
                "nhưng thiếu thông tin dòng đích.",
                row=m.source_row,
            )
            continue

        coord = (m.target_file, m.target_sheet, m.target_row)
        matched_update_coords.add(coord)
        u_hit = update_by_coord.get(coord)
        if u_hit is None:
            res.reconcile_to_plan_ok = False
            add_error(
                VerificationCode.RECONCILE_PLAN_MISMATCH,
                f"Dòng nguồn {m.source_row} (Số thẻ '{m.card_no}', TK '{m.account_no}') "
                f"không có CellUpdate tại dòng đích {m.target_row}.",
                file_path=m.target_file,
                sheet=m.target_sheet,
                row=m.target_row,
            )
            continue
        u = u_hit

        if u.card_no != m.card_no or u.account_no != m.account_no:
            res.reconcile_to_plan_ok = False
            add_error(
                VerificationCode.RECONCILE_PLAN_MISMATCH,
                f"Định danh tại ô {u.cell} (Số thẻ '{u.card_no}', TK '{u.account_no}') "
                f"lệch với MatchItem (Số thẻ '{m.card_no}', TK '{m.account_no}').",
                file_path=u.file_path,
                sheet=u.sheet_name,
                row=u.row,
                column=u.column,
            )

        if u.new_value != m.new_amount or u.old_value != m.old_amount:
            res.reconcile_to_plan_ok = False
            add_error(
                VerificationCode.RECONCILE_PLAN_MISMATCH,
                f"Giá trị tại ô {u.cell} (old={u.old_value}, new={u.new_value}) "
                f"lệch với MatchItem (old={m.old_amount}, new={m.new_amount}).",
                file_path=u.file_path,
                sheet=u.sheet_name,
                row=u.row,
                column=u.column,
            )

        if u.reason in ("MATCHED", "CONFIRMED_MISMATCH"):
            if u.new_value != m.source_amount:
                res.reconcile_to_plan_ok = False
                add_error(
                    VerificationCode.RECONCILE_PLAN_MISMATCH,
                    f"Ô {u.cell} (Số thẻ '{m.card_no}') có new_value={u.new_value} "
                    f"khác với số tiền lương nguồn ({m.source_amount}).",
                    file_path=u.file_path,
                    sheet=u.sheet_name,
                    row=u.row,
                    column=u.column,
                )
        elif u.reason in _ZERO_OR_CLEAR_REASONS:
            zero_or_clear_from_source_rows += 1
            if is_transferable_candidate and m.source_amount is not None:
                skipped_transferable_amount += m.source_amount
        elif u.reason == "OVERRIDE_AMOUNT":
            base_amt = m.source_amount if (is_transferable_candidate and m.source_amount is not None) else 0
            override_delta += (u.new_value or 0) - base_amt

    # Kiểm tra các dòng đích không có trong lương mới (X5)
    for m in target_only_matches:
        if m.status in (MatchStatus.PENDING, MatchStatus.ERROR):
            res.reconcile_to_plan_ok = False
            add_error(
                VerificationCode.RECONCILE_PLAN_MISMATCH,
                f"Dòng đích {m.target_row} (TK '{m.account_no}') vẫn đang ở trạng thái {m.status.value}.",
                file_path=m.target_file,
                sheet=m.target_sheet,
                row=m.target_row,
            )
            continue
        if m.target_file and m.target_sheet and m.target_row is not None:
            coord = (m.target_file, m.target_sheet, m.target_row)
            if coord in update_by_coord:
                matched_update_coords.add(coord)
                u_x5 = update_by_coord[coord]
                if u_x5.reason not in _ZERO_OR_CLEAR_REASONS:
                    res.reconcile_to_plan_ok = False
                    add_error(
                        VerificationCode.RECONCILE_PLAN_MISMATCH,
                        f"Ô {u_x5.cell} của dòng đích dư (X5) có lý do không hợp lệ: '{u_x5.reason}'.",
                        file_path=u_x5.file_path,
                        sheet=u_x5.sheet_name,
                        row=u_x5.row,
                        column=u_x5.column,
                    )

    # Kiểm tra CellUpdate mồ côi
    orphan_coords = set(update_by_coord.keys()) - matched_update_coords
    for fpath, sname, r in sorted(orphan_coords):
        res.reconcile_to_plan_ok = False
        u_orp = update_by_coord[(fpath, sname, r)]
        add_error(
            VerificationCode.RECONCILE_PLAN_MISMATCH,
            f"Ô {u_orp.cell} tại '{Path(fpath).name}' nằm trong WritePlan "
            "nhưng không khớp với bất kỳ MatchItem nào.",
            file_path=fpath,
            sheet=sname,
            row=r,
            column=u_orp.column,
        )

    res.total_skipped_transferable_amount = skipped_transferable_amount
    res.total_override_delta = override_delta

    # --- Chênh lệch số dòng (row_discrepancy) & tổng tiền (amount_discrepancy) ---
    accounted_source_rows = (
        res.total_source_written_records
        + zero_or_clear_from_source_rows
        + res.total_skipped_records
    )
    res.row_discrepancy = res.total_source_records - accounted_source_rows
    if res.row_discrepancy != 0:
        add_error(
            VerificationCode.ROW_DISCREPANCY,
            f"Chênh lệch số dòng nguồn: tổng dòng nguồn ({res.total_source_records}) - "
            f"[dòng ghi ({res.total_source_written_records}) + đặt 0/xóa ({zero_or_clear_from_source_rows}) + "
            f"bỏ qua ({res.total_skipped_records})] = {res.row_discrepancy}.",
        )

    expected_total_amount = res.total_source_transferable_amount + res.total_override_delta
    actual_total_amount = res.total_planned_write_amount + res.total_skipped_transferable_amount
    res.amount_discrepancy = expected_total_amount - actual_total_amount
    if res.amount_discrepancy != 0:
        add_error(
            VerificationCode.AMOUNT_DISCREPANCY,
            f"Chênh lệch tổng tiền: [tổng tiền đủ điều kiện ({res.total_source_transferable_amount:,} đ) + "
            f"chênh lệch sửa tay ({res.total_override_delta:,} đ)] - "
            f"[tổng tiền ghi ({res.total_planned_write_amount:,} đ) + "
            f"tổng tiền bỏ qua ({res.total_skipped_transferable_amount:,} đ)] = {res.amount_discrepancy:,} đ.",
        )


# =============================================================================
# 3. Hậu kiểm từng ô giữa File Kết Quả Mới và File Ngân Hàng Gốc (K6)
# =============================================================================

def _verify_disk_outputs(
    *,
    plan: WritePlan,
    exec_result: Optional[WriteExecutionResult],
    run_result: Optional[RunResult],
    res: VerificationResult,
    add_error: Callable[..., None],
) -> None:
    target_outputs = _collect_target_outputs(plan, exec_result, run_result, add_error, res)

    for tp, out_meta in target_outputs:
        orig_path = Path(out_meta.source_target_file)
        out_path = Path(out_meta.output_file)
        bak_path = Path(out_meta.backup_file) if out_meta.backup_file else None

        if not orig_path.exists():
            res.post_write_diff_ok = False
            add_error(
                VerificationCode.MISSING_OUTPUT_FILE,
                f"Không tìm thấy file ngân hàng gốc để đối chiếu hậu kiểm: '{orig_path}'.",
                file_path=str(orig_path),
            )
            continue

        if not out_path.exists():
            res.post_write_diff_ok = False
            add_error(
                VerificationCode.MISSING_OUTPUT_FILE,
                f"Không tìm thấy file kết quả mới để hậu kiểm: '{out_path}'.",
                file_path=str(out_path),
            )
            continue

        if bak_path is not None and not bak_path.exists():
            res.post_write_diff_ok = False
            add_error(
                VerificationCode.MISSING_BACKUP_FILE,
                f"Không tìm thấy file sao lưu (backup) đã khai báo: '{bak_path}'.",
                file_path=str(bak_path),
            )

        if out_meta.original_sha256:
            current_orig_hash = _sha256_file(orig_path)
            if current_orig_hash != out_meta.original_sha256:
                res.post_write_diff_ok = False
                add_error(
                    VerificationCode.ORIGINAL_FILE_MODIFIED,
                    f"File ngân hàng gốc '{orig_path.name}' đã bị thay đổi nội dung (SHA-256 lệch)!",
                    file_path=str(orig_path),
                )

        diffs = verify_output_file_diff(
            original_path=orig_path,
            output_path=out_path,
            target_plan=tp,
            res=res,
            add_error=add_error,
        )
        res.cell_diffs.extend(diffs)
        res.verified_files.append(str(out_path))


def verify_output_file_diff(
    original_path: PathLike,
    output_path: PathLike,
    target_plan: TargetWritePlan,
    *,
    res: Optional[VerificationResult] = None,
    add_error: Optional[Callable[..., None]] = None,
) -> list[CellDiff]:
    """Đọc lại `output_path`, so toàn bộ các sheet và từng ô với `original_path`.

    Quy tắc K6:
    - Danh sách sheet và trạng thái ẩn/hiện phải giống hệt bản gốc.
    - Các ô nằm trong `target_plan.updates` phải mang đúng giá trị `u.new_value`.
    - Mọi ô dữ liệu khác trên toàn bộ các sheet phải giữ nguyên giá trị như bản gốc.
    """
    diffs: list[CellDiff] = []
    orig_p = Path(original_path)
    out_p = Path(output_path)

    def report_err(
        code: VerificationCode,
        msg: str,
        sheet: Optional[str] = None,
        row: Optional[int] = None,
        col: Optional[str] = None,
    ) -> None:
        if res is not None:
            res.post_write_diff_ok = False
        if add_error is not None:
            add_error(code, msg, file_path=str(out_p), sheet=sheet, row=row, column=col)

    try:
        with WorkbookReader(orig_p) as wb_orig, WorkbookReader(out_p) as wb_out:
            if wb_orig.sheet_infos != wb_out.sheet_infos:
                report_err(
                    VerificationCode.SHEET_STRUCTURE_CHANGED,
                    f"Cấu trúc danh sách sheet của '{out_p.name}' ({wb_out.sheet_infos}) "
                    f"khác với bản gốc ({wb_orig.sheet_infos}).",
                )
                return diffs

            planned_by_rc: dict[tuple[int, int], CellUpdate] = {
                (u.row, u.col_index): u for u in target_plan.updates
            }

            for s_info in wb_orig.sheet_infos:
                sname = s_info.name
                s_orig = wb_orig.read_sheet(sname)
                s_out = wb_out.read_sheet(sname)
                is_target_sheet = sname == target_plan.sheet_name

                all_coords = set(s_orig.cells.keys()) | set(s_out.cells.keys())
                if is_target_sheet:
                    all_coords |= set(planned_by_rc.keys())

                for r, c in sorted(all_coords):
                    col_str = column_letter(c)
                    c_orig = s_orig.cell(r, c)
                    c_out = s_out.cell(r, c)

                    if is_target_sheet and (r, c) in planned_by_rc:
                        u = planned_by_rc[(r, c)]
                        if not _planned_cell_matches(c_out, u.new_value):
                            msg = (
                                f"Ô dự kiến ghi {col_str}{r} tại sheet '{sname}' ('{out_p.name}') "
                                f"có giá trị thực tế {c_out.value!r} khác với giá trị kế hoạch {u.new_value!r}."
                            )
                            diffs.append(
                                CellDiff(
                                    file_path=str(out_p),
                                    sheet_name=sname,
                                    row=r,
                                    column=col_str,
                                    expected_value=u.new_value,
                                    actual_value=c_out.value,
                                    is_planned_cell=True,
                                    message=msg,
                                )
                            )
                            report_err(
                                VerificationCode.POST_WRITE_VALUE_MISMATCH,
                                msg,
                                sheet=sname,
                                row=r,
                                col=col_str,
                            )
                    else:
                        if not _unplanned_cells_equal(c_orig, c_out):
                            msg = (
                                f"Ô ngoài kế hoạch {col_str}{r} tại sheet '{sname}' ('{out_p.name}') "
                                f"bị thay đổi ngoài ý muốn: gốc={c_orig.value!r} -> mới={c_out.value!r}."
                            )
                            diffs.append(
                                CellDiff(
                                    file_path=str(out_p),
                                    sheet_name=sname,
                                    row=r,
                                    column=col_str,
                                    expected_value=c_orig.value,
                                    actual_value=c_out.value,
                                    is_planned_cell=False,
                                    message=msg,
                                )
                            )
                            report_err(
                                VerificationCode.UNEXPECTED_CELL_MODIFIED,
                                msg,
                                sheet=sname,
                                row=r,
                                col=col_str,
                            )
    except ExcelReadError as exc:
        report_err(VerificationCode.MISSING_OUTPUT_FILE, exc.user_message)

    return diffs


# =============================================================================
# 4. Các hàm phụ trợ so sánh giá trị ô
# =============================================================================

def _planned_cell_matches(actual: CellData, expected: Optional[int]) -> bool:
    if actual.is_formula:
        return False
    if expected is None:
        return actual.is_empty
    if actual.is_empty or actual.kind == CellKind.ERROR:
        return False
    norm = normalize_amount(actual.value)
    return not norm.is_error and norm.value == expected


def _unplanned_cells_equal(orig: CellData, out: CellData) -> bool:
    if orig.is_empty and out.is_empty:
        return True
    # Với .xlsx: nếu ô gốc là công thức và ô mới vẫn là công thức -> công thức được bảo toàn
    # (giá trị cache của công thức có thể được Excel tự động tính lại khi lưu)
    if orig.is_formula and out.is_formula:
        return True
    if orig.is_formula != out.is_formula:
        return False
    if orig.kind == CellKind.ERROR and out.kind == CellKind.ERROR:
        return True
    if orig.kind != out.kind:
        return False
    if orig.kind == CellKind.NUMBER:
        try:
            f1 = float(orig.value)
            f2 = float(out.value)
            if math.isnan(f1) and math.isnan(f2):
                return True
            return abs(f1 - f2) <= 1e-6
        except (TypeError, ValueError):
            return False
    if orig.kind == CellKind.TEXT:
        return str(orig.value) == str(out.value) or clean_text(orig.value) == clean_text(out.value)
    return bool(orig.value == out.value)


def _extract_inputs(
    *,
    run_result: Optional[Union[RunResult, WriteExecutionResult, WritePlan]],
    validation_result: Optional[ValidationResult],
    write_execution_result: Optional[WriteExecutionResult],
    write_plan: Optional[WritePlan],
) -> tuple[
    WritePlan,
    Optional[ValidationResult],
    Optional[WriteExecutionResult],
    Optional[RunResult],
]:
    eff_plan: Optional[WritePlan] = write_plan
    eff_val: Optional[ValidationResult] = validation_result
    eff_exec: Optional[WriteExecutionResult] = write_execution_result
    eff_run: Optional[RunResult] = None

    if isinstance(run_result, WritePlan):
        eff_plan = eff_plan or run_result
    elif isinstance(run_result, WriteExecutionResult):
        eff_exec = eff_exec or run_result
        eff_plan = eff_plan or run_result.write_plan
        eff_val = eff_val or run_result.validation
    elif isinstance(run_result, RunResult):
        eff_run = run_result
        eff_plan = eff_plan or run_result.write_plan

    if eff_exec is not None:
        eff_plan = eff_plan or eff_exec.write_plan
        eff_val = eff_val or eff_exec.validation

    if eff_plan is None:
        raise VerificationError("Phải cung cấp WritePlan, WriteExecutionResult hoặc RunResult để hậu kiểm.")

    return eff_plan, eff_val, eff_exec, eff_run


def _collect_target_outputs(
    plan: WritePlan,
    exec_result: Optional[WriteExecutionResult],
    run_result: Optional[RunResult],
    add_error: Callable[..., None],
    res: VerificationResult,
) -> list[tuple[TargetWritePlan, TargetWriteOutput]]:
    active_plans = [tp for tp in plan.target_plans if tp.updates]
    paired: list[tuple[TargetWritePlan, TargetWriteOutput]] = []

    if exec_result is not None:
        if len(exec_result.target_outputs) != len(active_plans):
            res.post_write_diff_ok = False
            add_error(
                VerificationCode.MISSING_OUTPUT_FILE,
                f"Số file đã xuất ({len(exec_result.target_outputs)}) không khớp với "
                f"số kế hoạch file đích ({len(active_plans)}).",
            )
        by_src = {
            Path(o.source_target_file).resolve(): o for o in exec_result.target_outputs
        }
        for tp in active_plans:
            hit = by_src.get(Path(tp.file_path).resolve())
            if hit is None:
                res.post_write_diff_ok = False
                add_error(
                    VerificationCode.MISSING_OUTPUT_FILE,
                    f"Không tìm thấy kết quả ghi (TargetWriteOutput) cho file đích '{tp.file_path}'.",
                    file_path=tp.file_path,
                )
            else:
                paired.append((tp, hit))
        return paired

    if run_result is not None and run_result.output_files:
        if len(run_result.output_files) != len(active_plans):
            res.post_write_diff_ok = False
            add_error(
                VerificationCode.MISSING_OUTPUT_FILE,
                f"Số lượng output_files trong RunResult ({len(run_result.output_files)}) "
                f"không khớp với số kế hoạch file đích ({len(active_plans)}).",
            )
        for idx, tp in enumerate(active_plans):
            if idx >= len(run_result.output_files):
                break
            out_file = run_result.output_files[idx]
            paired.append(
                (
                    tp,
                    TargetWriteOutput(
                        source_target_file=tp.file_path,
                        backup_file="",
                        output_file=out_file,
                        sheet_name=tp.sheet_name,
                        file_format=tp.file_format,
                        updated_cells=len(tp.updates),
                        changed_cells=len(tp.changed_updates),
                        total_new_amount=tp.total_new_amount,
                    ),
                )
            )
    return paired


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()
