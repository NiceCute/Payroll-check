"""
safety.py - Lớp chốt chặn an toàn (Safety Layer) trước khi ghi file ngân hàng.

Chức năng:
1. Kiểm tra WritePlan, ExceptionItem, RunResult trước khi cho phép ghi file.
2. Ngăn chặn tuyệt đối việc ghi dữ liệu nguy hiểm.
3. Phát hiện:
   - Mapping mơ hồ (lỗi tiêu đề, cột nghi lệch, nhiều kế hoạch trùng file, trùng tọa độ ô, lệch cột).
   - Nhiều người trùng định danh (trùng Số thẻ, trùng Số tài khoản ở nguồn/đích/WritePlan).
   - Thiếu số tiền hoặc số tiền không hợp lệ (trống, <= 0, không nguyên).
   - Thiếu Số tài khoản hoặc Số tài khoản không hợp lệ tại ô dự kiến ghi.
   - Thay đổi bất thường (số tiền > 1 tỷ, biến động đột biến so với kỳ trước, đặt 0/xóa trắng, ghi đè file gốc).
   - Dữ liệu không nhất quán (còn Exception chưa xử lý, lệch cân đối số dòng K8, lệch cân đối tổng tiền K8).

Module THUẦN KIỂM TRA: tuyệt đối không ghi file, không sửa file.
"""

from __future__ import annotations

import os
import stat
from collections import defaultdict
from pathlib import Path
from typing import Callable, Iterable, Optional, Sequence, Union

from .models import (
    CellUpdate,
    ExceptionCode,
    ExceptionItem,
    MatchItem,
    MatchStatus,
    ResolutionAction,
    RunResult,
    SafetyCode,
    SafetyError,
    SafetyFinding,
    Severity,
    ValidationResult,
    WritePlan,
    column_index,
    column_letter,
)

__all__ = [
    "ABNORMAL_DECREASE_RATIO",
    "ABNORMAL_INCREASE_RATIO",
    "ABNORMAL_TOTAL_DELTA_RATIO",
    "AMOUNT_WARNING_THRESHOLD",
    "SafetyCode",
    "SafetyError",
    "SafetyFinding",
    "ValidationResult",
    "check_file_locks",
    "ensure_safe_to_write",
    "is_file_locked",
    "validate_safety",
    "validate_write_plan",
]

# Ngưỡng cảnh báo thay đổi bất thường
AMOUNT_WARNING_THRESHOLD: int = 1_000_000_000   # > 1 tỷ VND / người -> cảnh báo
ABNORMAL_INCREASE_RATIO: float = 3.0            # Tăng gấp >= 3 lần so với kỳ trước -> cảnh báo
ABNORMAL_DECREASE_RATIO: float = 0.2            # Giảm còn <= 20% so với kỳ trước -> cảnh báo
ABNORMAL_TOTAL_DELTA_RATIO: float = 0.5         # Tổng tiền cả file lệch >= 50% so với kỳ trước -> cảnh báo

_ZERO_OR_CLEAR_REASONS: frozenset[str] = frozenset({
    "X5_SET_ZERO",
    "X5_CLEAR",
    "X7_SET_ZERO",
    "X7_CLEAR",
})
_CLEAR_REASONS: frozenset[str] = frozenset({"X5_CLEAR", "X7_CLEAR"})
_ZERO_REASONS: frozenset[str] = frozenset({"X5_SET_ZERO", "X7_SET_ZERO"})


# =============================================================================
# 1. Giao diện kiểm tra chính
# =============================================================================

def validate_write_plan(
    plan_or_run: Union[WritePlan, RunResult],
    exceptions: Optional[Sequence[ExceptionItem]] = None,
    run_result: Optional[RunResult] = None,
    *,
    check_disk: bool = False,
) -> ValidationResult:
    """Kiểm tra toàn diện `WritePlan` (kèm `ExceptionItem` và `RunResult` nếu có) trước khi ghi file.

    Args:
        plan_or_run : `WritePlan` hoặc `RunResult`.
        exceptions  : Danh sách `ExceptionItem` (nếu None sẽ lấy từ `plan.exceptions`).
        run_result  : `RunResult` (để kiểm tra đường dẫn backup/output không ghi đè file gốc R8).
        check_disk  : Nếu True, kiểm tra thêm khóa file trên đĩa (file đang mở trong Excel - K9).

    Returns:
        `ValidationResult` chứa `errors` (Error List) và `warnings` (Warning List).
    """
    effective_run: Optional[RunResult]
    if isinstance(plan_or_run, RunResult):
        plan = plan_or_run.write_plan
        effective_run = run_result or plan_or_run
    else:
        plan = plan_or_run
        effective_run = run_result

    effective_exceptions = list(exceptions) if exceptions is not None else list(plan.exceptions)
    result = ValidationResult()

    def add_error(
        code: SafetyCode,
        message: str,
        *,
        file_path: Optional[str] = None,
        sheet: Optional[str] = None,
        row: Optional[int] = None,
        column: Optional[str] = None,
        card_no: Optional[str] = None,
        account_no: Optional[str] = None,
    ) -> None:
        result.errors.append(
            SafetyFinding(
                code=code,
                severity=Severity.ERROR,
                message=message,
                file_path=file_path,
                sheet=sheet,
                row=row,
                column=column,
                card_no=card_no,
                account_no=account_no,
            )
        )

    def add_warning(
        code: SafetyCode,
        message: str,
        *,
        file_path: Optional[str] = None,
        sheet: Optional[str] = None,
        row: Optional[int] = None,
        column: Optional[str] = None,
        card_no: Optional[str] = None,
        account_no: Optional[str] = None,
    ) -> None:
        result.warnings.append(
            SafetyFinding(
                code=code,
                severity=Severity.WARNING,
                message=message,
                file_path=file_path,
                sheet=sheet,
                row=row,
                column=column,
                card_no=card_no,
                account_no=account_no,
            )
        )

    # Chốt 1: Kiểm tra mapping mơ hồ & cấu trúc kế hoạch ghi
    _check_mapping_and_structure(plan, add_error)

    # Chốt 2: Kiểm tra Exception chưa xử lý & lỗi chặn cứng
    _check_exceptions(plan, effective_exceptions, add_error, add_warning)

    # Chốt 3: Kiểm tra nhiều người trùng định danh (Số thẻ / Số tài khoản)
    _check_duplicate_identities(plan, effective_exceptions, add_error)

    # Chốt 4: Kiểm tra thiếu Số tài khoản & thiếu/sai Số tiền trên từng CellUpdate & MatchItem
    _check_updates_and_matches(plan, add_error)

    # Chốt 5: Kiểm tra tính nhất quán & cân đối số dòng + tổng tiền (K8)
    row_ok, amt_ok = _check_consistency_and_balance(plan, add_error)
    result.row_balance_ok = row_ok
    result.amount_balance_ok = amt_ok

    # Chốt 6: Phát hiện thay đổi bất thường
    _check_abnormal_changes(plan, add_error, add_warning)

    # Chốt 7: Kiểm tra an toàn đường dẫn xuất file / backup (R8) & khóa file (K9)
    if effective_run is not None:
        _check_run_result_paths(plan, effective_run, add_error, add_warning)

    if check_disk:
        _check_disk_locks(plan, effective_run, result)

    return result


# Bí danh ngắn gọn
validate_safety = validate_write_plan


def ensure_safe_to_write(
    plan_or_run: Union[WritePlan, RunResult],
    exceptions: Optional[Sequence[ExceptionItem]] = None,
    run_result: Optional[RunResult] = None,
    *,
    check_disk: bool = False,
) -> ValidationResult:
    """Kiểm tra an toàn và ném `SafetyError` ngay nếu có bất kỳ lỗi nào."""
    res = validate_write_plan(
        plan_or_run,
        exceptions=exceptions,
        run_result=run_result,
        check_disk=check_disk,
    )
    res.ensure_safe()
    return res


# =============================================================================
# 2. Chi tiết từng chốt kiểm tra
# =============================================================================

def _check_mapping_and_structure(plan: WritePlan, add_error: Callable[..., None]) -> None:
    # 1a. Lỗi cấu trúc / mapping mơ hồ từ bước đọc file (HEADER_MISMATCH, COLUMN_CONTENT_SUSPECT, NO_RECORDS)
    for issue in plan.blocking_issues:
        add_error(
            SafetyCode.AMBIGUOUS_MAPPING,
            f"Lỗi cấu trúc/mapping chưa khắc phục: {issue.message}",
            sheet=issue.sheet,
            row=issue.row,
            column=issue.column,
            card_no=issue.card_no,
        )

    if not plan.source_file or not plan.source_file.strip():
        add_error(SafetyCode.AMBIGUOUS_MAPPING, "Thiếu thông tin đường dẫn file lương nguồn.")
    if not plan.source_sheet or not plan.source_sheet.strip():
        add_error(SafetyCode.AMBIGUOUS_MAPPING, "Thiếu thông tin tên sheet bảng lương nguồn.")

    if not plan.target_plans:
        add_error(SafetyCode.AMBIGUOUS_MAPPING, "WritePlan không có kế hoạch ghi cho bất kỳ file ngân hàng nào.")
        return

    seen_target_files: set[str] = set()
    flattened_from_targets: list[CellUpdate] = []

    for tp in plan.target_plans:
        if not tp.file_path or not tp.file_path.strip():
            add_error(SafetyCode.AMBIGUOUS_MAPPING, "Có kế hoạch file đích bị trống đường dẫn file.")
            continue
        if not tp.sheet_name or not tp.sheet_name.strip():
            add_error(
                SafetyCode.AMBIGUOUS_MAPPING,
                f"File đích '{tp.file_path}' bị trống tên sheet nhập liệu.",
                file_path=tp.file_path,
            )
        # Kiểm tra cột Số tiền hợp lệ
        norm_col = ""
        try:
            norm_col = column_letter(column_index(tp.amount_column))
        except Exception:
            add_error(
                SafetyCode.AMBIGUOUS_MAPPING,
                f"File đích '{tp.file_path}' có cột Số tiền không hợp lệ: {tp.amount_column!r}.",
                file_path=tp.file_path,
                sheet=tp.sheet_name,
            )

        norm_path = _norm_path_key(tp.file_path)
        if norm_path in seen_target_files:
            add_error(
                SafetyCode.AMBIGUOUS_MAPPING,
                f"File ngân hàng '{tp.file_path}' xuất hiện nhiều lần trong danh sách kế hoạch ghi.",
                file_path=tp.file_path,
                sheet=tp.sheet_name,
            )
        seen_target_files.add(norm_path)

        for u in tp.updates:
            flattened_from_targets.append(u)
            if u.file_path != tp.file_path or u.sheet_name != tp.sheet_name:
                add_error(
                    SafetyCode.INCONSISTENT_DATA,
                    f"Ô {u.cell} có file/sheet ('{u.file_path}' / '{u.sheet_name}') "
                    f"không khớp với TargetWritePlan ('{tp.file_path}' / '{tp.sheet_name}').",
                    file_path=u.file_path,
                    sheet=u.sheet_name,
                    row=u.row,
                    column=u.column,
                )
            if norm_col and u.column != norm_col:
                add_error(
                    SafetyCode.AMBIGUOUS_MAPPING,
                    f"Ô {u.cell} trong '{Path(tp.file_path).name}' ghi vào cột '{u.column}' "
                    f"khác với cột Số tiền đã cấu hình ('{norm_col}').",
                    file_path=u.file_path,
                    sheet=u.sheet_name,
                    row=u.row,
                    column=u.column,
                )

    # Kiểm tra plan.updates và tổng hợp tp.updates phải chứa đúng cùng tập hợp các ô cập nhật
    if len(flattened_from_targets) != len(plan.updates) or set(flattened_from_targets) != set(plan.updates):
        add_error(
            SafetyCode.INCONSISTENT_DATA,
            f"Danh sách ô cập nhật tổng (`plan.updates`: {len(plan.updates)} ô) không khớp "
            f"với các `target_plans` ({len(flattened_from_targets)} ô).",
        )

    # Kiểm tra trùng tọa độ ô đích (2 lệnh cùng ghi vào 1 ô Excel)
    by_coord: dict[tuple[str, str, int, str], list[CellUpdate]] = defaultdict(list)
    for u in plan.updates:
        if u.row < 1:
            add_error(
                SafetyCode.AMBIGUOUS_MAPPING,
                f"Ô cập nhật có số dòng không hợp lệ ({u.row}) trong '{u.file_path}'.",
                file_path=u.file_path,
                sheet=u.sheet_name,
                row=u.row,
                column=u.column,
            )
        by_coord[(u.file_path, u.sheet_name, u.row, u.column)].append(u)

    for (fpath, sname, r, col), items in by_coord.items():
        if len(items) > 1:
            add_error(
                SafetyCode.AMBIGUOUS_MAPPING,
                f"Ô {col}{r} tại sheet '{sname}' ('{Path(fpath).name}') bị lên kế hoạch ghi "
                f"{len(items)} lần (mapping/đối chiếu mơ hồ).",
                file_path=fpath,
                sheet=sname,
                row=r,
                column=col,
            )


def _check_exceptions(
    plan: WritePlan,
    exceptions: list[ExceptionItem],
    add_error: Callable[..., None],
    add_warning: Callable[..., None],
) -> None:
    seen_ids: set[str] = set()
    for exc in exceptions:
        if exc.id in seen_ids:
            add_error(
                SafetyCode.INCONSISTENT_DATA,
                f"Trùng mã Exception ID '{exc.id}' trong danh sách ngoại lệ.",
            )
        seen_ids.add(exc.id)

        if exc.code == ExceptionCode.X8_DUPLICATE_KEY:
            # X8 được báo cáo chi tiết trong _check_duplicate_identities
            continue

        if exc.is_blocking:
            add_error(
                SafetyCode.UNRESOLVED_EXCEPTION,
                f"Ngoại lệ chặn cứng [{exc.id} - {exc.code.value}]: {exc.message}",
                file_path=exc.target_file,
                sheet=exc.target_sheet,
                row=exc.target_row or exc.source_row,
                card_no=exc.card_no,
                account_no=exc.account_no,
            )
        elif not exc.is_resolved:
            add_error(
                SafetyCode.UNRESOLVED_EXCEPTION,
                f"Ngoại lệ [{exc.id} - {exc.code.value}] chưa được người dùng xử lý: {exc.message}",
                file_path=exc.target_file,
                sheet=exc.target_sheet,
                row=exc.target_row or exc.source_row,
                card_no=exc.card_no,
                account_no=exc.account_no,
            )
        elif exc.action in (ResolutionAction.CONFIRM_WRITE, ResolutionAction.OVERRIDE_AMOUNT):
            add_warning(
                SafetyCode.ABNORMAL_CHANGE,
                f"Ngoại lệ [{exc.id} - {exc.code.value}] đã được người dùng xác nhận ghi "
                f"({exc.action.value}): {exc.message}",
                file_path=exc.target_file,
                sheet=exc.target_sheet,
                row=exc.target_row or exc.source_row,
                card_no=exc.card_no,
                account_no=exc.account_no,
            )

    # Kiểm tra các dòng MatchItem còn đang ở trạng thái PENDING hoặc ERROR
    for m in plan.matches:
        if m.status == MatchStatus.PENDING:
            add_error(
                SafetyCode.UNRESOLVED_EXCEPTION,
                f"Dòng đối chiếu (Số thẻ: {m.card_no or '-'}, TK: {m.account_no or '-'}) "
                "vẫn đang ở trạng thái chờ xác nhận (PENDING).",
                file_path=m.target_file,
                sheet=m.target_sheet,
                row=m.target_row or m.source_row,
                card_no=m.card_no,
                account_no=m.account_no,
            )
        elif m.status == MatchStatus.ERROR:
            add_error(
                SafetyCode.INCONSISTENT_DATA,
                f"Dòng đối chiếu (Số thẻ: {m.card_no or '-'}, TK: {m.account_no or '-'}) "
                f"đang ở trạng thái lỗi (ERROR): {m.note}",
                file_path=m.target_file,
                sheet=m.target_sheet,
                row=m.target_row or m.source_row,
                card_no=m.card_no,
                account_no=m.account_no,
            )


def _check_duplicate_identities(
    plan: WritePlan,
    exceptions: list[ExceptionItem],
    add_error: Callable[..., None],
) -> None:
    # 1. Từ danh sách Exception X8
    for exc in exceptions:
        if exc.code == ExceptionCode.X8_DUPLICATE_KEY:
            add_error(
                SafetyCode.DUPLICATE_IDENTITY,
                f"Trùng định danh [{exc.id}]: {exc.message}",
                file_path=exc.target_file,
                sheet=exc.target_sheet,
                row=exc.target_row or exc.source_row,
                card_no=exc.card_no,
                account_no=exc.account_no,
            )

    # 2. Kiểm tra trực tiếp trên danh sách CellUpdate (chống lọt trùng Số thẻ / Số TK khi ghi)
    by_card: dict[str, list[CellUpdate]] = defaultdict(list)
    by_acc: dict[str, list[CellUpdate]] = defaultdict(list)
    for u in plan.updates:
        if u.card_no:
            by_card[u.card_no].append(u)
        if u.account_no:
            by_acc[u.account_no].append(u)

    for card, u_list in by_card.items():
        if len(u_list) > 1:
            cells = ", ".join(f"{Path(u.file_path).name}!{u.cell}" for u in u_list)
            add_error(
                SafetyCode.DUPLICATE_IDENTITY,
                f"Số thẻ '{card}' được lên kế hoạch ghi vào nhiều ô đích: {cells}.",
                card_no=card,
            )

    for acc, u_list in by_acc.items():
        if len(u_list) > 1:
            cells = ", ".join(f"{Path(u.file_path).name}!{u.cell}" for u in u_list)
            add_error(
                SafetyCode.DUPLICATE_IDENTITY,
                f"Số tài khoản '{acc}' được lên kế hoạch ghi vào nhiều ô đích: {cells}.",
                account_no=acc,
            )

    # 3. Kiểm tra trùng Số thẻ / Số TK trong danh sách MatchItem phía nguồn
    src_cards: dict[str, list[MatchItem]] = defaultdict(list)
    src_accs: dict[str, list[MatchItem]] = defaultdict(list)
    for m in plan.matches:
        if m.source_row is not None:
            if m.card_no:
                src_cards[m.card_no].append(m)
            if m.account_no:
                src_accs[m.account_no].append(m)

    already_reported_cards = {
        e.card_no for e in exceptions if e.code == ExceptionCode.X8_DUPLICATE_KEY and e.card_no
    }
    already_reported_accs = {
        e.account_no for e in exceptions if e.code == ExceptionCode.X8_DUPLICATE_KEY and e.account_no
    }

    for card, m_list in src_cards.items():
        if len(m_list) > 1 and card not in already_reported_cards:
            rows = [m.source_row for m in m_list]
            add_error(
                SafetyCode.DUPLICATE_IDENTITY,
                f"Nhiều dòng bảng lương có cùng Số thẻ '{card}': dòng {rows}.",
                card_no=card,
            )

    for acc, m_list in src_accs.items():
        if len(m_list) > 1 and acc not in already_reported_accs:
            cards = [m.card_no for m in m_list]
            add_error(
                SafetyCode.DUPLICATE_IDENTITY,
                f"Nhiều nhân sự có cùng Số tài khoản '{acc}': Số thẻ {cards}.",
                account_no=acc,
            )


def _check_updates_and_matches(plan: WritePlan, add_error: Callable[..., None]) -> None:
    for u in plan.updates:
        # 1. Phải có Số tài khoản hợp lệ (toàn chữ số) ở mọi ô dự kiến ghi
        acc = (u.account_no or "").strip()
        if not acc:
            add_error(
                SafetyCode.MISSING_ACCOUNT,
                f"Ô {u.cell} tại '{Path(u.file_path).name}' dự kiến ghi nhưng thiếu Số tài khoản.",
                file_path=u.file_path,
                sheet=u.sheet_name,
                row=u.row,
                column=u.column,
                card_no=u.card_no,
            )
        elif not acc.isdigit():
            add_error(
                SafetyCode.MISSING_ACCOUNT,
                f"Ô {u.cell} tại '{Path(u.file_path).name}' có Số tài khoản không hợp lệ: '{u.account_no}'.",
                file_path=u.file_path,
                sheet=u.sheet_name,
                row=u.row,
                column=u.column,
                card_no=u.card_no,
                account_no=u.account_no,
            )

        # 2. Kiểm tra số tiền theo lý do cập nhật
        if u.reason in _CLEAR_REASONS:
            if u.new_value is not None:
                add_error(
                    SafetyCode.INCONSISTENT_DATA,
                    f"Ô {u.cell} có lý do '{u.reason}' (xóa trắng) nhưng new_value={u.new_value!r}.",
                    file_path=u.file_path,
                    sheet=u.sheet_name,
                    row=u.row,
                    column=u.column,
                    account_no=u.account_no,
                )
        elif u.reason in _ZERO_REASONS:
            if u.new_value != 0:
                add_error(
                    SafetyCode.INCONSISTENT_DATA,
                    f"Ô {u.cell} có lý do '{u.reason}' (đặt 0) nhưng new_value={u.new_value!r}.",
                    file_path=u.file_path,
                    sheet=u.sheet_name,
                    row=u.row,
                    column=u.column,
                    account_no=u.account_no,
                )
        else:
            # Lệnh chuyển khoản thông thường: bắt buộc có Số thẻ và số tiền nguyên > 0
            if not u.card_no or not u.card_no.strip():
                add_error(
                    SafetyCode.INCONSISTENT_DATA,
                    f"Ô {u.cell} ('{Path(u.file_path).name}') cập nhật lương nhưng thiếu Số thẻ nguồn.",
                    file_path=u.file_path,
                    sheet=u.sheet_name,
                    row=u.row,
                    column=u.column,
                    account_no=u.account_no,
                )
            if (
                u.new_value is None
                or isinstance(u.new_value, bool)
                or not isinstance(u.new_value, int)
                or u.new_value <= 0
            ):
                add_error(
                    SafetyCode.MISSING_OR_INVALID_AMOUNT,
                    f"Ô {u.cell} ('{Path(u.file_path).name}', TK '{u.account_no or ''}') "
                    f"thiếu số tiền hoặc số tiền không lớn hơn 0: {u.new_value!r}.",
                    file_path=u.file_path,
                    sheet=u.sheet_name,
                    row=u.row,
                    column=u.column,
                    card_no=u.card_no,
                    account_no=u.account_no,
                )

    # 3. Kiểm tra các dòng MatchItem phía nguồn ở trạng thái cập nhật nhưng thiếu TK hoặc thiếu tiền
    for m in plan.matches:
        if m.source_row is not None and m.status in (MatchStatus.WILL_UPDATE, MatchStatus.UNCHANGED):
            if not m.account_no or not m.account_no.strip():
                add_error(
                    SafetyCode.MISSING_ACCOUNT,
                    f"Dòng nguồn {m.source_row} (Số thẻ '{m.card_no}') ở trạng thái '{m.status.value}' "
                    "nhưng thiếu Số tài khoản.",
                    row=m.source_row,
                    card_no=m.card_no,
                )
            if m.new_amount is not None and m.new_amount < 0:
                add_error(
                    SafetyCode.MISSING_OR_INVALID_AMOUNT,
                    f"Dòng nguồn {m.source_row} (Số thẻ '{m.card_no}') có số tiền âm: {m.new_amount}.",
                    row=m.source_row,
                    card_no=m.card_no,
                    account_no=m.account_no,
                )


def _check_consistency_and_balance(plan: WritePlan, add_error: Callable[..., None]) -> tuple[bool, bool]:
    """Kiểm tra nhất quán giữa MatchItem <-> CellUpdate và cân đối số dòng + tổng tiền (K8)."""
    row_ok = True
    amt_ok = True

    update_by_target_row: dict[tuple[str, str, int], CellUpdate] = {
        (u.file_path, u.sheet_name, u.row): u for u in plan.updates
    }
    matched_target_coords: set[tuple[str, str, int]] = set()

    source_matches = [m for m in plan.matches if m.source_row is not None]
    written_source_count = 0
    skipped_source_count = 0
    written_source_base_amount = 0
    skipped_transferable_amount = 0

    for m in source_matches:
        has_valid_transferable_amt = (
            m.account_no is not None
            and m.source_amount is not None
            and isinstance(m.source_amount, int)
            and m.source_amount > 0
        )
        if m.status == MatchStatus.SKIPPED:
            skipped_source_count += 1
            if has_valid_transferable_amt and m.source_amount is not None:
                skipped_transferable_amount += m.source_amount
            if m.target_file and m.target_sheet and m.target_row is not None:
                coord = (m.target_file, m.target_sheet, m.target_row)
                if coord in update_by_target_row:
                    row_ok = False
                    add_error(
                        SafetyCode.INCONSISTENT_DATA,
                        f"Dòng nguồn {m.source_row} (Số thẻ '{m.card_no}') đã chọn SKIP "
                        f"nhưng vẫn có lệnh ghi tại ô {update_by_target_row[coord].cell}.",
                        file_path=m.target_file,
                        sheet=m.target_sheet,
                        row=m.target_row,
                        card_no=m.card_no,
                        account_no=m.account_no,
                    )
            continue

        if m.status in (MatchStatus.PENDING, MatchStatus.ERROR):
            continue

        # Trạng thái WILL_UPDATE hoặc UNCHANGED cho dòng nguồn
        if not m.target_file or not m.target_sheet or m.target_row is None:
            row_ok = False
            add_error(
                SafetyCode.INCONSISTENT_DATA,
                f"Dòng nguồn {m.source_row} (Số thẻ '{m.card_no}') ở trạng thái '{m.status.value}' "
                "nhưng không gắn với dòng file đích nào.",
                row=m.source_row,
                card_no=m.card_no,
                account_no=m.account_no,
            )
            continue

        coord = (m.target_file, m.target_sheet, m.target_row)
        matched_target_coords.add(coord)
        u = update_by_target_row.get(coord)
        if u is None:
            row_ok = False
            add_error(
                SafetyCode.INCONSISTENT_DATA,
                f"Dòng nguồn {m.source_row} (Số thẻ '{m.card_no}', TK '{m.account_no}') "
                f"ở trạng thái '{m.status.value}' nhưng không có CellUpdate tương ứng.",
                file_path=m.target_file,
                sheet=m.target_sheet,
                row=m.target_row,
                card_no=m.card_no,
                account_no=m.account_no,
            )
            continue

        written_source_count += 1
        if u.card_no != m.card_no or u.account_no != m.account_no:
            row_ok = False
            add_error(
                SafetyCode.INCONSISTENT_DATA,
                f"Ô {u.cell} có định danh (Số thẻ '{u.card_no}', TK '{u.account_no}') "
                f"không khớp với MatchItem (Số thẻ '{m.card_no}', TK '{m.account_no}').",
                file_path=u.file_path,
                sheet=u.sheet_name,
                row=u.row,
                column=u.column,
            )

        if u.new_value != m.new_amount:
            amt_ok = False
            add_error(
                SafetyCode.AMOUNT_BALANCE_MISMATCH,
                f"Ô {u.cell} có new_value={u.new_value} không khớp với MatchItem.new_amount={m.new_amount}.",
                file_path=u.file_path,
                sheet=u.sheet_name,
                row=u.row,
                column=u.column,
                card_no=m.card_no,
                account_no=m.account_no,
            )

        if u.reason in ("MATCHED", "CONFIRMED_MISMATCH"):
            if u.new_value != m.source_amount:
                amt_ok = False
                add_error(
                    SafetyCode.AMOUNT_BALANCE_MISMATCH,
                    f"Ô {u.cell} (Số thẻ '{m.card_no}') ghi số tiền {u.new_value} "
                    f"lệch với số tiền gốc trong bảng lương ({m.source_amount}).",
                    file_path=u.file_path,
                    sheet=u.sheet_name,
                    row=u.row,
                    column=u.column,
                    card_no=m.card_no,
                    account_no=m.account_no,
                )
            if has_valid_transferable_amt and m.source_amount is not None:
                written_source_base_amount += m.source_amount
        elif u.reason in _ZERO_OR_CLEAR_REASONS:
            if has_valid_transferable_amt and m.source_amount is not None:
                skipped_transferable_amount += m.source_amount
        elif u.reason == "OVERRIDE_AMOUNT":
            if has_valid_transferable_amt and m.source_amount is not None:
                written_source_base_amount += m.source_amount

    # Kiểm tra các dòng đích X5 (source_row is None) có CellUpdate
    for m in plan.matches:
        if m.source_row is None and m.target_file and m.target_sheet and m.target_row is not None:
            coord = (m.target_file, m.target_sheet, m.target_row)
            if coord in update_by_target_row:
                matched_target_coords.add(coord)

    # Mọi CellUpdate trong plan.updates đều phải bắt nguồn từ 1 MatchItem hợp lệ
    orphan_coords = set(update_by_target_row.keys()) - matched_target_coords
    for fpath, sname, r in sorted(orphan_coords):
        row_ok = False
        u_orp = update_by_target_row[(fpath, sname, r)]
        add_error(
            SafetyCode.INCONSISTENT_DATA,
            f"Ô {u_orp.cell} tại '{Path(fpath).name}' nằm trong WritePlan "
            "nhưng không tương ứng với bất kỳ dòng đối chiếu (MatchItem) nào.",
            file_path=fpath,
            sheet=sname,
            row=r,
            column=u_orp.column,
        )

    # Kiểm tra cân đối tổng số dòng nguồn (K8): nguồn = cập nhật + bỏ qua
    if plan.total_source_records != len(source_matches):
        row_ok = False
        add_error(
            SafetyCode.ROW_BALANCE_MISMATCH,
            f"Lệch số dòng nguồn: bảng lương có {plan.total_source_records} dòng có Số thẻ, "
            f"nhưng bảng đối chiếu có {len(source_matches)} dòng nguồn.",
        )
    elif not any(m.status in (MatchStatus.PENDING, MatchStatus.ERROR) for m in source_matches):
        accounted_rows = written_source_count + skipped_source_count
        if accounted_rows != plan.total_source_records:
            row_ok = False
            add_error(
                SafetyCode.ROW_BALANCE_MISMATCH,
                f"Lệch cân đối số dòng (K8): tổng dòng nguồn ({plan.total_source_records}) != "
                f"số dòng ghi ({written_source_count}) + số dòng bỏ qua ({skipped_source_count}).",
            )

    # Kiểm tra cân đối tổng tiền chuyển khoản nguồn (K8): Σ đủ điều kiện = Σ gốc đã ghi + Σ gốc bỏ qua
    if not any(m.status in (MatchStatus.PENDING, MatchStatus.ERROR) for m in source_matches):
        accounted_amount = written_source_base_amount + skipped_transferable_amount
        if accounted_amount != plan.total_source_transferable_amount:
            amt_ok = False
            add_error(
                SafetyCode.AMOUNT_BALANCE_MISMATCH,
                f"Lệch cân đối tổng tiền (K8): tổng tiền chuyển khoản bảng lương "
                f"({plan.total_source_transferable_amount:,} đ) != "
                f"tổng tiền gốc đã phân bổ ghi ({written_source_base_amount:,} đ) + "
                f"tổng tiền ngoại lệ bỏ qua ({skipped_transferable_amount:,} đ).",
            )

    return row_ok, amt_ok


def _check_abnormal_changes(
    plan: WritePlan,
    add_error: Callable[..., None],
    add_warning: Callable[..., None],
) -> None:
    if not plan.updates:
        add_error(
            SafetyCode.ABNORMAL_CHANGE,
            "Không có ô dữ liệu nào được lên kế hoạch ghi vào file ngân hàng.",
        )
        return

    for u in plan.updates:
        # 1. Cảnh báo dòng được đặt 0 hoặc xóa trắng (X5/X7)
        if u.reason in _ZERO_OR_CLEAR_REASONS:
            action_label = "đặt số tiền = 0" if u.reason in _ZERO_REASONS else "xóa trắng số tiền"
            add_warning(
                SafetyCode.ABNORMAL_CHANGE,
                f"Ô {u.cell} ('{Path(u.file_path).name}', TK '{u.account_no or ''}') được {action_label}; "
                "lưu ý một số biểu mẫu ngân hàng có thể từ chối dòng 0 đồng/trống khi upload.",
                file_path=u.file_path,
                sheet=u.sheet_name,
                row=u.row,
                column=u.column,
                account_no=u.account_no,
            )
            continue

        if u.new_value is None:
            continue

        # 2. Số tiền 1 người vượt ngưỡng 1 tỷ VND
        if u.new_value > AMOUNT_WARNING_THRESHOLD:
            add_warning(
                SafetyCode.ABNORMAL_CHANGE,
                f"Ô {u.cell} ('{Path(u.file_path).name}', Số thẻ '{u.card_no}', TK '{u.account_no}') "
                f"có số tiền lớn bất thường: {u.new_value:,} đ (> {AMOUNT_WARNING_THRESHOLD:,} đ).",
                file_path=u.file_path,
                sheet=u.sheet_name,
                row=u.row,
                column=u.column,
                card_no=u.card_no,
                account_no=u.account_no,
            )

        # 3. Biến động đột biến của từng cá nhân so với số tiền cũ trong file đích
        if u.old_value is not None and u.old_value > 0 and u.new_value > 0:
            ratio = u.new_value / u.old_value
            if ratio >= ABNORMAL_INCREASE_RATIO:
                add_warning(
                    SafetyCode.ABNORMAL_CHANGE,
                    f"Ô {u.cell} ('{Path(u.file_path).name}', TK '{u.account_no}') tăng đột biến "
                    f"gấp {ratio:.1f} lần so với số tiền cũ ({u.old_value:,} đ -> {u.new_value:,} đ).",
                    file_path=u.file_path,
                    sheet=u.sheet_name,
                    row=u.row,
                    column=u.column,
                    card_no=u.card_no,
                    account_no=u.account_no,
                )
            elif ratio <= ABNORMAL_DECREASE_RATIO:
                add_warning(
                    SafetyCode.ABNORMAL_CHANGE,
                    f"Ô {u.cell} ('{Path(u.file_path).name}', TK '{u.account_no}') giảm đột biến "
                    f"còn {ratio * 100:.1f}% so với số tiền cũ ({u.old_value:,} đ -> {u.new_value:,} đ).",
                    file_path=u.file_path,
                    sheet=u.sheet_name,
                    row=u.row,
                    column=u.column,
                    card_no=u.card_no,
                    account_no=u.account_no,
                )

    # 4. Biến động tổng tiền của từng file đích
    for tp in plan.target_plans:
        if tp.updates and tp.total_old_amount > 0 and tp.total_new_amount > 0:
            delta_ratio = abs(tp.total_new_amount - tp.total_old_amount) / tp.total_old_amount
            if delta_ratio >= ABNORMAL_TOTAL_DELTA_RATIO:
                add_warning(
                    SafetyCode.ABNORMAL_CHANGE,
                    f"Tổng tiền của file '{Path(tp.file_path).name}' thay đổi {delta_ratio * 100:.1f}% "
                    f"so với kỳ trước ({tp.total_old_amount:,} đ -> {tp.total_new_amount:,} đ).",
                    file_path=tp.file_path,
                    sheet=tp.sheet_name,
                )


def _check_run_result_paths(
    plan: WritePlan,
    run_result: RunResult,
    add_error: Callable[..., None],
    add_warning: Callable[..., None],
) -> None:
    """Kiểm tra quy tắc R8: Không bao giờ ghi đè file gốc, phải có backup trước khi xuất."""
    original_paths: set[str] = set()
    if plan.source_file:
        original_paths.add(_norm_path_key(plan.source_file))
    for tp in plan.target_plans:
        if tp.file_path:
            original_paths.add(_norm_path_key(tp.file_path))

    seen_outputs: set[str] = set()
    for out_file in run_result.output_files:
        norm_out = _norm_path_key(out_file)
        if norm_out in original_paths:
            add_error(
                SafetyCode.OVERWRITE_ORIGINAL_FORBIDDEN,
                f"Vi phạm quy tắc an toàn R8: đường dẫn xuất file '{out_file}' trùng với file gốc.",
                file_path=out_file,
            )
        if norm_out in seen_outputs:
            add_error(
                SafetyCode.AMBIGUOUS_MAPPING,
                f"Đường dẫn xuất file '{out_file}' bị trùng lặp cho nhiều file đích.",
                file_path=out_file,
            )
        seen_outputs.add(norm_out)

    if run_result.output_files and not run_result.backup_files:
        add_warning(
            SafetyCode.MISSING_BACKUP,
            "Chưa ghi nhận đường dẫn file backup cho lần chạy này (R8).",
        )


# =============================================================================
# 3. Kiểm tra khóa file trên đĩa (K9 - File đang mở trong Excel)
# =============================================================================

def is_file_locked(path: Union[str, os.PathLike[str]]) -> bool:
    """Trả về True nếu file đang bị khóa bởi tiến trình khác (ví dụ đang mở trong Excel).

    Hàm chỉ kiểm tra quyền truy cập và file tạm `~$`, tuyệt đối không sửa nội dung file.
    """
    p = Path(path)
    if not p.exists() or p.is_dir():
        return False

    # Excel tạo file khóa tạm ~$<ten_file> cùng thư mục khi đang mở
    lock_file = p.parent / f"~${p.name}"
    if lock_file.exists():
        return True

    try:
        st_mode = p.stat().st_mode
        flags = os.O_RDWR if (st_mode & stat.S_IWRITE) else os.O_RDONLY
        fd = os.open(str(p), flags)
    except OSError:
        return True
    else:
        os.close(fd)
        return False


def check_file_locks(paths: Iterable[Union[str, os.PathLike[str]]]) -> list[SafetyFinding]:
    """Kiểm tra danh sách đường dẫn file và trả về các lỗi `FILE_LOCKED` nếu file đang mở."""
    findings: list[SafetyFinding] = []
    for raw in paths:
        p = Path(raw)
        if is_file_locked(p):
            findings.append(
                SafetyFinding(
                    code=SafetyCode.FILE_LOCKED,
                    severity=Severity.ERROR,
                    message=f"File '{p.name}' đang được mở trong Excel hoặc bị khóa bởi chương trình khác. "
                    "Hãy đóng file trước khi tiếp tục.",
                    file_path=str(p),
                )
            )
    return findings


def _check_disk_locks(
    plan: WritePlan,
    run_result: Optional[RunResult],
    result: ValidationResult,
) -> None:
    paths_to_check: list[str] = [tp.file_path for tp in plan.target_plans if tp.file_path]
    if run_result is not None:
        paths_to_check.extend(run_result.output_files)
    result.errors.extend(check_file_locks(paths_to_check))


def _norm_path_key(path_str: str) -> str:
    try:
        return os.path.normcase(str(Path(path_str).resolve()))
    except OSError:
        return os.path.normcase(os.path.abspath(path_str))
