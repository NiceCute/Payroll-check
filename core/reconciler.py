"""
reconciler.py - Đối chiếu dữ liệu bảng lương tháng mới với 1 hoặc nhiều file mẫu ngân hàng.

Tuân thủ nghiêm ngặt:
- R1: Chỉ đối chiếu các dòng nguồn có Số thẻ (SourceRecord).
- R2: Khóa chính nhân sự = Số thẻ -> tra Master (nếu có) hoặc lấy Số TK từ dòng lương.
- R3: Định vị dòng trong file đích bằng Số tài khoản. Tên nhân viên CHỈ dùng để kiểm tra chéo.
- R4: Định danh (Số thẻ, Số TK, Mã NH) luôn là chuỗi, giữ số 0 đầu, KHÔNG BAO GIỜ ghi đè.
- R5: Số tiền phải là số nguyên > 0. Số tiền trống / <= 0 / lẻ / lỗi -> Exception X7 (không làm tròn).
- R6: WritePlan chỉ cập nhật cột Số tiền đã mapping; tuyệt đối không ghi đè ô chứa công thức.
- R7: Mọi bất thường -> ExceptionItem; chặn xuất file cho tới khi xử lý hết.
- R8: Thuần bộ nhớ, không chạm vào file gốc trên đĩa.
"""

from __future__ import annotations

import logging
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Mapping, Optional, Sequence, Union

from .models import (
    BLOCKING_CODES,
    CellUpdate,
    ConfigError,
    ExceptionCode,
    ExceptionItem,
    Issue,
    IssueCode,
    MasterRecord,
    MatchItem,
    MatchStatus,
    ReconcileResult,
    ResolutionAction,
    Severity,
    SourceReadResult,
    SourceRecord,
    TargetReadResult,
    TargetRow,
    TargetRules,
    TargetTemplate,
    TargetWritePlan,
    ToolError,
    WritePlan,
    column_index,
    column_letter,
)
from .normalize import clean_text, normalize_name_key

log = logging.getLogger(__name__)

__all__ = [
    "CellUpdate",
    "ExceptionCode",
    "ExceptionItem",
    "MasterRecord",
    "MatchItem",
    "MatchStatus",
    "ReconcileError",
    "ReconcileResult",
    "ResolutionAction",
    "TargetInput",
    "TargetSpec",
    "TargetWritePlan",
    "WritePlan",
    "apply_resolutions",
    "rebuild_plan",
    "reconcile",
    "resolve_exception",
]

# Cột Số tiền mặc định cho 2 mẫu chuẩn khi người gọi chỉ truyền TargetReadResult
_KNOWN_TEMPLATE_AMOUNT_COLUMNS: dict[str, str] = {
    "ACB - Chi hộ lương qua tài khoản ACB": "D",
    "Thông thường - Chi hộ lương liên ngân hàng": "E",
}


class ReconcileError(ToolError):
    """Lỗi khi thực hiện đối chiếu hoặc cập nhật phương án xử lý Exception."""


@dataclass(frozen=True)
class TargetInput:
    """Gói thông tin 1 file đích đưa vào bộ đối chiếu."""

    target: TargetReadResult
    template: Optional[TargetTemplate] = None
    amount_column: str = ""

    def resolved_amount_column(self) -> str:
        """Xác định chữ cái cột Số tiền cần ghi và kiểm tra an toàn R4/R6."""
        if self.template is not None:
            self.template.ensure_valid()
            col = self.template.amount.letter
            # R4 & R6: Cột ghi số tiền không bao giờ được trùng với cột định danh
            forbidden = {
                self.template.account_no.letter,
                self.template.full_name.letter,
            }
            if self.template.bank_code is not None:
                forbidden.add(self.template.bank_code.letter)
            if col in forbidden:
                raise ConfigError(
                    f"Mẫu '{self.template.name}': cột Số tiền ({col}) trùng với cột định danh."
                )
            return col

        if self.amount_column.strip():
            return column_letter(column_index(self.amount_column))

        by_name = _KNOWN_TEMPLATE_AMOUNT_COLUMNS.get(self.target.template_name)
        if by_name:
            return by_name

        # Suy luận an toàn từ cấu trúc dòng đã đọc (có bank_code -> mẫu liên NH cột E, ngược lại cột D)
        if any(r.bank_code is not None for r in self.target.rows):
            return "E"
        return "D"

    @property
    def rules(self) -> TargetRules:
        return self.template.rules if self.template is not None else TargetRules()


TargetSpec = Union[
    TargetReadResult,
    TargetInput,
    tuple[TargetReadResult, TargetTemplate],
    tuple[TargetReadResult, str],
]

ResolutionValue = Union[ResolutionAction, tuple[ResolutionAction, Optional[int]]]


# =============================================================================
# 1. Hàm đối chiếu chính
# =============================================================================

def reconcile(
    source: SourceReadResult,
    targets: Union[TargetSpec, Sequence[TargetSpec]],
    *,
    templates: Optional[Mapping[str, TargetTemplate]] = None,
    master: Optional[Mapping[str, MasterRecord]] = None,
    resolutions: Optional[Mapping[str, ResolutionValue]] = None,
) -> WritePlan:
    """Đối chiếu dữ liệu từ bảng lương (`source`) vào 1 hoặc nhiều file ngân hàng (`targets`).

    Thứ tự ưu tiên định danh (tuyệt đối không dùng tên làm khóa duy nhất):
    1. Số thẻ (`card_no`) -> lọc dòng hợp lệ (R1) & tra Master Dataset nếu có (R2).
    2. Số tài khoản (`account_no`) -> định vị dòng trong file ngân hàng (R3).
    3. Họ tên (`name_key`) & Mã ngân hàng (`bank_code`) -> chỉ dùng kiểm tra chéo (R3, X6).

    Trả về `WritePlan` (chứa `.updates`, `.target_plans`, `.matches`, `.exceptions`).
    """
    target_inputs = _normalize_target_specs(targets, templates)
    if not target_inputs:
        raise ConfigError("Phải cung cấp ít nhất 1 file ngân hàng (file đích) để đối chiếu.")

    plan = WritePlan(
        source_file=source.file_path,
        source_sheet=source.sheet_name,
        target_plans=[
            TargetWritePlan(
                file_path=ti.target.file_path,
                file_format=ti.target.file_format,
                sheet_name=ti.target.sheet_name,
                template_name=ti.target.template_name,
                amount_column=ti.resolved_amount_column(),
                total_target_rows=len(ti.target.rows),
            )
            for ti in target_inputs
        ],
        total_source_records=len(source.records),
        total_source_transferable_amount=source.total_transferable_amount,
    )

    # 1. Thu thập lỗi cấu trúc mức file (sai tiêu đề cột, cột nghi lệch, không có dữ liệu)
    _collect_structural_issues(source, target_inputs, plan)

    exc_counter = 0

    def add_exception(
        code: ExceptionCode,
        severity: Severity,
        message: str,
        *,
        card_no: Optional[str] = None,
        account_no: Optional[str] = None,
        source_name: str = "",
        target_name: str = "",
        source_row: Optional[int] = None,
        target_file: Optional[str] = None,
        target_sheet: Optional[str] = None,
        target_row: Optional[int] = None,
        old_amount: Optional[int] = None,
        source_amount: Optional[int] = None,
        allowed_actions: tuple[ResolutionAction, ...] = (),
        action: Optional[ResolutionAction] = None,
    ) -> ExceptionItem:
        nonlocal exc_counter
        exc_counter += 1
        item = ExceptionItem(
            id=f"EX-{exc_counter:04d}",
            code=code,
            severity=severity,
            message=message,
            card_no=card_no,
            account_no=account_no,
            source_name=source_name,
            target_name=target_name,
            source_row=source_row,
            target_file=target_file,
            target_sheet=target_sheet,
            target_row=target_row,
            old_amount=old_amount,
            source_amount=source_amount,
            allowed_actions=allowed_actions,
            action=action,
        )
        plan.exceptions.append(item)
        return item

    # 2. Phát hiện trùng khóa định danh X8 (trùng Số thẻ nguồn, trùng TK nguồn, trùng TK đích)
    dup_cards, dup_src_accs, dup_tgt_accs, x8_by_card, x8_by_src_acc, x8_by_tgt_acc = (
        _detect_duplicates(source, target_inputs, master, add_exception)
    )

    # 3. Lập chỉ mục dòng đích theo Số tài khoản (R3)
    target_by_acc: dict[str, tuple[TargetInput, TargetRow, str]] = {}
    numeric_stripped_lookup: dict[str, tuple[TargetInput, TargetRow]] = {}
    for ti, tp in zip(target_inputs, plan.target_plans):
        for trow in ti.target.rows:
            if trow.account_no is None:
                continue
            if trow.account_stored_as_number:
                numeric_stripped_lookup.setdefault(trow.account_no, (ti, trow))
            if trow.account_no in dup_tgt_accs:
                continue
            target_by_acc[trow.account_no] = (ti, trow, tp.amount_column)

    # Lập chỉ mục lỗi theo dòng của nguồn và đích để lấy thông báo chi tiết
    src_issues_by_row: dict[int, list[Issue]] = defaultdict(list)
    for issue in source.issues:
        if issue.row is not None:
            src_issues_by_row[issue.row].append(issue)

    tgt_issues_by_row: dict[tuple[str, str, int], list[Issue]] = defaultdict(list)
    for ti in target_inputs:
        for issue in ti.target.issues:
            if issue.row is not None:
                tgt_issues_by_row[(ti.target.file_path, ti.target.sheet_name, issue.row)].append(issue)

    matched_target_keys: set[tuple[str, str, int]] = set()

    # 4. Duyệt từng dòng nguồn có Số thẻ (R1 -> R7)
    for srec in source.records:
        row_exc_ids: list[str] = []

        # 4a. Kiểm tra trùng khóa ở nguồn (X8)
        if srec.card_no in dup_cards or (srec.account_no is not None and srec.account_no in dup_src_accs):
            if srec.card_no in x8_by_card:
                row_exc_ids.append(x8_by_card[srec.card_no].id)
            if srec.account_no is not None and srec.account_no in x8_by_src_acc:
                row_exc_ids.append(x8_by_src_acc[srec.account_no].id)
            if srec.account_no is not None and srec.account_no in target_by_acc:
                ti_dup, trow_dup, _ = target_by_acc[srec.account_no]
                matched_target_keys.add((ti_dup.target.file_path, ti_dup.target.sheet_name, trow_dup.row))
            plan.matches.append(MatchItem(
                card_no=srec.card_no,
                account_no=srec.account_no,
                source_name=srec.full_name,
                target_name="",
                bank_code=None,
                source_row=srec.row,
                target_file=None,
                target_sheet=None,
                target_row=None,
                target_cell=None,
                old_amount=None,
                new_amount=None,
                source_amount=srec.amount,
                status=MatchStatus.ERROR,
                exception_ids=tuple(dict.fromkeys(row_exc_ids)),
                note="Trùng khóa định danh trong bảng lương (X8).",
            ))
            continue

        # 4b. Tra cứu Master Dataset (R2 - chỗ cắm cho V2 khi master != None)
        effective_account = srec.account_no
        expected_name_key = srec.name_key
        expected_bank_code: Optional[str] = None
        expected_target_type: Optional[str] = None

        if master is not None:
            mrec = master.get(srec.card_no)
            if mrec is None:
                ex_x1 = add_exception(
                    ExceptionCode.X1_NOT_IN_MASTER,
                    Severity.WARNING,
                    f"Nhân viên '{srec.full_name}' (Số thẻ '{srec.card_no}', dòng {srec.row}) "
                    "chưa có trong Master Dataset.",
                    card_no=srec.card_no,
                    account_no=srec.account_no,
                    source_name=srec.full_name,
                    source_row=srec.row,
                    source_amount=srec.amount,
                    allowed_actions=(ResolutionAction.CONFIRM_WRITE, ResolutionAction.SKIP),
                    action=None,
                )
                row_exc_ids.append(ex_x1.id)
            else:
                if srec.account_no is not None and srec.account_no != mrec.account_no:
                    ex_x2 = add_exception(
                        ExceptionCode.X2_ACCOUNT_MISMATCH_MASTER,
                        Severity.WARNING,
                        f"Số tài khoản của '{srec.full_name}' (Số thẻ '{srec.card_no}', dòng {srec.row}) "
                        f"trong bảng lương ('{srec.account_no}') khác với Master Dataset ('{mrec.account_no}').",
                        card_no=srec.card_no,
                        account_no=mrec.account_no,
                        source_name=srec.full_name,
                        source_row=srec.row,
                        source_amount=srec.amount,
                        allowed_actions=(ResolutionAction.CONFIRM_WRITE, ResolutionAction.SKIP),
                        action=None,
                    )
                    row_exc_ids.append(ex_x2.id)
                effective_account = mrec.account_no or srec.account_no
                if mrec.full_name:
                    expected_name_key = normalize_name_key(mrec.full_name)
                expected_bank_code = mrec.bank_code
                expected_target_type = mrec.target_type

        # 4c. Kiểm tra có Số tài khoản hợp lệ hay không (X3)
        if effective_account is None:
            acc_detail = _describe_source_account_issue(src_issues_by_row.get(srec.row, []))
            ex_x3 = add_exception(
                ExceptionCode.X3_MISSING_ACCOUNT,
                Severity.WARNING,
                f"Nhân viên '{srec.full_name or '(trống tên)'}' (Số thẻ '{srec.card_no}', dòng {srec.row}): "
                f"{acc_detail}. Không thể đối chiếu vào file ngân hàng.",
                card_no=srec.card_no,
                account_no=None,
                source_name=srec.full_name,
                source_row=srec.row,
                source_amount=srec.amount,
                allowed_actions=(ResolutionAction.SKIP,),
                action=ResolutionAction.SKIP,
            )
            row_exc_ids.append(ex_x3.id)
            plan.matches.append(MatchItem(
                card_no=srec.card_no,
                account_no=None,
                source_name=srec.full_name,
                target_name="",
                bank_code=None,
                source_row=srec.row,
                target_file=None,
                target_sheet=None,
                target_row=None,
                target_cell=None,
                old_amount=None,
                new_amount=None,
                source_amount=srec.amount,
                status=MatchStatus.SKIPPED,
                exception_ids=tuple(row_exc_ids),
                note=ex_x3.message,
            ))
            continue

        # 4d. Kiểm tra Số tài khoản bị trùng trong các file đích (X8)
        if effective_account in dup_tgt_accs:
            ex_x8_tgt = x8_by_tgt_acc[effective_account]
            row_exc_ids.append(ex_x8_tgt.id)
            plan.matches.append(MatchItem(
                card_no=srec.card_no,
                account_no=effective_account,
                source_name=srec.full_name,
                target_name="",
                bank_code=None,
                source_row=srec.row,
                target_file=ex_x8_tgt.target_file,
                target_sheet=ex_x8_tgt.target_sheet,
                target_row=ex_x8_tgt.target_row,
                target_cell=None,
                old_amount=None,
                new_amount=None,
                source_amount=srec.amount,
                status=MatchStatus.ERROR,
                exception_ids=tuple(row_exc_ids),
                note="Số tài khoản bị trùng trong file ngân hàng (X8).",
            ))
            continue

        # 4e. Định vị dòng trong file ngân hàng bằng Số tài khoản (R3)
        row_src_issues = src_issues_by_row.get(srec.row, [])
        target_hit = target_by_acc.get(effective_account)
        if target_hit is None:
            # Nếu đồng thời số tiền nguồn không hợp lệ -> báo X7; nếu số tiền hợp lệ -> báo X4
            if not _has_valid_positive_amount(srec, row_src_issues):
                amt_detail = _describe_source_amount_issue(srec, row_src_issues)
                ex_x7_no_tgt = add_exception(
                    ExceptionCode.X7_INVALID_AMOUNT,
                    Severity.WARNING,
                    f"Nhân viên '{srec.full_name}' (Số thẻ '{srec.card_no}', TK '{effective_account}', "
                    f"dòng {srec.row}): {amt_detail} (đồng thời không có dòng trong file ngân hàng).",
                    card_no=srec.card_no,
                    account_no=effective_account,
                    source_name=srec.full_name,
                    source_row=srec.row,
                    source_amount=srec.amount,
                    allowed_actions=(ResolutionAction.SKIP,),
                    action=ResolutionAction.SKIP,
                )
                row_exc_ids.append(ex_x7_no_tgt.id)
                note_msg = ex_x7_no_tgt.message
            else:
                hint = _leading_zero_hint(effective_account, numeric_stripped_lookup)
                ex_x4 = add_exception(
                    ExceptionCode.X4_NOT_IN_TARGET,
                    Severity.WARNING,
                    f"Nhân viên '{srec.full_name}' (Số thẻ '{srec.card_no}', TK '{effective_account}', "
                    f"số tiền {srec.amount:,} đ, dòng {srec.row}) không có dòng trong file ngân hàng.{hint}",
                    card_no=srec.card_no,
                    account_no=effective_account,
                    source_name=srec.full_name,
                    source_row=srec.row,
                    source_amount=srec.amount,
                    allowed_actions=(ResolutionAction.SKIP,),
                    action=ResolutionAction.SKIP,
                )
                row_exc_ids.append(ex_x4.id)
                note_msg = ex_x4.message

            plan.matches.append(MatchItem(
                card_no=srec.card_no,
                account_no=effective_account,
                source_name=srec.full_name,
                target_name="",
                bank_code=None,
                source_row=srec.row,
                target_file=None,
                target_sheet=None,
                target_row=None,
                target_cell=None,
                old_amount=None,
                new_amount=None,
                source_amount=srec.amount,
                status=MatchStatus.SKIPPED,
                exception_ids=tuple(row_exc_ids),
                note=note_msg,
            ))
            continue

        # 4f. Đã tìm thấy dòng đích duy nhất theo Số tài khoản
        ti, trow, amt_col = target_hit
        tkey = (ti.target.file_path, ti.target.sheet_name, trow.row)
        matched_target_keys.add(tkey)
        target_cell = f"{amt_col}{trow.row}"
        t_issues = tgt_issues_by_row.get(tkey, [])

        # Kiểm tra dòng đích có lỗi cứng (công thức ở ô Số tiền, thiếu Mã ngân hàng...) -> X10
        if trow.has_error:
            err_desc = "; ".join(i.message for i in t_issues if i.severity == Severity.ERROR) or (
                "Dòng trong file ngân hàng có lỗi định danh hoặc công thức"
            )
            ex_x10_err = add_exception(
                ExceptionCode.X10_RULE_VIOLATION,
                Severity.ERROR,
                f"Dòng {trow.row} trong '{Path(ti.target.file_path).name}' (TK '{effective_account}'): "
                f"{err_desc}. Tool không ghi đè vào dòng lỗi này.",
                card_no=srec.card_no,
                account_no=effective_account,
                source_name=srec.full_name,
                target_name=trow.full_name,
                source_row=srec.row,
                target_file=ti.target.file_path,
                target_sheet=ti.target.sheet_name,
                target_row=trow.row,
                old_amount=trow.current_amount,
                source_amount=srec.amount,
                allowed_actions=(ResolutionAction.SKIP,),
                action=None,
            )
            row_exc_ids.append(ex_x10_err.id)

        # Kiểm tra quy tắc nhập liệu trên dòng đích (RULE_VIOLATION) -> X10
        rule_warnings = [i.message for i in t_issues if i.code == IssueCode.RULE_VIOLATION]
        if rule_warnings and not trow.has_error:
            ex_x10_warn = add_exception(
                ExceptionCode.X10_RULE_VIOLATION,
                Severity.WARNING,
                f"Dòng {trow.row} trong '{Path(ti.target.file_path).name}' (TK '{effective_account}'): "
                f"{'; '.join(rule_warnings)}",
                card_no=srec.card_no,
                account_no=effective_account,
                source_name=srec.full_name,
                target_name=trow.full_name,
                source_row=srec.row,
                target_file=ti.target.file_path,
                target_sheet=ti.target.sheet_name,
                target_row=trow.row,
                old_amount=trow.current_amount,
                source_amount=srec.amount,
                allowed_actions=(ResolutionAction.CONFIRM_WRITE, ResolutionAction.SKIP),
                action=None,
            )
            row_exc_ids.append(ex_x10_warn.id)

        # Kiểm tra số tiền nguồn hợp lệ (> 0, số nguyên, không lỗi) -> R5, X7
        if not _has_valid_positive_amount(srec, row_src_issues):
            amt_detail = _describe_source_amount_issue(srec, row_src_issues)
            allowed_x7 = (
                (ResolutionAction.SKIP,)
                if trow.has_error
                else (
                    ResolutionAction.SKIP,
                    ResolutionAction.SET_ZERO,
                    ResolutionAction.CLEAR,
                    ResolutionAction.OVERRIDE_AMOUNT,
                )
            )
            ex_x7 = add_exception(
                ExceptionCode.X7_INVALID_AMOUNT,
                Severity.WARNING,
                f"Nhân viên '{srec.full_name}' (Số thẻ '{srec.card_no}', TK '{effective_account}', "
                f"dòng lương {srec.row} -> dòng NH {trow.row}): {amt_detail}.",
                card_no=srec.card_no,
                account_no=effective_account,
                source_name=srec.full_name,
                target_name=trow.full_name,
                source_row=srec.row,
                target_file=ti.target.file_path,
                target_sheet=ti.target.sheet_name,
                target_row=trow.row,
                old_amount=trow.current_amount,
                source_amount=srec.amount,
                allowed_actions=allowed_x7,
                action=ResolutionAction.SKIP,
            )
            row_exc_ids.append(ex_x7.id)

        # Kiểm tra chéo Họ tên & Mã ngân hàng / Loại file đích -> R3, X6
        mismatch_reasons = _cross_check_reasons(
            srec=srec,
            trow=trow,
            target_input=ti,
            expected_name_key=expected_name_key,
            expected_bank_code=expected_bank_code,
            expected_target_type=expected_target_type,
            source_row_issues=src_issues_by_row.get(srec.row, []),
        )
        if mismatch_reasons:
            allowed_x6 = (
                (ResolutionAction.SKIP,)
                if trow.has_error
                else (ResolutionAction.CONFIRM_WRITE, ResolutionAction.SKIP)
            )
            ex_x6 = add_exception(
                ExceptionCode.X6_TARGET_MISMATCH,
                Severity.WARNING,
                f"Khớp Số TK '{effective_account}' (Số thẻ '{srec.card_no}', dòng lương {srec.row} -> "
                f"dòng NH {trow.row}) nhưng: {'; '.join(mismatch_reasons)}.",
                card_no=srec.card_no,
                account_no=effective_account,
                source_name=srec.full_name,
                target_name=trow.full_name,
                source_row=srec.row,
                target_file=ti.target.file_path,
                target_sheet=ti.target.sheet_name,
                target_row=trow.row,
                old_amount=trow.current_amount,
                source_amount=srec.amount,
                allowed_actions=allowed_x6,
                action=None,
            )
            row_exc_ids.append(ex_x6.id)

        plan.matches.append(MatchItem(
            card_no=srec.card_no,
            account_no=effective_account,
            source_name=srec.full_name,
            target_name=trow.full_name,
            bank_code=trow.bank_code,
            source_row=srec.row,
            target_file=ti.target.file_path,
            target_sheet=ti.target.sheet_name,
            target_row=trow.row,
            target_cell=target_cell,
            old_amount=trow.current_amount,
            new_amount=srec.amount,
            source_amount=srec.amount,
            status=MatchStatus.PENDING,  # sẽ được tính chính xác trong rebuild_plan()
            exception_ids=tuple(row_exc_ids),
        ))

    # 5. Duyệt các dòng trong file ngân hàng KHÔNG có trong bảng lương mới -> X5 (hoặc X8/X10)
    for ti, tp in zip(target_inputs, plan.target_plans):
        for trow in ti.target.rows:
            tkey = (ti.target.file_path, ti.target.sheet_name, trow.row)
            if tkey in matched_target_keys:
                continue
            target_cell = f"{tp.amount_column}{trow.row}"

            # Dòng đích bị trùng Số tài khoản (X8)
            if trow.account_no is not None and trow.account_no in dup_tgt_accs:
                ex_x8_row = x8_by_tgt_acc[trow.account_no]
                plan.matches.append(MatchItem(
                    card_no=None,
                    account_no=trow.account_no,
                    source_name="",
                    target_name=trow.full_name,
                    bank_code=trow.bank_code,
                    source_row=None,
                    target_file=ti.target.file_path,
                    target_sheet=ti.target.sheet_name,
                    target_row=trow.row,
                    target_cell=target_cell,
                    old_amount=trow.current_amount,
                    new_amount=None,
                    source_amount=None,
                    status=MatchStatus.ERROR,
                    exception_ids=(ex_x8_row.id,),
                    note="Dòng trong file ngân hàng bị trùng Số tài khoản (X8).",
                ))
                continue

            # Dòng đích có dữ liệu nhưng thiếu Số tài khoản (dòng ghi chú / không hoàn chỉnh)
            if trow.account_no is None:
                ex_inc = add_exception(
                    ExceptionCode.X10_RULE_VIOLATION,
                    Severity.WARNING,
                    f"Dòng {trow.row} trong '{Path(ti.target.file_path).name}' "
                    f"(Tên: '{trow.full_name}') không có Số tài khoản hợp lệ -> không ghi vào dòng này.",
                    account_no=None,
                    target_name=trow.full_name,
                    target_file=ti.target.file_path,
                    target_sheet=ti.target.sheet_name,
                    target_row=trow.row,
                    old_amount=trow.current_amount,
                    allowed_actions=(ResolutionAction.SKIP,),
                    action=ResolutionAction.SKIP,
                )
                plan.matches.append(MatchItem(
                    card_no=None,
                    account_no=None,
                    source_name="",
                    target_name=trow.full_name,
                    bank_code=trow.bank_code,
                    source_row=None,
                    target_file=ti.target.file_path,
                    target_sheet=ti.target.sheet_name,
                    target_row=trow.row,
                    target_cell=target_cell,
                    old_amount=trow.current_amount,
                    new_amount=trow.current_amount,
                    source_amount=None,
                    status=MatchStatus.SKIPPED,
                    exception_ids=(ex_inc.id,),
                    note=ex_inc.message,
                ))
                continue

            # Dòng đích có Số tài khoản nhưng không có trong bảng lương tháng mới -> X5
            allowed_x5 = (
                (ResolutionAction.KEEP,)
                if trow.has_error
                else (ResolutionAction.KEEP, ResolutionAction.SET_ZERO, ResolutionAction.CLEAR)
            )
            old_display = f"{trow.current_amount:,} đ" if trow.current_amount is not None else "trống"
            ex_x5 = add_exception(
                ExceptionCode.X5_TARGET_NOT_IN_SOURCE,
                Severity.WARNING,
                f"Dòng {trow.row} trong '{Path(ti.target.file_path).name}' "
                f"(Tên: '{trow.full_name}', TK '{trow.account_no}', số tiền hiện tại: {old_display}) "
                "không có trong bảng lương tháng mới. Hãy chọn: Giữ nguyên / Đặt 0 / Xóa giá trị.",
                account_no=trow.account_no,
                target_name=trow.full_name,
                target_file=ti.target.file_path,
                target_sheet=ti.target.sheet_name,
                target_row=trow.row,
                old_amount=trow.current_amount,
                allowed_actions=allowed_x5,
                action=None,
            )
            plan.matches.append(MatchItem(
                card_no=None,
                account_no=trow.account_no,
                source_name="",
                target_name=trow.full_name,
                bank_code=trow.bank_code,
                source_row=None,
                target_file=ti.target.file_path,
                target_sheet=ti.target.sheet_name,
                target_row=trow.row,
                target_cell=target_cell,
                old_amount=trow.current_amount,
                new_amount=trow.current_amount,
                source_amount=None,
                status=MatchStatus.PENDING,
                exception_ids=(ex_x5.id,),
            ))

    # 6. Áp dụng các quyết định xử lý (nếu người gọi truyền sẵn) và dựng WritePlan
    if resolutions:
        apply_resolutions(plan, resolutions)
    else:
        rebuild_plan(plan)

    return plan


# =============================================================================
# 2. Xử lý Exception & Tính lại WritePlan
# =============================================================================

def resolve_exception(
    plan: WritePlan,
    exception_id: str,
    action: ResolutionAction,
    *,
    override_amount: Optional[int] = None,
) -> WritePlan:
    """Gán lựa chọn xử lý cho 1 `ExceptionItem` theo ID và tính lại toàn bộ `WritePlan`."""
    target_exc = next((e for e in plan.exceptions if e.id == exception_id), None)
    if target_exc is None:
        raise ReconcileError(f"Không tìm thấy mã ngoại lệ '{exception_id}'.")
    if target_exc.is_blocking:
        raise ReconcileError(
            f"Ngoại lệ '{exception_id}' ({target_exc.code.value}) là lỗi chặn cứng, "
            "phải sửa dữ liệu gốc chứ không thể bỏ qua."
        )
    if action not in target_exc.allowed_actions:
        allowed_str = ", ".join(a.value for a in target_exc.allowed_actions)
        raise ReconcileError(
            f"Hành động '{action.value}' không hợp lệ cho ngoại lệ '{exception_id}' "
            f"(chỉ chấp nhận: {allowed_str})."
        )
    if action == ResolutionAction.OVERRIDE_AMOUNT:
        if (
            override_amount is None
            or not isinstance(override_amount, int)
            or isinstance(override_amount, bool)
            or override_amount <= 0
        ):
            raise ReconcileError("Số tiền sửa tay (OVERRIDE_AMOUNT) phải là số nguyên lớn hơn 0.")
        target_exc.override_amount = override_amount
    else:
        target_exc.override_amount = None

    target_exc.action = action
    return rebuild_plan(plan)


def apply_resolutions(
    plan: WritePlan,
    resolutions: Mapping[str, ResolutionValue],
) -> WritePlan:
    """Áp dụng hàng loạt quyết định xử lý ngoại lệ `{exception_id: action | (action, override_amount)}`."""
    for exc_id, val in resolutions.items():
        if isinstance(val, tuple):
            act, override = val
            resolve_exception(plan, exc_id, act, override_amount=override)
        else:
            resolve_exception(plan, exc_id, val)
    return rebuild_plan(plan)


def rebuild_plan(plan: WritePlan) -> WritePlan:
    """Tính lại trạng thái `MatchItem` và danh sách `CellUpdate` từ trạng thái `ExceptionItem` hiện tại."""
    exc_by_id: dict[str, ExceptionItem] = {e.id: e for e in plan.exceptions}
    tp_by_file: dict[tuple[str, str], TargetWritePlan] = {
        (tp.file_path, tp.sheet_name): tp for tp in plan.target_plans
    }
    for tp in plan.target_plans:
        tp.updates.clear()
    plan.updates.clear()

    for match in plan.matches:
        attached = [exc_by_id[eid] for eid in match.exception_ids if eid in exc_by_id]

        # 1. Có lỗi chặn cứng (X8)
        if any(e.is_blocking for e in attached):
            match.status = MatchStatus.ERROR
            match.new_amount = None
            match.note = "; ".join(e.message for e in attached)
            continue

        # 2. Còn ngoại lệ chưa được người dùng chọn phương án xử lý
        if any(not e.is_resolved for e in attached):
            match.status = MatchStatus.PENDING
            match.new_amount = match.source_amount if match.source_row is not None else match.old_amount
            match.note = "; ".join(e.message for e in attached if not e.is_resolved)
            continue

        # 3. Dòng không gắn với dòng đích nào (X3 thiếu TK, X4 không có trong file đích)
        if match.target_file is None or match.target_sheet is None or match.target_row is None:
            match.status = MatchStatus.SKIPPED
            match.new_amount = None
            if attached and not match.note:
                match.note = "; ".join(e.message for e in attached)
            continue

        tp_target = tp_by_file.get((match.target_file, match.target_sheet))
        col_letter = tp_target.amount_column if tp_target is not None else (
            match.target_cell.rstrip("0123456789") if match.target_cell else "D"
        )

        # 4. Dòng đích dư không có trong bảng lương mới (X5) hoặc dòng đích thiếu TK (X10)
        if match.source_row is None:
            x5_exc = next((e for e in attached if e.code == ExceptionCode.X5_TARGET_NOT_IN_SOURCE), None)
            if x5_exc is None or x5_exc.action is None or x5_exc.action == ResolutionAction.KEEP:
                match.new_amount = match.old_amount
                match.status = MatchStatus.UNCHANGED if x5_exc is not None else MatchStatus.SKIPPED
                match.note = "Giữ nguyên giá trị hiện tại trong file ngân hàng." if x5_exc else match.note
                continue

            if x5_exc.action == ResolutionAction.SET_ZERO:
                match.new_amount = 0
                match.status = MatchStatus.WILL_UPDATE if match.old_amount != 0 else MatchStatus.UNCHANGED
                match.note = "Đặt số tiền = 0 (nhân sự không có trong lương mới)."
                _append_update(
                    plan, tp_target,
                    CellUpdate(
                        file_path=match.target_file,
                        sheet_name=match.target_sheet,
                        row=match.target_row,
                        column=col_letter,
                        old_value=match.old_amount,
                        new_value=0,
                        card_no=None,
                        account_no=match.account_no,
                        full_name=match.target_name,
                        reason="X5_SET_ZERO",
                    ),
                )
                continue

            if x5_exc.action == ResolutionAction.CLEAR:
                match.new_amount = None
                match.status = MatchStatus.WILL_UPDATE if match.old_amount is not None else MatchStatus.UNCHANGED
                match.note = "Xóa trắng ô số tiền (nhân sự không có trong lương mới)."
                _append_update(
                    plan, tp_target,
                    CellUpdate(
                        file_path=match.target_file,
                        sheet_name=match.target_sheet,
                        row=match.target_row,
                        column=col_letter,
                        old_value=match.old_amount,
                        new_value=None,
                        card_no=None,
                        account_no=match.account_no,
                        full_name=match.target_name,
                        reason="X5_CLEAR",
                    ),
                )
                continue

        # 5. Dòng khớp giữa nguồn và đích: kiểm tra các hành động của Exception đính kèm
        if any(e.action == ResolutionAction.SKIP for e in attached):
            match.new_amount = match.old_amount
            match.status = MatchStatus.SKIPPED
            match.note = "Đã chọn bỏ qua không cập nhật dòng này."
            continue

        if any(e.action == ResolutionAction.SET_ZERO for e in attached):
            match.new_amount = 0
            match.status = MatchStatus.WILL_UPDATE if match.old_amount != 0 else MatchStatus.UNCHANGED
            match.note = "Đặt số tiền = 0 theo lựa chọn xử lý ngoại lệ."
            _append_update(
                plan, tp_target,
                CellUpdate(
                    file_path=match.target_file,
                    sheet_name=match.target_sheet,
                    row=match.target_row,
                    column=col_letter,
                    old_value=match.old_amount,
                    new_value=0,
                    card_no=match.card_no,
                    account_no=match.account_no,
                    full_name=match.target_name or match.source_name,
                    reason="X7_SET_ZERO",
                ),
            )
            continue

        if any(e.action == ResolutionAction.CLEAR for e in attached):
            match.new_amount = None
            match.status = MatchStatus.WILL_UPDATE if match.old_amount is not None else MatchStatus.UNCHANGED
            match.note = "Xóa trắng ô số tiền theo lựa chọn xử lý ngoại lệ."
            _append_update(
                plan, tp_target,
                CellUpdate(
                    file_path=match.target_file,
                    sheet_name=match.target_sheet,
                    row=match.target_row,
                    column=col_letter,
                    old_value=match.old_amount,
                    new_value=None,
                    card_no=match.card_no,
                    account_no=match.account_no,
                    full_name=match.target_name or match.source_name,
                    reason="X7_CLEAR",
                ),
            )
            continue

        override_exc = next(
            (e for e in attached if e.action == ResolutionAction.OVERRIDE_AMOUNT and e.override_amount),
            None,
        )
        if override_exc is not None and override_exc.override_amount is not None:
            final_amount: Optional[int] = override_exc.override_amount
            reason = "OVERRIDE_AMOUNT"
        else:
            final_amount = match.source_amount
            reason = "CONFIRMED_MISMATCH" if attached else "MATCHED"

        if final_amount is None or final_amount <= 0:
            match.new_amount = match.old_amount
            match.status = MatchStatus.SKIPPED
            continue

        match.new_amount = final_amount
        match.status = (
            MatchStatus.WILL_UPDATE if match.old_amount != final_amount else MatchStatus.UNCHANGED
        )
        match.note = (
            "Đã xác nhận cập nhật." if attached else (
                "Sẽ cập nhật số tiền mới." if match.old_amount != final_amount else "Số tiền không đổi."
            )
        )
        _append_update(
            plan, tp_target,
            CellUpdate(
                file_path=match.target_file,
                sheet_name=match.target_sheet,
                row=match.target_row,
                column=col_letter,
                old_value=match.old_amount,
                new_value=final_amount,
                card_no=match.card_no,
                account_no=match.account_no,
                full_name=match.target_name or match.source_name,
                reason=reason,
            ),
        )

    return plan


# =============================================================================
# 3. Các hàm nội bộ
# =============================================================================

def _append_update(
    plan: WritePlan,
    tp: Optional[TargetWritePlan],
    update: CellUpdate,
) -> None:
    if tp is not None:
        tp.updates.append(update)
    plan.updates.append(update)


def _normalize_target_specs(
    targets: Union[TargetSpec, Sequence[TargetSpec]],
    templates: Optional[Mapping[str, TargetTemplate]],
) -> list[TargetInput]:
    raw_list: list[TargetSpec]
    if isinstance(targets, (TargetReadResult, TargetInput)):
        raw_list = [targets]
    elif (
        isinstance(targets, tuple)
        and len(targets) == 2
        and isinstance(targets[0], TargetReadResult)
        and isinstance(targets[1], (TargetTemplate, str))
    ):
        raw_list = [targets]
    else:
        raw_list = list(targets)

    result: list[TargetInput] = []
    seen_files: set[tuple[str, str]] = set()
    for spec in raw_list:
        if isinstance(spec, TargetInput):
            ti = spec
        elif isinstance(spec, TargetReadResult):
            tpl = None
            if templates is not None:
                tpl = templates.get(spec.template_name) or templates.get(spec.file_path)
            ti = TargetInput(target=spec, template=tpl)
        elif isinstance(spec, tuple) and len(spec) == 2 and isinstance(spec[0], TargetReadResult):
            second = spec[1]
            if isinstance(second, TargetTemplate):
                ti = TargetInput(target=spec[0], template=second)
            elif isinstance(second, str):
                ti = TargetInput(target=spec[0], amount_column=second)
            else:
                raise ConfigError(f"Cấu hình file đích không hợp lệ: {spec!r}")
        else:
            raise ConfigError(f"Cấu hình file đích không hợp lệ: {spec!r}")

        # Kiểm tra cột Số tiền hợp lệ ngay từ đầu
        ti.resolved_amount_column()
        file_sheet_key = (ti.target.file_path, ti.target.sheet_name)
        if file_sheet_key in seen_files:
            raise ConfigError(
                f"File ngân hàng '{ti.target.file_path}' (sheet '{ti.target.sheet_name}') "
                "được đưa vào danh sách đối chiếu nhiều lần."
            )
        seen_files.add(file_sheet_key)
        result.append(ti)
    return result


def _collect_structural_issues(
    source: SourceReadResult,
    target_inputs: list[TargetInput],
    plan: WritePlan,
) -> None:
    structural_codes = BLOCKING_CODES - {IssueCode.DUPLICATE_CARD, IssueCode.DUPLICATE_ACCOUNT}
    for issue in source.issues:
        if issue.code in structural_codes:
            plan.blocking_issues.append(issue)
    for ti in target_inputs:
        for issue in ti.target.issues:
            if issue.code in structural_codes:
                plan.blocking_issues.append(issue)


def _detect_duplicates(
    source: SourceReadResult,
    target_inputs: list[TargetInput],
    master: Optional[Mapping[str, MasterRecord]],
    add_exception: Callable[..., ExceptionItem],
) -> tuple[
    set[str],
    set[str],
    set[str],
    dict[str, ExceptionItem],
    dict[str, ExceptionItem],
    dict[str, ExceptionItem],
]:
    dup_cards: set[str] = set()
    dup_src_accs: set[str] = set()
    dup_tgt_accs: set[str] = set()
    x8_by_card: dict[str, ExceptionItem] = {}
    x8_by_src_acc: dict[str, ExceptionItem] = {}
    x8_by_tgt_acc: dict[str, ExceptionItem] = {}

    # 1. Trùng Số thẻ trong bảng lương
    by_card: dict[str, list[SourceRecord]] = defaultdict(list)
    for rec in source.records:
        by_card[rec.card_no].append(rec)
    for card, recs in by_card.items():
        if len(recs) > 1:
            dup_cards.add(card)
            rows_str = ", ".join(str(r.row) for r in recs)
            x8_by_card[card] = add_exception(
                ExceptionCode.X8_DUPLICATE_KEY,
                Severity.ERROR,
                f"Trùng Số thẻ '{card}' tại các dòng [{rows_str}] trong bảng lương. "
                "Phải sửa dữ liệu gốc trước khi xuất file.",
                card_no=card,
                source_name=recs[0].full_name,
                source_row=recs[0].row,
                source_amount=recs[0].amount,
                allowed_actions=(),
            )

    # 2. Trùng Số tài khoản trong bảng lương (nhiều dòng nguồn trỏ vào cùng 1 số TK)
    by_src_acc: dict[str, list[SourceRecord]] = defaultdict(list)
    for rec in source.records:
        seen_for_row: set[str] = set()
        if rec.account_no is not None:
            by_src_acc[rec.account_no].append(rec)
            seen_for_row.add(rec.account_no)
        if master is not None and rec.card_no in master and master[rec.card_no].account_no:
            m_acc = master[rec.card_no].account_no
            if m_acc not in seen_for_row:
                by_src_acc[m_acc].append(rec)
    for acc, recs in by_src_acc.items():
        if len(recs) > 1:
            dup_src_accs.add(acc)
            detail = ", ".join(f"dòng {r.row} (Số thẻ '{r.card_no}')" for r in recs)
            x8_by_src_acc[acc] = add_exception(
                ExceptionCode.X8_DUPLICATE_KEY,
                Severity.ERROR,
                f"Trùng Số tài khoản '{acc}' trong bảng lương tại: {detail}. "
                "Phải sửa dữ liệu gốc trước khi xuất file.",
                card_no=recs[0].card_no,
                account_no=acc,
                source_name=recs[0].full_name,
                source_row=recs[0].row,
                source_amount=recs[0].amount,
                allowed_actions=(),
            )

    # 3. Trùng Số tài khoản trong 1 file đích hoặc giữa nhiều file đích
    by_tgt_acc: dict[str, list[tuple[TargetInput, TargetRow]]] = defaultdict(list)
    for ti in target_inputs:
        for trow in ti.target.rows:
            if trow.account_no is not None:
                by_tgt_acc[trow.account_no].append((ti, trow))
    for acc, locs in by_tgt_acc.items():
        if len(locs) > 1:
            dup_tgt_accs.add(acc)
            loc_desc = ", ".join(
                f"'{Path(ti.target.file_path).name}' (sheet '{ti.target.sheet_name}', dòng {tr.row})"
                for ti, tr in locs
            )
            first_ti, first_tr = locs[0]
            x8_by_tgt_acc[acc] = add_exception(
                ExceptionCode.X8_DUPLICATE_KEY,
                Severity.ERROR,
                f"Trùng Số tài khoản '{acc}' trong file ngân hàng tại: {loc_desc}. "
                "Phải sửa dữ liệu gốc trước khi xuất file.",
                account_no=acc,
                target_name=first_tr.full_name,
                target_file=first_ti.target.file_path,
                target_sheet=first_ti.target.sheet_name,
                target_row=first_tr.row,
                old_amount=first_tr.current_amount,
                allowed_actions=(),
            )

    return dup_cards, dup_src_accs, dup_tgt_accs, x8_by_card, x8_by_src_acc, x8_by_tgt_acc


_AMOUNT_ERROR_CODES: frozenset[IssueCode] = frozenset({
    IssueCode.AMOUNT_MISSING,
    IssueCode.AMOUNT_NOT_NUMBER,
    IssueCode.AMOUNT_NOT_INTEGER,
    IssueCode.AMOUNT_NEGATIVE,
    IssueCode.AMOUNT_ZERO,
    IssueCode.FORMULA_NOT_CALCULATED,
    IssueCode.EXCEL_ERROR_VALUE,
})


def _has_valid_positive_amount(rec: SourceRecord, row_issues: list[Issue]) -> bool:
    if rec.amount is None or rec.amount <= 0:
        return False
    if any(i.code in _AMOUNT_ERROR_CODES and i.severity == Severity.ERROR for i in row_issues):
        return False
    if rec.has_error and rec.account_no is not None and not row_issues:
        return False
    return True


def _describe_source_account_issue(row_issues: list[Issue]) -> str:
    acc_codes = {
        IssueCode.ACCOUNT_MISSING,
        IssueCode.ACCOUNT_INVALID,
    }
    for issue in row_issues:
        if issue.code in acc_codes:
            return issue.message
    return "Thiếu Số tài khoản"


def _describe_source_amount_issue(rec: SourceRecord, row_issues: list[Issue]) -> str:
    amt_codes = {
        IssueCode.AMOUNT_MISSING,
        IssueCode.AMOUNT_NOT_NUMBER,
        IssueCode.AMOUNT_NOT_INTEGER,
        IssueCode.AMOUNT_NEGATIVE,
        IssueCode.AMOUNT_ZERO,
        IssueCode.FORMULA_NOT_CALCULATED,
        IssueCode.EXCEL_ERROR_VALUE,
    }
    for issue in row_issues:
        if issue.code in amt_codes:
            return issue.message
    if rec.amount == 0:
        return "Số tiền chuyển khoản bằng 0"
    if rec.amount is None:
        return "Thiếu số tiền chuyển khoản"
    return f"Số tiền chuyển khoản không hợp lệ ({rec.amount})"


def _leading_zero_hint(
    account_no: str,
    numeric_stripped_lookup: Mapping[str, tuple[TargetInput, TargetRow]],
) -> str:
    if not account_no.startswith("0"):
        return ""
    stripped = account_no.lstrip("0")
    if not stripped or stripped not in numeric_stripped_lookup:
        return ""
    ti, trow = numeric_stripped_lookup[stripped]
    return (
        f" (Lưu ý: phát hiện dòng {trow.row} trong '{Path(ti.target.file_path).name}' "
        f"có số TK '{trow.account_no}' lưu dạng số -> có thể đã bị mất số 0 ở đầu; "
        "tool không tự ghép để đảm bảo an toàn)."
    )


def _cross_check_reasons(
    *,
    srec: SourceRecord,
    trow: TargetRow,
    target_input: TargetInput,
    expected_name_key: str,
    expected_bank_code: Optional[str],
    expected_target_type: Optional[str],
    source_row_issues: list[Issue],
) -> list[str]:
    reasons: list[str] = []
    has_legacy_font = any(i.code == IssueCode.NAME_LEGACY_FONT for i in source_row_issues)

    if not expected_name_key or not trow.name_key:
        reasons.append(
            f"thiếu tên nhân viên để kiểm tra chéo (bảng lương: '{srec.full_name or '(trống)'}', "
            f"file ngân hàng: '{trow.full_name or '(trống)'}')"
        )
    elif expected_name_key != trow.name_key:
        font_note = " (cột tên bảng lương đang dùng font VNI/TCVN3)" if has_legacy_font else ""
        reasons.append(
            f"lệch tên nhân viên (bảng lương: '{srec.full_name}'{font_note} != "
            f"file ngân hàng: '{trow.full_name}')"
        )

    if expected_bank_code is not None and trow.bank_code is not None:
        if clean_text(expected_bank_code) != clean_text(trow.bank_code):
            reasons.append(
                f"lệch Mã ngân hàng (Master: '{expected_bank_code}' != file ngân hàng: '{trow.bank_code}')"
            )

    if expected_target_type is not None:
        tpl_name = target_input.target.template_name
        if clean_text(expected_target_type).casefold() != clean_text(tpl_name).casefold():
            reasons.append(
                f"lệch loại mẫu ngân hàng (Master: '{expected_target_type}' != đang chọn: '{tpl_name}')"
            )

    return reasons
