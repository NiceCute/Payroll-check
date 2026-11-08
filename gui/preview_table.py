"""
preview_table.py - Bảng xem trước dữ liệu đối chiếu trước khi Execute (GUI Module 2).

Dữ liệu hiển thị trên từng dòng:
- Họ và tên (`full_name`)
- Số thẻ (`card_no`)
- Số tài khoản (`account_no`)
- Mã ngân hàng (`bank_code`)
- Số tiền (`amount`)
- Trạng thái match (`match_status`)
- Trạng thái validation (`validation_status`)
- Exception (`exception`)

Chức năng:
- Filter (lọc theo trạng thái Match, trạng thái Validation, hoặc chỉ dòng có Exception / Warning)
- Search (tìm kiếm tức thì theo Họ tên, Số thẻ, Số tài khoản, Mã NH, Exception — hỗ trợ bỏ dấu tiếng Việt)
- Sort (sắp xếp tăng/giảm theo bất kỳ cột nào khi bấm tiêu đề cột hoặc gọi `sort_by`)
- Highlight Exception (tô màu nổi bật các dòng có Exception)
- Highlight Warning (tô màu nổi bật các dòng có Warning)
- Hỗ trợ chọn phương án xử lý (`ResolutionAction`) cho dòng ngoại lệ qua callback.

Nguyên tắc:
- Không ghi file, không xử lý workflow, không xử lý mapping dialog.
- Chỉ phụ thuộc `core.models` (+ `core.normalize` để hỗ trợ tìm kiếm không dấu).
"""

from __future__ import annotations

import tkinter as tk
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from tkinter import ttk
from typing import Any, Callable, Optional, Sequence, Union

from core.models import (
    ExceptionItem,
    MatchItem,
    MatchStatus,
    ResolutionAction,
    SafetyFinding,
    Severity,
    ValidationResult,
    WorkflowResult,
    WritePlan,
)
from core.normalize import remove_diacritics

__all__ = [
    "COLOR_EXCEPTION_BG",
    "COLOR_EXCEPTION_FG",
    "COLOR_SKIPPED_BG",
    "COLOR_SKIPPED_FG",
    "COLOR_UNCHANGED_BG",
    "COLOR_UNCHANGED_FG",
    "COLOR_UPDATE_BG",
    "COLOR_UPDATE_FG",
    "COLOR_WARNING_BG",
    "COLOR_WARNING_FG",
    "PreviewColumn",
    "PreviewFilter",
    "PreviewRow",
    "PreviewTable",
    "TAG_EXCEPTION",
    "TAG_SKIPPED",
    "TAG_UNCHANGED",
    "TAG_UPDATE",
    "TAG_WARNING",
    "ValidationFilter",
    "build_preview_rows",
]

# Mã tag hiển thị màu trên Treeview
TAG_EXCEPTION: str = "tag_exception"
TAG_WARNING: str = "tag_warning"
TAG_UPDATE: str = "tag_update"
TAG_UNCHANGED: str = "tag_unchanged"
TAG_SKIPPED: str = "tag_skipped"

# Bảng màu chuẩn theo mục 5 tài liệu V1 (🟢 cập nhật · ⚪ không đổi · 🟡 cảnh báo/cần xác nhận · 🔴 lỗi/exception)
COLOR_EXCEPTION_BG: str = "#FCE4D6"
COLOR_EXCEPTION_FG: str = "#9C0006"
COLOR_WARNING_BG: str = "#FFF2CC"
COLOR_WARNING_FG: str = "#7F6000"
COLOR_UPDATE_BG: str = "#E2EFDA"
COLOR_UPDATE_FG: str = "#1E4620"
COLOR_UNCHANGED_BG: str = "#FFFFFF"
COLOR_UNCHANGED_FG: str = "#333333"
COLOR_SKIPPED_BG: str = "#F2F2F2"
COLOR_SKIPPED_FG: str = "#666666"


class PreviewColumn(str, Enum):
    """Định danh các cột trên bảng Preview Table."""

    STT = "stt"
    FULL_NAME = "full_name"
    CARD_NO = "card_no"
    ACCOUNT_NO = "account_no"
    BANK_CODE = "bank_code"
    AMOUNT = "amount"
    MATCH_STATUS = "match_status"
    VALIDATION_STATUS = "validation_status"
    EXCEPTION = "exception"


class PreviewFilter(str, Enum):
    """Các chế độ lọc nhanh theo trạng thái Match / Exception."""

    ALL = "ALL"
    WILL_UPDATE = "WILL_UPDATE"
    UNCHANGED = "UNCHANGED"
    PENDING = "PENDING"
    SKIPPED = "SKIPPED"
    ERROR = "ERROR"
    EXCEPTIONS_ONLY = "EXCEPTIONS_ONLY"
    UNRESOLVED_ONLY = "UNRESOLVED_ONLY"
    WARNINGS_ONLY = "WARNINGS_ONLY"


class ValidationFilter(str, Enum):
    """Các chế độ lọc theo trạng thái Validation."""

    ALL = "ALL"
    OK = "OK"
    WARNING = "WARNING"
    ERROR = "ERROR"


_COLUMN_HEADERS: dict[str, str] = {
    PreviewColumn.STT.value: "#",
    PreviewColumn.FULL_NAME.value: "Họ và tên",
    PreviewColumn.CARD_NO.value: "Số thẻ",
    PreviewColumn.ACCOUNT_NO.value: "Số tài khoản",
    PreviewColumn.BANK_CODE.value: "Mã ngân hàng",
    PreviewColumn.AMOUNT.value: "Số tiền (VND)",
    PreviewColumn.MATCH_STATUS.value: "Trạng thái Match",
    PreviewColumn.VALIDATION_STATUS.value: "Trạng thái Validation",
    PreviewColumn.EXCEPTION.value: "Exception",
}


# =============================================================================
# 1. Dữ liệu 1 dòng hiển thị trên Preview Table
# =============================================================================

@dataclass(frozen=True)
class PreviewRow:
    """1 dòng dữ liệu đã chuẩn hóa để hiển thị, lọc, tìm kiếm, sắp xếp và tô màu trên Preview Table."""

    index: int
    full_name: str
    source_name: str
    target_name: str
    card_no: str
    account_no: str
    bank_code: str
    old_amount: Optional[int]
    source_amount: Optional[int]
    new_amount: Optional[int]
    display_amount: Optional[int]
    match_status: MatchStatus
    validation_status: str          # "OK", "WARNING", "ERROR"
    validation_detail: str
    exception_text: str
    exception_codes: tuple[str, ...]
    exception_ids: tuple[str, ...]
    allowed_actions: tuple[ResolutionAction, ...]
    selected_action: Optional[ResolutionAction]
    override_amount: Optional[int]
    has_exception: bool
    has_warning: bool
    is_blocking: bool
    is_unresolved: bool
    row_tag: str
    source_row: Optional[int] = None
    target_file: str = ""
    target_sheet: str = ""
    target_row: Optional[int] = None
    target_cell: str = ""
    note: str = ""

    @property
    def amount_display_str(self) -> str:
        if self.display_amount is None:
            return "-"
        if (
            self.old_amount is not None
            and self.new_amount is not None
            and self.old_amount != self.new_amount
        ):
            return f"{self.new_amount:,} (cũ: {self.old_amount:,})"
        return f"{self.display_amount:,}"

    @property
    def match_status_display(self) -> str:
        icons = {
            MatchStatus.WILL_UPDATE: "🟢 WILL_UPDATE",
            MatchStatus.UNCHANGED: "⚪ UNCHANGED",
            MatchStatus.PENDING: "🟡 PENDING",
            MatchStatus.SKIPPED: "⚪ SKIPPED",
            MatchStatus.ERROR: "🔴 ERROR",
        }
        return icons.get(self.match_status, self.match_status.value)

    @property
    def validation_status_display(self) -> str:
        if self.validation_status == "ERROR":
            return f"🔴 ERROR ({self.validation_detail})" if self.validation_detail else "🔴 ERROR"
        if self.validation_status == "WARNING":
            return f"🟡 WARNING ({self.validation_detail})" if self.validation_detail else "🟡 WARNING"
        return "🟢 OK"

    def values_tuple(self) -> tuple[Any, ...]:
        return (
            self.index,
            self.full_name,
            self.card_no,
            self.account_no,
            self.bank_code,
            self.amount_display_str,
            self.match_status_display,
            self.validation_status_display,
            self.exception_text,
        )


# =============================================================================
# 2. Hàm chuyển đổi từ WritePlan / ValidationResult sang danh sách PreviewRow
# =============================================================================

def build_preview_rows(
    write_plan: Optional[WritePlan],
    *,
    validation_result: Optional[ValidationResult] = None,
    exceptions: Optional[Sequence[ExceptionItem]] = None,
) -> list[PreviewRow]:
    """Chuyển đổi `WritePlan` + `ValidationResult` + `ExceptionItem` thành danh sách `PreviewRow`."""
    if write_plan is None:
        return []

    exc_list = list(exceptions) if exceptions is not None else list(write_plan.exceptions)
    exc_by_id: dict[str, ExceptionItem] = {e.id: e for e in exc_list}

    # Ánh xạ các phát hiện từ ValidationResult (Safety Layer) theo dòng / Số thẻ / Số TK
    safety_by_card: dict[str, list[SafetyFinding]] = {}
    safety_by_acc: dict[str, list[SafetyFinding]] = {}
    safety_by_target_row: dict[tuple[str, str, int], list[SafetyFinding]] = {}

    if validation_result is not None:
        for f in [*validation_result.errors, *validation_result.warnings]:
            if f.card_no:
                safety_by_card.setdefault(f.card_no, []).append(f)
            if f.account_no:
                safety_by_acc.setdefault(f.account_no, []).append(f)
            if f.file_path and f.sheet and f.row is not None:
                safety_by_target_row.setdefault((f.file_path, f.sheet, f.row), []).append(f)

    rows: list[PreviewRow] = []
    for idx, m in enumerate(write_plan.matches, start=1):
        attached_excs = [exc_by_id[eid] for eid in m.exception_ids if eid in exc_by_id]
        row_findings = _collect_row_safety_findings(
            m,
            safety_by_card=safety_by_card,
            safety_by_acc=safety_by_acc,
            safety_by_target_row=safety_by_target_row,
        )

        has_exc = bool(attached_excs)
        is_blocking = any(e.is_blocking for e in attached_excs)
        is_unresolved = any(not e.is_resolved for e in attached_excs)
        has_exc_error = any(e.severity == Severity.ERROR for e in attached_excs)
        has_exc_warn = any(e.severity == Severity.WARNING for e in attached_excs)

        has_safety_error = any(f.severity == Severity.ERROR for f in row_findings)
        has_safety_warn = any(f.severity == Severity.WARNING for f in row_findings)

        # Trạng thái Validation của dòng
        if is_blocking or is_unresolved or has_safety_error or m.status in (MatchStatus.ERROR, MatchStatus.PENDING):
            val_status = "ERROR"
        elif has_safety_warn or has_exc_warn or has_exc:
            val_status = "WARNING"
        else:
            val_status = "OK"

        val_details: list[str] = []
        for f in row_findings:
            if f.code.value not in val_details:
                val_details.append(f.code.value)
        if not val_details and is_unresolved:
            val_details.append("UNRESOLVED_EXCEPTION")
        val_detail_str = ", ".join(val_details)

        # Cờ Highlight Exception & Highlight Warning
        has_warning = bool(has_safety_warn or has_exc_warn or val_status == "WARNING")
        row_tag = _determine_row_tag(
            match_status=m.status,
            has_exception=has_exc,
            is_blocking=is_blocking,
            is_unresolved=is_unresolved,
            has_exc_error=has_exc_error,
            has_safety_error=has_safety_error,
            has_warning=has_warning,
        )

        # Chuỗi hiển thị cột Exception
        exc_parts: list[str] = []
        first_exc = attached_excs[0] if attached_excs else None
        for e in attached_excs:
            act_str = e.action.value if e.action is not None else "CHƯA XỬ LÝ"
            if e.action == ResolutionAction.OVERRIDE_AMOUNT and e.override_amount is not None:
                act_str = f"{act_str}={e.override_amount:,}"
            exc_parts.append(f"[{e.code.value}: {act_str}] {e.message}")
        if not exc_parts and m.note:
            exc_parts.append(m.note)

        full_name = m.source_name or m.target_name
        if m.source_name and m.target_name and m.source_name != m.target_name:
            full_name = f"{m.source_name} (Đích: {m.target_name})"

        disp_amt = (
            m.new_amount
            if m.new_amount is not None
            else (m.source_amount if m.source_amount is not None else m.old_amount)
        )

        rows.append(
            PreviewRow(
                index=idx,
                full_name=full_name,
                source_name=m.source_name,
                target_name=m.target_name,
                card_no=m.card_no or "",
                account_no=m.account_no or "",
                bank_code=m.bank_code or "",
                old_amount=m.old_amount,
                source_amount=m.source_amount,
                new_amount=m.new_amount,
                display_amount=disp_amt,
                match_status=m.status,
                validation_status=val_status,
                validation_detail=val_detail_str,
                exception_text=" | ".join(exc_parts),
                exception_codes=tuple(e.code.value for e in attached_excs),
                exception_ids=m.exception_ids,
                allowed_actions=first_exc.allowed_actions if first_exc else (),
                selected_action=first_exc.action if first_exc else None,
                override_amount=first_exc.override_amount if first_exc else None,
                has_exception=has_exc,
                has_warning=has_warning,
                is_blocking=is_blocking,
                is_unresolved=is_unresolved,
                row_tag=row_tag,
                source_row=m.source_row,
                target_file=Path(m.target_file).name if m.target_file else "",
                target_sheet=m.target_sheet or "",
                target_row=m.target_row,
                target_cell=m.target_cell or "",
                note=m.note,
            )
        )

    return rows


def _collect_row_safety_findings(
    m: MatchItem,
    *,
    safety_by_card: dict[str, list[SafetyFinding]],
    safety_by_acc: dict[str, list[SafetyFinding]],
    safety_by_target_row: dict[tuple[str, str, int], list[SafetyFinding]],
) -> list[SafetyFinding]:
    found: list[SafetyFinding] = []
    if m.card_no and m.card_no in safety_by_card:
        for f in safety_by_card[m.card_no]:
            if f not in found:
                found.append(f)
    if m.account_no and m.account_no in safety_by_acc:
        for f in safety_by_acc[m.account_no]:
            if f not in found:
                found.append(f)
    if m.target_file and m.target_sheet and m.target_row is not None:
        key = (m.target_file, m.target_sheet, m.target_row)
        for f in safety_by_target_row.get(key, []):
            if f not in found:
                found.append(f)
    return found


def _determine_row_tag(
    *,
    match_status: MatchStatus,
    has_exception: bool,
    is_blocking: bool,
    is_unresolved: bool,
    has_exc_error: bool,
    has_safety_error: bool,
    has_warning: bool,
) -> str:
    """Quyết định tag màu hiển thị cho dòng:
    - `TAG_EXCEPTION`: Dòng có lỗi chặn, ngoại lệ chưa xử lý, ngoại lệ mức ERROR, hoặc lỗi Safety.
    - `TAG_WARNING`: Dòng có cảnh báo (Warning) hoặc ngoại lệ đã xử lý / mức WARNING.
    - `TAG_UPDATE`: Dòng sẽ cập nhật bình thường (`WILL_UPDATE`).
    - `TAG_SKIPPED`: Dòng đã bỏ qua (`SKIPPED`).
    - `TAG_UNCHANGED`: Dòng giữ nguyên (`UNCHANGED`).
    """
    if (
        is_blocking
        or is_unresolved
        or has_exc_error
        or has_safety_error
        or match_status in (MatchStatus.ERROR, MatchStatus.PENDING)
    ):
        return TAG_EXCEPTION
    if has_warning or has_exception:
        return TAG_WARNING
    if match_status == MatchStatus.WILL_UPDATE:
        return TAG_UPDATE
    if match_status == MatchStatus.SKIPPED:
        return TAG_SKIPPED
    return TAG_UNCHANGED


# =============================================================================
# 3. Widget bảng PreviewTable
# =============================================================================

ResolveCallback = Callable[[str, ResolutionAction, Optional[int]], Any]


class PreviewTable(ttk.Frame):
    """Bảng xem trước dữ liệu đối chiếu với Filter, Search, Sort, Highlight Exception & Highlight Warning."""

    def __init__(
        self,
        master: Optional[tk.Misc] = None,
        *,
        on_resolve_exception: Optional[ResolveCallback] = None,
    ) -> None:
        super().__init__(master, padding=6)
        self.on_resolve_exception: Optional[ResolveCallback] = on_resolve_exception

        self._write_plan: Optional[WritePlan] = None
        self._validation_result: Optional[ValidationResult] = None
        self._exceptions: list[ExceptionItem] = []
        self._all_rows: list[PreviewRow] = []
        self._visible_rows: list[PreviewRow] = []
        self._row_by_iid: dict[str, PreviewRow] = {}

        # Trạng thái Sort hiện tại
        self._sort_column: Optional[str] = None
        self._sort_descending: bool = False

        # Biến điều khiển Filter & Search
        self.search_var = tk.StringVar(master=self, value="")
        self.filter_var = tk.StringVar(master=self, value=PreviewFilter.ALL.value)
        self.validation_filter_var = tk.StringVar(master=self, value=ValidationFilter.ALL.value)
        self.only_exceptions_var = tk.BooleanVar(master=self, value=False)
        self.only_warnings_var = tk.BooleanVar(master=self, value=False)
        self.count_label_var = tk.StringVar(master=self, value="Hiển thị 0 / 0 dòng")

        # Biến điều khiển xử lý nhanh Exception của dòng đang chọn
        self.selected_exc_info_var = tk.StringVar(
            master=self,
            value="Chọn 1 dòng có Exception trên bảng để xem chi tiết hoặc đổi phương án xử lý.",
        )
        self.selected_action_var = tk.StringVar(master=self, value="")
        self.override_amount_var = tk.StringVar(master=self, value="")
        self._selected_exc_id: Optional[str] = None

        self._build_ui()

    # ------------------------------------------------------------------ Dựng giao diện
    def _build_ui(self) -> None:
        # 1. Thanh công cụ Filter & Search
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(toolbar, text="Tìm kiếm:").pack(side=tk.LEFT, padx=(0, 4))
        self.entry_search = ttk.Entry(toolbar, textvariable=self.search_var, width=26)
        self.entry_search.pack(side=tk.LEFT, padx=(0, 4))
        self.entry_search.bind("<KeyRelease>", lambda _e: self.apply_filter_and_sort())

        self.btn_clear_search = ttk.Button(
            toolbar,
            text="Xóa tìm",
            width=8,
            command=self.clear_search,
        )
        self.btn_clear_search.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(toolbar, text="Lọc Match:").pack(side=tk.LEFT, padx=(0, 4))
        self.cmb_filter = ttk.Combobox(
            toolbar,
            textvariable=self.filter_var,
            values=[f.value for f in PreviewFilter],
            width=18,
            state="readonly",
        )
        self.cmb_filter.pack(side=tk.LEFT, padx=(0, 12))
        self.cmb_filter.bind("<<ComboboxSelected>>", lambda _e: self.apply_filter_and_sort())

        ttk.Label(toolbar, text="Lọc Validation:").pack(side=tk.LEFT, padx=(0, 4))
        self.cmb_validation_filter = ttk.Combobox(
            toolbar,
            textvariable=self.validation_filter_var,
            values=[f.value for f in ValidationFilter],
            width=12,
            state="readonly",
        )
        self.cmb_validation_filter.pack(side=tk.LEFT, padx=(0, 12))
        self.cmb_validation_filter.bind("<<ComboboxSelected>>", lambda _e: self.apply_filter_and_sort())

        self.chk_only_exceptions = ttk.Checkbutton(
            toolbar,
            text="Chỉ Exception",
            variable=self.only_exceptions_var,
            command=self.apply_filter_and_sort,
        )
        self.chk_only_exceptions.pack(side=tk.LEFT, padx=(0, 8))

        self.chk_only_warnings = ttk.Checkbutton(
            toolbar,
            text="Chỉ Warning",
            variable=self.only_warnings_var,
            command=self.apply_filter_and_sort,
        )
        self.chk_only_warnings.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_reset_filters = ttk.Button(
            toolbar,
            text="Đặt lại bộ lọc",
            command=self.reset_filters,
        )
        self.btn_reset_filters.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(toolbar, textvariable=self.count_label_var).pack(side=tk.RIGHT)

        # 2. Bảng Treeview hiển thị 8 nhóm cột dữ liệu
        table_frame = ttk.Frame(self)
        table_frame.pack(fill=tk.BOTH, expand=True)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        col_ids = tuple(c.value for c in PreviewColumn)
        self.tree = ttk.Treeview(
            table_frame,
            columns=col_ids,
            show="headings",
            selectmode="browse",
            height=14,
        )
        self.tree.grid(row=0, column=0, sticky=tk.NSEW)

        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky=tk.NS)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        hsb.grid(row=1, column=0, sticky=tk.EW)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        col_widths: dict[str, tuple[int, Any]] = {
            PreviewColumn.STT.value: (48, tk.CENTER),
            PreviewColumn.FULL_NAME.value: (210, tk.W),
            PreviewColumn.CARD_NO.value: (95, tk.W),
            PreviewColumn.ACCOUNT_NO.value: (125, tk.W),
            PreviewColumn.BANK_CODE.value: (100, tk.CENTER),
            PreviewColumn.AMOUNT.value: (135, tk.E),
            PreviewColumn.MATCH_STATUS.value: (135, tk.CENTER),
            PreviewColumn.VALIDATION_STATUS.value: (155, tk.CENTER),
            PreviewColumn.EXCEPTION.value: (320, tk.W),
        }
        for cid in col_ids:
            width, anchor = col_widths[cid]
            self.tree.heading(
                cid,
                text=_COLUMN_HEADERS[cid],
                command=self._make_sort_command(cid),
            )
            self.tree.column(cid, width=width, anchor=anchor, stretch=(cid == PreviewColumn.EXCEPTION.value))

        # Cấu hình màu Highlight Exception, Highlight Warning và các trạng thái
        self.tree.tag_configure(
            TAG_EXCEPTION,
            background=COLOR_EXCEPTION_BG,
            foreground=COLOR_EXCEPTION_FG,
        )
        self.tree.tag_configure(
            TAG_WARNING,
            background=COLOR_WARNING_BG,
            foreground=COLOR_WARNING_FG,
        )
        self.tree.tag_configure(
            TAG_UPDATE,
            background=COLOR_UPDATE_BG,
            foreground=COLOR_UPDATE_FG,
        )
        self.tree.tag_configure(
            TAG_UNCHANGED,
            background=COLOR_UNCHANGED_BG,
            foreground=COLOR_UNCHANGED_FG,
        )
        self.tree.tag_configure(
            TAG_SKIPPED,
            background=COLOR_SKIPPED_BG,
            foreground=COLOR_SKIPPED_FG,
        )

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        # 3. Thanh chú giải màu & khung chọn xử lý Exception cho dòng đang chọn
        bottom_frame = ttk.LabelFrame(
            self,
            text=" Chú giải màu & Xử lý dòng Exception đang chọn ",
            padding=6,
        )
        bottom_frame.pack(fill=tk.X, pady=(6, 0))

        legend_row = ttk.Frame(bottom_frame)
        legend_row.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(
            legend_row,
            text=(
                "Chú giải:  🔴 Đỏ nhạt = Exception / Chặn ghi   |   "
                "🟡 Vàng nhạt = Warning / Ngoại lệ đã xác nhận   |   "
                "🟢 Xanh nhạt = Sẽ cập nhật   |   "
                "⚪ Trắng/Xám = Không đổi / Đã bỏ qua"
            ),
        ).pack(side=tk.LEFT)

        exc_row = ttk.Frame(bottom_frame)
        exc_row.pack(fill=tk.X)
        ttk.Label(exc_row, textvariable=self.selected_exc_info_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8)
        )

        ttk.Label(exc_row, text="Xử lý:").pack(side=tk.LEFT, padx=(0, 4))
        self.cmb_resolution = ttk.Combobox(
            exc_row,
            textvariable=self.selected_action_var,
            width=18,
            state="disabled",
        )
        self.cmb_resolution.pack(side=tk.LEFT, padx=(0, 6))

        ttk.Label(exc_row, text="Số tiền sửa tay:").pack(side=tk.LEFT, padx=(0, 4))
        self.entry_override_amount = ttk.Entry(
            exc_row,
            textvariable=self.override_amount_var,
            width=14,
            state="disabled",
        )
        self.entry_override_amount.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_apply_resolution = ttk.Button(
            exc_row,
            text="Áp dụng xử lý",
            command=self.apply_selected_resolution,
            state=tk.DISABLED,
        )
        self.btn_apply_resolution.pack(side=tk.LEFT)

    # =========================================================================
    # 4. Nạp dữ liệu vào bảng Preview
    # =========================================================================

    def load_data(
        self,
        write_plan: Optional[WritePlan] = None,
        *,
        validation_result: Optional[ValidationResult] = None,
        exceptions: Optional[Sequence[ExceptionItem]] = None,
    ) -> list[PreviewRow]:
        """Nạp `WritePlan` (kèm `ValidationResult` và `ExceptionItem`) vào bảng Preview."""
        self._write_plan = write_plan
        self._validation_result = validation_result
        if exceptions is not None:
            self._exceptions = list(exceptions)
        elif write_plan is not None:
            self._exceptions = list(write_plan.exceptions)
        else:
            self._exceptions = []

        self._all_rows = build_preview_rows(
            write_plan,
            validation_result=validation_result,
            exceptions=self._exceptions,
        )
        self.apply_filter_and_sort()
        return list(self._visible_rows)

    def load_plan(
        self,
        plan_or_result: Union[WritePlan, WorkflowResult, None],
        validation_result: Optional[ValidationResult] = None,
    ) -> list[PreviewRow]:
        """Nạp trực tiếp từ `WritePlan` hoặc `WorkflowResult`."""
        if isinstance(plan_or_result, WorkflowResult):
            return self.set_workflow_result(plan_or_result)
        return self.load_data(plan_or_result, validation_result=validation_result)

    def set_workflow_result(self, workflow_result: WorkflowResult) -> list[PreviewRow]:
        """Tương thích với `MainWindow._refresh_embedded_preview`."""
        return self.load_data(
            workflow_result.write_plan,
            validation_result=workflow_result.validation_result,
            exceptions=workflow_result.exceptions,
        )

    def clear(self) -> None:
        """Xóa toàn bộ dữ liệu trên bảng Preview."""
        self.load_data(None)

    # =========================================================================
    # 5. Chức năng Filter, Search và Sort
    # =========================================================================

    def set_search_query(self, query: str) -> list[PreviewRow]:
        """Đặt từ khóa tìm kiếm và cập nhật bảng."""
        self.search_var.set(query)
        return self.apply_filter_and_sort()

    def clear_search(self) -> list[PreviewRow]:
        """Xóa từ khóa tìm kiếm."""
        self.search_var.set("")
        return self.apply_filter_and_sort()

    def set_filter(self, filter_mode: Union[PreviewFilter, str]) -> list[PreviewRow]:
        """Đặt bộ lọc theo trạng thái Match / Exception."""
        val = filter_mode.value if isinstance(filter_mode, PreviewFilter) else str(filter_mode)
        self.filter_var.set(val)
        return self.apply_filter_and_sort()

    def set_validation_filter(self, status: Union[ValidationFilter, str]) -> list[PreviewRow]:
        """Đặt bộ lọc theo trạng thái Validation (`ALL`, `OK`, `WARNING`, `ERROR`)."""
        val = status.value if isinstance(status, ValidationFilter) else str(status)
        self.validation_filter_var.set(val)
        return self.apply_filter_and_sort()

    def set_highlight_filters(
        self,
        *,
        only_exceptions: Optional[bool] = None,
        only_warnings: Optional[bool] = None,
    ) -> list[PreviewRow]:
        """Bật/tắt chế độ chỉ hiển thị dòng Exception hoặc chỉ dòng Warning."""
        if only_exceptions is not None:
            self.only_exceptions_var.set(only_exceptions)
        if only_warnings is not None:
            self.only_warnings_var.set(only_warnings)
        return self.apply_filter_and_sort()

    def reset_filters(self) -> list[PreviewRow]:
        """Đặt lại toàn bộ bộ lọc, tìm kiếm và thứ tự sắp xếp về mặc định."""
        self.search_var.set("")
        self.filter_var.set(PreviewFilter.ALL.value)
        self.validation_filter_var.set(ValidationFilter.ALL.value)
        self.only_exceptions_var.set(False)
        self.only_warnings_var.set(False)
        self._sort_column = None
        self._sort_descending = False
        self._update_heading_indicators()
        return self.apply_filter_and_sort()

    def sort_by(
        self,
        column: Union[PreviewColumn, str],
        *,
        descending: Optional[bool] = None,
    ) -> list[PreviewRow]:
        """Sắp xếp bảng theo cột chỉ định. Nếu không truyền `descending`, tự động đảo chiều khi bấm lặp lại."""
        col_name = column.value if isinstance(column, PreviewColumn) else str(column)
        if col_name not in _COLUMN_HEADERS:
            raise ValueError(f"Cột sắp xếp không hợp lệ: {col_name!r}")

        if descending is not None:
            self._sort_descending = descending
        elif self._sort_column == col_name:
            self._sort_descending = not self._sort_descending
        else:
            self._sort_descending = False

        self._sort_column = col_name
        self._update_heading_indicators()
        return self.apply_filter_and_sort()

    def apply_filter_and_sort(self) -> list[PreviewRow]:
        """Áp dụng đồng thời Search, Filter và Sort lên danh sách dòng và vẽ lại Treeview."""
        query = self.search_var.get().strip()
        query_norm = remove_diacritics(query).casefold() if query else ""

        mode = self.filter_var.get().strip() or PreviewFilter.ALL.value
        val_mode = self.validation_filter_var.get().strip() or ValidationFilter.ALL.value
        only_exc = bool(self.only_exceptions_var.get())
        only_warn = bool(self.only_warnings_var.get())

        filtered: list[PreviewRow] = []
        for r in self._all_rows:
            if not _matches_filter_mode(r, mode):
                continue
            if not _matches_validation_filter(r, val_mode):
                continue
            if only_exc and not (r.has_exception or r.row_tag == TAG_EXCEPTION):
                continue
            if only_warn and not (r.has_warning or r.row_tag == TAG_WARNING):
                continue
            if query_norm and not _matches_search_query(r, query_norm):
                continue
            filtered.append(r)

        if self._sort_column is not None:
            col = self._sort_column
            filtered.sort(
                key=lambda item: _sort_key_for_column(item, col),
                reverse=self._sort_descending,
            )

        self._visible_rows = filtered
        self._populate_tree()
        return list(self._visible_rows)

    # =========================================================================
    # 6. Truy vấn dữ liệu & Xử lý sự kiện chọn dòng
    # =========================================================================

    def all_rows(self) -> list[PreviewRow]:
        return list(self._all_rows)

    def visible_rows(self) -> list[PreviewRow]:
        return list(self._visible_rows)

    def exception_rows(self) -> list[PreviewRow]:
        return [r for r in self._visible_rows if r.has_exception or r.row_tag == TAG_EXCEPTION]

    def warning_rows(self) -> list[PreviewRow]:
        return [r for r in self._visible_rows if r.has_warning or r.row_tag == TAG_WARNING]

    def get_row_tags(self, row_index: int) -> tuple[str, ...]:
        """Lấy danh sách tag màu của dòng thứ `row_index` (0-based) đang hiển thị trên Treeview."""
        children = self.tree.get_children()
        if not 0 <= row_index < len(children):
            raise IndexError(f"Chỉ số dòng nằm ngoài phạm vi hiển thị: {row_index}")
        iid = children[row_index]
        raw_tags: Any = self.tree.item(iid, "tags")
        if isinstance(raw_tags, (list, tuple)):
            return tuple(str(t) for t in raw_tags)
        if isinstance(raw_tags, str) and raw_tags:
            return (raw_tags,)
        return ()

    def apply_selected_resolution(self) -> bool:
        """Áp dụng lựa chọn xử lý ngoại lệ cho dòng đang chọn trên bảng."""
        if not self._selected_exc_id:
            return False
        act_str = self.selected_action_var.get().strip()
        if not act_str:
            return False

        action = ResolutionAction(act_str)
        override_val: Optional[int] = None
        if action == ResolutionAction.OVERRIDE_AMOUNT:
            raw_amt = self.override_amount_var.get().replace(",", "").replace(".", "").strip()
            if not raw_amt.isdigit() or int(raw_amt) <= 0:
                return False
            override_val = int(raw_amt)

        if self.on_resolve_exception is not None:
            self.on_resolve_exception(self._selected_exc_id, action, override_val)
            return True

        if self._write_plan is not None:
            self._write_plan.resolve(
                self._selected_exc_id,
                action,
                override_amount=override_val,
            )
            self.load_data(
                self._write_plan,
                validation_result=self._validation_result,
                exceptions=self._write_plan.exceptions,
            )
            return True

        return False

    # =========================================================================
    # 7. Các hàm nội bộ
    # =========================================================================

    def _make_sort_command(self, col_id: str) -> Callable[[], None]:
        def _cmd() -> None:
            self.sort_by(col_id)

        return _cmd

    def _populate_tree(self) -> None:
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self._row_by_iid.clear()

        for row in self._visible_rows:
            iid = f"row_{row.index}"
            self._row_by_iid[iid] = row
            self.tree.insert(
                "",
                tk.END,
                iid=iid,
                values=row.values_tuple(),
                tags=(row.row_tag,),
            )

        self.count_label_var.set(
            f"Hiển thị {len(self._visible_rows):,} / {len(self._all_rows):,} dòng"
        )

    def _update_heading_indicators(self) -> None:
        for col in PreviewColumn:
            cid = col.value
            base_text = _COLUMN_HEADERS[cid]
            if self._sort_column == cid:
                arrow = " ▼" if self._sort_descending else " ▲"
                self.tree.heading(cid, text=f"{base_text}{arrow}")
            else:
                self.tree.heading(cid, text=base_text)

    def _on_tree_select(self, _event: Any = None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        row = self._row_by_iid.get(sel[0])
        if row is None or not row.exception_ids:
            self._selected_exc_id = None
            self.selected_exc_info_var.set(
                f"Dòng #{row.index}: Không có ngoại lệ cần xử lý." if row else ""
            )
            self.cmb_resolution.configure(state="disabled", values=[])
            self.entry_override_amount.configure(state="disabled")
            self.btn_apply_resolution.configure(state=tk.DISABLED)
            return

        self._selected_exc_id = row.exception_ids[0]
        self.selected_exc_info_var.set(
            f"Dòng #{row.index} [{self._selected_exc_id}]: {row.exception_text}"
        )
        allowed = [a.value for a in row.allowed_actions]
        if allowed and not row.is_blocking:
            self.cmb_resolution.configure(state="readonly", values=allowed)
            self.selected_action_var.set(
                row.selected_action.value if row.selected_action else allowed[0]
            )
            self.entry_override_amount.configure(state="normal")
            self.override_amount_var.set(
                str(row.override_amount) if row.override_amount is not None else ""
            )
            self.btn_apply_resolution.configure(state=tk.NORMAL)
        else:
            self.cmb_resolution.configure(state="disabled", values=[])
            self.selected_action_var.set("")
            self.entry_override_amount.configure(state="disabled")
            self.btn_apply_resolution.configure(state=tk.DISABLED)


# =============================================================================
# 4. Các hàm phụ trợ Lọc, Tìm kiếm, Sắp xếp
# =============================================================================

def _matches_filter_mode(row: PreviewRow, mode: str) -> bool:
    if mode == PreviewFilter.ALL.value:
        return True
    if mode == PreviewFilter.WILL_UPDATE.value:
        return row.match_status == MatchStatus.WILL_UPDATE
    if mode == PreviewFilter.UNCHANGED.value:
        return row.match_status == MatchStatus.UNCHANGED
    if mode == PreviewFilter.PENDING.value:
        return row.match_status == MatchStatus.PENDING
    if mode == PreviewFilter.SKIPPED.value:
        return row.match_status == MatchStatus.SKIPPED
    if mode == PreviewFilter.ERROR.value:
        return row.match_status == MatchStatus.ERROR
    if mode == PreviewFilter.EXCEPTIONS_ONLY.value:
        return row.has_exception or row.row_tag == TAG_EXCEPTION
    if mode == PreviewFilter.UNRESOLVED_ONLY.value:
        return row.is_unresolved or row.is_blocking
    if mode == PreviewFilter.WARNINGS_ONLY.value:
        return row.has_warning or row.row_tag == TAG_WARNING
    return True


def _matches_validation_filter(row: PreviewRow, val_mode: str) -> bool:
    if val_mode == ValidationFilter.ALL.value:
        return True
    return row.validation_status == val_mode


def _matches_search_query(row: PreviewRow, query_norm: str) -> bool:
    haystack = " | ".join([
        row.full_name,
        row.source_name,
        row.target_name,
        row.card_no,
        row.account_no,
        row.bank_code,
        row.match_status.value,
        row.validation_status,
        row.exception_text,
        row.note,
        str(row.display_amount or ""),
    ])
    return query_norm in remove_diacritics(haystack).casefold()


def _sort_key_for_column(row: PreviewRow, column: str) -> tuple[int, Any]:
    if column == PreviewColumn.STT.value:
        return (0, row.index)
    if column == PreviewColumn.FULL_NAME.value:
        return (0, remove_diacritics(row.full_name).casefold())
    if column == PreviewColumn.CARD_NO.value:
        return (1 if not row.card_no else 0, row.card_no.casefold())
    if column == PreviewColumn.ACCOUNT_NO.value:
        return (1 if not row.account_no else 0, row.account_no.casefold())
    if column == PreviewColumn.BANK_CODE.value:
        return (1 if not row.bank_code else 0, row.bank_code.casefold())
    if column == PreviewColumn.AMOUNT.value:
        return (1 if row.display_amount is None else 0, row.display_amount or 0)
    if column == PreviewColumn.MATCH_STATUS.value:
        return (0, row.match_status.value)
    if column == PreviewColumn.VALIDATION_STATUS.value:
        order = {"ERROR": 0, "WARNING": 1, "OK": 2}.get(row.validation_status, 3)
        return (0, order)
    if column == PreviewColumn.EXCEPTION.value:
        return (1 if not row.exception_text else 0, row.exception_text.casefold())
    return (0, row.index)
