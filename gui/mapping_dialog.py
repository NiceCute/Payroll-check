"""
mapping_dialog.py - Hộp thoại cấu hình Mapping Cột Nguồn → Cột Đích (GUI Module 3).

Chức năng:
- Hiển thị danh sách cột nguồn (quét trực tiếp từ file lương Excel hoặc chọn chữ cái cột A..Z).
- Hiển thị danh sách cột đích (quét trực tiếp từ file mẫu ngân hàng Excel hoặc chọn chữ cái cột A..Z).
- Cho phép ánh xạ trực quan Source → Target cho các trường nghiệp vụ:
  + Số thẻ (`card_no` - khóa lọc nhân sự nguồn)
  + Số tài khoản (`account_no` - khóa đối chiếu Source → Target)
  + Họ và tên (`full_name` - kiểm tra chéo Source → Target)
  + Số tiền (`amount` - giá trị chuyển khoản Source → ô cập nhật Target)
  + Mã ngân hàng (`bank_code` - tùy chọn trên Target)
- Tạo mới / Sửa / Xóa cấu hình mapping.
- Lưu mapping vào `templates/*.json` thông qua `TemplateStore`.
- Tải lại các mapping đã lưu từ `templates/*.json`.
- Validate đầy đủ các trường bắt buộc, dòng tiêu đề, dòng bắt đầu dữ liệu và phát hiện trùng/thiếu cột trước khi lưu.

Nguyên tắc:
- Không xử lý workflow, không xử lý excel writer, không xử lý business logic đối chiếu.
- Chỉ phụ thuộc `core.template_store`, `core.excel_reader`, `core.models`.
"""

from __future__ import annotations

import re
import tkinter as tk
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Optional, Sequence

from core.excel_reader import (
    ExcelReadError,
    PathLike,
    column_headers,
    list_sheets,
    read_sheet,
)
from core.models import (
    ColumnRef,
    ConfigError,
    SOURCE_FIELDS,
    SourceProfile,
    TARGET_FIELDS,
    TargetRules,
    TargetTemplate,
    column_index,
    column_letter,
)
from core.template_store import TemplateStore, TemplateStoreError

__all__ = [
    "ColumnOption",
    "MappingDialog",
    "MappingValidationResult",
    "SourceTargetLink",
    "extract_sheet_column_options",
    "open_mapping_dialog",
    "parse_column_choice",
    "parse_header_rows_text",
]

DEFAULT_COLUMN_LETTERS: tuple[str, ...] = tuple(column_letter(i) for i in range(1, 53))  # A..AZ


# =============================================================================
# 1. Cấu trúc dữ liệu hỗ trợ hiển thị & kiểm tra Mapping
# =============================================================================

@dataclass(frozen=True)
class ColumnOption:
    """1 cột trong sheet Excel (chữ cái cột + số thứ tự + tiêu đề thực tế)."""

    letter: str
    index: int
    header: str = ""

    @property
    def display(self) -> str:
        if self.header:
            return f"{self.letter} - {self.header}"
        return self.letter


@dataclass(frozen=True)
class SourceTargetLink:
    """1 dòng liên kết ánh xạ giữa Trường Nguồn (Bảng lương) và Trường Đích (Mẫu ngân hàng)."""

    field_key: str
    field_label: str
    source_column: str
    source_header: str
    target_column: str
    target_header: str
    is_required_source: bool
    is_required_target: bool
    status: str  # "OK", "MISSING", "DUPLICATE"


@dataclass
class MappingValidationResult:
    """Kết quả kiểm tra hợp lệ của cấu hình Mapping trước khi lưu."""

    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        return not self.errors

    @property
    def ok(self) -> bool:
        return not self.errors


def parse_column_choice(raw: str) -> str:
    """Trích chữ cái cột từ chuỗi chọn trên Combobox (ví dụ `'P - Hồng Lâm Chuyển Khoản'` -> `'P'`)."""
    text = raw.strip()
    if not text or text.startswith("("):
        return ""
    if " - " in text:
        text = text.split(" - ", 1)[0].strip()
    return text.upper()


def parse_header_rows_text(raw: str) -> tuple[int, ...]:
    """Chuyển chuỗi dòng tiêu đề (ví dụ `'3, 4'` hoặc `'3-4'` hoặc `'2'`) thành `tuple[int, ...]`."""
    text = raw.strip()
    if not text:
        return ()
    result: list[int] = []
    for part in re.split(r"[,;\s]+", text):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            bounds = part.split("-", 1)
            start_s, end_s = bounds[0].strip(), bounds[1].strip()
            if not (start_s.isdigit() and end_s.isdigit()):
                raise ValueError(f"Khoảng dòng tiêu đề không hợp lệ: {part!r}")
            start_r, end_r = int(start_s), int(end_s)
            if start_r < 1 or end_r < start_r:
                raise ValueError(f"Khoảng dòng tiêu đề không hợp lệ: {part!r}")
            for r in range(start_r, end_r + 1):
                if r not in result:
                    result.append(r)
        else:
            if not part.isdigit() or int(part) < 1:
                raise ValueError(f"Số dòng tiêu đề phải là số nguyên dương: {part!r}")
            val = int(part)
            if val not in result:
                result.append(val)
    return tuple(result)


def extract_sheet_column_options(
    file_path: PathLike,
    sheet_name: str,
    header_rows: Sequence[int],
) -> list[ColumnOption]:
    """Đọc sheet Excel và trả về danh sách các cột kèm tiêu đề ghép từ `header_rows`."""
    sheet = read_sheet(file_path, sheet_name)
    hdr_map = column_headers(sheet, list(header_rows))
    max_col = max(sheet.max_col, 26)
    options: list[ColumnOption] = []
    for col_idx in range(1, max_col + 1):
        letter = column_letter(col_idx)
        hdr = hdr_map.get(col_idx, "")
        options.append(ColumnOption(letter=letter, index=col_idx, header=hdr))
    return options


# =============================================================================
# 2. Hộp thoại cấu hình MappingDialog
# =============================================================================

class MappingDialog(tk.Toplevel):
    """Hộp thoại trực quan cho kế toán xem cột nguồn/đích, ánh xạ Source → Target và quản lý JSON template."""

    def __init__(
        self,
        parent: Optional[tk.Misc] = None,
        *,
        template_store: Optional[TemplateStore] = None,
        templates_dir: PathLike = "templates",
        source_excel_path: Optional[PathLike] = None,
        target_excel_path: Optional[PathLike] = None,
        interactive: bool = True,
    ) -> None:
        super().__init__(master=parent)
        self.title("Cấu hình Mapping Cột Bảng Lương Nguồn → Mẫu Ngân Hàng Đích")
        self.geometry("1160x780")
        self.minsize(980, 660)

        self.interactive: bool = interactive
        self.store: TemplateStore = template_store or TemplateStore(templates_dir)
        self.store.ensure_defaults()

        # Danh sách cột quét từ file Excel mẫu (nếu có)
        self._source_col_options: list[ColumnOption] = [
            ColumnOption(letter=ch, index=column_index(ch)) for ch in DEFAULT_COLUMN_LETTERS
        ]
        self._target_col_options: list[ColumnOption] = [
            ColumnOption(letter=ch, index=column_index(ch)) for ch in DEFAULT_COLUMN_LETTERS
        ]
        self._source_header_by_letter: dict[str, str] = {}
        self._target_header_by_letter: dict[str, str] = {}

        # Tên gốc khi đang sửa 1 cấu hình đã lưu (để hỗ trợ đổi tên)
        self._loaded_source_name: Optional[str] = None
        self._loaded_target_name: Optional[str] = None

        # ---------------- Biến Tkinter cho Bảng lương nguồn (SourceProfile) ----------------
        self.saved_source_var = tk.StringVar(master=self, value="")
        self.source_sample_file_var = tk.StringVar(master=self, value="")
        self.source_name_var = tk.StringVar(master=self, value="")
        self.source_sheet_var = tk.StringVar(master=self, value="")
        self.source_sheet_patterns_var = tk.StringVar(master=self, value="HLCK*, *LUONG*")
        self.source_header_rows_var = tk.StringVar(master=self, value="3, 4")
        self.source_data_start_row_var = tk.StringVar(master=self, value="5")
        self.source_data_end_row_var = tk.StringVar(master=self, value="")

        # Các trường cột nguồn: card_no, full_name, account_no, amount
        self.src_col_vars: dict[str, tk.StringVar] = {
            k: tk.StringVar(master=self, value="") for k in SOURCE_FIELDS
        }
        self.src_hdr_vars: dict[str, tk.StringVar] = {
            k: tk.StringVar(master=self, value="") for k in SOURCE_FIELDS
        }

        # ---------------- Biến Tkinter cho Mẫu ngân hàng đích (TargetTemplate) ----------------
        self.saved_target_var = tk.StringVar(master=self, value="")
        self.target_sample_file_var = tk.StringVar(master=self, value="")
        self.target_name_var = tk.StringVar(master=self, value="")
        self.target_desc_var = tk.StringVar(master=self, value="")
        self.target_file_patterns_var = tk.StringVar(master=self, value="")
        self.target_sheet_var = tk.StringVar(master=self, value="")
        self.target_sheet_patterns_var = tk.StringVar(master=self, value="")
        self.target_header_rows_var = tk.StringVar(master=self, value="2")
        self.target_data_start_row_var = tk.StringVar(master=self, value="3")
        self.target_data_end_row_var = tk.StringVar(master=self, value="")

        # Quy tắc nhập liệu TargetRules
        self.rule_no_diacritics_var = tk.BooleanVar(master=self, value=False)
        self.rule_no_extra_spaces_var = tk.BooleanVar(master=self, value=False)
        self.rule_max_len_var = tk.StringVar(master=self, value="")
        self.rule_forbidden_chars_var = tk.StringVar(master=self, value="")

        # Các trường cột đích: account_no, full_name, amount, bank_code
        self.tgt_col_vars: dict[str, tk.StringVar] = {
            k: tk.StringVar(master=self, value="") for k in TARGET_FIELDS
        }
        self.tgt_hdr_vars: dict[str, tk.StringVar] = {
            k: tk.StringVar(master=self, value="") for k in TARGET_FIELDS
        }

        # Biến hiển thị cảnh báo / trạng thái kiểm tra
        self.validation_message_var = tk.StringVar(
            master=self,
            value="Sẵn sàng. Chọn mapping đã lưu hoặc nạp file Excel mẫu để xem danh sách cột.",
        )

        self._src_col_combos: dict[str, ttk.Combobox] = {}
        self._tgt_col_combos: dict[str, ttk.Combobox] = {}

        self._build_ui()
        self.reload_saved_mappings()

        # Tự nạp cấu hình đầu tiên nếu có
        sources = self.store.list_sources()
        if sources:
            self.load_source_mapping(sources[0].name)
        targets = self.store.list_targets()
        if targets:
            self.load_target_mapping(targets[0].name)

        if source_excel_path:
            self.inspect_source_excel(source_excel_path)
        if target_excel_path:
            self.inspect_target_excel(target_excel_path)

        self.refresh_mapping_preview()

    # ------------------------------------------------------------------ Dựng giao diện
    def _build_ui(self) -> None:
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # Thanh công cụ trên cùng
        top_bar = ttk.Frame(main_frame)
        top_bar.grid(row=0, column=0, columnspan=2, sticky=tk.EW, pady=(0, 8))
        ttk.Label(
            top_bar,
            text="CẤU HÌNH MAPPING CỘT NGUỒN (BẢNG LƯƠNG) → CỘT ĐÍCH (MẪU NGÂN HÀNG)",
            font=("Arial", 12, "bold"),
            foreground="#1F4E78",
        ).pack(side=tk.LEFT)

        self.btn_reload_all = ttk.Button(
            top_bar,
            text="🔄 Tải lại Mapping đã lưu",
            command=self.reload_saved_mappings,
        )
        self.btn_reload_all.pack(side=tk.RIGHT, padx=(6, 0))

        self.btn_validate_now = ttk.Button(
            top_bar,
            text="✔ Kiểm tra hợp lệ (Validate)",
            command=self.run_interactive_validation,
        )
        self.btn_validate_now.pack(side=tk.RIGHT)

        # ==============================================================
        # NỬA TRÁI: CẤU HÌNH CỘT NGUỒN (SOURCE PROFILE)
        # ==============================================================
        src_frame = ttk.LabelFrame(
            main_frame,
            text=" 1. Cột Nguồn — File Bảng Lương (SourceProfile) ",
            padding=8,
        )
        src_frame.grid(row=1, column=0, sticky=tk.NSEW, padx=(0, 6))
        src_frame.columnconfigure(1, weight=1)

        # Chọn cấu hình nguồn đã lưu + Tạo mới / Lưu / Xóa
        src_saved_bar = ttk.Frame(src_frame)
        src_saved_bar.grid(row=0, column=0, columnspan=3, sticky=tk.EW, pady=(0, 6))
        ttk.Label(src_saved_bar, text="Mapping nguồn đã lưu:").pack(side=tk.LEFT, padx=(0, 4))
        self.cmb_saved_source = ttk.Combobox(
            src_saved_bar,
            textvariable=self.saved_source_var,
            width=24,
            state="readonly",
        )
        self.cmb_saved_source.pack(side=tk.LEFT, padx=(0, 6))
        self.cmb_saved_source.bind(
            "<<ComboboxSelected>>",
            lambda _e: self._on_select_saved_source(),
        )
        self.btn_new_source = ttk.Button(
            src_saved_bar,
            text="Tạo mới",
            width=8,
            command=self.new_source_mapping,
        )
        self.btn_new_source.pack(side=tk.LEFT, padx=2)
        self.btn_save_source = ttk.Button(
            src_saved_bar,
            text="Lưu nguồn",
            width=10,
            command=self.save_source_mapping_action,
        )
        self.btn_save_source.pack(side=tk.LEFT, padx=2)
        self.btn_delete_source = ttk.Button(
            src_saved_bar,
            text="Xóa",
            width=6,
            command=self.delete_source_mapping_action,
        )
        self.btn_delete_source.pack(side=tk.LEFT, padx=2)

        # Chọn file Excel mẫu nguồn để đọc danh sách cột
        ttk.Label(src_frame, text="File lương mẫu (để quét cột):").grid(
            row=1, column=0, sticky=tk.W, pady=2
        )
        ttk.Entry(src_frame, textvariable=self.source_sample_file_var, state="readonly").grid(
            row=1, column=1, sticky=tk.EW, padx=4, pady=2
        )
        ttk.Button(
            src_frame,
            text="Chọn file...",
            command=self.browse_source_sample_excel,
        ).grid(row=1, column=2, sticky=tk.EW, pady=2)

        # Tên cấu hình nguồn, Sheet, Header rows, Data start row
        ttk.Label(src_frame, text="Tên cấu hình nguồn (*):").grid(row=2, column=0, sticky=tk.W, pady=2)
        ttk.Entry(src_frame, textvariable=self.source_name_var).grid(
            row=2, column=1, columnspan=2, sticky=tk.EW, padx=4, pady=2
        )

        ttk.Label(src_frame, text="Sheet cố định / Mẫu tên sheet:").grid(
            row=3, column=0, sticky=tk.W, pady=2
        )
        src_sheet_box = ttk.Frame(src_frame)
        src_sheet_box.grid(row=3, column=1, columnspan=2, sticky=tk.EW, padx=4, pady=2)
        self.cmb_source_sheet = ttk.Combobox(
            src_sheet_box,
            textvariable=self.source_sheet_var,
            width=18,
        )
        self.cmb_source_sheet.pack(side=tk.LEFT, padx=(0, 6))
        self.cmb_source_sheet.bind("<<ComboboxSelected>>", lambda _e: self._rescan_source_columns())
        ttk.Label(src_sheet_box, text="Mẫu (*):").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(src_sheet_box, textvariable=self.source_sheet_patterns_var, width=22).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )

        ttk.Label(src_frame, text="Dòng tiêu đề / Dòng bắt đầu:").grid(
            row=4, column=0, sticky=tk.W, pady=2
        )
        src_row_box = ttk.Frame(src_frame)
        src_row_box.grid(row=4, column=1, columnspan=2, sticky=tk.EW, padx=4, pady=2)
        ttk.Entry(src_row_box, textvariable=self.source_header_rows_var, width=10).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Label(src_row_box, text="Bắt đầu dữ liệu:").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(src_row_box, textvariable=self.source_data_start_row_var, width=6).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(
            src_row_box,
            text="Quét lại cột",
            command=self._rescan_source_columns,
        ).pack(side=tk.RIGHT)

        # Bảng chọn cột cho 4 trường bắt buộc của Bảng lương nguồn
        src_cols_box = ttk.LabelFrame(
            src_frame,
            text=" Danh sách Trường Nguồn (Bắt buộc) → Chọn Cột & Tiêu đề mong đợi ",
            padding=6,
        )
        src_cols_box.grid(row=5, column=0, columnspan=3, sticky=tk.NSEW, pady=(6, 0))
        src_cols_box.columnconfigure(1, weight=1)
        src_cols_box.columnconfigure(2, weight=1)

        ttk.Label(src_cols_box, text="Trường dữ liệu", font=("Arial", 9, "bold")).grid(
            row=0, column=0, sticky=tk.W, padx=4
        )
        ttk.Label(src_cols_box, text="Cột Excel (Chữ cái)", font=("Arial", 9, "bold")).grid(
            row=0, column=1, sticky=tk.W, padx=4
        )
        ttk.Label(src_cols_box, text="Tiêu đề mong đợi (nếu có)", font=("Arial", 9, "bold")).grid(
            row=0, column=2, sticky=tk.W, padx=4
        )

        for r_idx, (f_key, f_label) in enumerate(SOURCE_FIELDS.items(), start=1):
            ttk.Label(src_cols_box, text=f"{f_label} (*):").grid(
                row=r_idx, column=0, sticky=tk.W, padx=4, pady=3
            )
            cmb = ttk.Combobox(
                src_cols_box,
                textvariable=self.src_col_vars[f_key],
                width=22,
            )
            cmb.grid(row=r_idx, column=1, sticky=tk.EW, padx=4, pady=3)
            cmb.bind(
                "<<ComboboxSelected>>",
                self._make_source_col_selected_handler(f_key),
            )
            self._src_col_combos[f_key] = cmb

            ent_hdr = ttk.Entry(src_cols_box, textvariable=self.src_hdr_vars[f_key])
            ent_hdr.grid(row=r_idx, column=2, sticky=tk.EW, padx=4, pady=3)

        # ==============================================================
        # NỬA PHẢI: CẤU HÌNH CỘT ĐÍCH (TARGET TEMPLATE)
        # ==============================================================
        tgt_frame = ttk.LabelFrame(
            main_frame,
            text=" 2. Cột Đích — Mẫu Ngân Hàng (TargetTemplate) ",
            padding=8,
        )
        tgt_frame.grid(row=1, column=1, sticky=tk.NSEW, padx=(6, 0))
        tgt_frame.columnconfigure(1, weight=1)

        # Chọn mẫu đích đã lưu + Tạo mới / Lưu / Xóa
        tgt_saved_bar = ttk.Frame(tgt_frame)
        tgt_saved_bar.grid(row=0, column=0, columnspan=3, sticky=tk.EW, pady=(0, 6))
        ttk.Label(tgt_saved_bar, text="Mapping đích đã lưu:").pack(side=tk.LEFT, padx=(0, 4))
        self.cmb_saved_target = ttk.Combobox(
            tgt_saved_bar,
            textvariable=self.saved_target_var,
            width=24,
            state="readonly",
        )
        self.cmb_saved_target.pack(side=tk.LEFT, padx=(0, 6))
        self.cmb_saved_target.bind(
            "<<ComboboxSelected>>",
            lambda _e: self._on_select_saved_target(),
        )
        self.btn_new_target = ttk.Button(
            tgt_saved_bar,
            text="Tạo mới",
            width=8,
            command=self.new_target_mapping,
        )
        self.btn_new_target.pack(side=tk.LEFT, padx=2)
        self.btn_save_target = ttk.Button(
            tgt_saved_bar,
            text="Lưu mẫu đích",
            width=12,
            command=self.save_target_mapping_action,
        )
        self.btn_save_target.pack(side=tk.LEFT, padx=2)
        self.btn_delete_target = ttk.Button(
            tgt_saved_bar,
            text="Xóa",
            width=6,
            command=self.delete_target_mapping_action,
        )
        self.btn_delete_target.pack(side=tk.LEFT, padx=2)

        # Chọn file Excel mẫu ngân hàng để đọc danh sách cột
        ttk.Label(tgt_frame, text="File NH mẫu (để quét cột):").grid(
            row=1, column=0, sticky=tk.W, pady=2
        )
        ttk.Entry(tgt_frame, textvariable=self.target_sample_file_var, state="readonly").grid(
            row=1, column=1, sticky=tk.EW, padx=4, pady=2
        )
        ttk.Button(
            tgt_frame,
            text="Chọn file...",
            command=self.browse_target_sample_excel,
        ).grid(row=1, column=2, sticky=tk.EW, pady=2)

        # Tên mẫu ngân hàng, Mẫu tên file, Sheet, Header rows, Data start row
        ttk.Label(tgt_frame, text="Tên mẫu ngân hàng (*):").grid(
            row=2, column=0, sticky=tk.W, pady=2
        )
        ttk.Entry(tgt_frame, textvariable=self.target_name_var).grid(
            row=2, column=1, columnspan=2, sticky=tk.EW, padx=4, pady=2
        )

        ttk.Label(tgt_frame, text="Mẫu nhận diện tên file:").grid(
            row=3, column=0, sticky=tk.W, pady=2
        )
        ttk.Entry(tgt_frame, textvariable=self.target_file_patterns_var).grid(
            row=3, column=1, columnspan=2, sticky=tk.EW, padx=4, pady=2
        )

        ttk.Label(tgt_frame, text="Sheet cố định / Mẫu tên sheet:").grid(
            row=4, column=0, sticky=tk.W, pady=2
        )
        tgt_sheet_box = ttk.Frame(tgt_frame)
        tgt_sheet_box.grid(row=4, column=1, columnspan=2, sticky=tk.EW, padx=4, pady=2)
        self.cmb_target_sheet = ttk.Combobox(
            tgt_sheet_box,
            textvariable=self.target_sheet_var,
            width=18,
        )
        self.cmb_target_sheet.pack(side=tk.LEFT, padx=(0, 6))
        self.cmb_target_sheet.bind("<<ComboboxSelected>>", lambda _e: self._rescan_target_columns())
        ttk.Label(tgt_sheet_box, text="Mẫu:").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(tgt_sheet_box, textvariable=self.target_sheet_patterns_var, width=22).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )

        ttk.Label(tgt_frame, text="Dòng tiêu đề / Dòng bắt đầu:").grid(
            row=5, column=0, sticky=tk.W, pady=2
        )
        tgt_row_box = ttk.Frame(tgt_frame)
        tgt_row_box.grid(row=5, column=1, columnspan=2, sticky=tk.EW, padx=4, pady=2)
        ttk.Entry(tgt_row_box, textvariable=self.target_header_rows_var, width=10).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Label(tgt_row_box, text="Bắt đầu dữ liệu:").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(tgt_row_box, textvariable=self.target_data_start_row_var, width=6).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(
            tgt_row_box,
            text="Quét lại cột",
            command=self._rescan_target_columns,
        ).pack(side=tk.RIGHT)

        # Bảng chọn cột cho 4 trường của Mẫu ngân hàng đích
        tgt_cols_box = ttk.LabelFrame(
            tgt_frame,
            text=" Danh sách Trường Đích → Chọn Cột & Tiêu đề mong đợi ",
            padding=6,
        )
        tgt_cols_box.grid(row=6, column=0, columnspan=3, sticky=tk.NSEW, pady=(6, 0))
        tgt_cols_box.columnconfigure(1, weight=1)
        tgt_cols_box.columnconfigure(2, weight=1)

        ttk.Label(tgt_cols_box, text="Trường dữ liệu", font=("Arial", 9, "bold")).grid(
            row=0, column=0, sticky=tk.W, padx=4
        )
        ttk.Label(tgt_cols_box, text="Cột Excel (Chữ cái)", font=("Arial", 9, "bold")).grid(
            row=0, column=1, sticky=tk.W, padx=4
        )
        ttk.Label(tgt_cols_box, text="Tiêu đề mong đợi (* với TK & Tiền)", font=("Arial", 9, "bold")).grid(
            row=0, column=2, sticky=tk.W, padx=4
        )

        for r_idx, (f_key, f_label) in enumerate(TARGET_FIELDS.items(), start=1):
            req_mark = "(*)" if f_key in ("account_no", "full_name", "amount") else "(Tùy chọn)"
            ttk.Label(tgt_cols_box, text=f"{f_label} {req_mark}:").grid(
                row=r_idx, column=0, sticky=tk.W, padx=4, pady=3
            )
            cmb = ttk.Combobox(
                tgt_cols_box,
                textvariable=self.tgt_col_vars[f_key],
                width=22,
            )
            cmb.grid(row=r_idx, column=1, sticky=tk.EW, padx=4, pady=3)
            cmb.bind(
                "<<ComboboxSelected>>",
                self._make_target_col_selected_handler(f_key),
            )
            self._tgt_col_combos[f_key] = cmb

            ent_hdr = ttk.Entry(tgt_cols_box, textvariable=self.tgt_hdr_vars[f_key])
            ent_hdr.grid(row=r_idx, column=2, sticky=tk.EW, padx=4, pady=3)

        # ==============================================================
        # PHẦN DƯỚI: BẢNG TỔNG HỢP ÁNH XẠ SOURCE → TARGET & CẢNH BÁO VALIDATE
        # ==============================================================
        link_frame = ttk.LabelFrame(
            main_frame,
            text=" 3. Sơ đồ Ánh xạ Source → Target & Cảnh báo Kiểm tra ",
            padding=8,
        )
        link_frame.grid(row=2, column=0, columnspan=2, sticky=tk.NSEW, pady=(8, 0))
        link_frame.columnconfigure(0, weight=1)

        cols = ("field", "source_col", "arrow", "target_col", "status")
        self.tree_links = ttk.Treeview(
            link_frame,
            columns=cols,
            show="headings",
            height=5,
        )
        self.tree_links.heading("field", text="Trường nghiệp vụ")
        self.tree_links.heading("source_col", text="Cột Nguồn (Bảng lương)")
        self.tree_links.heading("arrow", text="→")
        self.tree_links.heading("target_col", text="Cột Đích (Mẫu ngân hàng)")
        self.tree_links.heading("status", text="Trạng thái")

        self.tree_links.column("field", width=210, anchor=tk.W)
        self.tree_links.column("source_col", width=260, anchor=tk.W)
        self.tree_links.column("arrow", width=40, anchor=tk.CENTER)
        self.tree_links.column("target_col", width=260, anchor=tk.W)
        self.tree_links.column("status", width=220, anchor=tk.W)
        self.tree_links.grid(row=0, column=0, sticky=tk.EW)

        self.tree_links.tag_configure("ok", background="#E2EFDA", foreground="#1E4620")
        self.tree_links.tag_configure("warn", background="#FCE4D6", foreground="#9C0006")

        # Nhãn cảnh báo lỗi / trùng cột / thiếu mapping
        self.lbl_validation = ttk.Label(
            link_frame,
            textvariable=self.validation_message_var,
            font=("Arial", 10, "bold"),
            foreground="#1F4E78",
            wraplength=1080,
        )
        self.lbl_validation.grid(row=1, column=0, sticky=tk.EW, pady=(6, 0))

        self._refresh_combobox_choices()

    # =========================================================================
    # 3. Quét danh sách cột từ file Excel nguồn và đích thực tế
    # =========================================================================

    def browse_source_sample_excel(
        self,
        *,
        dialog_fn: Optional[Callable[..., str]] = None,
    ) -> Optional[str]:
        ask_fn = dialog_fn or filedialog.askopenfilename
        chosen = ask_fn(
            parent=self,
            title="Chọn file bảng lương mẫu để đọc danh sách cột",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if chosen:
            self.inspect_source_excel(chosen)
            return str(chosen)
        return None

    def browse_target_sample_excel(
        self,
        *,
        dialog_fn: Optional[Callable[..., str]] = None,
    ) -> Optional[str]:
        ask_fn = dialog_fn or filedialog.askopenfilename
        chosen = ask_fn(
            parent=self,
            title="Chọn file mẫu ngân hàng để đọc danh sách cột",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if chosen:
            self.inspect_target_excel(chosen)
            return str(chosen)
        return None

    def inspect_source_excel(
        self,
        file_path: PathLike,
        *,
        sheet_name: Optional[str] = None,
        header_rows: Optional[Sequence[int]] = None,
    ) -> list[ColumnOption]:
        """Đọc danh sách sheet và tiêu đề cột từ file bảng lương mẫu."""
        p = Path(file_path)
        self.source_sample_file_var.set(str(p))
        sheets = [s.name for s in list_sheets(p) if s.visible]
        self.cmb_source_sheet["values"] = sheets

        chosen_sheet = sheet_name or self.source_sheet_var.get().strip()
        if not chosen_sheet or chosen_sheet not in sheets:
            chosen_sheet = sheets[0] if sheets else ""
            if chosen_sheet:
                self.source_sheet_var.set(chosen_sheet)

        hdr_rows = (
            tuple(header_rows)
            if header_rows is not None
            else (parse_header_rows_text(self.source_header_rows_var.get()) or (3, 4))
        )
        if chosen_sheet:
            self._source_col_options = extract_sheet_column_options(p, chosen_sheet, hdr_rows)
            self._source_header_by_letter = {
                opt.letter: opt.header for opt in self._source_col_options if opt.header
            }
        self._refresh_combobox_choices()
        self.refresh_mapping_preview()
        return list(self._source_col_options)

    def inspect_target_excel(
        self,
        file_path: PathLike,
        *,
        sheet_name: Optional[str] = None,
        header_rows: Optional[Sequence[int]] = None,
    ) -> list[ColumnOption]:
        """Đọc danh sách sheet và tiêu đề cột từ file ngân hàng mẫu."""
        p = Path(file_path)
        self.target_sample_file_var.set(str(p))
        sheets = [s.name for s in list_sheets(p) if s.visible]
        self.cmb_target_sheet["values"] = sheets

        chosen_sheet = sheet_name or self.target_sheet_var.get().strip()
        if not chosen_sheet or chosen_sheet not in sheets:
            chosen_sheet = sheets[0] if sheets else ""
            if chosen_sheet:
                self.target_sheet_var.set(chosen_sheet)

        hdr_rows = (
            tuple(header_rows)
            if header_rows is not None
            else (parse_header_rows_text(self.target_header_rows_var.get()) or (2,))
        )
        if chosen_sheet:
            self._target_col_options = extract_sheet_column_options(p, chosen_sheet, hdr_rows)
            self._target_header_by_letter = {
                opt.letter: opt.header for opt in self._target_col_options if opt.header
            }
        self._refresh_combobox_choices()
        self.refresh_mapping_preview()
        return list(self._target_col_options)

    @property
    def source_column_options(self) -> list[ColumnOption]:
        return list(self._source_col_options)

    @property
    def target_column_options(self) -> list[ColumnOption]:
        return list(self._target_col_options)

    # =========================================================================
    # 4. Tải lại / Nạp / Tạo mới / Lưu / Xóa Mapping (templates/*.json)
    # =========================================================================

    def reload_saved_mappings(self) -> tuple[list[str], list[str]]:
        """Tải lại danh sách các mapping nguồn & đích đã lưu trong `templates/*.json`."""
        src_items = self.store.list_sources()
        tgt_items = self.store.list_targets()

        src_names = [s.name for s in src_items]
        tgt_names = [t.name for t in tgt_items]

        self.cmb_saved_source["values"] = src_names
        self.cmb_saved_target["values"] = tgt_names

        if self.saved_source_var.get() not in src_names and src_names:
            self.saved_source_var.set(src_names[0])
        if self.saved_target_var.get() not in tgt_names and tgt_names:
            self.saved_target_var.set(tgt_names[0])

        return src_names, tgt_names

    def load_source_mapping(self, name: str) -> SourceProfile:
        """Tải 1 `SourceProfile` đã lưu lên form."""
        prof = self.store.get_source(name)
        self._loaded_source_name = prof.name
        self.saved_source_var.set(prof.name)
        self.source_name_var.set(prof.name)
        self.source_sheet_var.set(prof.sheet_name or "")
        self.source_sheet_patterns_var.set(", ".join(prof.sheet_patterns))
        self.source_header_rows_var.set(", ".join(str(r) for r in prof.header_rows))
        self.source_data_start_row_var.set(str(prof.data_start_row))
        self.source_data_end_row_var.set(str(prof.data_end_row) if prof.data_end_row else "")

        for f_key, ref in prof.columns().items():
            self.src_col_vars[f_key].set(ref.letter)
            self.src_hdr_vars[f_key].set(ref.expected_header or "")

        self.refresh_mapping_preview()
        return prof

    def load_target_mapping(self, name: str) -> TargetTemplate:
        """Tải 1 `TargetTemplate` đã lưu lên form."""
        tmpl = self.store.get_target(name)
        self._loaded_target_name = tmpl.name
        self.saved_target_var.set(tmpl.name)
        self.target_name_var.set(tmpl.name)
        self.target_desc_var.set(tmpl.description)
        self.target_file_patterns_var.set(", ".join(tmpl.file_patterns))
        self.target_sheet_var.set(tmpl.sheet_name or "")
        self.target_sheet_patterns_var.set(", ".join(tmpl.sheet_patterns))
        self.target_header_rows_var.set(", ".join(str(r) for r in tmpl.header_rows))
        self.target_data_start_row_var.set(str(tmpl.data_start_row))
        self.target_data_end_row_var.set(str(tmpl.data_end_row) if tmpl.data_end_row else "")

        self.rule_no_diacritics_var.set(tmpl.rules.no_diacritics)
        self.rule_no_extra_spaces_var.set(tmpl.rules.no_extra_spaces)
        self.rule_max_len_var.set(
            str(tmpl.rules.max_text_length) if tmpl.rules.max_text_length else ""
        )
        self.rule_forbidden_chars_var.set(tmpl.rules.forbidden_chars)

        cols = tmpl.columns()
        for f_key in TARGET_FIELDS:
            ref = cols.get(f_key)
            self.tgt_col_vars[f_key].set(ref.letter if ref else "")
            self.tgt_hdr_vars[f_key].set((ref.expected_header or "") if ref else "")

        self.refresh_mapping_preview()
        return tmpl

    def new_source_mapping(self) -> None:
        """Khởi tạo form trống để tạo mới cấu hình nguồn."""
        self._loaded_source_name = None
        self.saved_source_var.set("")
        self.source_name_var.set("Bảng lương mới")
        self.source_sheet_var.set("")
        self.source_sheet_patterns_var.set("HLCK*, *LUONG*")
        self.source_header_rows_var.set("3, 4")
        self.source_data_start_row_var.set("5")
        self.source_data_end_row_var.set("")
        for f_key in SOURCE_FIELDS:
            self.src_col_vars[f_key].set("")
            self.src_hdr_vars[f_key].set("")
        self.refresh_mapping_preview()

    def new_target_mapping(self) -> None:
        """Khởi tạo form trống để tạo mới mẫu ngân hàng đích."""
        self._loaded_target_name = None
        self.saved_target_var.set("")
        self.target_name_var.set("Mẫu ngân hàng mới")
        self.target_desc_var.set("")
        self.target_file_patterns_var.set("*MauChiHoLuong*")
        self.target_sheet_var.set("")
        self.target_sheet_patterns_var.set("")
        self.target_header_rows_var.set("2")
        self.target_data_start_row_var.set("3")
        self.target_data_end_row_var.set("")
        for f_key in TARGET_FIELDS:
            self.tgt_col_vars[f_key].set("")
            self.tgt_hdr_vars[f_key].set("")
        self.refresh_mapping_preview()

    def set_source_field_mapping(
        self,
        field_key: str,
        column_letter_or_choice: str,
        expected_header: Optional[str] = None,
    ) -> None:
        """Thiết lập cột cho 1 trường nguồn (`card_no`, `full_name`, `account_no`, `amount`)."""
        if field_key not in self.src_col_vars:
            raise KeyError(f"Trường nguồn không hợp lệ: {field_key!r}")
        col = parse_column_choice(column_letter_or_choice)
        self.src_col_vars[field_key].set(col)
        if expected_header is not None:
            self.src_hdr_vars[field_key].set(expected_header)
        elif col in self._source_header_by_letter and not self.src_hdr_vars[field_key].get():
            self.src_hdr_vars[field_key].set(self._source_header_by_letter[col])
        self.refresh_mapping_preview()

    def set_target_field_mapping(
        self,
        field_key: str,
        column_letter_or_choice: str,
        expected_header: Optional[str] = None,
    ) -> None:
        """Thiết lập cột cho 1 trường đích (`account_no`, `full_name`, `amount`, `bank_code`)."""
        if field_key not in self.tgt_col_vars:
            raise KeyError(f"Trường đích không hợp lệ: {field_key!r}")
        col = parse_column_choice(column_letter_or_choice)
        self.tgt_col_vars[field_key].set(col)
        if expected_header is not None:
            self.tgt_hdr_vars[field_key].set(expected_header)
        elif col in self._target_header_by_letter:
            self.tgt_hdr_vars[field_key].set(self._target_header_by_letter[col])
        self.refresh_mapping_preview()

    def save_source_mapping(self, *, overwrite: bool = True) -> Path:
        """Validate và lưu `SourceProfile` hiện tại vào `templates/source_*.json`."""
        val = self.validate_source_mapping()
        if not val.is_valid:
            self._update_validation_banner(val)
            raise ConfigError("Cấu hình nguồn chưa hợp lệ:\n- " + "\n- ".join(val.errors))

        profile = self.build_source_profile()
        saved_path = self.store.save_source(
            profile,
            overwrite=overwrite,
            previous_name=self._loaded_source_name,
        )
        self._loaded_source_name = profile.name
        self.reload_saved_mappings()
        self.saved_source_var.set(profile.name)
        self.refresh_mapping_preview()
        return saved_path

    def save_target_mapping(self, *, overwrite: bool = True) -> Path:
        """Validate và lưu `TargetTemplate` hiện tại vào `templates/target_*.json`."""
        val = self.validate_target_mapping()
        if not val.is_valid:
            self._update_validation_banner(val)
            raise ConfigError("Cấu hình mẫu đích chưa hợp lệ:\n- " + "\n- ".join(val.errors))

        template = self.build_target_template()
        saved_path = self.store.save_target(
            template,
            overwrite=overwrite,
            previous_name=self._loaded_target_name,
        )
        self._loaded_target_name = template.name
        self.reload_saved_mappings()
        self.saved_target_var.set(template.name)
        self.refresh_mapping_preview()
        return saved_path

    def delete_source_mapping(self, name: Optional[str] = None) -> Path:
        """Xóa 1 cấu hình nguồn trong `templates/` (có bản sao lưu tự động)."""
        target_name = (name or self.saved_source_var.get() or self.source_name_var.get()).strip()
        if not target_name:
            raise TemplateStoreError("Chưa chọn cấu hình nguồn cần xóa.")
        backup_path = self.store.delete("source", target_name)
        self._loaded_source_name = None
        src_names, _ = self.reload_saved_mappings()
        if src_names:
            self.load_source_mapping(src_names[0])
        else:
            self.new_source_mapping()
        return backup_path

    def delete_target_mapping(self, name: Optional[str] = None) -> Path:
        """Xóa 1 mẫu ngân hàng đích trong `templates/` (có bản sao lưu tự động)."""
        target_name = (name or self.saved_target_var.get() or self.target_name_var.get()).strip()
        if not target_name:
            raise TemplateStoreError("Chưa chọn mẫu ngân hàng cần xóa.")
        backup_path = self.store.delete("target", target_name)
        self._loaded_target_name = None
        _, tgt_names = self.reload_saved_mappings()
        if tgt_names:
            self.load_target_mapping(tgt_names[0])
        else:
            self.new_target_mapping()
        return backup_path

    # =========================================================================
    # 5. Validate Mapping bắt buộc & Phát hiện trùng / thiếu cột
    # =========================================================================

    def validate_source_mapping(self) -> MappingValidationResult:
        """Kiểm tra cấu hình cột nguồn (bắt buộc đủ 4 trường, không trùng cột, dòng dữ liệu hợp lệ)."""
        res = MappingValidationResult()
        name = self.source_name_var.get().strip()
        if not name:
            res.errors.append("Tên cấu hình nguồn không được để trống.")

        sheet_exact = self.source_sheet_var.get().strip()
        sheet_pats = _split_patterns(self.source_sheet_patterns_var.get())
        if not sheet_exact and not sheet_pats:
            res.errors.append("Nguồn: Phải khai báo Tên sheet cố định hoặc Mẫu tên sheet.")

        try:
            hdr_rows = parse_header_rows_text(self.source_header_rows_var.get())
            if not hdr_rows:
                res.errors.append("Nguồn: Phải khai báo ít nhất 1 dòng tiêu đề.")
        except ValueError as exc:
            hdr_rows = ()
            res.errors.append(f"Nguồn: {exc}")

        start_s = self.source_data_start_row_var.get().strip()
        if not start_s.isdigit() or int(start_s) < 1:
            res.errors.append("Nguồn: Dòng bắt đầu dữ liệu phải là số nguyên dương.")
        elif hdr_rows and int(start_s) <= max(hdr_rows):
            res.errors.append("Nguồn: Dòng bắt đầu dữ liệu phải lớn hơn các dòng tiêu đề.")

        used_cols: dict[str, str] = {}
        for f_key, f_label in SOURCE_FIELDS.items():
            col = parse_column_choice(self.src_col_vars[f_key].get())
            if not col:
                res.errors.append(f"Nguồn: Thiếu mapping bắt buộc cho trường '{f_label}'.")
                continue
            try:
                column_index(col)
            except ConfigError as exc:
                res.errors.append(f"Nguồn ({f_label}): {exc.user_message}")
                continue
            if col in used_cols:
                other_label = SOURCE_FIELDS[used_cols[col]]
                res.errors.append(
                    f"Nguồn: Cột {col} bị trùng giữa '{other_label}' và '{f_label}'."
                )
            else:
                used_cols[col] = f_key

            if not self.src_hdr_vars[f_key].get().strip():
                res.warnings.append(
                    f"Nguồn: Trường '{f_label}' (cột {col}) chưa khai báo tiêu đề mong đợi."
                )

        return res

    def validate_target_mapping(self) -> MappingValidationResult:
        """Kiểm tra cấu hình mẫu ngân hàng đích (bắt buộc đủ Số TK, Tên NV, Số tiền; không trùng cột)."""
        res = MappingValidationResult()
        name = self.target_name_var.get().strip()
        if not name:
            res.errors.append("Tên mẫu ngân hàng đích không được để trống.")

        sheet_exact = self.target_sheet_var.get().strip()
        sheet_pats = _split_patterns(self.target_sheet_patterns_var.get())
        if not sheet_exact and not sheet_pats:
            res.errors.append("Đích: Phải khai báo Tên sheet cố định hoặc Mẫu tên sheet.")

        try:
            hdr_rows = parse_header_rows_text(self.target_header_rows_var.get())
            if not hdr_rows:
                res.errors.append("Đích: Phải khai báo ít nhất 1 dòng tiêu đề.")
        except ValueError as exc:
            hdr_rows = ()
            res.errors.append(f"Đích: {exc}")

        start_s = self.target_data_start_row_var.get().strip()
        if not start_s.isdigit() or int(start_s) < 1:
            res.errors.append("Đích: Dòng bắt đầu dữ liệu phải là số nguyên dương.")
        elif hdr_rows and int(start_s) <= max(hdr_rows):
            res.errors.append("Đích: Dòng bắt đầu dữ liệu phải lớn hơn các dòng tiêu đề.")

        required_target_keys = ("account_no", "full_name", "amount")
        used_cols: dict[str, str] = {}
        for f_key, f_label in TARGET_FIELDS.items():
            col = parse_column_choice(self.tgt_col_vars[f_key].get())
            if not col:
                if f_key in required_target_keys:
                    res.errors.append(f"Đích: Thiếu mapping bắt buộc cho trường '{f_label}'.")
                continue
            try:
                column_index(col)
            except ConfigError as exc:
                res.errors.append(f"Đích ({f_label}): {exc.user_message}")
                continue
            if col in used_cols:
                other_label = TARGET_FIELDS[used_cols[col]]
                res.errors.append(
                    f"Đích: Cột {col} bị trùng giữa '{other_label}' và '{f_label}'."
                )
            else:
                used_cols[col] = f_key

            hdr = self.tgt_hdr_vars[f_key].get().strip()
            if f_key in ("account_no", "amount") and not hdr:
                res.errors.append(
                    f"Đích: Cột '{f_label}' (cột {col}) bắt buộc phải khai báo tiêu đề mong đợi."
                )

        return res

    def validate_all(self) -> MappingValidationResult:
        """Kiểm tra đồng thời cả cấu hình nguồn và cấu hình đích."""
        s_val = self.validate_source_mapping()
        t_val = self.validate_target_mapping()
        combined = MappingValidationResult(
            errors=[*s_val.errors, *t_val.errors],
            warnings=[*s_val.warnings, *t_val.warnings],
        )
        self._update_validation_banner(combined)
        return combined

    def build_source_profile(self) -> SourceProfile:
        """Dựng đối tượng `SourceProfile` từ dữ liệu đang nhập trên form."""
        hdr_rows = parse_header_rows_text(self.source_header_rows_var.get())
        start_row = int(self.source_data_start_row_var.get().strip())
        end_raw = self.source_data_end_row_var.get().strip()
        end_row = int(end_raw) if end_raw else None

        def make_ref(key: str) -> ColumnRef:
            letter = parse_column_choice(self.src_col_vars[key].get())
            hdr = self.src_hdr_vars[key].get().strip() or None
            return ColumnRef(letter=letter, expected_header=hdr, label=SOURCE_FIELDS[key])

        profile = SourceProfile(
            name=self.source_name_var.get().strip(),
            sheet_name=self.source_sheet_var.get().strip() or None,
            sheet_patterns=_split_patterns(self.source_sheet_patterns_var.get()),
            header_rows=hdr_rows,
            data_start_row=start_row,
            data_end_row=end_row,
            card_no=make_ref("card_no"),
            full_name=make_ref("full_name"),
            account_no=make_ref("account_no"),
            amount=make_ref("amount"),
        )
        profile.ensure_valid()
        return profile

    def build_target_template(self) -> TargetTemplate:
        """Dựng đối tượng `TargetTemplate` từ dữ liệu đang nhập trên form."""
        hdr_rows = parse_header_rows_text(self.target_header_rows_var.get())
        start_row = int(self.target_data_start_row_var.get().strip())
        end_raw = self.target_data_end_row_var.get().strip()
        end_row = int(end_raw) if end_raw else None

        def make_ref(key: str) -> ColumnRef:
            letter = parse_column_choice(self.tgt_col_vars[key].get())
            hdr = self.tgt_hdr_vars[key].get().strip() or None
            return ColumnRef(letter=letter, expected_header=hdr, label=TARGET_FIELDS[key])

        bank_col = parse_column_choice(self.tgt_col_vars["bank_code"].get())
        bank_ref = make_ref("bank_code") if bank_col else None

        max_len_raw = self.rule_max_len_var.get().strip()
        rules = TargetRules(
            no_diacritics=bool(self.rule_no_diacritics_var.get()),
            no_extra_spaces=bool(self.rule_no_extra_spaces_var.get()),
            max_text_length=int(max_len_raw) if max_len_raw else None,
            forbidden_chars=self.rule_forbidden_chars_var.get(),
        )

        template = TargetTemplate(
            name=self.target_name_var.get().strip(),
            description=self.target_desc_var.get().strip(),
            file_patterns=_split_patterns(self.target_file_patterns_var.get()),
            sheet_name=self.target_sheet_var.get().strip() or None,
            sheet_patterns=_split_patterns(self.target_sheet_patterns_var.get()),
            header_rows=hdr_rows,
            data_start_row=start_row,
            data_end_row=end_row,
            account_no=make_ref("account_no"),
            full_name=make_ref("full_name"),
            amount=make_ref("amount"),
            bank_code=bank_ref,
            rules=rules,
        )
        template.ensure_valid()
        return template

    def get_source_target_links(self) -> list[SourceTargetLink]:
        """Trả về danh sách ánh xạ trực quan `Source → Target` hiện tại."""
        src_used: list[str] = [
            parse_column_choice(self.src_col_vars[k].get())
            for k in SOURCE_FIELDS
            if parse_column_choice(self.src_col_vars[k].get())
        ]
        tgt_used: list[str] = [
            parse_column_choice(self.tgt_col_vars[k].get())
            for k in TARGET_FIELDS
            if parse_column_choice(self.tgt_col_vars[k].get())
        ]

        specs = [
            ("card_no", "Số thẻ (Khóa lọc dòng)", True, False),
            ("account_no", "Số tài khoản (Khóa đối chiếu)", True, True),
            ("full_name", "Họ và tên (Kiểm tra chéo)", True, True),
            ("amount", "Số tiền chuyển khoản (Ô ghi)", True, True),
            ("bank_code", "Mã ngân hàng (Tùy chọn)", False, False),
        ]

        links: list[SourceTargetLink] = []
        for f_key, label, req_src, req_tgt in specs:
            s_col = parse_column_choice(self.src_col_vars[f_key].get()) if f_key in self.src_col_vars else ""
            s_hdr = self.src_hdr_vars[f_key].get().strip() if f_key in self.src_hdr_vars else ""
            t_col = parse_column_choice(self.tgt_col_vars[f_key].get()) if f_key in self.tgt_col_vars else ""
            t_hdr = self.tgt_hdr_vars[f_key].get().strip() if f_key in self.tgt_hdr_vars else ""

            status = "OK"
            if (req_src and not s_col) or (req_tgt and not t_col):
                status = "MISSING"
            elif (s_col and src_used.count(s_col) > 1) or (t_col and tgt_used.count(t_col) > 1):
                status = "DUPLICATE"
            elif f_key in ("account_no", "amount") and req_tgt and not t_hdr:
                status = "MISSING"

            links.append(
                SourceTargetLink(
                    field_key=f_key,
                    field_label=label,
                    source_column=s_col,
                    source_header=s_hdr,
                    target_column=t_col,
                    target_header=t_hdr,
                    is_required_source=req_src,
                    is_required_target=req_tgt,
                    status=status,
                )
            )
        return links

    def refresh_mapping_preview(self) -> list[SourceTargetLink]:
        """Cập nhật bảng sơ đồ ánh xạ `Source → Target` và thanh cảnh báo hợp lệ."""
        for iid in self.tree_links.get_children():
            self.tree_links.delete(iid)

        links = self.get_source_target_links()
        for link in links:
            src_txt = (
                f"Cột {link.source_column} ({link.source_header or 'không tiêu đề'})"
                if link.source_column
                else ("(Thiếu cột nguồn!)" if link.is_required_source else "(Không dùng ở nguồn)")
            )
            tgt_txt = (
                f"Cột {link.target_column} ({link.target_header or 'chưa nhập tiêu đề'})"
                if link.target_column
                else ("(Thiếu cột đích!)" if link.is_required_target else "(Chỉ dùng ở nguồn / Tùy chọn)")
            )
            status_txt = {
                "OK": "🟢 Hợp lệ",
                "MISSING": "🔴 Thiếu thông tin bắt buộc",
                "DUPLICATE": "🔴 Trùng cột",
            }.get(link.status, link.status)
            tag = "ok" if link.status == "OK" else "warn"
            self.tree_links.insert(
                "",
                tk.END,
                values=(link.field_label, src_txt, "→", tgt_txt, status_txt),
                tags=(tag,),
            )

        self.validate_all()
        return links

    # =========================================================================
    # 6. Sự kiện nút bấm trên giao diện
    # =========================================================================

    def save_source_mapping_action(self) -> Optional[Path]:
        try:
            path = self.save_source_mapping(overwrite=True)
            msg = f"Đã lưu cấu hình nguồn vào: {path.name}"
            self.validation_message_var.set(f"✅ {msg}")
            if self.interactive:
                messagebox.showinfo("Lưu thành công", msg, parent=self)
            return path
        except (ConfigError, TemplateStoreError) as exc:
            if self.interactive:
                messagebox.showerror("Không thể lưu cấu hình nguồn", exc.user_message, parent=self)
            return None

    def save_target_mapping_action(self) -> Optional[Path]:
        try:
            path = self.save_target_mapping(overwrite=True)
            msg = f"Đã lưu mẫu ngân hàng vào: {path.name}"
            self.validation_message_var.set(f"✅ {msg}")
            if self.interactive:
                messagebox.showinfo("Lưu thành công", msg, parent=self)
            return path
        except (ConfigError, TemplateStoreError) as exc:
            if self.interactive:
                messagebox.showerror("Không thể lưu mẫu ngân hàng", exc.user_message, parent=self)
            return None

    def delete_source_mapping_action(self) -> Optional[Path]:
        try:
            backup = self.delete_source_mapping()
            msg = f"Đã xóa cấu hình nguồn (bản sao lưu: {backup.name})."
            self.validation_message_var.set(msg)
            return backup
        except TemplateStoreError as exc:
            if self.interactive:
                messagebox.showerror("Không thể xóa", exc.user_message, parent=self)
            return None

    def delete_target_mapping_action(self) -> Optional[Path]:
        try:
            backup = self.delete_target_mapping()
            msg = f"Đã xóa mẫu ngân hàng (bản sao lưu: {backup.name})."
            self.validation_message_var.set(msg)
            return backup
        except TemplateStoreError as exc:
            if self.interactive:
                messagebox.showerror("Không thể xóa", exc.user_message, parent=self)
            return None

    def run_interactive_validation(self) -> MappingValidationResult:
        val = self.validate_all()
        if self.interactive:
            if val.is_valid:
                messagebox.showinfo(
                    "Kiểm tra Mapping",
                    "Cấu hình Mapping Nguồn và Đích hoàn toàn hợp lệ!",
                    parent=self,
                )
            else:
                messagebox.showwarning(
                    "Cảnh báo Mapping chưa hợp lệ",
                    "\n- ".join(["Phát hiện lỗi cấu hình:", *val.errors]),
                    parent=self,
                )
        return val

    # =========================================================================
    # 7. Các hàm phụ trợ nội bộ
    # =========================================================================

    def _on_select_saved_source(self) -> None:
        name = self.saved_source_var.get().strip()
        if name:
            self.load_source_mapping(name)

    def _on_select_saved_target(self) -> None:
        name = self.saved_target_var.get().strip()
        if name:
            self.load_target_mapping(name)

    def _rescan_source_columns(self) -> None:
        sample = self.source_sample_file_var.get().strip()
        if sample and Path(sample).exists():
            try:
                self.inspect_source_excel(sample)
            except (ExcelReadError, ValueError):
                pass

    def _rescan_target_columns(self) -> None:
        sample = self.target_sample_file_var.get().strip()
        if sample and Path(sample).exists():
            try:
                self.inspect_target_excel(sample)
            except (ExcelReadError, ValueError):
                pass

    def _refresh_combobox_choices(self) -> None:
        src_choices = [opt.display for opt in self._source_col_options]
        for cmb in self._src_col_combos.values():
            cmb["values"] = src_choices

        tgt_choices = ["", *[opt.display for opt in self._target_col_options]]
        for cmb in self._tgt_col_combos.values():
            cmb["values"] = tgt_choices

    def _make_source_col_selected_handler(self, field_key: str) -> Callable[[Any], None]:
        def _handler(_event: Any = None) -> None:
            raw = self.src_col_vars[field_key].get()
            col = parse_column_choice(raw)
            self.src_col_vars[field_key].set(col)
            if col in self._source_header_by_letter:
                self.src_hdr_vars[field_key].set(self._source_header_by_letter[col])
            self.refresh_mapping_preview()

        return _handler

    def _make_target_col_selected_handler(self, field_key: str) -> Callable[[Any], None]:
        def _handler(_event: Any = None) -> None:
            raw = self.tgt_col_vars[field_key].get()
            col = parse_column_choice(raw)
            self.tgt_col_vars[field_key].set(col)
            if col in self._target_header_by_letter:
                self.tgt_hdr_vars[field_key].set(self._target_header_by_letter[col])
            self.refresh_mapping_preview()

        return _handler

    def _update_validation_banner(self, val: MappingValidationResult) -> None:
        if not val.is_valid:
            self.validation_message_var.set("⚠️ CẢNH BÁO: " + " | ".join(val.errors))
            self.lbl_validation.configure(foreground="#9C0006")
        elif val.warnings:
            self.validation_message_var.set("ℹ️ Lưu ý: " + " | ".join(val.warnings))
            self.lbl_validation.configure(foreground="#B45F06")
        else:
            self.validation_message_var.set("🟢 Cấu hình Mapping Nguồn → Đích hợp lệ, sẵn sàng lưu.")
            self.lbl_validation.configure(foreground="#1E4620")


def open_mapping_dialog(
    parent: Optional[tk.Misc] = None,
    *,
    template_store: Optional[TemplateStore] = None,
    templates_dir: PathLike = "templates",
    source_excel_path: Optional[PathLike] = None,
    target_excel_path: Optional[PathLike] = None,
    interactive: bool = True,
) -> MappingDialog:
    """Hàm mở hộp thoại cấu hình Mapping (được gọi từ `MainWindow.open_mapping_dialog`)."""
    return MappingDialog(
        parent,
        template_store=template_store,
        templates_dir=templates_dir,
        source_excel_path=source_excel_path,
        target_excel_path=target_excel_path,
        interactive=interactive,
    )


def _split_patterns(raw: str) -> tuple[str, ...]:
    parts = [p.strip() for p in re.split(r"[,;]+", raw) if p.strip()]
    return tuple(parts)
