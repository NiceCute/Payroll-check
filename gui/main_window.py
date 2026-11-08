"""
main_window.py - Màn hình chính 1 trang cho người dùng kế toán (GUI Module 1).

Thiết kế theo mục 5 tài liệu V1 (1 màn hình, 3 bước):
    Bước ① Chọn file & Template (File bảng lương nguồn + 1..n File ngân hàng đích + chọn Template)
    Bước ② Phân tích & Preview (Nút Analyze, Nút Preview, bảng chỉ số thống kê dòng / exception / tổng tiền)
    Bước ③ Xuất kết quả (Nút Execute: Backup -> Ghi file mới -> Hậu kiểm -> Xuất báo cáo .xlsx)

Quy tắc phụ thuộc:
    `gui/main_window -> workflow (+ template_store, reporter, models)`.
    Không chứa logic đối chiếu hay ghi file ngân hàng trực tiếp.
"""

from __future__ import annotations

import os
import tkinter as tk
from dataclasses import asdict, dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Mapping, Optional, Sequence, Union

from core.excel_reader import PathLike, list_sheets
from core.models import (
    MatchStatus,
    ReportResult,
    ResolutionAction,
    RuntimeConfig,
    WorkflowResult,
    WorkflowStage,
)
from core.reconciler import ResolutionValue
from core.reporter import Reporter
from core.template_store import TemplateStore
from core.workflow import Workflow

__all__ = [
    "AUTO_TEMPLATE_OPTION",
    "MainWindow",
    "MainWindowStats",
    "format_vnd",
    "launch_main_window",
]

AUTO_TEMPLATE_OPTION: str = "Tự động nhận diện (Auto)"


# =============================================================================
# 1. Cấu trúc dữ liệu thống kê hiển thị trên Main Window
# =============================================================================

@dataclass(frozen=True)
class MainWindowStats:
    """Các chỉ số thống kê hiển thị trên Main Window cho kế toán."""

    # --- Số lượng record ---
    total_source_records: int = 0
    total_matched_records: int = 0
    total_written_records: int = 0
    will_update_records: int = 0
    unchanged_records: int = 0
    skipped_records: int = 0
    pending_records: int = 0
    error_records: int = 0

    # --- Số lượng exception ---
    total_exceptions: int = 0
    resolved_exceptions: int = 0
    unresolved_exceptions: int = 0
    blocking_exceptions: int = 0

    # --- Tổng tiền (VND) ---
    total_source_amount: int = 0
    total_source_transferable_amount: int = 0
    total_old_amount: int = 0
    total_planned_write_amount: int = 0
    delta_amount: int = 0

    # --- Trạng thái ---
    stage: str = WorkflowStage.IDLE.value
    is_safe: bool = False
    is_verified: bool = False

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def format_vnd(amount: Optional[int]) -> str:
    """Định dạng số nguyên VND có dấu phẩy phân cách hàng nghìn (ví dụ: '445,138,601 đ')."""
    if amount is None:
        return "0 đ"
    return f"{amount:,} đ"


# =============================================================================
# 2. Lớp giao diện chính MainWindow
# =============================================================================

class MainWindow(ttk.Frame):
    """Màn hình chính điều khiển quy trình Chi hộ lương V1."""

    def __init__(
        self,
        master: Optional[tk.Misc] = None,
        *,
        config: Optional[RuntimeConfig] = None,
        workflow: Optional[Workflow] = None,
        reporter: Optional[Reporter] = None,
        interactive: bool = True,
    ) -> None:
        if master is None:
            root = tk.Tk()
            root.title("Công cụ Tự động Chi hộ Lương - Hồng Lâm (V1)")
            root.geometry("1120x760")
            root.minsize(960, 640)
            master = root
            self._owns_root = True
        else:
            self._owns_root = False

        super().__init__(master, padding=12)
        self.pack(fill=tk.BOTH, expand=True)

        self.config_data: RuntimeConfig = config or RuntimeConfig()
        self.workflow: Workflow = workflow or Workflow(self.config_data)
        self.reporter: Reporter = reporter or Reporter()
        self.interactive: bool = interactive

        self._template_store = TemplateStore(self.config_data.templates_dir)
        if self.config_data.auto_create_default_templates:
            self._template_store.ensure_defaults()

        # Trạng thái dữ liệu hiện tại trên giao diện
        self._source_file: str = ""
        self._target_files: list[str] = []
        self._stats: MainWindowStats = MainWindowStats()
        self._last_report: Optional[ReportResult] = None
        self._preview_visible: bool = False
        self._embedded_preview_widget: Optional[Any] = None

        # Biến liên kết Tkinter
        self.source_file_var = tk.StringVar(master=self, value="")
        self.source_sheet_var = tk.StringVar(master=self, value="")
        self.source_template_var = tk.StringVar(master=self, value=AUTO_TEMPLATE_OPTION)
        self.target_template_var = tk.StringVar(master=self, value=AUTO_TEMPLATE_OPTION)
        self.output_dir_var = tk.StringVar(master=self, value=self.config_data.output_dir)
        self.status_var = tk.StringVar(
            master=self,
            value="Sẵn sàng. Vui lòng chọn file bảng lương nguồn và các file mẫu ngân hàng đích.",
        )

        # Biến hiển thị thống kê
        self.stat_records_var = tk.StringVar(master=self, value="0 dòng")
        self.stat_records_detail_var = tk.StringVar(
            master=self,
            value="Khớp: 0 | Ghi: 0 | Bỏ qua: 0",
        )
        self.stat_exceptions_var = tk.StringVar(master=self, value="0 ngoại lệ")
        self.stat_exceptions_detail_var = tk.StringVar(
            master=self,
            value="Chưa xử lý: 0 | Đã xử lý: 0 | Chặn cứng: 0",
        )
        self.stat_amount_var = tk.StringVar(master=self, value="0 đ")
        self.stat_amount_detail_var = tk.StringVar(
            master=self,
            value="Nguồn đủ ĐK: 0 đ | Kỳ trước: 0 đ | Chênh lệch: 0 đ",
        )
        self.stat_safety_var = tk.StringVar(master=self, value="Chưa kiểm tra")

        self._build_styles()
        self._build_ui()
        self.refresh_templates()
        self._update_button_states()

    # ------------------------------------------------------------------ Thuộc tính truy cập nhanh
    @property
    def root_window(self) -> tk.Misc:
        return self.master

    @property
    def source_file(self) -> str:
        return self._source_file

    @property
    def target_files(self) -> list[str]:
        return list(self._target_files)

    @property
    def selected_source_template(self) -> str:
        return self.source_template_var.get().strip()

    @property
    def selected_target_template(self) -> str:
        return self.target_template_var.get().strip()

    @property
    def stats(self) -> MainWindowStats:
        return self._stats

    @property
    def workflow_result(self) -> WorkflowResult:
        return self.workflow.result

    @property
    def last_report(self) -> Optional[ReportResult]:
        return self._last_report

    # ------------------------------------------------------------------ Dựng giao diện (3 bước trên 1 màn hình)
    def _build_styles(self) -> None:
        style = ttk.Style(self)
        style.configure("Header.TLabel", font=("Arial", 14, "bold"), foreground="#1F4E78")
        style.configure("SubHeader.TLabel", font=("Arial", 10, "bold"), foreground="#1F4E78")
        style.configure("StatTitle.TLabel", font=("Arial", 10, "bold"), foreground="#333333")
        style.configure("StatValue.TLabel", font=("Arial", 14, "bold"), foreground="#1F4E78")
        style.configure("StatWarn.TLabel", font=("Arial", 14, "bold"), foreground="#B45F06")
        style.configure("StatDetail.TLabel", font=("Arial", 9), foreground="#555555")
        style.configure("Action.TButton", font=("Arial", 10, "bold"), padding=(12, 6))
        style.configure("Status.TLabel", font=("Arial", 10), padding=(6, 4))

    def _build_ui(self) -> None:
        # Tiêu đề ứng dụng
        header_frame = ttk.Frame(self)
        header_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(
            header_frame,
            text="CÔNG CỤ TỰ ĐỘNG ĐỐI CHIẾU & CHI HỘ LƯƠNG NGÂN HÀNG (V1)",
            style="Header.TLabel",
        ).pack(side=tk.LEFT)

        # ==============================================================
        # BƯỚC 1: CHỌN FILE NGUỒN, FILE ĐÍCH & TEMPLATE
        # ==============================================================
        step1_frame = ttk.LabelFrame(
            self,
            text=" ① Chọn File Lương Nguồn · File Ngân Hàng Đích · Cấu hình Template ",
            padding=10,
        )
        step1_frame.pack(fill=tk.X, pady=(0, 8))
        step1_frame.columnconfigure(1, weight=1)

        # Dòng 1: File nguồn
        ttk.Label(step1_frame, text="File bảng lương nguồn:").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 8), pady=4
        )
        self.entry_source_file = ttk.Entry(
            step1_frame,
            textvariable=self.source_file_var,
            state="readonly",
        )
        self.entry_source_file.grid(row=0, column=1, sticky=tk.EW, pady=4)
        self.btn_browse_source = ttk.Button(
            step1_frame,
            text="Chọn file nguồn...",
            command=self.browse_source_file,
        )
        self.btn_browse_source.grid(row=0, column=2, sticky=tk.EW, padx=(6, 0), pady=4)

        # Dòng 2: Sheet nguồn & Template nguồn
        src_opt_frame = ttk.Frame(step1_frame)
        src_opt_frame.grid(row=1, column=1, columnspan=2, sticky=tk.EW, pady=2)
        ttk.Label(src_opt_frame, text="Sheet lương:").pack(side=tk.LEFT, padx=(0, 4))
        self.cmb_source_sheet = ttk.Combobox(
            src_opt_frame,
            textvariable=self.source_sheet_var,
            width=22,
            state="readonly",
        )
        self.cmb_source_sheet.pack(side=tk.LEFT, padx=(0, 16))

        ttk.Label(src_opt_frame, text="Template nguồn:").pack(side=tk.LEFT, padx=(0, 4))
        self.cmb_source_template = ttk.Combobox(
            src_opt_frame,
            textvariable=self.source_template_var,
            width=28,
            state="readonly",
        )
        self.cmb_source_template.pack(side=tk.LEFT, padx=(0, 8))

        # Dòng 3: Danh sách file đích ngân hàng
        ttk.Label(step1_frame, text="File mẫu ngân hàng đích:").grid(
            row=2, column=0, sticky=tk.NW, padx=(0, 8), pady=4
        )
        target_list_frame = ttk.Frame(step1_frame)
        target_list_frame.grid(row=2, column=1, sticky=tk.EW, pady=4)
        target_list_frame.columnconfigure(0, weight=1)

        self.lst_target_files = tk.Listbox(
            target_list_frame,
            height=4,
            selectmode=tk.EXTENDED,
            activestyle="dotbox",
        )
        self.lst_target_files.grid(row=0, column=0, sticky=tk.EW)
        scr_targets = ttk.Scrollbar(
            target_list_frame,
            orient=tk.VERTICAL,
            command=self.lst_target_files.yview,
        )
        scr_targets.grid(row=0, column=1, sticky=tk.NS)
        self.lst_target_files.configure(yscrollcommand=scr_targets.set)

        target_btn_frame = ttk.Frame(step1_frame)
        target_btn_frame.grid(row=2, column=2, sticky=tk.N, padx=(6, 0), pady=4)
        self.btn_browse_targets = ttk.Button(
            target_btn_frame,
            text="Chọn file đích...",
            command=self.browse_target_files,
        )
        self.btn_browse_targets.pack(fill=tk.X, pady=(0, 4))
        self.btn_remove_target = ttk.Button(
            target_btn_frame,
            text="Xóa file đã chọn",
            command=self.remove_selected_target_files,
        )
        self.btn_remove_target.pack(fill=tk.X, pady=(0, 4))
        self.btn_clear_targets = ttk.Button(
            target_btn_frame,
            text="Xóa tất cả",
            command=self.clear_target_files,
        )
        self.btn_clear_targets.pack(fill=tk.X)

        # Dòng 4: Chọn Template đích & Thư mục xuất
        tgt_opt_frame = ttk.Frame(step1_frame)
        tgt_opt_frame.grid(row=3, column=1, columnspan=2, sticky=tk.EW, pady=(4, 0))
        ttk.Label(tgt_opt_frame, text="Template ngân hàng:").pack(side=tk.LEFT, padx=(0, 4))
        self.cmb_target_template = ttk.Combobox(
            tgt_opt_frame,
            textvariable=self.target_template_var,
            width=30,
            state="readonly",
        )
        self.cmb_target_template.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_refresh_templates = ttk.Button(
            tgt_opt_frame,
            text="Làm mới Template",
            command=self.refresh_templates,
        )
        self.btn_refresh_templates.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_open_mapping = ttk.Button(
            tgt_opt_frame,
            text="Cấu hình Mapping...",
            command=self.open_mapping_dialog,
        )
        self.btn_open_mapping.pack(side=tk.LEFT, padx=(0, 16))

        ttk.Label(tgt_opt_frame, text="Thư mục xuất:").pack(side=tk.LEFT, padx=(0, 4))
        self.entry_output_dir = ttk.Entry(
            tgt_opt_frame,
            textvariable=self.output_dir_var,
            width=24,
        )
        self.entry_output_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.btn_browse_output = ttk.Button(
            tgt_opt_frame,
            text="...",
            width=3,
            command=self.browse_output_dir,
        )
        self.btn_browse_output.pack(side=tk.LEFT)

        # ==============================================================
        # THANH NÚT ĐIỀU KHIỂN CHÍNH: ANALYZE · PREVIEW · EXECUTE
        # ==============================================================
        action_bar = ttk.Frame(self, padding=(0, 4))
        action_bar.pack(fill=tk.X, pady=(0, 8))

        self.btn_analyze = ttk.Button(
            action_bar,
            text="🔍 1. Analyze (Phân tích & Đối chiếu)",
            style="Action.TButton",
            command=self.analyze,
        )
        self.btn_analyze.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_preview = ttk.Button(
            action_bar,
            text="👁 2. Preview (Xem bảng & Ngoại lệ)",
            style="Action.TButton",
            command=self.preview,
        )
        self.btn_preview.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_execute = ttk.Button(
            action_bar,
            text="✅ 3. Execute (Xuất file kết quả & Báo cáo)",
            style="Action.TButton",
            command=self.execute,
        )
        self.btn_execute.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_open_output_dir = ttk.Button(
            action_bar,
            text="📂 Mở thư mục kết quả",
            command=self.open_output_folder,
        )
        self.btn_open_output_dir.pack(side=tk.RIGHT)

        # ==============================================================
        # BẢNG HIỂN THỊ THỐNG KÊ: SỐ RECORD · EXCEPTION · TỔNG TIỀN
        # ==============================================================
        stats_frame = ttk.LabelFrame(
            self,
            text=" ② Thống kê Đối chiếu (Số Record · Exception · Tổng Tiền) ",
            padding=10,
        )
        stats_frame.pack(fill=tk.X, pady=(0, 8))
        for col in range(4):
            stats_frame.columnconfigure(col, weight=1)

        # Thẻ 1: Số Record
        card_rec = ttk.Frame(stats_frame, padding=6, relief=tk.GROOVE)
        card_rec.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 6))
        ttk.Label(card_rec, text="SỐ RECORD NGUỒN", style="StatTitle.TLabel").pack(anchor=tk.W)
        self.lbl_stat_records = ttk.Label(
            card_rec,
            textvariable=self.stat_records_var,
            style="StatValue.TLabel",
        )
        self.lbl_stat_records.pack(anchor=tk.W, pady=(2, 2))
        ttk.Label(
            card_rec,
            textvariable=self.stat_records_detail_var,
            style="StatDetail.TLabel",
        ).pack(anchor=tk.W)

        # Thẻ 2: Exception
        card_exc = ttk.Frame(stats_frame, padding=6, relief=tk.GROOVE)
        card_exc.grid(row=0, column=1, sticky=tk.NSEW, padx=6)
        ttk.Label(card_exc, text="NGOẠI LỆ (EXCEPTION)", style="StatTitle.TLabel").pack(anchor=tk.W)
        self.lbl_stat_exceptions = ttk.Label(
            card_exc,
            textvariable=self.stat_exceptions_var,
            style="StatWarn.TLabel",
        )
        self.lbl_stat_exceptions.pack(anchor=tk.W, pady=(2, 2))
        ttk.Label(
            card_exc,
            textvariable=self.stat_exceptions_detail_var,
            style="StatDetail.TLabel",
        ).pack(anchor=tk.W)

        # Thẻ 3: Tổng tiền
        card_amt = ttk.Frame(stats_frame, padding=6, relief=tk.GROOVE)
        card_amt.grid(row=0, column=2, sticky=tk.NSEW, padx=6)
        ttk.Label(card_amt, text="TỔNG TIỀN GHI (VND)", style="StatTitle.TLabel").pack(anchor=tk.W)
        self.lbl_stat_amount = ttk.Label(
            card_amt,
            textvariable=self.stat_amount_var,
            style="StatValue.TLabel",
        )
        self.lbl_stat_amount.pack(anchor=tk.W, pady=(2, 2))
        ttk.Label(
            card_amt,
            textvariable=self.stat_amount_detail_var,
            style="StatDetail.TLabel",
        ).pack(anchor=tk.W)

        # Thẻ 4: Trạng thái An toàn & Hậu kiểm
        card_safe = ttk.Frame(stats_frame, padding=6, relief=tk.GROOVE)
        card_safe.grid(row=0, column=3, sticky=tk.NSEW, padx=(6, 0))
        ttk.Label(card_safe, text="CHỐT AN TOÀN & HẬU KIỂM", style="StatTitle.TLabel").pack(anchor=tk.W)
        self.lbl_stat_safety = ttk.Label(
            card_safe,
            textvariable=self.stat_safety_var,
            style="StatValue.TLabel",
        )
        self.lbl_stat_safety.pack(anchor=tk.W, pady=(2, 2))

        # ==============================================================
        # KHUNG CHỨA PREVIEW / TÓM TẮT KẾT QUẢ (Nơi cắm preview_table ở GUI Module 3)
        # ==============================================================
        self.preview_container = ttk.LabelFrame(
            self,
            text=" ③ Khung Preview & Chi tiết Ngoại lệ ",
            padding=8,
        )
        self.preview_container.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.txt_summary_view = tk.Text(
            self.preview_container,
            height=12,
            wrap=tk.WORD,
            font=("Consolas", 10),
            state=tk.DISABLED,
        )
        scr_summary = ttk.Scrollbar(
            self.preview_container,
            orient=tk.VERTICAL,
            command=self.txt_summary_view.yview,
        )
        self.txt_summary_view.configure(yscrollcommand=scr_summary.set)
        self.txt_summary_view.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scr_summary.pack(side=tk.RIGHT, fill=tk.Y)

        # Thanh trạng thái dưới cùng
        status_bar = ttk.Frame(self, relief=tk.SUNKEN)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Label(
            status_bar,
            textvariable=self.status_var,
            style="Status.TLabel",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

    # =========================================================================
    # 3. Các thao tác chọn File Nguồn, File Đích và Template
    # =========================================================================

    def browse_source_file(
        self,
        *,
        dialog_fn: Optional[Callable[..., str]] = None,
    ) -> Optional[str]:
        """Mở hộp thoại chọn file bảng lương nguồn (`.xlsx` / `.xls`)."""
        ask_fn = dialog_fn or filedialog.askopenfilename
        chosen = ask_fn(
            parent=self,
            title="Chọn file bảng lương nguồn (.xlsx / .xls)",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if chosen:
            self.set_source_file(chosen)
            return str(chosen)
        return None

    def set_source_file(self, path: PathLike, *, sheet_name: Optional[str] = None) -> None:
        """Thiết lập đường dẫn file bảng lương nguồn và nạp danh sách sheet."""
        p = Path(path)
        self._source_file = str(p)
        self.source_file_var.set(self._source_file)

        sheets: list[str] = []
        if p.exists():
            try:
                sheets = [s.name for s in list_sheets(p) if s.visible]
            except Exception:
                sheets = []

        self.cmb_source_sheet["values"] = ["(Tự động)", *sheets]
        if sheet_name and sheet_name in sheets:
            self.source_sheet_var.set(sheet_name)
        else:
            self.source_sheet_var.set("(Tự động)")

        self.status_var.set(f"Đã chọn file lương nguồn: {p.name}")
        self._update_button_states()

    def browse_target_files(
        self,
        *,
        dialog_fn: Optional[Callable[..., Sequence[str]]] = None,
    ) -> list[str]:
        """Mở hộp thoại chọn 1 hoặc nhiều file mẫu ngân hàng đích (`.xls` / `.xlsx`)."""
        ask_fn = dialog_fn or filedialog.askopenfilenames
        chosen = ask_fn(
            parent=self,
            title="Chọn 1 hoặc nhiều file mẫu ngân hàng (.xls / .xlsx)",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if chosen:
            self.add_target_files(list(chosen))
            return list(chosen)
        return []

    def set_target_files(self, paths: Sequence[PathLike]) -> None:
        """Thiết lập mới toàn bộ danh sách file ngân hàng đích."""
        self._target_files.clear()
        self.add_target_files(paths)

    def add_target_files(self, paths: Sequence[PathLike]) -> None:
        """Thêm 1 hoặc nhiều file ngân hàng đích vào danh sách (không trùng lặp)."""
        existing_norm = {str(Path(x).resolve()).casefold() for x in self._target_files}
        for p in paths:
            p_str = str(Path(p))
            norm = str(Path(p).resolve()).casefold()
            if norm not in existing_norm:
                self._target_files.append(p_str)
                existing_norm.add(norm)
        self._sync_target_listbox()
        self.status_var.set(f"Đã chọn {len(self._target_files)} file ngân hàng đích.")
        self._update_button_states()

    def remove_selected_target_files(self) -> None:
        """Xóa các file đích đang được bôi chọn trong Listbox."""
        raw_sel: Any = self.lst_target_files.curselection()  # type: ignore[no-untyped-call]
        selected_indices = sorted((int(i) for i in raw_sel), reverse=True)
        for idx in selected_indices:
            if 0 <= idx < len(self._target_files):
                self._target_files.pop(idx)
        self._sync_target_listbox()
        self._update_button_states()

    def clear_target_files(self) -> None:
        """Xóa toàn bộ danh sách file đích."""
        self._target_files.clear()
        self._sync_target_listbox()
        self._update_button_states()

    def browse_output_dir(
        self,
        *,
        dialog_fn: Optional[Callable[..., str]] = None,
    ) -> Optional[str]:
        ask_fn = dialog_fn or filedialog.askdirectory
        chosen = ask_fn(parent=self, title="Chọn thư mục lưu file kết quả")
        if chosen:
            self.output_dir_var.set(str(chosen))
            self.config_data.output_dir = str(chosen)
            return str(chosen)
        return None

    def refresh_templates(self) -> None:
        """Nạp danh sách cấu hình nguồn và mẫu ngân hàng từ `TemplateStore`."""
        self._template_store = TemplateStore(self.config_data.templates_dir)
        if self.config_data.auto_create_default_templates:
            self._template_store.ensure_defaults()

        src_names = [AUTO_TEMPLATE_OPTION, *[s.name for s in self._template_store.list_sources()]]
        tgt_names = [AUTO_TEMPLATE_OPTION, *[t.name for t in self._template_store.list_targets()]]

        self.cmb_source_template["values"] = src_names
        if self.source_template_var.get() not in src_names:
            self.source_template_var.set(AUTO_TEMPLATE_OPTION)

        self.cmb_target_template["values"] = tgt_names
        if self.target_template_var.get() not in tgt_names:
            self.target_template_var.set(AUTO_TEMPLATE_OPTION)

    def select_source_template(self, template_name: str) -> None:
        """Chọn cấu hình mapping cho file lương nguồn."""
        self.source_template_var.set(template_name)

    def select_target_template(self, template_name: str) -> None:
        """Chọn mẫu ngân hàng đích (hoặc `AUTO_TEMPLATE_OPTION`)."""
        self.target_template_var.set(template_name)

    def open_mapping_dialog(self) -> Optional[Any]:
        """Mở hộp thoại cấu hình Mapping (nếu module `gui.mapping_dialog` đã được cài đặt)."""
        try:
            from . import mapping_dialog

            open_fn = getattr(mapping_dialog, "open_mapping_dialog", None) or getattr(
                mapping_dialog, "MappingDialog", None
            )
            if callable(open_fn):
                dlg = open_fn(self, template_store=self._template_store)
                self.refresh_templates()
                return dlg
        except ImportError:
            pass

        msg = "Chức năng Mapping Dialog sẽ được cung cấp bởi module gui/mapping_dialog.py."
        self.status_var.set(msg)
        if self.interactive:
            messagebox.showinfo("Cấu hình Mapping", msg, parent=self)
        return None

    # =========================================================================
    # 4. Ba nút hành động chính: Analyze · Preview · Execute
    # =========================================================================

    def analyze(
        self,
        *,
        resolutions: Optional[Mapping[str, ResolutionValue]] = None,
    ) -> WorkflowResult:
        """Nút Analyze: Đọc file nguồn, nhận diện/áp dụng template, đọc file đích và chạy đối chiếu."""
        if not self._source_file:
            msg = "Vui lòng chọn file bảng lương nguồn trước khi phân tích (Analyze)."
            self.status_var.set(msg)
            if self.interactive:
                messagebox.showwarning("Thiếu file nguồn", msg, parent=self)
            return self.workflow.result

        if not self._target_files:
            msg = "Vui lòng chọn ít nhất 1 file mẫu ngân hàng đích trước khi phân tích (Analyze)."
            self.status_var.set(msg)
            if self.interactive:
                messagebox.showwarning("Thiếu file đích", msg, parent=self)
            return self.workflow.result

        self._sync_runtime_config()
        wf_res = self.workflow.prepare(
            source_file=self._source_file,
            target_files=self._target_files,
            config=self.config_data,
            resolutions=resolutions,
        )

        self.update_statistics(wf_res)
        self._render_summary_text(wf_res)
        self._refresh_embedded_preview(wf_res)
        self._update_button_states()

        if wf_res.stage == WorkflowStage.FAILED:
            reason = wf_res.stopped_reason or "Phân tích thất bại."
            self.status_var.set(f"Lỗi phân tích: {reason}")
            if self.interactive:
                messagebox.showerror("Phân tích không thành công", reason, parent=self)
        elif wf_res.has_unresolved_exceptions:
            self.status_var.set(
                f"Đã phân tích xong: khớp {self._stats.total_matched_records}/{self._stats.total_source_records} dòng. "
                f"Còn {self._stats.unresolved_exceptions} ngoại lệ cần xử lý trong Preview."
            )
        else:
            self.status_var.set(
                f"Đã phân tích xong: {self._stats.total_written_records} ô sẵn sàng ghi, "
                f"tổng tiền {format_vnd(self._stats.total_planned_write_amount)}."
            )

        return wf_res

    def preview(self) -> WorkflowResult:
        """Nút Preview: Hiển thị chi tiết kết quả đối chiếu và danh sách ngoại lệ."""
        if self.workflow.result.write_plan is None and self._source_file and self._target_files:
            self.analyze()

        wf_res = self.workflow.result
        self._preview_visible = True
        self.update_statistics(wf_res)
        self._render_summary_text(wf_res, detailed=True)
        self._refresh_embedded_preview(wf_res)
        self._update_button_states()

        if wf_res.write_plan is not None:
            self.status_var.set(
                f"Đang hiển thị Preview: {len(wf_res.write_plan.matches)} dòng đối chiếu, "
                f"{len(wf_res.exceptions)} ngoại lệ."
            )
        return wf_res

    def execute(
        self,
        *,
        resolutions: Optional[Mapping[str, ResolutionValue]] = None,
        output_paths: Optional[Any] = None,
    ) -> WorkflowResult:
        """Nút Execute: Thực thi kiểm tra an toàn -> sao lưu -> ghi file mới -> hậu kiểm -> xuất báo cáo."""
        if self.workflow.result.write_plan is None:
            if self._source_file and self._target_files:
                prep = self.analyze(resolutions=resolutions)
                if prep.stage == WorkflowStage.FAILED:
                    return prep
            else:
                msg = "Vui lòng chạy Analyze trước khi thực thi (Execute)."
                self.status_var.set(msg)
                if self.interactive:
                    messagebox.showwarning("Chưa có kế hoạch ghi", msg, parent=self)
                return self.workflow.result
        elif resolutions:
            self.workflow.apply_resolutions(resolutions)

        self._sync_runtime_config()
        wf_res = self.workflow.execute(output_paths=output_paths)
        self.update_statistics(wf_res)

        # Sinh báo cáo kết quả (.xlsx) nếu đã có WritePlan
        if wf_res.write_plan is not None:
            try:
                self._last_report = self.reporter.generate(
                    run_result=wf_res.run_result,
                    workflow_result=wf_res,
                    verification_result=wf_res.verification_result,
                    validation_result=wf_res.validation_result,
                )
            except Exception:
                self._last_report = None

        self._render_summary_text(wf_res, detailed=True)
        self._refresh_embedded_preview(wf_res)
        self._update_button_states()

        if wf_res.is_success:
            out_count = len(wf_res.output_files)
            rep_name = Path(self._last_report.report_file).name if (self._last_report and self._last_report.report_file) else ""
            msg = (
                f"Đã xuất thành công {out_count} file ngân hàng mới và hậu kiểm đạt 100%!\n"
                f"Tổng số ô đã ghi: {self._stats.total_written_records}\n"
                f"Tổng số tiền đã ghi: {format_vnd(self._stats.total_planned_write_amount)}"
                + (f"\nBáo cáo: {rep_name}" if rep_name else "")
            )
            self.status_var.set(msg.replace("\n", " | "))
            if self.interactive:
                messagebox.showinfo("Hoàn tất xuất file", msg, parent=self)
        else:
            reason = wf_res.stopped_reason or "Quy trình bị dừng do không đạt kiểm tra an toàn hoặc hậu kiểm."
            self.status_var.set(f"Đã dừng thực thi: {reason}")
            if self.interactive:
                messagebox.showerror("Không thể xuất file", reason, parent=self)

        return wf_res

    def resolve_exception(
        self,
        exception_id: str,
        action: Union[ResolutionAction, str],
        *,
        override_amount: Optional[int] = None,
    ) -> WorkflowResult:
        """Cập nhật lựa chọn xử lý cho 1 ngoại lệ từ giao diện và làm mới thống kê."""
        wf_res = self.workflow.resolve_exception(
            exception_id,
            action,
            override_amount=override_amount,
        )
        self.update_statistics(wf_res)
        self._render_summary_text(wf_res, detailed=self._preview_visible)
        self._update_button_states()
        return wf_res

    def open_output_folder(self) -> Optional[Path]:
        """Mở thư mục chứa file kết quả trong Windows Explorer."""
        folder = Path(self.output_dir_var.get().strip() or self.config_data.output_dir)
        folder.mkdir(parents=True, exist_ok=True)
        if self.interactive and hasattr(os, "startfile"):
            try:
                os.startfile(str(folder.resolve()))
            except OSError:
                pass
        return folder.resolve()

    # =========================================================================
    # 5. Tính toán & Cập nhật thống kê hiển thị
    # =========================================================================

    def update_statistics(self, wf_res: Optional[WorkflowResult] = None) -> MainWindowStats:
        """Tính toán và cập nhật các nhãn thống kê (số record, exception, tổng tiền) trên giao diện."""
        res = wf_res or self.workflow.result
        plan = res.write_plan
        ver = res.verification_result or res.pre_verification_result
        val = res.validation_result

        if plan is None:
            self._stats = MainWindowStats(stage=res.stage.value)
        else:
            p_sum = plan.summary()
            total_src = ver.total_source_records if ver else plan.total_source_records
            matched = (
                ver.total_matched_records
                if ver
                else sum(1 for m in plan.matches if m.source_row is not None and m.target_row is not None)
            )
            written = ver.total_written_records if ver else len(plan.updates)
            src_amt = (
                ver.total_source_amount
                if ver
                else sum(
                    m.source_amount
                    for m in plan.matches
                    if m.source_row is not None and isinstance(m.source_amount, int) and m.source_amount > 0
                )
            )
            src_trans_amt = plan.total_source_transferable_amount
            old_amt = plan.total_old_amount
            new_amt = plan.total_new_amount

            excs = res.exceptions if res.exceptions else plan.exceptions
            total_exc = len(excs)
            resolved_exc = sum(1 for e in excs if e.is_resolved)
            unresolved_exc = sum(1 for e in excs if not e.is_resolved)
            blocking_exc = sum(1 for e in excs if e.is_blocking)

            is_safe = val.is_safe if val is not None else plan.is_ready
            is_verified = res.verified

            self._stats = MainWindowStats(
                total_source_records=total_src,
                total_matched_records=matched,
                total_written_records=written,
                will_update_records=p_sum["will_update"],
                unchanged_records=p_sum["unchanged"],
                skipped_records=p_sum["skipped"],
                pending_records=p_sum["pending"],
                error_records=p_sum["error"],
                total_exceptions=total_exc,
                resolved_exceptions=resolved_exc,
                unresolved_exceptions=unresolved_exc,
                blocking_exceptions=blocking_exc,
                total_source_amount=src_amt,
                total_source_transferable_amount=src_trans_amt,
                total_old_amount=old_amt,
                total_planned_write_amount=new_amt,
                delta_amount=new_amt - old_amt,
                stage=res.stage.value,
                is_safe=is_safe,
                is_verified=is_verified,
            )

        s = self._stats
        self.stat_records_var.set(f"{s.total_source_records:,} dòng nguồn")
        self.stat_records_detail_var.set(
            f"Khớp: {s.total_matched_records:,} | Ghi: {s.total_written_records:,} "
            f"(🟢 {s.will_update_records:,} · ⚪ {s.unchanged_records:,}) | Bỏ qua: {s.skipped_records:,}"
        )

        self.stat_exceptions_var.set(f"{s.total_exceptions:,} ngoại lệ")
        self.stat_exceptions_detail_var.set(
            f"🟡 Chưa xử lý: {s.unresolved_exceptions:,} | Đã xử lý: {s.resolved_exceptions:,} | "
            f"🔴 Chặn: {s.blocking_exceptions:,}"
        )

        self.stat_amount_var.set(format_vnd(s.total_planned_write_amount))
        self.stat_amount_detail_var.set(
            f"Nguồn đủ ĐK: {format_vnd(s.total_source_transferable_amount)} | "
            f"Kỳ trước: {format_vnd(s.total_old_amount)}"
        )

        if s.stage == WorkflowStage.COMPLETED.value and s.is_verified:
            self.stat_safety_var.set("✅ ĐÃ XUẤT & HẬU KIỂM")
        elif s.is_safe:
            self.stat_safety_var.set("🟢 AN TOÀN (Sẵn sàng ghi)")
        elif s.blocking_exceptions > 0 or s.stage in (
            WorkflowStage.SAFETY_FAILED.value,
            WorkflowStage.VERIFICATION_FAILED.value,
            WorkflowStage.FAILED.value,
        ):
            self.stat_safety_var.set("🔴 CHẶN GHI")
        elif s.unresolved_exceptions > 0:
            self.stat_safety_var.set("🟡 CẦN XÁC NHẬN")
        else:
            self.stat_safety_var.set("⚪ CHƯA PHÂN TÍCH")

        return self._stats

    def get_statistics(self) -> MainWindowStats:
        """Trả về đối tượng `MainWindowStats` hiện tại."""
        return self._stats

    # =========================================================================
    # 6. Các hàm hỗ trợ nội bộ
    # =========================================================================

    def _sync_target_listbox(self) -> None:
        self.lst_target_files.delete(0, tk.END)
        for p in self._target_files:
            self.lst_target_files.insert(tk.END, p)

    def _sync_runtime_config(self) -> None:
        cfg = self.config_data
        out_dir = self.output_dir_var.get().strip()
        if out_dir:
            cfg.output_dir = out_dir

        sheet_sel = self.source_sheet_var.get().strip()
        cfg.source_sheet = None if (not sheet_sel or sheet_sel == "(Tự động)") else sheet_sel

        src_tmpl = self.source_template_var.get().strip()
        cfg.source_profile_name = (
            None if (not src_tmpl or src_tmpl == AUTO_TEMPLATE_OPTION) else src_tmpl
        )

        tgt_tmpl = self.target_template_var.get().strip()
        cfg.target_templates = (
            None if (not tgt_tmpl or tgt_tmpl == AUTO_TEMPLATE_OPTION) else tgt_tmpl
        )

    def _update_button_states(self) -> None:
        has_files = bool(self._source_file and self._target_files)
        has_plan = self.workflow.result.write_plan is not None
        can_exec = has_files and (not has_plan or self.workflow.result.can_proceed_to_write)

        self.btn_analyze.configure(state=tk.NORMAL if has_files else tk.DISABLED)
        self.btn_preview.configure(state=tk.NORMAL if (has_files or has_plan) else tk.DISABLED)
        self.btn_execute.configure(state=tk.NORMAL if can_exec else tk.DISABLED)

    def _refresh_embedded_preview(self, wf_res: WorkflowResult) -> None:
        """Nếu module `gui.preview_table` đã có, tự động đẩy dữ liệu sang bảng Preview."""
        if self._embedded_preview_widget is not None:
            setter = getattr(self._embedded_preview_widget, "set_workflow_result", None) or getattr(
                self._embedded_preview_widget, "load_plan", None
            )
            if callable(setter):
                setter(wf_res)

    def _render_summary_text(self, wf_res: WorkflowResult, *, detailed: bool = False) -> None:
        lines: list[str] = []
        s = self._stats
        lines.append(f"Trạng thái quy trình : {s.stage}")
        lines.append(f"File bảng lương nguồn: {self._source_file or '(Chưa chọn)'}")
        lines.append(f"Số file ngân hàng    : {len(self._target_files)}")
        lines.append("-" * 78)
        lines.append(
            f"• Tổng record nguồn  : {s.total_source_records:,} dòng "
            f"(Khớp đích: {s.total_matched_records:,} | Sẽ ghi: {s.total_written_records:,} | Bỏ qua: {s.skipped_records:,})"
        )
        lines.append(
            f"• Ngoại lệ (Exception): {s.total_exceptions:,} "
            f"(Chưa xử lý: {s.unresolved_exceptions:,} | Đã xử lý: {s.resolved_exceptions:,} | Chặn cứng: {s.blocking_exceptions:,})"
        )
        lines.append(
            f"• Tổng tiền lương đủ ĐK: {format_vnd(s.total_source_transferable_amount)} | "
            f"Tổng tiền chuẩn bị ghi: {format_vnd(s.total_planned_write_amount)} "
            f"(Chênh lệch kỳ trước: {format_vnd(s.delta_amount)})"
        )

        if wf_res.write_plan is not None and wf_res.write_plan.target_plans:
            lines.append("-" * 78)
            lines.append("Chi tiết theo từng file ngân hàng đích:")
            for tp in wf_res.write_plan.target_plans:
                lines.append(
                    f"  - {Path(tp.file_path).name} [{tp.sheet_name}, cột {tp.amount_column}]: "
                    f"{len(tp.updates)} ô ghi ({len(tp.changed_updates)} ô đổi giá trị), "
                    f"tổng tiền mới = {format_vnd(tp.total_new_amount)}"
                )

        if wf_res.exceptions:
            lines.append("-" * 78)
            lines.append(f"Danh sách ngoại lệ ({len(wf_res.exceptions)}):")
            for exc in wf_res.exceptions:
                status_tag = (
                    "🔴 CHẶN CỨNG"
                    if exc.is_blocking
                    else (f"🟢 Đã chọn: {exc.action.value}" if exc.is_resolved and exc.action else "🟡 CHƯA XỬ LÝ")
                )
                lines.append(
                    f"  [{exc.id} - {exc.code.value}] ({status_tag}) "
                    f"Thẻ={exc.card_no or '-'}, TK={exc.account_no or '-'}: {exc.message}"
                )

        if detailed and wf_res.write_plan is not None and wf_res.write_plan.matches:
            lines.append("-" * 78)
            lines.append(f"Bảng đối chiếu chi tiết ({len(wf_res.write_plan.matches)} dòng):")
            for idx, m in enumerate(wf_res.write_plan.matches[:100], start=1):
                icon = {
                    MatchStatus.WILL_UPDATE: "🟢",
                    MatchStatus.UNCHANGED: "⚪",
                    MatchStatus.PENDING: "🟡",
                    MatchStatus.SKIPPED: "⚪",
                    MatchStatus.ERROR: "🔴",
                }.get(m.status, "•")
                lines.append(
                    f"  {idx:02d}. {icon} [{m.status.value}] Thẻ={m.card_no or '-'} | "
                    f"TK={m.account_no or '-'} | {m.source_name or m.target_name} | "
                    f"Ô={m.target_cell or '-'} | Cũ={format_vnd(m.old_amount)} -> Mới={format_vnd(m.new_amount)}"
                )

        if wf_res.output_files:
            lines.append("-" * 78)
            lines.append("File kết quả đã xuất:")
            for out_f in wf_res.output_files:
                lines.append(f"  + {out_f}")
            if self._last_report and self._last_report.report_file:
                lines.append(f"File báo cáo Excel (.xlsx): {self._last_report.report_file}")

        if wf_res.errors:
            lines.append("-" * 78)
            lines.append("Thông báo lỗi / lý do dừng:")
            for err in wf_res.errors:
                lines.append(f"  ! {err}")

        self.txt_summary_view.configure(state=tk.NORMAL)
        self.txt_summary_view.delete("1.0", tk.END)
        self.txt_summary_view.insert("1.0", "\n".join(lines))
        self.txt_summary_view.configure(state=tk.DISABLED)

    def destroy(self) -> None:
        owns = getattr(self, "_owns_root", False)
        root = self.master
        super().destroy()
        if owns and root is not None:
            try:
                root.destroy()
            except tk.TclError:
                pass


def launch_main_window(config: Optional[RuntimeConfig] = None) -> None:
    """Khởi chạy cửa sổ chính của ứng dụng."""
    app = MainWindow(config=config, interactive=True)
    app.root_window.mainloop()
