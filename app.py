"""
app.py - Điểm khởi chạy chính của Công cụ Tự động Đối chiếu & Chi hộ Lương Ngân hàng (V1).

Chức năng:
1. Chuẩn bị các thư mục vận hành (`templates/`, `data/backup/`, `data/output/`, `data/reports/`, `data/logs/`).
2. Cấu hình ghi log theo ngày (`data/logs/chi_ho_luong_YYYYMMDD.log`).
3. Khởi tạo `MainWindow` kết nối sẵn bảng `PreviewTable` và `TemplateStore`, mở giao diện cho kế toán.
"""

from __future__ import annotations

import json
import logging
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from typing import Any, Optional, Sequence

from core.models import ResolutionAction, RuntimeConfig
from core.reporter import Reporter
from core.template_store import TemplateStore
from gui.main_window import MainWindow
from gui.preview_table import PreviewTable

__all__ = [
    "APP_NAME",
    "APP_VERSION",
    "configure_logging",
    "create_app_window",
    "ensure_runtime_directories",
    "get_app_root",
    "main",
]

APP_NAME: str = "ChiHoLuongTool"
APP_VERSION: str = "1.0.0"


def get_app_root() -> Path:
    """Trả về thư mục gốc của ứng dụng (hỗ trợ cả chạy trực tiếp bằng Python và bản đóng gói .exe)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def ensure_runtime_directories(base_dir: Optional[Path] = None) -> RuntimeConfig:
    """Đảm bảo các thư mục dữ liệu và template mặc định tồn tại trước khi chạy."""
    root = base_dir or get_app_root()
    templates_dir = root / "templates"
    backup_dir = root / "data" / "backup"
    output_dir = root / "data" / "output"
    reports_dir = root / "data" / "reports"
    logs_dir = root / "data" / "logs"

    for folder in (templates_dir, backup_dir, output_dir, reports_dir, logs_dir):
        folder.mkdir(parents=True, exist_ok=True)

    store = TemplateStore(templates_dir)
    store.ensure_defaults()

    return RuntimeConfig(
        templates_dir=str(templates_dir),
        output_dir=str(output_dir),
        backup_dir=str(backup_dir),
        auto_create_default_templates=True,
    )


def configure_logging(base_dir: Optional[Path] = None) -> Path:
    """Thiết lập ghi log UTF-8 theo ngày tại `data/logs/chi_ho_luong_YYYYMMDD.log`."""
    root = base_dir or get_app_root()
    logs_dir = root / "data" / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)

    date_tag = datetime.now().strftime("%Y%m%d")
    log_file = logs_dir / f"chi_ho_luong_{date_tag}.log"

    root_logger = logging.getLogger()
    if not root_logger.handlers:
        root_logger.setLevel(logging.INFO)
        fmt = logging.Formatter(
            fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(fmt)
        root_logger.addHandler(file_handler)

    return log_file


def create_app_window(
    master: Optional[tk.Misc] = None,
    *,
    config: Optional[RuntimeConfig] = None,
    interactive: bool = True,
) -> MainWindow:
    """Tạo cửa sổ chính `MainWindow` và gắn kết bảng `PreviewTable` vào khung Preview."""
    runtime_cfg = config or ensure_runtime_directories()
    reports_dir = Path(runtime_cfg.output_dir).parent / "reports"
    reporter = Reporter(report_dir=reports_dir)
    window = MainWindow(
        master,
        config=runtime_cfg,
        reporter=reporter,
        interactive=interactive,
    )

    def _handle_resolve(
        exc_id: str,
        action: ResolutionAction,
        override_amount: Optional[int],
    ) -> None:
        wf_res = window.resolve_exception(
            exc_id,
            action,
            override_amount=override_amount,
        )
        preview_widget.set_workflow_result(wf_res)

    preview_widget = PreviewTable(
        window.preview_container,
        on_resolve_exception=_handle_resolve,
    )
    preview_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 6))
    window.txt_summary_view.configure(height=6)
    window._embedded_preview_widget = preview_widget

    return window


def _run_smoke_test(args: Sequence[str], cfg: RuntimeConfig) -> int:
    """Chạy kiểm thử tự động từ dòng lệnh (`--smoke-test <json_out> [source_file] [target1 ...]`)."""
    out_json = Path(args[0])
    root = tk.Tk()
    root.withdraw()
    window = create_app_window(root, config=cfg, interactive=False)
    try:
        if len(args) >= 3:
            window.set_source_file(args[1], sheet_name="HLCK05.26 (2)")
            window.set_target_files(list(args[2:]))
            window.analyze()
            window.preview()
            wf_res = window.execute()
            payload: dict[str, Any] = {
                "app_name": APP_NAME,
                "app_version": APP_VERSION,
                "frozen": bool(getattr(sys, "frozen", False)),
                "stage": wf_res.stage.value,
                "is_success": wf_res.is_success,
                "stats": window.stats.to_dict(),
                "output_files": [Path(p).name for p in wf_res.output_files],
                "backup_count": len(wf_res.backup_files),
                "has_report": bool(window.last_report and window.last_report.report_file),
                "errors": list(wf_res.errors),
            }
        else:
            payload = {
                "app_name": APP_NAME,
                "app_version": APP_VERSION,
                "frozen": bool(getattr(sys, "frozen", False)),
                "stage": window.stats.stage,
                "is_success": True,
                "stats": window.stats.to_dict(),
                "output_files": [],
                "backup_count": 0,
                "has_report": False,
                "errors": [],
            }
        out_json.parent.mkdir(parents=True, exist_ok=True)
        out_json.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return 0 if payload["is_success"] else 1
    finally:
        window.destroy()
        try:
            root.destroy()
        except tk.TclError:
            pass


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Hàm khởi chạy chính của ứng dụng."""
    for stream in (sys.stdout, sys.stderr):
        reconf = getattr(stream, "reconfigure", None)
        if callable(reconf):
            try:
                reconf(encoding="utf-8", errors="replace")
            except Exception:
                pass

    args = list(sys.argv[1:] if argv is None else argv)
    base_dir = get_app_root()
    log_path = configure_logging(base_dir)
    cfg = ensure_runtime_directories(base_dir)
    logging.getLogger(__name__).info(
        "Khởi động %s v%s (log: %s)",
        APP_NAME,
        APP_VERSION,
        log_path,
    )

    if args and args[0] == "--smoke-test" and len(args) >= 2:
        return _run_smoke_test(args[1:], cfg)

    window = create_app_window(config=cfg, interactive=True)
    window.root_window.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
