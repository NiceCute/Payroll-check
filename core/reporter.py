"""
reporter.py - Lớp tạo dữ liệu báo cáo và xuất file báo cáo Excel (.xlsx) (Reporting Layer).

Đầu vào:
- `RunResult`
- `WorkflowResult`
- `VerificationResult`
- `ValidationResult`
(hoặc `WritePlan`)

Chức năng:
1. Sinh báo cáo tổng hợp (`SummaryReportData` + `DetailReportRow`).
2. Sinh `Exception Report` (`ExceptionReportRow`).
3. Sinh `Validation Report` (`ValidationReportRow` từ `ValidationResult`, `VerificationResult`, `CellDiff`).
4. Xuất file báo cáo Excel `.xlsx` gồm 4 sheet:
   - `Tổng hợp`
   - `Chi tiết`
   - `Exception`
   - `Kiểm tra & Hậu kiểm`

Nguyên tắc:
- Không xử lý GUI, không xử lý workflow, không ghi vào file ngân hàng.
- Các cột định danh (`Số thẻ`, `Số tài khoản`, `Mã NH`) trong file báo cáo `.xlsx` luôn định dạng Text (`@`) để giữ số 0 đầu.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Optional, Sequence, Union

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .excel_reader import PathLike
from .models import (
    DetailReportRow,
    ExceptionItem,
    ExceptionReportRow,
    MatchStatus,
    ReportError,
    ReportResult,
    RunResult,
    SummaryReportData,
    TargetFileSummaryRow,
    ValidationReportRow,
    ValidationResult,
    VerificationResult,
    WorkflowResult,
    WriteExecutionResult,
    WritePlan,
)
from .verifier import verify

DEFAULT_REPORT_DIR: str = "data/reports"
DEFAULT_REPORT_PREFIX: str = "BaoCao_ChiHoLuong"

SHEET_SUMMARY: str = "Tổng hợp"
SHEET_DETAIL: str = "Chi tiết"
SHEET_EXCEPTION: str = "Exception"
SHEET_VALIDATION: str = "Kiểm tra & Hậu kiểm"

__all__ = [
    "DEFAULT_REPORT_DIR",
    "DEFAULT_REPORT_PREFIX",
    "DetailReportRow",
    "ExceptionReportRow",
    "ReportError",
    "ReportResult",
    "Reporter",
    "SHEET_DETAIL",
    "SHEET_EXCEPTION",
    "SHEET_SUMMARY",
    "SHEET_VALIDATION",
    "SummaryReportData",
    "TargetFileSummaryRow",
    "ValidationReportRow",
    "build_detail_report",
    "build_exception_report",
    "build_report_data",
    "build_summary_report",
    "build_validation_report",
    "create_report",
    "export_report_xlsx",
    "generate_report",
]

ReportInput = Union[RunResult, WorkflowResult, WritePlan]


# =============================================================================
# 1. Giao diện chính (Class & Function API)
# =============================================================================

class Reporter:
    """Bộ sinh dữ liệu báo cáo và xuất file Excel (.xlsx)."""

    def __init__(
        self,
        *,
        report_dir: PathLike = DEFAULT_REPORT_DIR,
        report_prefix: str = DEFAULT_REPORT_PREFIX,
        write_excel_file: bool = True,
    ) -> None:
        self.report_dir = Path(report_dir)
        self.report_prefix = report_prefix
        self.write_excel_file = write_excel_file

    def generate(
        self,
        run_result: Optional[ReportInput] = None,
        workflow_result: Optional[WorkflowResult] = None,
        verification_result: Optional[VerificationResult] = None,
        validation_result: Optional[ValidationResult] = None,
        *,
        write_plan: Optional[WritePlan] = None,
        output_path: Optional[PathLike] = None,
        write_excel_file: Optional[bool] = None,
    ) -> ReportResult:
        should_write = self.write_excel_file if write_excel_file is None else write_excel_file
        return generate_report(
            run_result=run_result,
            workflow_result=workflow_result,
            verification_result=verification_result,
            validation_result=validation_result,
            write_plan=write_plan,
            output_path=output_path,
            report_dir=self.report_dir,
            report_prefix=self.report_prefix,
            write_excel_file=should_write,
        )


def generate_report(
    run_result: Optional[ReportInput] = None,
    workflow_result: Optional[WorkflowResult] = None,
    verification_result: Optional[VerificationResult] = None,
    validation_result: Optional[ValidationResult] = None,
    *,
    write_plan: Optional[WritePlan] = None,
    output_path: Optional[PathLike] = None,
    report_dir: PathLike = DEFAULT_REPORT_DIR,
    report_prefix: str = DEFAULT_REPORT_PREFIX,
    write_excel_file: bool = True,
) -> ReportResult:
    """Tổng hợp dữ liệu báo cáo và (mặc định) xuất ra file Excel `.xlsx`.

    Cập nhật luôn `run_result.report_file` nếu có xuất file `.xlsx`.
    """
    eff_plan, eff_run, eff_wf, eff_ver, eff_val = _resolve_inputs(
        run_result=run_result,
        workflow_result=workflow_result,
        verification_result=verification_result,
        validation_result=validation_result,
        write_plan=write_plan,
    )

    report = build_report_data(
        run_result=eff_run,
        workflow_result=eff_wf,
        verification_result=eff_ver,
        validation_result=eff_val,
        write_plan=eff_plan,
    )

    if write_excel_file:
        target_path = (
            Path(output_path)
            if output_path is not None
            else _build_default_report_path(report_dir, report_prefix)
        )
        saved_path = export_report_xlsx(report, target_path)
        report.report_file = str(saved_path)
        if eff_run is not None:
            eff_run.report_file = str(saved_path)
        if eff_wf is not None and eff_wf.run_result is not None:
            eff_wf.run_result.report_file = str(saved_path)

    return report


create_report = generate_report


# =============================================================================
# 2. Xây dựng các cấu trúc dữ liệu báo cáo
# =============================================================================

def build_report_data(
    run_result: Optional[ReportInput] = None,
    workflow_result: Optional[WorkflowResult] = None,
    verification_result: Optional[VerificationResult] = None,
    validation_result: Optional[ValidationResult] = None,
    *,
    write_plan: Optional[WritePlan] = None,
) -> ReportResult:
    """Tạo đối tượng `ReportResult` hoàn chỉnh trên bộ nhớ (không ghi file)."""
    eff_plan, eff_run, eff_wf, eff_ver, eff_val = _resolve_inputs(
        run_result=run_result,
        workflow_result=workflow_result,
        verification_result=verification_result,
        validation_result=validation_result,
        write_plan=write_plan,
    )

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    summary_data = build_summary_report(
        write_plan=eff_plan,
        run_result=eff_run,
        workflow_result=eff_wf,
        verification_result=eff_ver,
        validation_result=eff_val,
        created_at=now_str,
    )
    detail_rows = build_detail_report(eff_plan)
    exceptions = eff_wf.exceptions if (eff_wf is not None and eff_wf.exceptions) else eff_plan.exceptions
    exception_rows = build_exception_report(exceptions)
    validation_rows = build_validation_report(
        write_plan=eff_plan,
        validation_result=eff_val,
        verification_result=eff_ver,
    )

    return ReportResult(
        summary_report=summary_data,
        detail_rows=detail_rows,
        exception_report=exception_rows,
        validation_report=validation_rows,
        report_file=eff_run.report_file if eff_run is not None else None,
        created_at=now_str,
    )


def build_summary_report(
    *,
    write_plan: WritePlan,
    run_result: Optional[RunResult] = None,
    workflow_result: Optional[WorkflowResult] = None,
    verification_result: Optional[VerificationResult] = None,
    validation_result: Optional[ValidationResult] = None,
    created_at: Optional[str] = None,
) -> SummaryReportData:
    """Tạo `SummaryReportData` (báo cáo tổng hợp dòng, tiền, chênh lệch và kết quả từng file đích)."""
    now_str = created_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Nếu chưa truyền VerificationResult, tự tính chỉ số nhất quán trên bộ nhớ
    ver = verification_result or verify(
        run_result=write_plan,
        validation_result=validation_result,
        source=workflow_result.source_result if workflow_result is not None else None,
        verify_files_on_disk=False,
    )

    is_safe = validation_result.is_safe if validation_result is not None else write_plan.is_ready
    is_verified = (
        verification_result.is_verified
        if verification_result is not None
        else (run_result.verified if run_result is not None else ver.is_verified)
    )
    stage_str = (
        workflow_result.stage.value
        if workflow_result is not None
        else ("COMPLETED" if (run_result and run_result.verified) else "PREPARED")
    )

    output_files: list[str] = []
    backup_files: list[str] = []
    if run_result is not None:
        output_files = list(run_result.output_files)
        backup_files = list(run_result.backup_files)
    elif workflow_result is not None:
        output_files = list(workflow_result.output_files)
        backup_files = list(workflow_result.backup_files)

    exec_res: Optional[WriteExecutionResult] = (
        workflow_result.write_execution_result if workflow_result is not None else None
    )
    out_by_target: dict[str, tuple[str, str]] = {}
    if exec_res is not None:
        for t_out in exec_res.target_outputs:
            out_by_target[str(Path(t_out.source_target_file).resolve())] = (
                t_out.output_file,
                t_out.backup_file,
            )

    target_summaries: list[TargetFileSummaryRow] = []
    for idx, tp in enumerate(write_plan.target_plans):
        key_resolved = str(Path(tp.file_path).resolve())
        if key_resolved in out_by_target:
            out_f, bak_f = out_by_target[key_resolved]
        else:
            out_f = output_files[idx] if idx < len(output_files) else ""
            bak_f = ""
        old_amt = tp.total_old_amount
        new_amt = tp.total_new_amount
        target_summaries.append(
            TargetFileSummaryRow(
                target_file=tp.file_path,
                sheet_name=tp.sheet_name,
                template_name=tp.template_name,
                amount_column=tp.amount_column,
                total_target_rows=tp.total_target_rows,
                scheduled_updates=len(tp.updates),
                changed_cells=len(tp.changed_updates),
                total_old_amount=old_amt,
                total_new_amount=new_amt,
                delta_amount=new_amt - old_amt,
                output_file=out_f,
                backup_file=bak_f,
            )
        )

    return SummaryReportData(
        created_at=now_str,
        stage=stage_str,
        is_safe=is_safe,
        is_verified=is_verified,
        source_file=write_plan.source_file,
        source_sheet=write_plan.source_sheet,
        total_source_records=ver.total_source_records,
        total_matched_records=ver.total_matched_records,
        total_unmatched_source_records=ver.total_unmatched_source_records,
        total_unmatched_target_records=ver.total_unmatched_target_records,
        total_written_records=ver.total_written_records,
        total_source_written_records=ver.total_source_written_records,
        total_zero_or_clear_records=ver.total_zero_or_clear_records,
        total_skipped_records=ver.total_skipped_records,
        total_exceptions=ver.total_exceptions,
        resolved_exceptions=ver.resolved_exceptions,
        unresolved_exceptions=ver.unresolved_exceptions,
        blocking_exceptions=ver.blocking_exceptions,
        total_source_amount=ver.total_source_amount,
        total_source_transferable_amount=ver.total_source_transferable_amount,
        total_old_target_amount=write_plan.total_old_amount,
        total_planned_write_amount=ver.total_planned_write_amount,
        total_skipped_transferable_amount=ver.total_skipped_transferable_amount,
        total_override_delta=ver.total_override_delta,
        row_discrepancy=ver.row_discrepancy,
        amount_discrepancy=ver.amount_discrepancy,
        safety_errors_count=len(validation_result.errors) if validation_result else 0,
        safety_warnings_count=len(validation_result.warnings) if validation_result else 0,
        verification_errors_count=len(ver.errors),
        verification_warnings_count=len(ver.warnings),
        cell_diffs_count=len(ver.cell_diffs),
        output_files=output_files,
        backup_files=backup_files,
        target_summaries=target_summaries,
    )


def build_detail_report(write_plan: WritePlan) -> list[DetailReportRow]:
    """Tạo danh sách `DetailReportRow` từ `write_plan.matches`."""
    rows: list[DetailReportRow] = []
    for idx, m in enumerate(write_plan.matches, start=1):
        delta: Optional[int] = None
        if m.new_amount is not None and m.old_amount is not None:
            delta = m.new_amount - m.old_amount
        elif m.new_amount is not None:
            delta = m.new_amount
        rows.append(
            DetailReportRow(
                stt=idx,
                card_no=m.card_no or "",
                account_no=m.account_no or "",
                source_name=m.source_name,
                target_name=m.target_name,
                bank_code=m.bank_code or "",
                source_row=m.source_row,
                target_file=Path(m.target_file).name if m.target_file else "",
                target_sheet=m.target_sheet or "",
                target_row=m.target_row,
                target_cell=m.target_cell or "",
                old_amount=m.old_amount,
                source_amount=m.source_amount,
                new_amount=m.new_amount,
                delta_amount=delta,
                status=m.status.value if isinstance(m.status, MatchStatus) else str(m.status),
                exception_ids=", ".join(m.exception_ids),
                note=m.note,
            )
        )
    return rows


def build_exception_report(exceptions: Sequence[ExceptionItem]) -> list[ExceptionReportRow]:
    """Tạo danh sách `ExceptionReportRow` từ `Exception List`."""
    rows: list[ExceptionReportRow] = []
    for idx, exc in enumerate(exceptions, start=1):
        rows.append(
            ExceptionReportRow(
                stt=idx,
                exception_id=exc.id,
                code=exc.code.value,
                severity=exc.severity.value,
                card_no=exc.card_no or "",
                account_no=exc.account_no or "",
                source_name=exc.source_name,
                target_name=exc.target_name,
                source_row=exc.source_row,
                target_file=Path(exc.target_file).name if exc.target_file else "",
                target_sheet=exc.target_sheet or "",
                target_row=exc.target_row,
                old_amount=exc.old_amount,
                source_amount=exc.source_amount,
                action=exc.action.value if exc.action is not None else "CHƯA XỬ LÝ",
                override_amount=exc.override_amount,
                is_resolved=exc.is_resolved,
                is_blocking=exc.is_blocking,
                message=exc.message,
            )
        )
    return rows


def build_validation_report(
    *,
    write_plan: Optional[WritePlan] = None,
    validation_result: Optional[ValidationResult] = None,
    verification_result: Optional[VerificationResult] = None,
) -> list[ValidationReportRow]:
    """Tổng hợp toàn bộ phát hiện từ `blocking_issues`, `ValidationResult` (Safety) và `VerificationResult` (Verifier + CellDiff)."""
    rows: list[ValidationReportRow] = []
    counter = 0

    if write_plan is not None:
        for issue in write_plan.blocking_issues:
            counter += 1
            rows.append(
                ValidationReportRow(
                    stt=counter,
                    layer="STRUCTURAL",
                    code=issue.code.value,
                    severity=issue.severity.value,
                    file_path="",
                    sheet=issue.sheet or "",
                    row=issue.row,
                    column=issue.column or "",
                    card_no=issue.card_no or "",
                    account_no="",
                    message=issue.message,
                )
            )

    if validation_result is not None:
        for finding in [*validation_result.errors, *validation_result.warnings]:
            counter += 1
            rows.append(
                ValidationReportRow(
                    stt=counter,
                    layer="SAFETY",
                    code=finding.code.value,
                    severity=finding.severity.value,
                    file_path=Path(finding.file_path).name if finding.file_path else "",
                    sheet=finding.sheet or "",
                    row=finding.row,
                    column=finding.column or "",
                    card_no=finding.card_no or "",
                    account_no=finding.account_no or "",
                    message=finding.message,
                )
            )

    if verification_result is not None:
        for v_finding in [*verification_result.errors, *verification_result.warnings]:
            counter += 1
            rows.append(
                ValidationReportRow(
                    stt=counter,
                    layer="VERIFIER",
                    code=v_finding.code.value,
                    severity=v_finding.severity.value,
                    file_path=Path(v_finding.file_path).name if v_finding.file_path else "",
                    sheet=v_finding.sheet or "",
                    row=v_finding.row,
                    column=v_finding.column or "",
                    card_no="",
                    account_no="",
                    message=v_finding.message,
                )
            )
        for diff in verification_result.cell_diffs:
            counter += 1
            rows.append(
                ValidationReportRow(
                    stt=counter,
                    layer="CELL_DIFF",
                    code=(
                        "PLANNED_CELL_MISMATCH"
                        if diff.is_planned_cell
                        else "UNEXPECTED_CELL_MODIFIED"
                    ),
                    severity="ERROR",
                    file_path=Path(diff.file_path).name if diff.file_path else "",
                    sheet=diff.sheet_name,
                    row=diff.row,
                    column=diff.column,
                    card_no="",
                    account_no="",
                    expected_value=diff.expected_value,
                    actual_value=diff.actual_value,
                    message=diff.message,
                )
            )

    return rows


# =============================================================================
# 3. Xuất file báo cáo Excel (.xlsx)
# =============================================================================

_HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F4E78")
_SECTION_FILL = PatternFill(fill_type="solid", fgColor="D9E1F2")
_OK_FILL = PatternFill(fill_type="solid", fgColor="E2EFDA")
_WARN_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")
_ERR_FILL = PatternFill(fill_type="solid", fgColor="FCE4D6")

_HEADER_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
_TITLE_FONT = Font(name="Arial", size=13, bold=True, color="1F4E78")
_SECTION_FONT = Font(name="Arial", size=10, bold=True, color="1F4E78")
_BODY_FONT = Font(name="Arial", size=10)
_BOLD_FONT = Font(name="Arial", size=10, bold=True)

_THIN = Side(border_style="thin", color="BFBFBF")
_GRID_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def export_report_xlsx(report: ReportResult, output_path: PathLike) -> Path:
    """Ghi `ReportResult` ra file Excel `.xlsx` gồm 4 sheet chuẩn."""
    out_p = Path(output_path)
    try:
        out_p.parent.mkdir(parents=True, exist_ok=True)
        wb = openpyxl.Workbook()
        ws_summary = wb.active
        assert ws_summary is not None
        ws_summary.title = SHEET_SUMMARY

        ws_detail = wb.create_sheet(SHEET_DETAIL)
        ws_exc = wb.create_sheet(SHEET_EXCEPTION)
        ws_val = wb.create_sheet(SHEET_VALIDATION)

        _write_summary_sheet(ws_summary, report.summary_report)
        _write_detail_sheet(ws_detail, report.detail_rows)
        _write_exception_sheet(ws_exc, report.exception_report)
        _write_validation_sheet(ws_val, report.validation_report)

        wb.save(out_p)
        wb.close()
        return out_p.resolve()
    except OSError as exc:
        raise ReportError(
            f"Không thể ghi file báo cáo Excel tại '{out_p}'.",
            detail=repr(exc),
        ) from exc


def _write_summary_sheet(ws: Worksheet, s: SummaryReportData) -> None:
    ws["A1"] = "BÁO CÁO TỔNG HỢP ĐỐI CHIẾU & CHI HỘ LƯƠNG"
    ws["A1"].font = _TITLE_FONT

    ws["A3"] = "Chỉ tiêu"
    ws["B3"] = "Giá trị"
    _style_header_row(ws, 3, 2)

    kv_rows: list[tuple[str, Any, Optional[str]]] = [
        ("Thời điểm lập báo cáo", s.created_at, None),
        ("Trạng thái quy trình", s.stage, None),
        ("Kết quả chốt an toàn (Safety)", "ĐẠT" if s.is_safe else "KHÔNG ĐẠT", None),
        ("Kết quả hậu kiểm (Verifier)", "ĐẠT" if s.is_verified else "KHÔNG ĐẠT", None),
        ("File bảng lương nguồn", s.source_file, None),
        ("Sheet bảng lương nguồn", s.source_sheet, None),
        ("Tổng số dòng nhân sự nguồn (có Số thẻ)", s.total_source_records, "#,##0"),
        ("Tổng số dòng khớp file ngân hàng (Matched)", s.total_matched_records, "#,##0"),
        ("Tổng số dòng nguồn chưa khớp đích", s.total_unmatched_source_records, "#,##0"),
        ("Tổng số dòng đích không có trong lương mới (X5)", s.total_unmatched_target_records, "#,##0"),
        ("Tổng số ô thực hiện ghi (WritePlan)", s.total_written_records, "#,##0"),
        ("  - Số dòng lương nguồn được ghi", s.total_source_written_records, "#,##0"),
        ("  - Số dòng đặt 0 / xóa trắng", s.total_zero_or_clear_records, "#,##0"),
        ("Tổng số dòng lương nguồn bỏ qua (Skipped)", s.total_skipped_records, "#,##0"),
        ("Tổng số ngoại lệ (Exception)", s.total_exceptions, "#,##0"),
        ("  - Đã xử lý", s.resolved_exceptions, "#,##0"),
        ("  - Chưa xử lý", s.unresolved_exceptions, "#,##0"),
        ("  - Lỗi chặn cứng (X8)", s.blocking_exceptions, "#,##0"),
        ("Tổng tiền lương nguồn (VND)", s.total_source_amount, "#,##0"),
        ("Tổng tiền đủ điều kiện chuyển khoản (VND)", s.total_source_transferable_amount, "#,##0"),
        ("Tổng tiền kỳ trước trên các ô cập nhật (VND)", s.total_old_target_amount, "#,##0"),
        ("Tổng tiền chuẩn bị ghi / đã ghi (VND)", s.total_planned_write_amount, "#,##0"),
        ("Tổng tiền các dòng đủ ĐK nhưng chọn bỏ qua (VND)", s.total_skipped_transferable_amount, "#,##0"),
        ("Tổng chênh lệch do nhập tay Override (VND)", s.total_override_delta, "#,##0"),
        ("Chênh lệch cân đối số dòng (row_discrepancy)", s.row_discrepancy, "#,##0"),
        ("Chênh lệch cân đối tổng tiền (amount_discrepancy)", s.amount_discrepancy, "#,##0"),
        ("Số lỗi Safety / Số cảnh báo Safety", f"{s.safety_errors_count} lỗi / {s.safety_warnings_count} cảnh báo", None),
        ("Số lỗi Verifier / Số ô lệch hậu kiểm", f"{s.verification_errors_count} lỗi / {s.cell_diffs_count} ô lệch", None),
        ("Danh sách file kết quả đã xuất", "\n".join(s.output_files) if s.output_files else "(Chưa xuất)", None),
        ("Danh sách file sao lưu (Backup)", "\n".join(s.backup_files) if s.backup_files else "(Chưa tạo)", None),
    ]

    r = 4
    for label, val, num_fmt in kv_rows:
        c_label = ws.cell(row=r, column=1, value=label)
        c_val = ws.cell(row=r, column=2, value=val)
        c_label.font = _BODY_FONT
        c_val.font = _BOLD_FONT if "Tổng tiền" in label or "Kết quả" in label else _BODY_FONT
        c_label.border = _GRID_BORDER
        c_val.border = _GRID_BORDER
        if num_fmt and isinstance(val, int):
            c_val.number_format = num_fmt
            c_val.alignment = Alignment(horizontal="right", vertical="center")
        else:
            c_val.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if label in ("Kết quả chốt an toàn (Safety)", "Kết quả hậu kiểm (Verifier)"):
            c_val.fill = _OK_FILL if val == "ĐẠT" else _ERR_FILL
        r += 1

    # Bảng tổng hợp theo từng file ngân hàng đích
    r += 1
    ws.cell(row=r, column=1, value="TỔNG HỢP THEO TỪNG FILE NGÂN HÀNG ĐÍCH").font = _SECTION_FONT
    r += 1
    headers = [
        "STT",
        "File ngân hàng gốc",
        "Sheet cập nhật",
        "Mẫu (Template)",
        "Cột Số tiền",
        "Tổng dòng đích",
        "Số ô ghi",
        "Số ô thay đổi",
        "Tổng tiền cũ (VND)",
        "Tổng tiền mới (VND)",
        "Chênh lệch (VND)",
        "File kết quả mới",
    ]
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=r, column=col_idx, value=h)
    _style_header_row(ws, r, len(headers))

    for idx, ts in enumerate(s.target_summaries, start=1):
        r += 1
        row_vals: list[Any] = [
            idx,
            Path(ts.target_file).name,
            ts.sheet_name,
            ts.template_name,
            ts.amount_column,
            ts.total_target_rows,
            ts.scheduled_updates,
            ts.changed_cells,
            ts.total_old_amount,
            ts.total_new_amount,
            ts.delta_amount,
            Path(ts.output_file).name if ts.output_file else "",
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=r, column=col_idx, value=val)
            cell.font = _BODY_FONT
            cell.border = _GRID_BORDER
            if col_idx in (6, 7, 8, 9, 10, 11):
                cell.number_format = "#,##0"

    _auto_fit_columns(ws)


def _write_detail_sheet(ws: Worksheet, rows: Sequence[DetailReportRow]) -> None:
    headers = [
        "STT",
        "Số thẻ",
        "Số tài khoản",
        "Họ tên (Bảng lương)",
        "Họ tên (File đích)",
        "Mã NH",
        "Dòng nguồn",
        "File đích",
        "Sheet đích",
        "Dòng đích",
        "Ô ghi",
        "Tiền cũ (VND)",
        "Tiền lương nguồn (VND)",
        "Tiền mới ghi (VND)",
        "Chênh lệch (VND)",
        "Trạng thái",
        "Mã ngoại lệ",
        "Ghi chú",
    ]
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=h)
    _style_header_row(ws, 1, len(headers))
    ws.freeze_panes = "A2"

    for r_idx, item in enumerate(rows, start=2):
        vals: list[Any] = [
            item.stt,
            item.card_no,
            item.account_no,
            item.source_name,
            item.target_name,
            item.bank_code,
            item.source_row,
            item.target_file,
            item.target_sheet,
            item.target_row,
            item.target_cell,
            item.old_amount,
            item.source_amount,
            item.new_amount,
            item.delta_amount,
            item.status,
            item.exception_ids,
            item.note,
        ]
        row_fill = _status_fill(item.status)
        for col_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=r_idx, column=col_idx)
            # Cột Số thẻ (2), Số TK (3), Mã NH (6) bắt buộc định dạng Text '@' để giữ số 0 đầu
            if col_idx in (2, 3, 6):
                cell.number_format = "@"
                cell.value = str(val) if val is not None else ""
            else:
                cell.value = val
                if col_idx in (12, 13, 14, 15) and isinstance(val, int):
                    cell.number_format = "#,##0"
            cell.font = _BODY_FONT
            cell.border = _GRID_BORDER
            if row_fill is not None:
                cell.fill = row_fill

    _auto_fit_columns(ws)


def _write_exception_sheet(ws: Worksheet, rows: Sequence[ExceptionReportRow]) -> None:
    headers = [
        "STT",
        "Mã Exception",
        "Loại (X1-X10)",
        "Mức độ",
        "Số thẻ",
        "Số tài khoản",
        "Họ tên nguồn",
        "Họ tên đích",
        "Dòng nguồn",
        "File đích",
        "Sheet đích",
        "Dòng đích",
        "Tiền cũ (VND)",
        "Tiền nguồn (VND)",
        "Phương án xử lý",
        "Tiền ghi đè (VND)",
        "Đã xử lý",
        "Chặn cứng",
        "Nội dung chi tiết",
    ]
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=h)
    _style_header_row(ws, 1, len(headers))
    ws.freeze_panes = "A2"

    for r_idx, item in enumerate(rows, start=2):
        vals: list[Any] = [
            item.stt,
            item.exception_id,
            item.code,
            item.severity,
            item.card_no,
            item.account_no,
            item.source_name,
            item.target_name,
            item.source_row,
            item.target_file,
            item.target_sheet,
            item.target_row,
            item.old_amount,
            item.source_amount,
            item.action,
            item.override_amount,
            "Đã xử lý" if item.is_resolved else "Chưa xử lý",
            "Có" if item.is_blocking else "Không",
            item.message,
        ]
        row_fill = _ERR_FILL if (item.is_blocking or not item.is_resolved) else _WARN_FILL
        for col_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=r_idx, column=col_idx)
            if col_idx in (5, 6):
                cell.number_format = "@"
                cell.value = str(val) if val is not None else ""
            else:
                cell.value = val
                if col_idx in (13, 14, 16) and isinstance(val, int):
                    cell.number_format = "#,##0"
            cell.font = _BODY_FONT
            cell.border = _GRID_BORDER
            cell.fill = row_fill

    _auto_fit_columns(ws)


def _write_validation_sheet(ws: Worksheet, rows: Sequence[ValidationReportRow]) -> None:
    headers = [
        "STT",
        "Lớp kiểm tra",
        "Mã kiểm tra",
        "Mức độ",
        "File",
        "Sheet",
        "Dòng",
        "Cột",
        "Số thẻ",
        "Số tài khoản",
        "Giá trị dự kiến",
        "Giá trị thực tế",
        "Thông báo chi tiết",
    ]
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=h)
    _style_header_row(ws, 1, len(headers))
    ws.freeze_panes = "A2"

    for r_idx, item in enumerate(rows, start=2):
        vals: list[Any] = [
            item.stt,
            item.layer,
            item.code,
            item.severity,
            item.file_path,
            item.sheet,
            item.row,
            item.column,
            item.card_no,
            item.account_no,
            item.expected_value,
            item.actual_value,
            item.message,
        ]
        row_fill = _ERR_FILL if item.severity == "ERROR" else _WARN_FILL
        for col_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=r_idx, column=col_idx)
            if col_idx in (9, 10):
                cell.number_format = "@"
                cell.value = str(val) if val is not None else ""
            else:
                cell.value = val
                if col_idx in (11, 12) and isinstance(val, int):
                    cell.number_format = "#,##0"
            cell.font = _BODY_FONT
            cell.border = _GRID_BORDER
            cell.fill = row_fill

    _auto_fit_columns(ws)


# =============================================================================
# 4. Các hàm phụ trợ nội bộ
# =============================================================================

def _resolve_inputs(
    *,
    run_result: Optional[ReportInput],
    workflow_result: Optional[WorkflowResult],
    verification_result: Optional[VerificationResult],
    validation_result: Optional[ValidationResult],
    write_plan: Optional[WritePlan],
) -> tuple[
    WritePlan,
    Optional[RunResult],
    Optional[WorkflowResult],
    Optional[VerificationResult],
    Optional[ValidationResult],
]:
    eff_plan: Optional[WritePlan] = write_plan
    eff_run: Optional[RunResult] = None
    eff_wf: Optional[WorkflowResult] = workflow_result
    eff_ver: Optional[VerificationResult] = verification_result
    eff_val: Optional[ValidationResult] = validation_result

    if isinstance(run_result, WorkflowResult):
        eff_wf = eff_wf or run_result
    elif isinstance(run_result, RunResult):
        eff_run = run_result
        eff_plan = eff_plan or run_result.write_plan
    elif isinstance(run_result, WritePlan):
        eff_plan = eff_plan or run_result

    if eff_wf is not None:
        eff_plan = eff_plan or eff_wf.write_plan
        eff_run = eff_run or eff_wf.run_result
        eff_val = eff_val or eff_wf.validation_result
        eff_ver = eff_ver or eff_wf.verification_result or eff_wf.pre_verification_result

    if eff_run is not None and eff_plan is None:
        eff_plan = eff_run.write_plan

    if eff_plan is None:
        raise ReportError(
            "Phải cung cấp ít nhất một trong: RunResult, WorkflowResult hoặc WritePlan để tạo báo cáo."
        )

    return eff_plan, eff_run, eff_wf, eff_ver, eff_val


def _build_default_report_path(report_dir: PathLike, prefix: str) -> Path:
    folder = Path(report_dir)
    folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    candidate = folder / f"{prefix}_{stamp}.xlsx"
    idx = 2
    while candidate.exists():
        candidate = folder / f"{prefix}_{stamp}_{idx}.xlsx"
        idx += 1
    return candidate


def _style_header_row(ws: Worksheet, row_idx: int, col_count: int) -> None:
    ws.row_dimensions[row_idx].height = 24.0
    for c in range(1, col_count + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.border = _GRID_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _status_fill(status: str) -> Optional[PatternFill]:
    if status == MatchStatus.WILL_UPDATE.value:
        return _OK_FILL
    if status == MatchStatus.PENDING.value:
        return _WARN_FILL
    if status == MatchStatus.ERROR.value:
        return _ERR_FILL
    return None


def _auto_fit_columns(ws: Worksheet) -> None:
    for col_cells in ws.columns:
        first = col_cells[0]
        col_letter = get_column_letter(first.column or 1)
        max_len = 10
        for c in col_cells[:200]:
            if c.row == 1 and ws.title == SHEET_SUMMARY:
                continue
            val_str = str(c.value) if c.value is not None else ""
            line_max = max((len(line) for line in val_str.splitlines()), default=0)
            if line_max > max_len:
                max_len = line_max
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 55)
