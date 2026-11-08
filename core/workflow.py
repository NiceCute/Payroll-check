"""
workflow.py - Lớp điều phối quy trình (Workflow Layer) cho ChiHoLuongTool.

Pipeline chuẩn:
    Source Reader
    → Template Store (nhận diện mẫu & Target Reader)
    → Reconciler
    → Safety
    → Verifier (kiểm tra tính nhất quán Source -> Reconcile -> WritePlan)
    → Excel Writer (sao lưu file gốc & ghi file kết quả mới)
    → Verifier (hậu kiểm diff từng ô trên file kết quả mới; tự động thu hồi file mới nếu sai lệch)

Nguyên tắc:
- Chỉ điều phối các module trong `core`, quản lý trạng thái chạy (`WorkflowStage`) và luân chuyển `Exception List`.
- Dừng ngay lập tức khi `Safety` hoặc `Verifier` thất bại.
- Không chứa logic nghiệp vụ riêng, không chứa GUI, không chứa code ghi log chi tiết.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping, Optional, Sequence, Union, cast

from .excel_reader import PathLike
from .excel_writer import (
    DEFAULT_BACKUP_DIR,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_OUTPUT_SUFFIX,
    WriterEngine,
    execute_write_plan,
)
from .models import (
    ConfigError,
    ExceptionItem,
    ResolutionAction,
    RunResult,
    RuntimeConfig,
    SafetyError,
    SourceProfile,
    SourceReadResult,
    TargetReadResult,
    TargetTemplate,
    ToolError,
    ValidationResult,
    VerificationError,
    VerificationResult,
    WorkflowConfig,
    WorkflowError,
    WorkflowResult,
    WorkflowStage,
    WorkflowState,
    WritePlan,
)
from .reconciler import (
    ResolutionValue,
    TargetInput,
    apply_resolutions as reconcile_apply_resolutions,
    reconcile,
    resolve_exception as reconcile_resolve_exception,
)
from .safety import validate_write_plan
from .source_reader import DEFAULT_SOURCE_PROFILE, read_source
from .target_reader import DEFAULT_TARGET_TEMPLATES, read_target
from .template_store import TemplateStore, match_templates
from .verifier import verify

__all__ = [
    "RuntimeConfig",
    "Workflow",
    "WorkflowConfig",
    "WorkflowError",
    "WorkflowResult",
    "WorkflowStage",
    "WorkflowState",
    "execute_workflow",
    "prepare_workflow",
    "run_workflow",
]

SourceInput = Union[PathLike, SourceReadResult]
SingleTargetSpec = Union[
    PathLike,
    TargetReadResult,
    TargetInput,
    tuple[PathLike, TargetTemplate],
    tuple[TargetReadResult, TargetTemplate],
]
TemplateFileInput = Union[SingleTargetSpec, Sequence[SingleTargetSpec]]


# =============================================================================
# 1. Bộ điều phối quy trình (Workflow Class)
# =============================================================================

class Workflow:
    """Điều phối 1 phiên chạy từ lúc đọc file đến khi xuất kết quả và hậu kiểm."""

    def __init__(self, config: Optional[RuntimeConfig] = None) -> None:
        self.config: RuntimeConfig = config or RuntimeConfig()
        self._result: WorkflowResult = WorkflowResult()

    @property
    def stage(self) -> WorkflowStage:
        return self._result.stage

    @property
    def state(self) -> WorkflowStage:
        return self._result.stage

    @property
    def result(self) -> WorkflowResult:
        return self._result

    @property
    def run_result(self) -> Optional[RunResult]:
        return self._result.run_result

    @property
    def exceptions(self) -> list[ExceptionItem]:
        return self._result.exceptions

    # ------------------------------------------------------------------ Bước 1 & 2: Chuẩn bị & Đối chiếu (Preview)
    def prepare(
        self,
        source_file: SourceInput,
        template_file: Optional[TemplateFileInput] = None,
        config: Optional[RuntimeConfig] = None,
        *,
        target_files: Optional[TemplateFileInput] = None,
        resolutions: Optional[Mapping[str, ResolutionValue]] = None,
    ) -> WorkflowResult:
        """Thực thi chặng chuẩn bị:
        Source Reader → Template Store (+ Target Reader) → Reconciler → kiểm tra sơ bộ Safety & Verifier.
        """
        if config is not None:
            self.config = config
        cfg = self.config
        self._result = WorkflowResult(stage=WorkflowStage.IDLE)

        eff_targets = template_file if template_file is not None else target_files
        if eff_targets is None:
            return self._fail("Phải cung cấp ít nhất 1 file mẫu ngân hàng (template_file / target_files).")

        try:
            # 1. Khởi tạo Template Store & nạp cấu hình mapping
            store = self._init_template_store(cfg)
            src_profile = self._resolve_source_profile(store, cfg)
            self._result.source_profile = src_profile
            self._result.stage = WorkflowStage.TEMPLATES_LOADED

            # 2. Đọc file lương nguồn (Source Reader)
            if isinstance(source_file, SourceReadResult):
                src_res = source_file
            else:
                src_res = read_source(
                    source_file,
                    profile=src_profile,
                    sheet_name=cfg.source_sheet,
                )
            self._result.source_result = src_res
            self._result.stage = WorkflowStage.SOURCE_READ

            # 3. Nhận diện mẫu qua Template Store & đọc các file ngân hàng đích (Target Reader)
            target_pairs, unmatched = self._read_targets(eff_targets, store, cfg)
            self._result.unmatched_target_files = unmatched
            self._result.target_results = [r for r, _ in target_pairs]
            self._result.target_templates = [t for _, t in target_pairs]

            if unmatched:
                names = ", ".join(Path(p).name for p in unmatched)
                return self._fail(
                    f"Không tự nhận diện được mẫu ngân hàng duy nhất cho: {names}. "
                    "Vui lòng chọn hoặc tạo mapping phù hợp."
                )
            if not target_pairs:
                return self._fail("Không có file ngân hàng hợp lệ nào để đối chiếu.")

            self._result.stage = WorkflowStage.TARGETS_READ

            # 4. Đối chiếu (Reconciler) -> WritePlan + Exception List
            plan = reconcile(
                src_res,
                target_pairs,
                master=cfg.master_records,
            )
            self._result.write_plan = plan
            self._result.exceptions = list(plan.exceptions)
            self._result.stage = WorkflowStage.RECONCILED

            # 5. Áp dụng quyết định xử lý Exception (nếu được truyền sẵn)
            merged_resolutions: dict[str, ResolutionValue] = {}
            if cfg.resolutions:
                merged_resolutions.update(cfg.resolutions)
            if resolutions:
                merged_resolutions.update(resolutions)
            if merged_resolutions:
                reconcile_apply_resolutions(plan, merged_resolutions)
                self._result.exceptions = list(plan.exceptions)

            # 6. Cập nhật trạng thái dựa trên Exception List & kiểm tra sơ bộ
            self._refresh_validation_and_pre_verify(check_locks=False)
            if self._result.has_unresolved_exceptions:
                self._result.stage = WorkflowStage.AWAITING_RESOLUTION

            return self._result

        except ToolError as exc:
            return self._fail(exc.user_message, exc=exc)

    # ------------------------------------------------------------------ Luân chuyển & xử lý Exception List
    def resolve_exception(
        self,
        exception_id: str,
        action: Union[ResolutionAction, str],
        *,
        override_amount: Optional[int] = None,
    ) -> WorkflowResult:
        """Cập nhật phương án xử lý cho 1 mục trong `Exception List` và kiểm tra lại kế hoạch."""
        plan = self._require_plan()
        act = ResolutionAction(action) if isinstance(action, str) else action
        reconcile_resolve_exception(
            plan,
            exception_id,
            act,
            override_amount=override_amount,
        )
        self._result.exceptions = list(plan.exceptions)
        self._refresh_validation_and_pre_verify(check_locks=False)
        return self._result

    def apply_resolutions(
        self,
        resolutions: Mapping[str, ResolutionValue],
    ) -> WorkflowResult:
        """Áp dụng hàng loạt quyết định xử lý cho `Exception List` và kiểm tra lại kế hoạch."""
        plan = self._require_plan()
        reconcile_apply_resolutions(plan, resolutions)
        self._result.exceptions = list(plan.exceptions)
        self._refresh_validation_and_pre_verify(check_locks=False)
        return self._result

    # ------------------------------------------------------------------ Kiểm tra Safety & Verifier độc lập
    def validate(
        self,
        *,
        check_locks: bool = True,
        output_paths: Optional[Any] = None,
    ) -> ValidationResult:
        """Thực hiện chốt kiểm tra an toàn (`core.safety`)."""
        plan = self._require_plan()
        val_res = validate_write_plan(
            plan,
            exceptions=self._result.exceptions,
            run_result=self._result.run_result,
            check_disk=check_locks,
        )
        self._result.validation_result = val_res
        self._result.stage = (
            WorkflowStage.SAFETY_CHECKED if val_res.is_safe else WorkflowStage.SAFETY_FAILED
        )
        return val_res

    def verify_plan(self) -> VerificationResult:
        """Kiểm tra tính nhất quán `Source -> Reconcile -> WritePlan` trước khi ghi (`core.verifier`)."""
        plan = self._require_plan()
        pre_ver = verify(
            run_result=plan,
            validation_result=self._result.validation_result,
            source=self._result.source_result,
            verify_files_on_disk=False,
        )
        self._result.pre_verification_result = pre_ver
        return pre_ver

    # ------------------------------------------------------------------ Bước 3: Safety -> Verifier -> Excel Writer -> Verifier
    def execute(
        self,
        *,
        output_paths: Optional[Any] = None,
    ) -> WorkflowResult:
        """Thực thi chặng xuất kết quả:
        Safety → Verifier (tiền kiểm) → Excel Writer → Verifier (hậu kiểm diff trên đĩa).
        Dừng ngay lập tức nếu Safety hoặc Verifier thất bại.
        """
        plan = self._require_plan()
        cfg = self.config
        eff_out_paths = output_paths if output_paths is not None else cfg.output_paths

        # 1. Chốt chặn Safety trước khi ghi
        val_res = self.validate(
            check_locks=not cfg.dry_run,
            output_paths=eff_out_paths,
        )
        if not val_res.is_safe:
            self._result.stage = WorkflowStage.SAFETY_FAILED
            self._result.stopped_reason = (
                val_res.error_messages[0]
                if val_res.error_messages
                else "Chốt kiểm tra an toàn (Safety) không đạt."
            )
            self._result.run_result = RunResult(write_plan=plan, verified=False)
            if cfg.raise_on_failure:
                val_res.ensure_safe()
            return self._result

        # 2. Kiểm tra tính nhất quán chuỗi Source -> Reconcile -> WritePlan (Verifier)
        pre_ver = self.verify_plan()
        if not pre_ver.is_verified:
            self._result.stage = WorkflowStage.VERIFICATION_FAILED
            self._result.stopped_reason = (
                pre_ver.error_messages[0]
                if pre_ver.error_messages
                else "Kiểm tra tính nhất quán (Verifier) không đạt."
            )
            self._result.run_result = RunResult(write_plan=plan, verified=False)
            if cfg.raise_on_failure:
                pre_ver.ensure_verified()
            return self._result

        self._result.stage = WorkflowStage.PRE_VERIFIED

        # Nếu chạy chế độ dry_run: dừng sau khi Safety & Pre-Verifier đã đạt
        if cfg.dry_run:
            self._result.verification_result = pre_ver
            self._result.run_result = RunResult(write_plan=plan, verified=True)
            return self._result

        # 3. Thực thi ghi file mới & sao lưu file gốc (Excel Writer)
        try:
            exec_res = execute_write_plan(
                write_plan=plan,
                validation_result=val_res,
                template_profile=self._result.target_templates or None,
                target_workbook=self._result.target_results or None,
                output_dir=cfg.output_dir or DEFAULT_OUTPUT_DIR,
                backup_dir=cfg.backup_dir or DEFAULT_BACKUP_DIR,
                output_paths=eff_out_paths,
                output_suffix=cfg.output_suffix or DEFAULT_OUTPUT_SUFFIX,
                overwrite_output=cfg.overwrite_output,
                backup_source_file=cfg.backup_source_file,
                engine=cast(WriterEngine, cfg.writer_engine),
            )
        except ToolError as exc:
            self._result.run_result = RunResult(write_plan=plan, verified=False)
            return self._fail(exc.user_message, exc=exc)

        self._result.write_execution_result = exec_res
        self._result.stage = WorkflowStage.WRITTEN
        run_res = exec_res.to_run_result(verified=False)
        self._result.run_result = run_res

        # 4. Hậu kiểm file kết quả mới trên đĩa (Verifier)
        post_ver = verify(
            run_result=run_res,
            validation_result=val_res,
            write_execution_result=exec_res,
            source=self._result.source_result,
            verify_files_on_disk=True,
        )
        self._result.verification_result = post_ver

        if not post_ver.is_verified:
            self._result.stage = WorkflowStage.VERIFICATION_FAILED
            self._result.stopped_reason = (
                post_ver.error_messages[0]
                if post_ver.error_messages
                else "Hậu kiểm file kết quả (Verifier) không đạt."
            )
            run_res.verified = False

            # Theo sơ đồ mục 4 tài liệu V1: hậu kiểm thất bại -> xóa file mới vừa tạo + báo lỗi
            if cfg.cleanup_output_on_verify_failure:
                rolled_back: list[str] = []
                for out_file in list(exec_res.output_files):
                    out_p = Path(out_file)
                    if out_p.exists():
                        try:
                            out_p.unlink()
                            rolled_back.append(str(out_p))
                        except OSError:
                            pass
                self._result.rolled_back_files = rolled_back
                run_res.output_files.clear()

            if cfg.raise_on_failure:
                post_ver.ensure_verified()
            return self._result

        # 5. Hoàn tất toàn trình
        run_res.verified = True
        self._result.stage = WorkflowStage.COMPLETED
        self._result.stopped_reason = None
        return self._result

    # ------------------------------------------------------------------ Chạy trọn gói 1 lần (Full Pipeline)
    def run(
        self,
        source_file: SourceInput,
        template_file: Optional[TemplateFileInput] = None,
        config: Optional[RuntimeConfig] = None,
        *,
        target_files: Optional[TemplateFileInput] = None,
        resolutions: Optional[Mapping[str, ResolutionValue]] = None,
        output_paths: Optional[Any] = None,
    ) -> WorkflowResult:
        """Chạy trọn gói toàn bộ pipeline:
        Source Reader → Template Store → Reconciler → Safety → Verifier → Excel Writer → Verifier.
        """
        prep = self.prepare(
            source_file=source_file,
            template_file=template_file,
            config=config,
            target_files=target_files,
            resolutions=resolutions,
        )
        if prep.stage == WorkflowStage.FAILED:
            return prep
        return self.execute(output_paths=output_paths)

    # =========================================================================
    # Các hàm điều phối nội bộ
    # =========================================================================

    def _init_template_store(self, cfg: RuntimeConfig) -> TemplateStore:
        store = TemplateStore(cfg.templates_dir)
        if cfg.auto_create_default_templates:
            store.ensure_defaults()
        return store

    def _resolve_source_profile(
        self,
        store: TemplateStore,
        cfg: RuntimeConfig,
    ) -> SourceProfile:
        if cfg.source_profile is not None:
            return cfg.source_profile
        if cfg.source_profile_name:
            return store.get_source(cfg.source_profile_name)
        stored_sources = store.list_sources()
        if stored_sources:
            return stored_sources[0].item
        return DEFAULT_SOURCE_PROFILE

    def _read_targets(
        self,
        target_spec: TemplateFileInput,
        store: TemplateStore,
        cfg: RuntimeConfig,
    ) -> tuple[list[tuple[TargetReadResult, TargetTemplate]], list[str]]:
        items = _normalize_target_items(target_spec)
        available_templates = [s.item for s in store.list_targets()] or list(DEFAULT_TARGET_TEMPLATES)

        explicit_by_key, explicit_seq = _parse_explicit_target_templates(
            cfg.target_templates,
            store,
        )

        pairs: list[tuple[TargetReadResult, TargetTemplate]] = []
        unmatched: list[str] = []

        for idx, item in enumerate(items):
            if isinstance(item, TargetInput):
                tmpl = item.template or self._find_template_by_name(
                    item.target.template_name,
                    available_templates,
                )
                pairs.append((item.target, tmpl))
                continue

            if isinstance(item, tuple) and len(item) == 2:
                first, second = item
                if isinstance(second, TargetTemplate):
                    if isinstance(first, TargetReadResult):
                        pairs.append((first, second))
                    else:
                        pairs.append((read_target(first, second), second))
                    continue

            if isinstance(item, TargetReadResult):
                tmpl = self._resolve_template_for_read_result(
                    item,
                    idx,
                    explicit_by_key,
                    explicit_seq,
                    available_templates,
                )
                pairs.append((item, tmpl))
                continue

            # item là đường dẫn file đích (.xls / .xlsx)
            path = Path(str(item))
            tmpl_opt = self._match_template_for_path(
                path,
                idx,
                explicit_by_key,
                explicit_seq,
                available_templates,
            )
            if tmpl_opt is None:
                unmatched.append(str(path))
                continue

            t_res = read_target(path, tmpl_opt)
            pairs.append((t_res, tmpl_opt))

        return pairs, unmatched

    def _resolve_template_for_read_result(
        self,
        res: TargetReadResult,
        idx: int,
        explicit_by_key: dict[str, TargetTemplate],
        explicit_seq: list[TargetTemplate],
        available_templates: list[TargetTemplate],
    ) -> TargetTemplate:
        for key in (
            str(Path(res.file_path).resolve()).casefold(),
            str(Path(res.file_path)).casefold(),
            Path(res.file_path).name.casefold(),
            res.template_name.strip().casefold(),
        ):
            if key in explicit_by_key:
                return explicit_by_key[key]
        if idx < len(explicit_seq):
            return explicit_seq[idx]
        return self._find_template_by_name(res.template_name, available_templates)

    def _match_template_for_path(
        self,
        path: Path,
        idx: int,
        explicit_by_key: dict[str, TargetTemplate],
        explicit_seq: list[TargetTemplate],
        available_templates: list[TargetTemplate],
    ) -> Optional[TargetTemplate]:
        for key in (
            str(path.resolve()).casefold(),
            str(path).casefold(),
            path.name.casefold(),
        ):
            if key in explicit_by_key:
                return explicit_by_key[key]
        if idx < len(explicit_seq):
            return explicit_seq[idx]

        matches = match_templates(path, available_templates)
        if len(matches) == 1:
            return matches[0].template
        by_name = [m for m in matches if m.by_file_name]
        if len(by_name) == 1:
            return by_name[0].template
        return None

    def _find_template_by_name(
        self,
        name: str,
        available_templates: list[TargetTemplate],
    ) -> TargetTemplate:
        key = name.strip().casefold()
        for t in available_templates:
            if t.name.strip().casefold() == key:
                return t
        for t in DEFAULT_TARGET_TEMPLATES:
            if t.name.strip().casefold() == key:
                return t
        raise ConfigError(f"Không tìm thấy mẫu ngân hàng '{name}'.")

    def _refresh_validation_and_pre_verify(self, *, check_locks: bool) -> None:
        plan = self._require_plan()
        val_res = validate_write_plan(
            plan,
            exceptions=self._result.exceptions,
            run_result=self._result.run_result,
            check_disk=check_locks,
        )
        self._result.validation_result = val_res
        pre_ver = verify(
            run_result=plan,
            validation_result=val_res,
            source=self._result.source_result,
            verify_files_on_disk=False,
        )
        self._result.pre_verification_result = pre_ver
        if val_res.is_safe and pre_ver.is_verified:
            self._result.stage = WorkflowStage.PRE_VERIFIED
        elif self._result.has_unresolved_exceptions:
            self._result.stage = WorkflowStage.AWAITING_RESOLUTION
        else:
            self._result.stage = WorkflowStage.SAFETY_FAILED

    def _require_plan(self) -> WritePlan:
        if self._result.write_plan is None:
            raise WorkflowError("Chưa có WritePlan. Hãy gọi prepare() trước.")
        return self._result.write_plan

    def _fail(self, reason: str, *, exc: Optional[ToolError] = None) -> WorkflowResult:
        self._result.stage = WorkflowStage.FAILED
        self._result.stopped_reason = reason
        if self.config.raise_on_failure:
            if isinstance(exc, (SafetyError, VerificationError, WorkflowError)):
                raise exc
            raise WorkflowError(reason, detail=exc.detail if exc else "")
        return self._result


# =============================================================================
# 2. Các hàm tiện ích cấp Module
# =============================================================================

def prepare_workflow(
    source_file: SourceInput,
    template_file: Optional[TemplateFileInput] = None,
    config: Optional[RuntimeConfig] = None,
    *,
    target_files: Optional[TemplateFileInput] = None,
    resolutions: Optional[Mapping[str, ResolutionValue]] = None,
) -> WorkflowResult:
    """Chuẩn bị dữ liệu và đối chiếu (dùng cho bước Preview trên GUI)."""
    wf = Workflow(config=config)
    return wf.prepare(
        source_file=source_file,
        template_file=template_file,
        config=config,
        target_files=target_files,
        resolutions=resolutions,
    )


def execute_workflow(
    workflow_or_result: Union[Workflow, WorkflowResult],
    *,
    config: Optional[RuntimeConfig] = None,
    resolutions: Optional[Mapping[str, ResolutionValue]] = None,
    output_paths: Optional[Any] = None,
) -> WorkflowResult:
    """Tiếp tục thực thi ghi file và hậu kiểm từ `Workflow` hoặc `WorkflowResult` đã chuẩn bị."""
    if isinstance(workflow_or_result, Workflow):
        wf = workflow_or_result
        if config is not None:
            wf.config = config
    else:
        wf = Workflow(config=config)
        wf._result = workflow_or_result

    if resolutions:
        wf.apply_resolutions(resolutions)

    return wf.execute(output_paths=output_paths)


def run_workflow(
    source_file: SourceInput,
    template_file: Optional[TemplateFileInput] = None,
    config: Optional[RuntimeConfig] = None,
    *,
    target_files: Optional[TemplateFileInput] = None,
    resolutions: Optional[Mapping[str, ResolutionValue]] = None,
    output_paths: Optional[Any] = None,
) -> WorkflowResult:
    """Chạy toàn trình từ đọc file đến xuất kết quả và hậu kiểm."""
    wf = Workflow(config=config)
    return wf.run(
        source_file=source_file,
        template_file=template_file,
        config=config,
        target_files=target_files,
        resolutions=resolutions,
        output_paths=output_paths,
    )


# =============================================================================
# 3. Chuẩn hóa tham số đầu vào
# =============================================================================

def _normalize_target_items(spec: TemplateFileInput) -> list[SingleTargetSpec]:
    if isinstance(spec, (str, Path, TargetReadResult, TargetInput)):
        return [spec]
    if isinstance(spec, tuple):
        if len(spec) == 2 and isinstance(spec[1], TargetTemplate):
            return [cast(SingleTargetSpec, spec)]
        return [cast(SingleTargetSpec, x) for x in spec]
    if isinstance(spec, Sequence):
        return [cast(SingleTargetSpec, x) for x in spec]
    return [spec]


def _parse_explicit_target_templates(
    raw: Optional[Any],
    store: TemplateStore,
) -> tuple[dict[str, TargetTemplate], list[TargetTemplate]]:
    by_key: dict[str, TargetTemplate] = {}
    seq: list[TargetTemplate] = []
    if raw is None:
        return by_key, seq

    def resolve_one(val: Any) -> TargetTemplate:
        if isinstance(val, TargetTemplate):
            return val
        if isinstance(val, str):
            return store.get_target(val)
        raise ConfigError(f"Cấu hình mẫu đích không hợp lệ: {val!r}")

    if isinstance(raw, TargetTemplate):
        t = raw
        by_key[t.name.strip().casefold()] = t
        seq.append(t)
    elif isinstance(raw, str):
        t = resolve_one(raw)
        by_key[t.name.strip().casefold()] = t
        seq.append(t)
    elif isinstance(raw, Mapping):
        for k, v in raw.items():
            t = resolve_one(v)
            k_str = str(k)
            by_key[k_str.strip().casefold()] = t
            by_key[Path(k_str).name.casefold()] = t
            try:
                by_key[str(Path(k_str).resolve()).casefold()] = t
            except OSError:
                pass
    elif isinstance(raw, Sequence):
        for item in raw:
            t = resolve_one(item)
            by_key[t.name.strip().casefold()] = t
            seq.append(t)

    return by_key, seq
