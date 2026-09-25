"""Conversion pipeline: validate → pdf2docx → structure fix → optional LLM."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from app.config import DEFAULT_OUTPUT_SUFFIX
from app.core.detector import DocTypeDetector, PdfInfo
from app.enhance.llm_enhancer import LLMEnhancer
from app.enhance.structure_fix import StructureFixer
from app.engines.pdf2docx_engine import Pdf2DocxEngine

ProgressCallback = Callable[[int, str], None]


@dataclass
class ConversionResult:
    ok: bool
    output_path: Path | None = None
    message: str = ""
    pdf_info: PdfInfo | None = None
    structure_stats: dict = field(default_factory=dict)
    llm_stats: dict = field(default_factory=dict)


class ConversionOrchestrator:
    def __init__(self) -> None:
        self.detector = DocTypeDetector()
        self.engine = Pdf2DocxEngine()
        self.fixer = StructureFixer()
        self.llm = LLMEnhancer()

    def convert(
        self,
        pdf_path: Path | str,
        output_path: Path | str | None = None,
        *,
        enable_llm: bool = False,
        on_progress: ProgressCallback | None = None,
    ) -> ConversionResult:
        def progress(pct: int, msg: str) -> None:
            if on_progress:
                on_progress(pct, msg)

        try:
            progress(5, "正在校验 PDF…")
            info = self.detector.inspect(Path(pdf_path))
            hint = "（检测到应急预案特征）" if info.is_likely_yuyan else ""
            progress(8, f"共 {info.page_count} 页{hint}")

            out = Path(output_path) if output_path else self._default_out(info.path)
            if out.suffix.lower() != ".docx":
                out = out.with_suffix(DEFAULT_OUTPUT_SUFFIX)

            self.engine.convert(info.path, out, on_progress=progress)

            progress(75, "正在应用预案结构规则…")
            structure_stats = self.fixer.apply(out)

            llm_stats: dict = {"skipped": True, "reason": "未启用"}
            if enable_llm:
                progress(85, "正在调用本地大模型做结构增强…")
                llm_stats = self.llm.enhance(out)
                if llm_stats.get("skipped"):
                    progress(90, llm_stats.get("reason", "LLM 已跳过"))

            progress(100, "转换完成")
            return ConversionResult(
                ok=True,
                output_path=out,
                message=f"已保存: {out}",
                pdf_info=info,
                structure_stats=structure_stats,
                llm_stats=llm_stats,
            )
        except Exception as exc:  # noqa: BLE001 — report to UI
            return ConversionResult(ok=False, message=str(exc))

    @staticmethod
    def _default_out(pdf_path: Path) -> Path:
        return pdf_path.with_suffix(DEFAULT_OUTPUT_SUFFIX)
