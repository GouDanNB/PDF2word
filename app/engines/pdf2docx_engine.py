"""pdf2docx conversion engine with progress callbacks."""

from __future__ import annotations

from pathlib import Path
from typing import Callable

from pdf2docx import Converter

ProgressCallback = Callable[[int, str], None]


class Pdf2DocxEngine:
    """Wrap pdf2docx Converter for layout-preserving PDF → DOCX."""

    def convert(
        self,
        pdf_path: Path,
        docx_path: Path,
        *,
        start: int = 0,
        end: int | None = None,
        on_progress: ProgressCallback | None = None,
    ) -> Path:
        pdf_path = Path(pdf_path)
        docx_path = Path(docx_path)
        docx_path.parent.mkdir(parents=True, exist_ok=True)

        if on_progress:
            on_progress(10, "正在打开 PDF…")

        cv = Converter(str(pdf_path))
        try:
            if on_progress:
                on_progress(35, "正在解析版面并写入 Word…")
            cv.convert(str(docx_path), start=start, end=end)
        finally:
            cv.close()

        if on_progress:
            on_progress(70, "基础转换完成")
        return docx_path
