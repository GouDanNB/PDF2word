"""PDF validation and light document-type hints for 应急预案."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pymupdf

from app.config import MAX_PAGES


@dataclass
class PdfInfo:
    path: Path
    page_count: int
    encrypted: bool
    is_likely_yuyan: bool
    sample_text: str


class DocTypeDetector:
    """Inspect PDF metadata / first-page text for conversion planning."""

    YUYAN_KEYWORDS = (
        "应急预案",
        "突发事件",
        "环境应急",
        "事故应急",
        "应急响应",
        "组织机构",
        "应急物资",
    )

    def inspect(self, pdf_path: Path) -> PdfInfo:
        path = Path(pdf_path)
        if not path.exists():
            raise FileNotFoundError(f"找不到文件: {path}")
        if path.suffix.lower() != ".pdf":
            raise ValueError("请选择 PDF 文件")

        doc = pymupdf.open(path)
        try:
            encrypted = bool(doc.is_encrypted)
            if encrypted and not doc.authenticate(""):
                raise ValueError("PDF 已加密且无法打开，请先解密")

            page_count = doc.page_count
            if page_count <= 0:
                raise ValueError("PDF 无有效页面")
            if page_count > MAX_PAGES:
                raise ValueError(
                    f"页数 {page_count} 超过上限 {MAX_PAGES}，请拆分后再转换"
                )

            sample_parts: list[str] = []
            for i in range(min(3, page_count)):
                sample_parts.append(doc.load_page(i).get_text("text") or "")
            sample_text = "\n".join(sample_parts)
            is_yuyan = any(k in sample_text for k in self.YUYAN_KEYWORDS)
            return PdfInfo(
                path=path,
                page_count=page_count,
                encrypted=encrypted,
                is_likely_yuyan=is_yuyan,
                sample_text=sample_text[:2000],
            )
        finally:
            doc.close()
