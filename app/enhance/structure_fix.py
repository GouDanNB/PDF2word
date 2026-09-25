"""Rule-based structure fixes for 应急预案-style DOCX after pdf2docx."""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.shared import Pt

from app.config import HEADING_PATTERNS


class StructureFixer:
    """Promote headings, merge tiny fragments, light table cleanup."""

    def __init__(self) -> None:
        self._heading_res = [re.compile(p) for p in HEADING_PATTERNS]

    def apply(self, docx_path: Path) -> dict:
        path = Path(docx_path)
        doc = Document(str(path))
        stats = {
            "headings_promoted": 0,
            "fragments_merged": 0,
            "empty_removed": 0,
        }

        self._promote_headings(doc, stats)
        self._merge_short_fragments(doc, stats)
        self._normalize_tables(doc)

        doc.save(str(path))
        return stats

    def _promote_headings(self, doc: Document, stats: dict) -> None:
        for para in doc.paragraphs:
            text = (para.text or "").strip()
            if not text or len(text) > 80:
                continue
            level = self._detect_heading_level(text)
            if level is None:
                continue
            style_name = f"Heading {level}"
            try:
                para.style = style_name
            except KeyError:
                # Fallback: bold + size
                for run in para.runs:
                    run.bold = True
                    run.font.size = Pt(18 - (level - 1) * 2)
            stats["headings_promoted"] += 1

    def _detect_heading_level(self, text: str) -> int | None:
        if re.match(r"^第[一二三四五六七八九十百零\d]+[章节编部分]", text):
            return 1
        m = re.match(r"^(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:\.(\d+))?\s+\S", text)
        if m:
            depth = sum(1 for g in m.groups() if g is not None)
            return min(depth, 3)
        if re.match(r"^[（(][一二三四五六七八九十]+[）)]", text):
            return 2
        if re.match(r"^[（(]\d+[）)]", text):
            return 3
        for cre in self._heading_res:
            if cre.match(text):
                return 2
        return None

    def _merge_short_fragments(self, doc: Document, stats: dict) -> None:
        """Merge consecutive 1–2 character body paragraphs into previous."""
        paragraphs = list(doc.paragraphs)
        i = 1
        while i < len(paragraphs):
            prev = paragraphs[i - 1]
            cur = paragraphs[i]
            prev_t = (prev.text or "").strip()
            cur_t = (cur.text or "").strip()
            if (
                cur_t
                and len(cur_t) <= 2
                and prev_t
                and not self._looks_like_heading(prev_t)
                and not self._looks_like_heading(cur_t)
                and self._is_body_style(cur)
                and self._is_body_style(prev)
            ):
                # Append to previous first run or create one
                if prev.runs:
                    prev.runs[-1].text = (prev.runs[-1].text or "") + cur_t
                else:
                    prev.add_run(cur_t)
                self._clear_paragraph(cur)
                stats["fragments_merged"] += 1
            i += 1

    def _normalize_tables(self, doc: Document) -> None:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        # Strip excessive trailing spaces inside cells
                        for run in para.runs:
                            if run.text:
                                run.text = run.text.replace("\u00a0", " ").strip()

    @staticmethod
    def _looks_like_heading(text: str) -> bool:
        return bool(
            re.match(r"^第[一二三四五六七八九十百零\d]+[章节]", text)
            or re.match(r"^\d+(\.\d+)+\s+", text)
        )

    @staticmethod
    def _is_body_style(para) -> bool:
        name = (para.style.name if para.style else "") or ""
        return not name.startswith("Heading")

    @staticmethod
    def _clear_paragraph(para) -> None:
        for run in para.runs:
            run.text = ""
