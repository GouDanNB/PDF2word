"""Rule-based structure fixes for 应急预案-style DOCX after pdf2docx."""

from __future__ import annotations

import copy
import re
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph

from app.config import (
    HEADING_AUTO_NUMBER,
    HEADING_CN_PAREN_MAX_LEN,
    HEADING_PATTERNS,
    HEADING_STRIP_PREFIX,
)
from app.enhance.heading_numbering import apply_heading_number, ensure_heading_numbering

# Soft line breaks / newlines that pdf2docx often leaves inside one <w:p>
_LINE_SPLIT_RE = re.compile(r"[\n\r\x0b]+")

# Split mid-paragraph only on clear heading starts (avoid chopping body like「见 2 条」)
_EMBEDDED_HEADING_RE = re.compile(
    r"(?<=\S)\s*(?="
    r"第[一二三四五六七八九十百零\d]+[章节编部分]"
    r"|\d+(?:\.\d+){1,3}\s+\S"
    r"|[（(][一二三四五六七八九十\d]+[）)]\s*\S"
    r")"
)

# Prefixes stripped when Word auto-numbering is enabled
_HEADING_PREFIX_RE = re.compile(
    r"^(?:"
    r"第[一二三四五六七八九十百零\d]+[章节编部分]\s*"
    r"|\d+(?:\.\d+){0,3}(?:\s+|(?=[\u4e00-\u9fffA-Za-z]))"
    r"|[（(][一二三四五六七八九十\d]+[）)]\s*"
    r")"
)


class StructureFixer:
    """Promote headings, merge tiny fragments, light table cleanup."""

    def __init__(self) -> None:
        self._heading_res = [re.compile(p) for p in HEADING_PATTERNS]

    def apply(self, docx_path: Path) -> dict:
        path = Path(docx_path)
        doc = Document(str(path))
        stats = {
            "paragraphs_split": 0,
            "headings_promoted": 0,
            "headings_numbered": 0,
            "prefixes_stripped": 0,
            "fragments_merged": 0,
            "empty_removed": 0,
        }

        # pdf2docx often packs title + body into one paragraph; split first
        # so Heading styles / outline levels apply per line.
        self._split_packed_paragraphs(doc, stats)

        num_id = None
        if HEADING_AUTO_NUMBER:
            try:
                num_id = ensure_heading_numbering(doc)
            except Exception:  # noqa: BLE001 — numbering is best-effort
                num_id = None

        self._promote_headings(doc, stats, num_id=num_id)
        self._merge_short_fragments(doc, stats)
        self._normalize_tables(doc)

        doc.save(str(path))
        return stats

    def _split_packed_paragraphs(self, doc: Document, stats: dict) -> None:
        for para in list(doc.paragraphs):
            lines = self._extract_logical_lines(para.text or "")
            if len(lines) <= 1:
                continue

            # Capture formatting BEFORE wiping runs — critical for Chinese fonts.
            run_formats = self._collect_run_formats(para)
            default_r_pr = run_formats[0][1] if run_formats else None
            p_pr_template = self._clone_p_pr(para)

            self._replace_paragraph_text(
                para, lines[0], self._r_pr_for_line(lines[0], run_formats, default_r_pr)
            )
            anchor = para
            for line in lines[1:]:
                r_pr = self._r_pr_for_line(line, run_formats, default_r_pr)
                anchor = self._insert_paragraph_after(
                    anchor, line, r_pr=r_pr, p_pr=p_pr_template
                )
            stats["paragraphs_split"] += len(lines) - 1

    def _extract_logical_lines(self, text: str) -> list[str]:
        raw = text.strip()
        if not raw:
            return []
        parts: list[str] = []
        for chunk in _LINE_SPLIT_RE.split(raw):
            chunk = chunk.strip()
            if not chunk:
                continue
            parts.extend(self._split_embedded_headings(chunk))
        return [p for p in parts if p]

    def _split_embedded_headings(self, text: str) -> list[str]:
        """Split '…正文 1.1 标题' into separate lines when safe."""
        pieces = _EMBEDDED_HEADING_RE.split(text)
        out: list[str] = []
        for piece in pieces:
            piece = piece.strip()
            if not piece:
                continue
            if out and self._detect_heading_level(piece) is None and len(piece) < 4:
                out[-1] = f"{out[-1]} {piece}".strip()
            else:
                out.append(piece)
        return out or [text]

    def _promote_headings(
        self, doc: Document, stats: dict, *, num_id: int | None
    ) -> None:
        for para in doc.paragraphs:
            text = (para.text or "").strip()
            if not text or len(text) > 80:
                continue
            level = self._detect_heading_level(text)
            if level is None:
                continue

            if HEADING_STRIP_PREFIX and HEADING_AUTO_NUMBER and num_id is not None:
                stripped = self._strip_heading_prefix(text)
                if stripped and stripped != text:
                    self._rewrite_para_keep_format(para, stripped)
                    stats["prefixes_stripped"] += 1
                    text = stripped
                if not text:
                    continue

            if self._apply_heading_style(para, level):
                stats["headings_promoted"] += 1
                if num_id is not None:
                    apply_heading_number(para, level, num_id)
                    stats["headings_numbered"] += 1

    @staticmethod
    def _strip_heading_prefix(text: str) -> str:
        stripped = _HEADING_PREFIX_RE.sub("", text, count=1).strip()
        # Keep original if stripping would leave nothing useful
        return stripped if stripped else text

    def _rewrite_para_keep_format(self, para: Paragraph, new_text: str) -> None:
        """Replace paragraph text while preserving the first run's rPr."""
        r_pr = None
        for run in para.runs:
            if (run.text or "").strip():
                existing = run._element.find(qn("w:rPr"))
                if existing is not None:
                    r_pr = copy.deepcopy(existing)
                break
        self._replace_paragraph_text(para, new_text, r_pr)

    def _apply_heading_style(self, para: Paragraph, level: int) -> bool:
        """Set Word heading style + outlineLvl, keep run-level Chinese fonts."""
        # Snapshot run fonts before style change (some styles reset char props)
        saved = [
            (run.text, self._clone_element(run._element.find(qn("w:rPr"))))
            for run in para.runs
            if (run.text or "").strip()
        ]

        style_candidates = [
            f"Heading {level}",
            f"标题 {level}",
        ]
        applied = False
        for name in style_candidates:
            try:
                para.style = name
                applied = True
                break
            except (KeyError, ValueError):
                continue

        if not applied:
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(18 - (level - 1) * 2)

        # Restore east-asia / size / etc. so正文观感接近 pdf2docx 原结果
        if saved:
            for run in para.runs:
                if not (run.text or "").strip():
                    continue
                for text, r_pr in saved:
                    if run.text == text and r_pr is not None:
                        self._set_run_r_pr(run, r_pr)
                        break

        self._set_outline_level(para, level)
        return True

    @staticmethod
    def _set_outline_level(para: Paragraph, level: int) -> None:
        """Word outlineLvl is 0-based (Heading 1 → 0)."""
        p_pr = para._p.get_or_add_pPr()
        outline = p_pr.find(qn("w:outlineLvl"))
        if outline is None:
            outline = OxmlElement("w:outlineLvl")
            p_pr.append(outline)
        outline.set(qn("w:val"), str(max(0, min(level, 9) - 1)))

    def _detect_heading_level(self, text: str) -> int | None:
        """Map PDF-like heading lines to outline levels; skip body list items."""
        if re.match(r"^第[一二三四五六七八九十百零\d]+[章节编部分]", text):
            return 1

        # Require at least one dotted level: 1.1 → 2, 1.1.1 → 3 (not bare「1 xxx」)
        num = re.match(r"^(\d+(?:\.\d+){1,3})(?:\s+|(?=[\u4e00-\u9fff]))\S", text)
        if num:
            depth = num.group(1).count(".") + 1
            return min(depth, 3)

        # （一）短标题 → H2；长句 / （1）阿拉伯列表 → 正文，避免多余大纲项
        if re.match(r"^[（(][一二三四五六七八九十]+[）)]", text):
            if len(text) <= HEADING_CN_PAREN_MAX_LEN and not re.search(
                r"[。；;]$", text
            ):
                return 2
            return None

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
                        for run in para.runs:
                            if run.text:
                                run.text = run.text.replace("\u00a0", " ").strip()

    @classmethod
    def _insert_paragraph_after(
        cls,
        paragraph: Paragraph,
        text: str,
        *,
        r_pr=None,
        p_pr=None,
    ) -> Paragraph:
        new_p = OxmlElement("w:p")
        if p_pr is not None:
            new_p.insert(0, copy.deepcopy(p_pr))
        paragraph._p.addnext(new_p)
        new_para = Paragraph(new_p, paragraph._parent)
        if text:
            run = new_para.add_run(text)
            if r_pr is not None:
                cls._set_run_r_pr(run, r_pr)
        return new_para

    @classmethod
    def _replace_paragraph_text(cls, para: Paragraph, text: str, r_pr=None) -> None:
        p = para._p
        for child in list(p):
            if child.tag == qn("w:r"):
                p.remove(child)
        run = para.add_run(text)
        if r_pr is not None:
            cls._set_run_r_pr(run, r_pr)

    @staticmethod
    def _collect_run_formats(para: Paragraph) -> list[tuple[str, object]]:
        """[(run_text, cloned rPr or None), ...] for non-empty runs."""
        out: list[tuple[str, object]] = []
        for run in para.runs:
            text = run.text or ""
            if not text.strip():
                continue
            r_pr = run._element.find(qn("w:rPr"))
            out.append((text, copy.deepcopy(r_pr) if r_pr is not None else None))
        return out

    @staticmethod
    def _r_pr_for_line(
        line: str,
        run_formats: list[tuple[str, object]],
        default_r_pr,
    ):
        """Pick the run format whose text best matches this logical line."""
        if not run_formats:
            return default_r_pr
        best = default_r_pr
        best_score = -1
        for text, r_pr in run_formats:
            if not text:
                continue
            if line == text.strip() or line in text or text.strip() in line:
                score = min(len(line), len(text.strip()))
                if score > best_score and r_pr is not None:
                    best_score = score
                    best = r_pr
        return best if best is not None else default_r_pr

    @staticmethod
    def _clone_p_pr(para: Paragraph):
        """Clone paragraph properties but drop numbering / outline (set later)."""
        p_pr = para._p.find(qn("w:pPr"))
        if p_pr is None:
            return None
        cloned = copy.deepcopy(p_pr)
        for tag in ("w:numPr", "w:outlineLvl", "w:pStyle"):
            node = cloned.find(qn(tag))
            if node is not None:
                cloned.remove(node)
        return cloned

    @staticmethod
    def _clone_element(el):
        return copy.deepcopy(el) if el is not None else None

    @staticmethod
    def _set_run_r_pr(run, r_pr) -> None:
        if r_pr is None:
            return
        r = run._element
        existing = r.find(qn("w:rPr"))
        if existing is not None:
            r.remove(existing)
        r.insert(0, copy.deepcopy(r_pr))

    @staticmethod
    def _looks_like_heading(text: str) -> bool:
        return bool(
            re.match(r"^第[一二三四五六七八九十百零\d]+[章节编部分]", text)
            or re.match(r"^\d+(\.\d+)+\s+", text)
            or re.match(r"^[（(][一二三四五六七八九十\d]+[）)]", text)
        )

    @staticmethod
    def _is_body_style(para) -> bool:
        name = (para.style.name if para.style else "") or ""
        return not (
            name.startswith("Heading")
            or name.startswith("标题")
        )

    @staticmethod
    def _clear_paragraph(para) -> None:
        for run in para.runs:
            run.text = ""
