"""Unit tests for heading heuristics."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from docx import Document

from app.enhance.structure_fix import StructureFixer
from app.enhance.llm_client import LLMClient


class StructureFixerTests(unittest.TestCase):
    def test_detect_levels(self) -> None:
        fixer = StructureFixer()
        self.assertEqual(fixer._detect_heading_level("第一章 总则"), 1)
        self.assertEqual(fixer._detect_heading_level("1.1 编制目的"), 2)
        self.assertEqual(fixer._detect_heading_level("1.1.1 细则"), 3)
        self.assertEqual(fixer._detect_heading_level("（一）指挥长职责"), 2)
        # Body list items must NOT become outline headings
        self.assertIsNone(fixer._detect_heading_level("（1）通过调查了解本厂区情况。"))
        self.assertIsNone(fixer._detect_heading_level("1 普通条目不应作一级标题"))
        self.assertIsNone(fixer._detect_heading_level("这是普通正文段落内容足够长。"))

    def test_apply_keeps_pdf_text_and_font(self) -> None:
        """Keep original heading text/fonts; no Word auto-number / strip."""
        from docx.oxml.ns import qn
        from docx.shared import Pt

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "t.docx"
            doc = Document()
            p = doc.add_paragraph()
            r = p.add_run("第一章 总则")
            r.font.name = "SimSun"
            r._element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")
            r.font.size = Pt(14)
            doc.add_paragraph("普通正文")
            p2 = doc.add_paragraph()
            r2 = p2.add_run("1.1 编制目的")
            r2.font.name = "SimSun"
            r2._element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")
            r2.font.size = Pt(12)
            doc.save(str(path))

            stats = StructureFixer().apply(path)
            self.assertGreaterEqual(stats["headings_promoted"], 2)
            self.assertEqual(stats.get("headings_numbered", 0), 0)
            self.assertEqual(stats.get("prefixes_stripped", 0), 0)

            doc2 = Document(str(path))
            texts = [(p.text or "").strip() for p in doc2.paragraphs]
            self.assertIn("第一章 总则", texts)
            self.assertIn("1.1 编制目的", texts)

            h1 = next(p for p in doc2.paragraphs if (p.text or "").strip() == "第一章 总则")
            self.assertEqual(h1.style.name, "Heading 1")
            self.assertIsNone(h1._p.get_or_add_pPr().find(qn("w:numPr")))
            run = next(r for r in h1.runs if (r.text or "").strip())
            r_fonts = run._element.find(qn("w:rPr")).find(qn("w:rFonts"))
            self.assertEqual(r_fonts.get(qn("w:eastAsia")), "SimSun")

    def test_split_multiline_then_outline(self) -> None:
        """pdf2docx-style packed paragraph must become separate outline levels."""
        from docx.oxml.ns import qn
        from docx.shared import Pt

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "packed.docx"
            doc = Document()
            para = doc.add_paragraph()
            run = para.add_run(
                "第一章 总则\n1.1 编制目的\n为有效应对突发环境事件，制定本预案。"
            )
            run.font.name = "SimSun"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")
            run.font.size = Pt(12)
            para2 = doc.add_paragraph()
            run2 = para2.add_run("第二章 组织机构\n2.1 应急指挥部")
            run2.font.name = "SimSun"
            run2._element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")
            run2.font.size = Pt(12)
            doc.save(str(path))

            stats = StructureFixer().apply(path)
            self.assertGreaterEqual(stats["paragraphs_split"], 3)

            doc2 = Document(str(path))
            by_text = {
                (p.text or "").strip(): p
                for p in doc2.paragraphs
                if (p.text or "").strip()
            }
            # Original PDF text preserved (no strip / no auto-number label)
            self.assertIn("第一章 总则", by_text)
            self.assertIn("1.1 编制目的", by_text)
            self.assertIn("为有效应对突发环境事件，制定本预案。", by_text)

            h1 = by_text["第一章 总则"]
            h2 = by_text["1.1 编制目的"]
            body = by_text["为有效应对突发环境事件，制定本预案。"]
            self.assertEqual(h1.style.name, "Heading 1")
            self.assertEqual(h2.style.name, "Heading 2")
            self.assertFalse((body.style.name or "").startswith("Heading"))
            self.assertIsNone(h1._p.get_or_add_pPr().find(qn("w:numPr")))

            ol1 = h1._p.get_or_add_pPr().find(qn("w:outlineLvl"))
            ol2 = h2._p.get_or_add_pPr().find(qn("w:outlineLvl"))
            self.assertEqual(ol1.get(qn("w:val")), "0")
            self.assertEqual(ol2.get(qn("w:val")), "1")

            body_run = next(r for r in body.runs if (r.text or "").strip())
            r_fonts = body_run._element.find(qn("w:rPr")).find(qn("w:rFonts"))
            self.assertEqual(r_fonts.get(qn("w:eastAsia")), "SimSun")
            self.assertEqual(body_run.font.size, Pt(12))


class LLMClientParseTests(unittest.TestCase):
    def test_parse_json(self) -> None:
        raw = '{"items":[{"index":0,"style":"Heading 1"},{"index":1,"style":"Body"}]}'
        items = LLMClient._parse_items(raw)
        self.assertEqual(len(items), 2)
        self.assertEqual(items[0]["style"], "Heading 1")

    def test_parse_fenced(self) -> None:
        raw = '```json\n{"items":[{"index":3,"style":"Heading 2"}]}\n```'
        items = LLMClient._parse_items(raw)
        self.assertEqual(items[0]["index"], 3)


if __name__ == "__main__":
    unittest.main()
