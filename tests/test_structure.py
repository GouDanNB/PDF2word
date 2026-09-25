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
        self.assertIsNone(fixer._detect_heading_level("这是普通正文段落内容足够长。"))

    def test_apply_promotes_headings(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "t.docx"
            doc = Document()
            doc.add_paragraph("第一章 总则")
            doc.add_paragraph("普通正文")
            doc.add_paragraph("1.1 编制目的")
            doc.save(str(path))
            stats = StructureFixer().apply(path)
            self.assertGreaterEqual(stats["headings_promoted"], 2)


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
