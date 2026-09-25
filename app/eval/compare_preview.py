"""Lightweight evaluation helpers and checklist text."""

from __future__ import annotations

from pathlib import Path

import pymupdf


CHECKLIST = """
# 金样评测清单（每次改引擎后人工对照）

对每份 fixtures 中的 PDF，转换后用 Word 打开，勾选：

1. [ ] 封面/标题是否可读、未严重错位
2. [ ] 「第 X 章」或「1 / 1.1」是否被识别为标题样式
3. [ ] 正文段落是否可连续选中（非大量碎文本框）
4. [ ] 关键表格（组织机构、联系方式、物资）网格是否完整
5. [ ] 页眉页脚是否未大面积侵入正文

通过率目标（第一版）：≥ 70% 样本「可编辑可用」。
"""


def pdf_page_count(pdf_path: Path) -> int:
    doc = pymupdf.open(pdf_path)
    try:
        return doc.page_count
    finally:
        doc.close()


def write_checklist(dest: Path) -> None:
    dest.write_text(CHECKLIST.strip() + "\n", encoding="utf-8")
