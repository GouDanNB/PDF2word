"""Generate a tiny synthetic 应急预案 PDF for smoke tests."""

from __future__ import annotations

from pathlib import Path

import pymupdf


def main() -> None:
    out = Path(__file__).resolve().parent.parent / "tests" / "fixtures" / "sample_yuyan.pdf"
    out.parent.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    y = 72
    lines = [
        ("某化工企业环境突发事件应急预案", 18),
        ("第一章 总则", 14),
        ("1.1 编制目的", 12),
        ("为有效应对突发环境事件，规范应急响应程序，制定本预案。", 11),
        ("1.2 适用范围", 12),
        ("本预案适用于本企业范围内突发环境事件的应急处置。", 11),
        ("第二章 组织机构", 14),
        ("2.1 应急指挥部", 12),
        ("应急指挥部负责统一领导、指挥应急救援工作。", 11),
        ("（一）指挥长职责", 12),
        ("决定启动和终止应急响应。", 11),
    ]
    for text, size in lines:
        page.insert_text((72, y), text, fontsize=size, fontname="china-s")
        y += size + 14

    # Simple table-like text block
    y += 20
    page.insert_text((72, y), "表1 应急物资清单", fontsize=12, fontname="china-s")
    y += 24
    page.insert_text((72, y), "名称          数量          存放地点", fontsize=11, fontname="china-s")
    y += 18
    page.insert_text((72, y), "吸附材料      20袋          仓库A", fontsize=11, fontname="china-s")
    y += 18
    page.insert_text((72, y), "防护服        10套          仓库B", fontsize=11, fontname="china-s")

    doc.save(out)
    doc.close()
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
