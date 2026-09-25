"""Multilevel heading numbering for Word (linked to Heading 1–3)."""

from __future__ import annotations

from docx.document import Document as DocumentObject
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from app.config import (
    HEADING_NUM_FMT,
    HEADING_NUM_ID_HINT,
)


def ensure_heading_numbering(doc: DocumentObject) -> int:
    """
    Ensure a multilevel list exists for Heading 1–3 and return its numId.

    Format (tunable via config):
      L1: 第%1章
      L2: %1.%2
      L3: %1.%2.%3
    """
    numbering = doc.part.numbering_part.element
    existing = _find_existing_heading_num_id(numbering)
    if existing is not None:
        _link_heading_styles(doc, existing)
        return existing

    abstract_id = _next_abstract_num_id(numbering)
    num_id = _next_num_id(numbering)

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))

    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "multilevel")
    abstract.append(multi)

    name = OxmlElement("w:name")
    name.set(qn("w:val"), HEADING_NUM_ID_HINT)
    abstract.append(name)

    for ilvl, (fmt, lvl_text, start) in enumerate(HEADING_NUM_FMT):
        abstract.append(_make_lvl(ilvl, fmt, lvl_text, start))

    # Insert abstractNum before any w:num elements (OOXML order requirement)
    first_num = numbering.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        numbering.append(abstract)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abs_ref = OxmlElement("w:abstractNumId")
    abs_ref.set(qn("w:val"), str(abstract_id))
    num.append(abs_ref)
    numbering.append(num)

    _link_heading_styles(doc, num_id)
    return num_id


def apply_heading_number(para: Paragraph, level: int, num_id: int) -> None:
    """Attach multilevel list reference to a heading paragraph (ilvl 0-based)."""
    p_pr = para._p.get_or_add_pPr()
    # Remove previous numPr if any
    old = p_pr.find(qn("w:numPr"))
    if old is not None:
        p_pr.remove(old)

    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), str(max(0, min(level, 9) - 1)))
    nid = OxmlElement("w:numId")
    nid.set(qn("w:val"), str(num_id))
    num_pr.append(ilvl)
    num_pr.append(nid)

    # Keep numPr near the start of pPr (after pStyle if present)
    p_style = p_pr.find(qn("w:pStyle"))
    if p_style is not None:
        p_style.addnext(num_pr)
    else:
        p_pr.insert(0, num_pr)


def _make_lvl(ilvl: int, num_fmt: str, lvl_text: str, start: int):
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), str(ilvl))

    start_el = OxmlElement("w:start")
    start_el.set(qn("w:val"), str(start))
    lvl.append(start_el)

    fmt_el = OxmlElement("w:numFmt")
    fmt_el.set(qn("w:val"), num_fmt)
    lvl.append(fmt_el)

    text_el = OxmlElement("w:lvlText")
    text_el.set(qn("w:val"), lvl_text)
    lvl.append(text_el)

    align = OxmlElement("w:lvlJc")
    align.set(qn("w:val"), "left")
    lvl.append(align)

    p_pr = OxmlElement("w:pPr")
    ind = OxmlElement("w:ind")
    # Slight hanging indent so wrapped heading lines align under title text
    left = 0 if ilvl == 0 else 210 * ilvl
    hanging = 0 if ilvl == 0 else 210
    ind.set(qn("w:left"), str(left))
    ind.set(qn("w:hanging"), str(hanging))
    p_pr.append(ind)
    lvl.append(p_pr)

    # Restart subordinate levels when this level increments
    if ilvl > 0:
        # lvlRestart optional; Word defaults restart after higher level
        pass

    return lvl


def _link_heading_styles(doc: DocumentObject, num_id: int) -> None:
    """Bind Heading 1–3 paragraph styles to the multilevel list."""
    for level in (1, 2, 3):
        for name in (f"Heading {level}", f"标题 {level}"):
            try:
                style = doc.styles[name]
            except KeyError:
                continue
            p_pr = style.element.get_or_add_pPr()
            old = p_pr.find(qn("w:numPr"))
            if old is not None:
                p_pr.remove(old)
            num_pr = OxmlElement("w:numPr")
            ilvl = OxmlElement("w:ilvl")
            ilvl.set(qn("w:val"), str(level - 1))
            nid = OxmlElement("w:numId")
            nid.set(qn("w:val"), str(num_id))
            num_pr.append(ilvl)
            num_pr.append(nid)
            p_style = p_pr.find(qn("w:pStyle"))
            if p_style is not None:
                p_style.addnext(num_pr)
            else:
                p_pr.insert(0, num_pr)


def _find_existing_heading_num_id(numbering) -> int | None:
    for abstract in numbering.findall(qn("w:abstractNum")):
        name_el = abstract.find(qn("w:name"))
        if name_el is not None and name_el.get(qn("w:val")) == HEADING_NUM_ID_HINT:
            abs_id = abstract.get(qn("w:abstractNumId"))
            for num in numbering.findall(qn("w:num")):
                ref = num.find(qn("w:abstractNumId"))
                if ref is not None and ref.get(qn("w:val")) == abs_id:
                    return int(num.get(qn("w:numId")))
    return None


def _next_abstract_num_id(numbering) -> int:
    ids = [
        int(el.get(qn("w:abstractNumId")))
        for el in numbering.findall(qn("w:abstractNum"))
        if el.get(qn("w:abstractNumId")) is not None
    ]
    return (max(ids) + 1) if ids else 0


def _next_num_id(numbering) -> int:
    ids = [
        int(el.get(qn("w:numId")))
        for el in numbering.findall(qn("w:num"))
        if el.get(qn("w:numId")) is not None
    ]
    return (max(ids) + 1) if ids else 1
