"""Multilevel heading numbering for Word (linked to Heading 1–3)."""

from __future__ import annotations

import copy
import hashlib

from docx.document import Document as DocumentObject
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from app.config import (
    HEADING_NUM_FMT,
    HEADING_NUM_ID_HINT,
)

# Cache on document part: (base_num_id, level, rpr_fp) -> override numId
_CACHE_ATTR = "_pdf2word_heading_num_rpr_cache"


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
    """
    Attach multilevel list to a heading paragraph.

    Number (list label) character format is synced to the paragraph's content
    run (font / size / bold / color / …) via a shared-abstract lvlOverride.
    """
    content_r_pr = _first_content_r_pr(para)
    effective_num_id = num_id
    if content_r_pr is not None:
        effective_num_id = _num_id_matching_content_rpr(
            para.part, num_id, level, content_r_pr
        )

    p_pr = para._p.get_or_add_pPr()
    old = p_pr.find(qn("w:numPr"))
    if old is not None:
        p_pr.remove(old)

    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), str(max(0, min(level, 9) - 1)))
    nid = OxmlElement("w:numId")
    nid.set(qn("w:val"), str(effective_num_id))
    num_pr.append(ilvl)
    num_pr.append(nid)

    p_style = p_pr.find(qn("w:pStyle"))
    if p_style is not None:
        p_style.addnext(num_pr)
    else:
        p_pr.insert(0, num_pr)


def _num_id_matching_content_rpr(part, base_num_id: int, level: int, r_pr) -> int:
    """
    Return a numId that shares the base abstract list but overrides this
    level's w:rPr to match the heading content formatting.
    """
    cache = getattr(part, _CACHE_ATTR, None)
    if cache is None:
        cache = {}
        setattr(part, _CACHE_ATTR, cache)

    fp = _rpr_fingerprint(r_pr)
    key = (base_num_id, level, fp)
    if key in cache:
        return cache[key]

    numbering = part.numbering_part.element
    abstract_id = _abstract_id_for_num(numbering, base_num_id)
    if abstract_id is None:
        cache[key] = base_num_id
        return base_num_id

    template = _lvl_template_from_abstract(numbering, abstract_id, level - 1)
    if template is None:
        # Fall back to config defaults
        fmt, lvl_text, start = HEADING_NUM_FMT[min(level, len(HEADING_NUM_FMT)) - 1]
        template = _make_lvl(level - 1, fmt, lvl_text, start)

    new_id = _next_num_id(numbering)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(new_id))
    abs_ref = OxmlElement("w:abstractNumId")
    abs_ref.set(qn("w:val"), str(abstract_id))
    num.append(abs_ref)

    override = OxmlElement("w:lvlOverride")
    override.set(qn("w:ilvl"), str(level - 1))
    lvl = copy.deepcopy(template)
    lvl.set(qn("w:ilvl"), str(level - 1))
    # Replace / set rPr from content
    old_r = lvl.find(qn("w:rPr"))
    if old_r is not None:
        lvl.remove(old_r)
    lvl.append(copy.deepcopy(r_pr))
    override.append(lvl)
    num.append(override)
    numbering.append(num)

    cache[key] = new_id
    return new_id


def _first_content_r_pr(para: Paragraph):
    for run in para.runs:
        if (run.text or "").strip():
            existing = run._element.find(qn("w:rPr"))
            if existing is not None:
                return copy.deepcopy(existing)
    return None


def _rpr_fingerprint(r_pr) -> str:
    """Stable hash of formatting-relevant rPr children/attrs."""
    if r_pr is None:
        return "none"
    parts: list[str] = []
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is not None:
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs", "w:hint"):
            val = r_fonts.get(qn(attr))
            if val:
                parts.append(f"font:{attr}:{val}")
    for tag in ("w:b", "w:bCs", "w:i", "w:iCs"):
        el = r_pr.find(qn(tag))
        if el is not None:
            parts.append(f"{tag}:{el.get(qn('w:val'))}")
    for tag in ("w:sz", "w:szCs"):
        el = r_pr.find(qn(tag))
        if el is not None and el.get(qn("w:val")):
            parts.append(f"{tag}:{el.get(qn('w:val'))}")
    color = r_pr.find(qn("w:color"))
    if color is not None:
        parts.append(f"color:{color.get(qn('w:val'))}:{color.get(qn('w:themeColor'))}")
    u = r_pr.find(qn("w:u"))
    if u is not None:
        parts.append(f"u:{u.get(qn('w:val'))}")
    raw = "|".join(parts) if parts else r_pr.xml if hasattr(r_pr, "xml") else str(r_pr)
    return hashlib.md5(raw.encode("utf-8", errors="ignore")).hexdigest()


def _abstract_id_for_num(numbering, num_id: int) -> str | None:
    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) == str(num_id):
            ref = num.find(qn("w:abstractNumId"))
            if ref is not None:
                return ref.get(qn("w:val"))
    return None


def _lvl_template_from_abstract(numbering, abstract_id: str, ilvl: int):
    for abstract in numbering.findall(qn("w:abstractNum")):
        if abstract.get(qn("w:abstractNumId")) != str(abstract_id):
            continue
        for lvl in abstract.findall(qn("w:lvl")):
            if lvl.get(qn("w:ilvl")) == str(ilvl):
                return copy.deepcopy(lvl)
    return None


def _make_lvl(ilvl: int, num_fmt: str, lvl_text: str, start: int, r_pr=None):
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

    if r_pr is not None:
        lvl.append(copy.deepcopy(r_pr))

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
                    # Prefer the base instance (no lvlOverride)
                    if num.find(qn("w:lvlOverride")) is None:
                        return int(num.get(qn("w:numId")))
            # Fallback: any num pointing at this abstract
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
