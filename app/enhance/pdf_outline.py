"""Extract PDF outline hints (bookmarks + tagged structure) for heading assist."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pymupdf

# Structure element roles that map to Word Heading 1–3
_ROLE_TO_LEVEL = {
    "H1": 1,
    "H2": 2,
    "H3": 3,
    "H4": 3,
    "H5": 3,
    "H6": 3,
    "H": 1,  # generic heading — treat as H1; TOC usually more precise
}

_WHITESPACE_RE = re.compile(r"\s+")
_TOC_LEADER_RE = re.compile(r"[\.．…·]{2,}.*$")
_HEADING_PREFIX_RE = re.compile(
    r"^(?:"
    r"第[一二三四五六七八九十百零〇\d]+章\s*"
    r"|\d+\.\d+(?:\.\d+)?(?:\s+|(?=[\u4e00-\u9fffA-Za-z]))"
    r")"
)


@dataclass(frozen=True)
class OutlineHint:
    """One heading candidate from PDF bookmarks or structure tags."""

    level: int  # 1–3
    title: str
    page: int | None = None
    source: str = "toc"  # "toc" | "struct"


def extract_pdf_outline(pdf_path: Path | str) -> list[OutlineHint]:
    """
    Read outline hints from a PDF when present.

    Priority / merge:
      1) Bookmarks (Outlines / get_toc) — most common for 预案 PDFs
      2) Tagged structure tree (StructTreeRoot H1–H3) when available
    """
    path = Path(pdf_path)
    if not path.is_file():
        return []

    hints: list[OutlineHint] = []
    try:
        doc = pymupdf.open(path)
    except Exception:  # noqa: BLE001
        return []

    try:
        hints.extend(_from_toc(doc))
        hints.extend(_from_struct_tree(doc))
    finally:
        doc.close()

    return _dedupe_hints(hints)


def normalize_outline_title(text: str) -> str:
    """Normalize for matching PDF outline titles to DOCX paragraph text."""
    t = (text or "").strip()
    if not t:
        return ""
    t = _TOC_LEADER_RE.sub("", t).strip()
    t = _WHITESPACE_RE.sub("", t)
    return t


def title_key_variants(text: str) -> set[str]:
    """Matching keys: full normalized + without numbering prefix."""
    norm = normalize_outline_title(text)
    if not norm:
        return set()
    keys = {norm}
    stripped = _HEADING_PREFIX_RE.sub("", (text or "").strip(), count=1).strip()
    stripped_norm = normalize_outline_title(stripped)
    if stripped_norm:
        keys.add(stripped_norm)
    return keys


def build_outline_index(hints: list[OutlineHint]) -> dict[str, int]:
    """
    Map normalized title → heading level (1–3).

    First occurrence wins (TOC order); struct tags fill gaps only.
    """
    index: dict[str, int] = {}
    # Prefer toc over struct when both present for same title
    ordered = sorted(hints, key=lambda h: 0 if h.source == "toc" else 1)
    for hint in ordered:
        level = max(1, min(3, int(hint.level)))
        for key in title_key_variants(hint.title):
            if key not in index:
                index[key] = level
    return index


def match_outline_level(text: str, index: dict[str, int]) -> int | None:
    """Return outline level if paragraph text matches an outline entry."""
    if not index:
        return None
    for key in title_key_variants(text):
        if key in index:
            return index[key]
    return None


def _from_toc(doc: pymupdf.Document) -> list[OutlineHint]:
    out: list[OutlineHint] = []
    try:
        toc = doc.get_toc() or []
    except Exception:  # noqa: BLE001
        return out
    for item in toc:
        if not item or len(item) < 2:
            continue
        level = int(item[0])
        title = str(item[1] or "").strip()
        page = int(item[2]) if len(item) > 2 and item[2] else None
        if not title or level < 1:
            continue
        if _TOC_LEADER_RE.search(title) and len(normalize_outline_title(title)) < 2:
            continue
        out.append(
            OutlineHint(
                level=max(1, min(3, level)),
                title=title,
                page=page,
                source="toc",
            )
        )
    return out


def _from_struct_tree(doc: pymupdf.Document) -> list[OutlineHint]:
    """Best-effort walk of /StructTreeRoot for H1–H3 (tagged PDFs)."""
    try:
        cat = doc.pdf_catalog()
        if not isinstance(cat, int):
            return []
        keys = set(doc.xref_get_keys(cat) or [])
        if "StructTreeRoot" not in keys:
            return []
        root_ref = doc.xref_get_key(cat, "StructTreeRoot")
        root_xref = _parse_xref_ref(root_ref)
        if root_xref is None:
            return []
    except Exception:  # noqa: BLE001
        return []

    out: list[OutlineHint] = []
    visited: set[int] = set()

    def walk(xref: int, depth: int = 0) -> None:
        if xref in visited or depth > 40:
            return
        visited.add(xref)
        try:
            role = _struct_role(doc, xref)
            if role in _ROLE_TO_LEVEL:
                title = _struct_text(doc, xref)
                if title and len(title) <= 80:
                    out.append(
                        OutlineHint(
                            level=_ROLE_TO_LEVEL[role],
                            title=title.strip(),
                            page=None,
                            source="struct",
                        )
                    )
            kids = _struct_kids(doc, xref)
            for kid in kids:
                walk(kid, depth + 1)
        except Exception:  # noqa: BLE001
            return

    walk(root_xref)
    return out


def _parse_xref_ref(ref) -> int | None:
    """Parse pymupdf xref_get_key result like ('xref', '12 0 R') → 12."""
    if ref is None:
        return None
    if isinstance(ref, int):
        return ref
    if isinstance(ref, tuple) and len(ref) >= 2:
        val = ref[1]
        if isinstance(val, int):
            return val
        m = re.search(r"(\d+)\s+0\s+R", str(val))
        if m:
            return int(m.group(1))
    m = re.search(r"(\d+)\s+0\s+R", str(ref))
    return int(m.group(1)) if m else None


def _struct_role(doc: pymupdf.Document, xref: int) -> str:
    try:
        keys = set(doc.xref_get_keys(xref) or [])
        if "S" in keys:
            kind, val = doc.xref_get_key(xref, "S")
            raw = str(val).strip()
            if raw.startswith("/"):
                return raw[1:].upper()
            return raw.upper()
    except Exception:  # noqa: BLE001
        pass
    return ""


def _struct_kids(doc: pymupdf.Document, xref: int) -> list[int]:
    kids: list[int] = []
    try:
        keys = set(doc.xref_get_keys(xref) or [])
        if "K" not in keys:
            return kids
        kind, val = doc.xref_get_key(xref, "K")
        raw = str(val)
        # Single kid: "12 0 R" or array "[12 0 R 13 0 R]"
        for m in re.finditer(r"(\d+)\s+0\s+R", raw):
            kids.append(int(m.group(1)))
        # Integer kid (MCID) — skip, text comes from parent leaf walk
    except Exception:  # noqa: BLE001
        return kids
    return kids


def _struct_text(doc: pymupdf.Document, xref: int) -> str:
    """
    Collect actual text for a structure element.

    Prefer /ActualText; else concatenate leaf content via page text is hard
    without full structure map — fall back to empty when only MCIDs.
    """
    try:
        keys = set(doc.xref_get_keys(xref) or [])
        if "ActualText" in keys:
            _kind, val = doc.xref_get_key(xref, "ActualText")
            text = _pdf_string(val)
            if text:
                return text
        if "T" in keys:  # sometimes used for name
            _kind, val = doc.xref_get_key(xref, "T")
            text = _pdf_string(val)
            if text:
                return text
    except Exception:  # noqa: BLE001
        return ""
    return ""


def _pdf_string(val) -> str:
    s = str(val or "").strip()
    if s.startswith("(") and s.endswith(")"):
        s = s[1:-1]
    if s.startswith("<") and s.endswith(">"):
        # hex string — skip heavy decode
        return ""
    s = s.replace("\\(", "(").replace("\\)", ")").replace("\\n", "\n")
    return s.strip()


def _dedupe_hints(hints: list[OutlineHint]) -> list[OutlineHint]:
    seen: set[tuple[str, int]] = set()
    out: list[OutlineHint] = []
    for h in hints:
        key = (normalize_outline_title(h.title), h.level)
        if not key[0] or key in seen:
            continue
        seen.add(key)
        out.append(h)
    return out
