"""Application constants — tune behaviour here, avoid magic numbers elsewhere."""

from pathlib import Path

APP_NAME = "PDF→Word 保真转换"
APP_VERSION = "0.1.0"

# Conversion limits
MAX_PAGES = 500
DEFAULT_OUTPUT_SUFFIX = ".docx"

# Ollama (OpenAI-compatible local API)
OLLAMA_BASE_URL = "http://127.0.0.1:11434/v1"
OLLAMA_MODEL = "qwen2.5:7b"
OLLAMA_TIMEOUT_SEC = 60
OLLAMA_MAX_CANDIDATES = 80

# Structure heuristics — ONLY line-leading 第X章 / X.Y / X.Y.Z
# Numeric titles need "emphasis" (real bold or clearly larger than body);
# also rejected if numbering conflicts with context.
HEADING_PATTERNS = [
    r"^第[一二三四五六七八九十百零〇\d]+章",
    r"^\d+\.\d+(?:\.\d+)?(?:\s+|(?=[\u4e00-\u9fff])|$)",
]

# Max sibling jump allowed (1.1 → 1.4 OK; 1.1 → 1.8 still OK if <= this)
HEADING_MAX_SIBLING_JUMP = 5
# pdf2docx often loses bold flags; larger-than-body size counts as emphasis
HEADING_SIZE_ABOVE_BODY = 1.5  # pt
# Fraction of non-space chars that must be bold to count the line as bold
HEADING_BOLD_CHAR_RATIO = 0.5
# Use PDF bookmarks / structure tags to assist heading promotion when present
HEADING_USE_PDF_OUTLINE = True

# pdf2docx often maps 首行缩进 as w:ind left=… (whole-paragraph indent).
# Convert moderate left-only indents on body paragraphs to firstLine.
BODY_INDENT_FIX = True
# Standard Chinese body first-line indent ≈ 2 chars at 12pt (480 twips)
BODY_FIRST_LINE_TWIPS = 480
# Only rewrite left indents at or below this (skip deep block quotes / covers)
BODY_LEFT_INDENT_MAX_TWIPS = 1000
# Skip when left≈right (cover / centered block)
BODY_INDENT_SYMMETRIC_MIN_TWIPS = 600

# Multilevel Word numbering for Heading 1–3 (insert/reorder auto-updates).
# Strip only numbering prefixes (第一章/1.1/1.1.1); keep title wording + fonts.
HEADING_AUTO_NUMBER = True
HEADING_STRIP_PREFIX = True
HEADING_NUM_ID_HINT = "pdf2word-heading-outline"
HEADING_NUM_FMT = [
    ("decimal", "第%1章 ", 1),  # Heading 1
    ("decimal", "%1.%2 ", 1),  # Heading 2
    ("decimal", "%1.%2.%3 ", 1),  # Heading 3
]

# Project paths
ROOT_DIR = Path(__file__).resolve().parent.parent
