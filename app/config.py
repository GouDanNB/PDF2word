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

# Structure heuristics for 应急预案
HEADING_PATTERNS = [
    # 第X章 / 第X节
    r"^第[一二三四五六七八九十百零\d]+[章节编部分]\s*.+",
    # 1 / 1.1 / 1.1.1
    r"^\d+(\.\d+){0,3}\s+.+",
    # （一） / (1)
    r"^[（(][一二三四五六七八九十\d]+[）)]\s*.+",
]

# Font size thresholds (pt) used when promoting headings after pdf2docx
HEADING1_MIN_SIZE = 16
HEADING2_MIN_SIZE = 14
HEADING3_MIN_SIZE = 12

# Project paths
ROOT_DIR = Path(__file__).resolve().parent.parent
