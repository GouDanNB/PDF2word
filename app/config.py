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

# Structure heuristics for 应急预案（与 PDF 大纲对齐，避免把正文条目抬成标题）
HEADING_PATTERNS = [
    # 第X章 / 第X节
    r"^第[一二三四五六七八九十百零\d]+[章节编部分]\s*.+",
    # 1.1 / 1.1.1（至少含一级小数点，避免「1 正文」误识别）
    r"^\d+(\.\d+){1,3}(?:\s+|(?=[\u4e00-\u9fff])).+",
    # （一）短标题；（1）列表项不在此提升
    r"^[（(][一二三四五六七八九十]+[）)]\s*.+",
]

# Font size thresholds (pt) used when promoting headings after pdf2docx
HEADING1_MIN_SIZE = 16
HEADING2_MIN_SIZE = 14
HEADING3_MIN_SIZE = 12

# Keep PDF heading text/fonts as-is. Do NOT inject Word auto-numbers
# (would create「第1章 + 第一章」duplicates and diverge from PDF outline).
HEADING_AUTO_NUMBER = False
HEADING_STRIP_PREFIX = False
HEADING_NUM_ID_HINT = "pdf2word-heading-outline"
HEADING_NUM_FMT = [
    ("decimal", "第%1章 ", 1),
    ("decimal", "%1.%2 ", 1),
    ("decimal", "%1.%2.%3 ", 1),
]

# Max length for Chinese-paren titles like「（一）职责」; longer → treat as body list
HEADING_CN_PAREN_MAX_LEN = 40

# Project paths
ROOT_DIR = Path(__file__).resolve().parent.parent
