"""Headless smoke test: PDF → DOCX via orchestrator."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from app.core.orchestrator import ConversionOrchestrator


def main() -> int:
    pdf = ROOT / "tests" / "fixtures" / "sample_yuyan.pdf"
    out = ROOT / "tests" / "fixtures" / "sample_yuyan.docx"
    if not pdf.exists():
        print(f"missing fixture: {pdf}")
        return 1

    def progress(pct: int, msg: str) -> None:
        print(f"[{pct:3d}%] {msg}")

    result = ConversionOrchestrator().convert(
        pdf, out, enable_llm=False, on_progress=progress
    )
    print(result.message)
    print("structure:", result.structure_stats)
    return 0 if result.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
