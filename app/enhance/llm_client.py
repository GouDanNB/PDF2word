"""Ollama OpenAI-compatible client for structure-only enhancement."""

from __future__ import annotations

import json
import re
from typing import Any

import httpx

from app.config import (
    OLLAMA_BASE_URL,
    OLLAMA_MAX_CANDIDATES,
    OLLAMA_MODEL,
    OLLAMA_TIMEOUT_SEC,
)

SYSTEM_PROMPT = """你是文档结构标注助手。用户会给出从 Word 中抽取的段落候选列表。
你的任务：判断每条应设为 Heading 1 / Heading 2 / Heading 3 / Body。
严格规则：
1. 只返回 JSON，不要解释，不要改写原文内容。
2. 禁止编造条款编号或新增文字。
3. JSON 格式：{"items":[{"index":0,"style":"Heading 1"}, ...]}
4. style 只能是 Heading 1、Heading 2、Heading 3、Body。
5. 只对明显是章节标题的条目标 Heading；拿不准标 Body。
"""


class LLMClient:
    def __init__(
        self,
        base_url: str = OLLAMA_BASE_URL,
        model: str = OLLAMA_MODEL,
        timeout: float = OLLAMA_TIMEOUT_SEC,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.timeout = timeout

    def is_available(self) -> bool:
        try:
            with httpx.Client(timeout=3.0) as client:
                # Ollama native health or OpenAI models list
                r = client.get(f"{self.base_url.rsplit('/v1', 1)[0]}/api/tags")
                return r.status_code == 200
        except Exception:
            return False

    def classify_headings(self, candidates: list[dict[str, Any]]) -> list[dict[str, Any]]:
        if not candidates:
            return []
        trimmed = candidates[:OLLAMA_MAX_CANDIDATES]
        user_payload = {
            "candidates": [
                {"index": c["index"], "text": c["text"][:120]} for c in trimmed
            ]
        }
        body = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {
                    "role": "user",
                    "content": json.dumps(user_payload, ensure_ascii=False),
                },
            ],
            "temperature": 0,
            "stream": False,
        }
        with httpx.Client(timeout=self.timeout) as client:
            resp = client.post(f"{self.base_url}/chat/completions", json=body)
            resp.raise_for_status()
            data = resp.json()
        content = data["choices"][0]["message"]["content"]
        return self._parse_items(content)

    @staticmethod
    def _parse_items(content: str) -> list[dict[str, Any]]:
        content = content.strip()
        # Strip markdown fences if present
        fence = re.search(r"```(?:json)?\s*([\s\S]*?)```", content)
        if fence:
            content = fence.group(1).strip()
        try:
            parsed = json.loads(content)
        except json.JSONDecodeError:
            # Try find first {...}
            m = re.search(r"\{[\s\S]*\}", content)
            if not m:
                return []
            parsed = json.loads(m.group(0))
        items = parsed.get("items", [])
        allowed = {"Heading 1", "Heading 2", "Heading 3", "Body"}
        out = []
        for it in items:
            if not isinstance(it, dict):
                continue
            style = it.get("style")
            idx = it.get("index")
            if style in allowed and isinstance(idx, int):
                out.append({"index": idx, "style": style})
        return out
