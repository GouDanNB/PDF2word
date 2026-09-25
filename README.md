# PDF→Word 保真转换（桌面版）

Windows 本地桌面应用：将 PDF（侧重环境突发事件应急预案）转为 Word，尽量保持版面，并可选通过本机 Ollama 做结构增强。

## 环境

- Python 3.11+（已在 3.14 验证安装流程）
- Windows 10/11
- （可选）[Ollama](https://ollama.com) + 模型如 `qwen2.5:7b`

## 安装

```bash
cd pdf2word-desktop
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 运行（推荐一键启动）

双击项目目录下的 **`start.bat`**（推荐；文件名纯英文，避免编码问题）。

也可双击 `启动.bat`（内部会调用 `start.bat`）。

脚本会自动使用 `.venv`；若不存在，会先创建并安装依赖再打开窗口。

命令行方式：

```bash
.\.venv\Scripts\python.exe run.py
```

或：

```bash
python -m app
```

## 使用

1. 拖入或浏览选择 PDF
2. 确认输出 `.docx` 路径
3. （可选）勾选「启用本地大模型增强」
4. 点击「开始转换」
5. 用 Word 打开结果，对照原 PDF

## 架构摘要

- UI：PySide6
- 主引擎：pdf2docx
- 预案规则：`app/enhance/structure_fix.py`（标题提升、碎段合并）
- LLM：Ollama OpenAI 兼容 API，只返回样式修正指令，不重写正文

## 评测

见 [tests/fixtures/CHECKLIST.md](tests/fixtures/CHECKLIST.md)。将脱敏预案 PDF 放入 `tests/fixtures/` 后按清单人工验收。

## 打包

```bash
.venv\Scripts\activate
pip install pyinstaller
python scripts/build.py
```

产物在 `dist/PDF2Word/`。

## 需求说明

见 [REQUIREMENTS.md](REQUIREMENTS.md)。
