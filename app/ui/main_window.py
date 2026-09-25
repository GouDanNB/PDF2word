"""Main window: drag-drop PDF, convert, progress, optional LLM."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)

from app.config import APP_NAME, APP_VERSION
from app.core.orchestrator import ConversionOrchestrator, ConversionResult


class ConvertWorker(QThread):
    progress = Signal(int, str)
    finished_ok = Signal(object)
    finished_err = Signal(str)

    def __init__(
        self,
        pdf_path: Path,
        output_path: Path,
        enable_llm: bool,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.output_path = output_path
        self.enable_llm = enable_llm

    def run(self) -> None:
        orch = ConversionOrchestrator()

        def on_progress(pct: int, msg: str) -> None:
            self.progress.emit(pct, msg)

        result = orch.convert(
            self.pdf_path,
            self.output_path,
            enable_llm=self.enable_llm,
            on_progress=on_progress,
        )
        if result.ok:
            self.finished_ok.emit(result)
        else:
            self.finished_err.emit(result.message)


class DropLineEdit(QLineEdit):
    file_dropped = Signal(str)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setPlaceholderText("拖入 PDF，或点击「浏览」选择…")
        self.setReadOnly(True)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        urls = event.mimeData().urls()
        if not urls:
            return
        path = urls[0].toLocalFile()
        if path.lower().endswith(".pdf"):
            self.setText(path)
            self.file_dropped.emit(path)
        else:
            QMessageBox.warning(self, "提示", "请拖入 PDF 文件")


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.resize(720, 520)
        self._worker: ConvertWorker | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setSpacing(10)

        title = QLabel(APP_NAME)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: 600; padding: 8px;")
        layout.addWidget(title)

        hint = QLabel("本地转换 · 尽量保持版面 · 可选本机 Ollama 结构增强")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint.setStyleSheet("color: #555;")
        layout.addWidget(hint)

        # PDF path
        pdf_row = QHBoxLayout()
        self.pdf_edit = DropLineEdit()
        self.pdf_edit.file_dropped.connect(self._on_pdf_chosen)
        browse_pdf = QPushButton("浏览…")
        browse_pdf.clicked.connect(self._browse_pdf)
        pdf_row.addWidget(QLabel("PDF"))
        pdf_row.addWidget(self.pdf_edit, stretch=1)
        pdf_row.addWidget(browse_pdf)
        layout.addLayout(pdf_row)

        # Output path
        out_row = QHBoxLayout()
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText("输出 .docx 路径（默认同目录）")
        browse_out = QPushButton("另存为…")
        browse_out.clicked.connect(self._browse_out)
        out_row.addWidget(QLabel("输出"))
        out_row.addWidget(self.out_edit, stretch=1)
        out_row.addWidget(browse_out)
        layout.addLayout(out_row)

        self.llm_check = QCheckBox("启用本地大模型增强（需本机 Ollama，失败自动跳过）")
        layout.addWidget(self.llm_check)

        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.setMinimumHeight(40)
        self.convert_btn.clicked.connect(self._start_convert)
        layout.addWidget(self.convert_btn)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        self.status = QLabel("就绪")
        layout.addWidget(self.status)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setPlaceholderText("日志输出…")
        layout.addWidget(self.log, stretch=1)

    def _log(self, msg: str) -> None:
        self.log.appendPlainText(msg)

    def _browse_pdf(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 PDF", "", "PDF 文件 (*.pdf)"
        )
        if path:
            self.pdf_edit.setText(path)
            self._on_pdf_chosen(path)

    def _on_pdf_chosen(self, path: str) -> None:
        p = Path(path)
        default_out = str(p.with_suffix(".docx"))
        if not self.out_edit.text().strip():
            self.out_edit.setText(default_out)
        self._log(f"已选择: {path}")

    def _browse_out(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self, "保存 Word", self.out_edit.text() or "", "Word 文档 (*.docx)"
        )
        if path:
            if not path.lower().endswith(".docx"):
                path += ".docx"
            self.out_edit.setText(path)

    def _start_convert(self) -> None:
        pdf = self.pdf_edit.text().strip()
        if not pdf:
            QMessageBox.warning(self, "提示", "请先选择 PDF 文件")
            return
        out = self.out_edit.text().strip()
        if not out:
            out = str(Path(pdf).with_suffix(".docx"))
            self.out_edit.setText(out)

        if self._worker and self._worker.isRunning():
            QMessageBox.information(self, "提示", "正在转换中，请稍候")
            return

        self.convert_btn.setEnabled(False)
        self.progress.setValue(0)
        self.status.setText("转换中…")
        self._log("—— 开始转换 ——")

        self._worker = ConvertWorker(
            Path(pdf), Path(out), self.llm_check.isChecked(), self
        )
        self._worker.progress.connect(self._on_progress)
        self._worker.finished_ok.connect(self._on_ok)
        self._worker.finished_err.connect(self._on_err)
        self._worker.start()

    def _on_progress(self, pct: int, msg: str) -> None:
        self.progress.setValue(pct)
        self.status.setText(msg)
        self._log(f"[{pct}%] {msg}")

    def _on_ok(self, result: ConversionResult) -> None:
        self.convert_btn.setEnabled(True)
        self.progress.setValue(100)
        self.status.setText("完成")
        stats = result.structure_stats
        self._log(
            f"结构规则: 标题提升 {stats.get('headings_promoted', 0)}，"
            f"碎段合并 {stats.get('fragments_merged', 0)}"
        )
        llm = result.llm_stats or {}
        if llm.get("skipped"):
            self._log(f"LLM: {llm.get('reason', '跳过')}")
        else:
            self._log(f"LLM: 应用样式 {llm.get('applied', 0)} 处")
        self._log(result.message)
        QMessageBox.information(
            self,
            "完成",
            f"{result.message}\n\n请用 Word 打开对照原 PDF。",
        )

    def _on_err(self, message: str) -> None:
        self.convert_btn.setEnabled(True)
        self.status.setText("失败")
        self._log(f"错误: {message}")
        QMessageBox.critical(self, "转换失败", message)
