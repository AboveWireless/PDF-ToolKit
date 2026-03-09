from __future__ import annotations

from pathlib import Path
import json
import sys
from typing import Any

import fitz

try:
    from PySide6.QtCore import QObject, QPoint, QRect, QRunnable, QSettings, QSize, Qt, QThreadPool, QUrl, Signal
    from PySide6.QtGui import QAction, QColor, QDesktopServices, QFont, QIcon, QImage, QPainter, QPen, QPixmap, QRegularExpressionValidator
    from PySide6.QtWidgets import (
        QApplication,
        QAbstractItemView,
        QCheckBox,
        QComboBox,
        QDoubleSpinBox,
        QFileDialog,
        QFormLayout,
        QFrame,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMessageBox,
        QPlainTextEdit,
        QPushButton,
        QScrollArea,
        QSpinBox,
        QSplitter,
        QTableWidget,
        QTableWidgetItem,
        QTabWidget,
        QTreeWidget,
        QTreeWidgetItem,
        QVBoxLayout,
        QWidget,
        QStyleFactory,
    )
    from PySide6.QtCore import QRegularExpression
except ImportError as exc:  # pragma: no cover - runtime dependency check
    QObject = object
    QPoint = object
    QRect = object
    QRunnable = object
    QSettings = object
    QSize = object
    Qt = object
    QThreadPool = object
    QUrl = object
    Signal = lambda *args, **kwargs: None
    QAction = object
    QColor = object
    QDesktopServices = object
    QFont = object
    QIcon = object
    QImage = object
    QPainter = object
    QPen = object
    QPixmap = object
    QRegularExpressionValidator = object
    QRegularExpression = object
    QWidget = object
    QApplication = None
    QAbstractItemView = object
    QCheckBox = object
    QComboBox = object
    QDoubleSpinBox = object
    QFileDialog = object
    QFormLayout = object
    QFrame = object
    QHBoxLayout = object
    QLabel = object
    QLineEdit = object
    QListWidget = object
    QListWidgetItem = object
    QMainWindow = object
    QMessageBox = object
    QPlainTextEdit = object
    QPushButton = object
    QScrollArea = object
    QSpinBox = object
    QSplitter = object
    QTableWidget = object
    QTableWidgetItem = object
    QTabWidget = object
    QTreeWidget = object
    QTreeWidgetItem = object
    QVBoxLayout = object
    QStyleFactory = object
    _PYSIDE_IMPORT_ERROR = exc
else:  # pragma: no cover - defined only when import works
    _PYSIDE_IMPORT_ERROR = None

from watchdog.observers import Observer

from pdf_toolkit.application import JobResult, OperationDefinition, OperationField, execute_job, get_operation_definitions, prepare_request
from pdf_toolkit.batch import WatchFolderHandler, build_file_batch_manifest, build_folder_batch_manifest, write_manifest
from pdf_toolkit.branding import APP_NAME, APP_TAGLINE, APP_VERSION, DOCS_URL, OCR_NOTE, ORGANIZATION_DOMAIN, ORGANIZATION_NAME, PROJECT_URL, RELEASES_URL, WELCOME_COPY
from pdf_toolkit.config import load_config
from pdf_toolkit.errors import ValidationError
from pdf_toolkit.environment import collect_doctor_status
from pdf_toolkit.ocr import scan_detect
from pdf_toolkit.workflow_templates import WorkflowTemplate, get_workflow_template, get_workflow_templates

PAGE_SPEC_VALIDATOR = QRegularExpressionValidator(QRegularExpression(r"^[0-9,\-\s]*$")) if QApplication is not None else None


def _open_path(path: Path) -> None:
    QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))


def _open_url(url: str) -> None:
    QDesktopServices.openUrl(QUrl(url))


def _pixmap_from_fitz(doc: fitz.Document, page_index: int, scale: float) -> QPixmap:
    page = doc.load_page(page_index)
    matrix = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888).copy()
    return QPixmap.fromImage(image)


def _pretty_json(data: object) -> str:
    return json.dumps(data, indent=2, sort_keys=True)


def _app_icon() -> QIcon:
    pixmap = QPixmap(256, 256)
    pixmap.fill(QColor("#081218"))
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    painter.setPen(QPen(QColor("#224a5e"), 6))
    painter.setBrush(QColor("#102532"))
    painter.drawRoundedRect(20, 20, 216, 216, 40, 40)
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor("#11b8f5"))
    painter.drawRoundedRect(52, 48, 52, 160, 18, 18)
    painter.setBrush(QColor("#ffd166"))
    painter.drawRoundedRect(112, 48, 92, 36, 18, 18)
    painter.setBrush(QColor("#ff6f61"))
    painter.drawRoundedRect(112, 96, 92, 36, 18, 18)
    painter.setBrush(QColor("#72e0a0"))
    painter.drawRoundedRect(112, 144, 92, 36, 18, 18)
    painter.setPen(QColor("#e8f6ff"))
    font = QFont("Segoe UI", 28)
    font.setBold(True)
    painter.setFont(font)
    painter.drawText(QRect(44, 182, 170, 42), Qt.AlignmentFlag.AlignCenter, "PDF")
    painter.end()
    return QIcon(pixmap)


def _panel_widget(title: str, body: QWidget, *, accent: str = "#204b57") -> QWidget:
    container = QFrame()
    container.setObjectName("PanelCard")
    layout = QVBoxLayout(container)
    layout.setContentsMargins(18, 18, 18, 18)
    layout.setSpacing(14)
    header = QFrame()
    header.setObjectName("PanelHeader")
    header_layout = QHBoxLayout(header)
    header_layout.setContentsMargins(0, 0, 0, 0)
    header_layout.setSpacing(10)
    accent_bar = QFrame()
    accent_bar.setObjectName("PanelAccent")
    accent_bar.setFixedWidth(6)
    title_stack = QVBoxLayout()
    title_stack.setContentsMargins(0, 0, 0, 0)
    title_stack.setSpacing(2)
    eyebrow = QLabel("Workspace Zone")
    eyebrow.setObjectName("PanelEyebrow")
    title_label = QLabel(title)
    title_label.setObjectName("PanelTitle")
    title_label.setProperty("accentColor", accent)
    title_stack.addWidget(eyebrow)
    title_stack.addWidget(title_label)
    header_layout.addWidget(accent_bar, 0)
    header_layout.addLayout(title_stack, 1)
    layout.addWidget(header)
    layout.addWidget(body, 1)
    return container


def _clear_layout(layout) -> None:
    while layout.count():
        item = layout.takeAt(0)
        widget = item.widget()
        child_layout = item.layout()
        if widget is not None:
            widget.deleteLater()
        elif child_layout is not None:
            _clear_layout(child_layout)


def _json_safe(value: Any) -> Any:
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, dict):
        return {str(key): _json_safe(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_json_safe(item) for item in value]
    return value


def _normalize_intake_paths(paths: list[str | Path]) -> list[Path]:
    normalized: list[Path] = []
    seen: set[str] = set()
    for item in paths:
        path = Path(str(item)).expanduser()
        key = str(path)
        if not key or key in seen:
            continue
        seen.add(key)
        normalized.append(path)
    return normalized


def _safe_pdf_page_count(path: Path) -> tuple[int, bool]:
    try:
        with fitz.open(path) as document:
            encrypted = bool(getattr(document, "needs_pass", False))
            return document.page_count, encrypted
    except Exception:
        return 0, False


def _default_output_stem(paths: list[Path]) -> str:
    if len(paths) == 1:
        return paths[0].stem
    if len(paths) > 1:
        return "merged"
    return "document"


def _suggest_output_values(operation_id: str, intake_paths: list[str | Path]) -> dict[str, Any]:
    paths = _normalize_intake_paths(intake_paths)
    stem = _default_output_stem(paths)
    if operation_id == "merge":
        return {"output": "merged.pdf"}
    if operation_id == "split":
        return {"output_dir": "split-pages"}
    if operation_id == "extract-text":
        return {"output": f"{stem}-text.txt"}
    if operation_id == "tables-extract":
        return {"output_dir": f"{stem}-tables"}
    if operation_id == "redact":
        return {"output": f"{stem}-redacted.pdf"}
    if operation_id == "ocr":
        return {"output": f"{stem}-searchable.pdf"}
    if operation_id == "extract-images":
        return {"output_dir": f"{stem}-images"}
    if operation_id == "render":
        return {"output_dir": f"{stem}-render"}
    if operation_id == "extract-llm":
        return {"output_dir": f"{stem}-llm"}
    if operation_id == "batch-run":
        return {
            "output_root": f"{stem}-processed",
            "report_path": f"{stem}-processed\\batch-report.json",
        }
    return {}


def _suggest_workflow_for_intake(intake_paths: list[str | Path]) -> dict[str, Any]:
    paths = [path for path in _normalize_intake_paths(intake_paths) if path.suffix.lower() == ".pdf"]
    file_count = len(paths)
    page_count = 0
    encrypted = False
    scan_summary: str | None = None
    warnings: list[str] = []
    alternatives: list[str] = []

    for path in paths[:10]:
        pages, is_encrypted = _safe_pdf_page_count(path)
        page_count += pages
        encrypted = encrypted or is_encrypted

    if file_count == 1 and paths[0].exists():
        try:
            detection = scan_detect(paths[0])
            scan_summary = str(detection.get("summary", ""))
            if not page_count:
                page_count = int(detection.get("page_count", 0))
        except Exception:
            scan_summary = None

    if file_count > 1:
        suggested_operation = "merge"
        suggested_label = "Combine PDFs"
        alternatives = ["Split Pages", "Export Text"]
    elif file_count == 1 and scan_summary == "image-based":
        suggested_operation = "ocr"
        suggested_label = "OCR Scans"
        alternatives = ["Export Text", "Redact", "Export Tables"]
    elif file_count == 1:
        suggested_operation = "extract-text"
        suggested_label = "Export Text"
        alternatives = ["Split Pages", "Redact", "Export Tables"]
    else:
        suggested_operation = "merge"
        suggested_label = "Combine PDFs"
        warnings.append("Drop one or more PDFs to get a suggested task.")

    if encrypted:
        warnings.append("One or more PDFs appear to be encrypted.")
    if scan_summary == "image-based":
        warnings.append("This PDF appears scan-heavy; OCR is likely the best next step.")

    return {
        "paths": [str(path) for path in paths],
        "file_count": file_count,
        "page_count": page_count,
        "encrypted": encrypted,
        "scan_summary": scan_summary,
        "suggested_operation": suggested_operation,
        "suggested_label": suggested_label,
        "alternatives": alternatives,
        "warnings": warnings,
    }


class IntakeDropZone(QFrame):
    files_dropped = Signal(list)

    def __init__(self, settings: QSettings, title: str, detail: str, *, button_label: str = "Choose PDF Files") -> None:
        super().__init__()
        self._settings = settings
        self.setObjectName("DropZoneCard")
        self.setAcceptDrops(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(8)
        eyebrow = QLabel(title)
        eyebrow.setObjectName("CardEyebrow")
        description = QLabel(detail)
        description.setWordWrap(True)
        description.setObjectName("MutedLabel")
        button = QPushButton(button_label)
        button.setProperty("secondary", True)
        button.clicked.connect(self._browse)
        layout.addWidget(eyebrow)
        layout.addWidget(description)
        layout.addWidget(button, 0, Qt.AlignmentFlag.AlignLeft)

    def _browse(self) -> None:
        start_dir = self._settings.value("paths/last_dir", str(Path.cwd()))
        files, _ = QFileDialog.getOpenFileNames(self, "Choose PDF Files", str(start_dir), "PDF Files (*.pdf);;All Files (*.*)")
        if files:
            self._settings.setValue("paths/last_dir", str(Path(files[0]).parent))
            self.files_dropped.emit(files)

    def dragEnterEvent(self, event) -> None:  # pragma: no cover - GUI interaction
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # pragma: no cover - GUI interaction
        files = [url.toLocalFile() for url in event.mimeData().urls() if url.toLocalFile()]
        if files:
            self.files_dropped.emit(files)


class SinglePathInput(QWidget):
    changed = Signal()

    def __init__(self, settings: QSettings, *, directory: bool = False, save_mode: bool = False) -> None:
        super().__init__()
        self._settings = settings
        self._directory = directory
        self._save_mode = save_mode
        self.setAcceptDrops(True)
        self._line_edit = QLineEdit()
        self._browse_button = QPushButton("Browse")
        self._browse_button.clicked.connect(self._browse)
        self._line_edit.textChanged.connect(lambda *_: self.changed.emit())
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._line_edit)
        layout.addWidget(self._browse_button)

    def value(self) -> str | None:
        text = self._line_edit.text().strip()
        return text or None

    def set_value(self, value: str | Path | None) -> None:
        self._line_edit.setText("" if value is None else str(value))

    def _browse(self) -> None:
        start_dir = self._settings.value("paths/last_dir", str(Path.cwd()))
        if self._directory:
            selected = QFileDialog.getExistingDirectory(self, "Select Folder", start_dir)
        elif self._save_mode:
            selected, _ = QFileDialog.getSaveFileName(self, "Select Output", start_dir)
        else:
            selected, _ = QFileDialog.getOpenFileName(self, "Select File", start_dir)
        if selected:
            self._settings.setValue("paths/last_dir", str(Path(selected).parent if not self._directory else Path(selected)))
            self._line_edit.setText(selected)

    def dragEnterEvent(self, event) -> None:  # pragma: no cover - GUI interaction
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # pragma: no cover - GUI interaction
        urls = event.mimeData().urls()
        if not urls:
            return
        path = Path(urls[0].toLocalFile())
        self._line_edit.setText(str(path))
        self.changed.emit()


class MultiPathInput(QWidget):
    changed = Signal()

    def __init__(self, settings: QSettings) -> None:
        super().__init__()
        self._settings = settings
        self.setAcceptDrops(True)
        self._list = QListWidget()
        self._list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        add_button = QPushButton("Add")
        remove_button = QPushButton("Remove")
        add_button.clicked.connect(self._add_items)
        remove_button.clicked.connect(self._remove_items)
        self._list.model().rowsInserted.connect(lambda *_: self.changed.emit())
        self._list.model().rowsRemoved.connect(lambda *_: self.changed.emit())
        buttons = QVBoxLayout()
        buttons.addWidget(add_button)
        buttons.addWidget(remove_button)
        buttons.addStretch(1)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._list, 1)
        layout.addLayout(buttons)

    def value(self) -> list[str]:
        return [self._list.item(index).text() for index in range(self._list.count())]

    def set_value(self, values: list[str] | list[Path] | None) -> None:
        self._list.clear()
        for value in values or []:
            self._list.addItem(str(value))

    def _add_items(self) -> None:
        start_dir = self._settings.value("paths/last_dir", str(Path.cwd()))
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files", start_dir)
        for file_name in files:
            self._list.addItem(file_name)
        if files:
            self._settings.setValue("paths/last_dir", str(Path(files[0]).parent))

    def _remove_items(self) -> None:
        for item in self._list.selectedItems():
            self._list.takeItem(self._list.row(item))
        self.changed.emit()

    def dragEnterEvent(self, event) -> None:  # pragma: no cover - GUI interaction
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # pragma: no cover - GUI interaction
        for url in event.mimeData().urls():
            local_file = url.toLocalFile()
            if local_file:
                self._list.addItem(local_file)
        self.changed.emit()


class MultiLineInput(QWidget):
    changed = Signal()

    def __init__(self) -> None:
        super().__init__()
        self._editor = QPlainTextEdit()
        self._editor.setPlaceholderText("One value per line")
        self._editor.textChanged.connect(lambda: self.changed.emit())
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._editor)

    def value(self) -> list[str]:
        return [line.strip() for line in self._editor.toPlainText().splitlines() if line.strip()]

    def set_value(self, values: list[str] | None) -> None:
        self._editor.setPlainText("\n".join(values or []))


class KeyValueTableInput(QWidget):
    changed = Signal()

    def __init__(self) -> None:
        super().__init__()
        self._table = QTableWidget(0, 2)
        self._table.setHorizontalHeaderLabels(["Key", "Value"])
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.itemChanged.connect(lambda *_: self.changed.emit())
        add_button = QPushButton("Add Row")
        remove_button = QPushButton("Remove Row")
        add_button.clicked.connect(self._add_row)
        remove_button.clicked.connect(self._remove_row)
        buttons = QHBoxLayout()
        buttons.addWidget(add_button)
        buttons.addWidget(remove_button)
        buttons.addStretch(1)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._table)
        layout.addLayout(buttons)

    def value(self) -> dict[str, str]:
        data: dict[str, str] = {}
        for row in range(self._table.rowCount()):
            key_item = self._table.item(row, 0)
            value_item = self._table.item(row, 1)
            key = key_item.text().strip() if key_item else ""
            if key:
                data[key] = value_item.text() if value_item else ""
        return data

    def set_value(self, values: dict[str, str] | None) -> None:
        self._table.setRowCount(0)
        for key, value in (values or {}).items():
            self._add_row(key, value)

    def _add_row(self, key: str = "", value: str = "") -> None:
        row = self._table.rowCount()
        self._table.insertRow(row)
        self._table.setItem(row, 0, QTableWidgetItem(key))
        self._table.setItem(row, 1, QTableWidgetItem(value))
        self.changed.emit()

    def _remove_row(self) -> None:
        rows = {item.row() for item in self._table.selectedItems()}
        for row in sorted(rows, reverse=True):
            self._table.removeRow(row)
        self.changed.emit()


class FolderBatchBuilder(QWidget):
    changed = Signal()

    def __init__(self, settings: QSettings) -> None:
        super().__init__()
        self._settings = settings

        self._source_mode = QComboBox()
        self._source_mode.addItem("Folder", "folder")
        self._source_mode.addItem("Selected Files", "files")
        self._input_dir = SinglePathInput(settings, directory=True)
        self._input_files = MultiPathInput(settings)
        self._output_dir = SinglePathInput(settings, directory=True)
        self._report_path = SinglePathInput(settings, save_mode=True)
        self._pattern = QLineEdit("*.pdf")
        self._recursive = QCheckBox("Include subfolders")
        self._recursive.setChecked(True)
        self._fail_fast = QCheckBox("Stop on first error")
        self._job_name = QLineEdit("folder-batch")

        self._compress = QCheckBox("Compress PDFs")
        self._extract_text = QCheckBox("Extract text files")
        self._extract_llm = QCheckBox("Extract LLM bundle")
        self._llm_chunk_size = QSpinBox()
        self._llm_chunk_size.setRange(200, 20000)
        self._llm_chunk_size.setValue(1200)
        self._llm_overlap = QSpinBox()
        self._llm_overlap.setRange(0, 5000)
        self._llm_overlap.setValue(200)
        self._llm_include_page_markers = QCheckBox("Keep page markers in chunks")
        self._llm_include_page_markers.setChecked(True)
        self._llm_include_metadata = QCheckBox("Include document metadata and bookmarks")
        self._llm_include_metadata.setChecked(True)
        self._analyze_llm = QCheckBox("Analyze with LLM")
        self._llm_preset = QComboBox()
        self._llm_preset.addItem("Summary", "summary")
        self._llm_preset.addItem("Entities", "entities")
        self._llm_preset.addItem("Q&A", "qa")
        self._llm_question = QLineEdit()
        self._llm_question.setPlaceholderText("Required for Q&A preset")
        self._llm_model = QLineEdit("gpt-5-mini")
        self._render = QCheckBox("Render page images")
        self._render_dpi = QSpinBox()
        self._render_dpi.setRange(72, 600)
        self._render_dpi.setValue(150)
        self._render_format = QComboBox()
        self._render_format.addItem("PNG", "png")
        self._render_format.addItem("JPG", "jpg")

        self._ocr = QCheckBox("Run OCR")
        self._ocr_language = QLineEdit("eng")
        self._ocr_skip_text = QCheckBox("Skip existing text")
        self._ocr_skip_text.setChecked(True)
        self._ocr_force = QCheckBox("Force OCR")

        self._tables = QCheckBox("Extract tables")
        self._tables_format = QComboBox()
        self._tables_format.addItem("CSV", "csv")
        self._tables_format.addItem("XLSX", "xlsx")
        self._tables_format.addItem("JSON", "json")
        self._tables_format.addItem("All", "all")
        self._tables_ocr_first = QCheckBox("OCR before table extraction")

        self._metadata = QCheckBox("Set metadata")
        self._metadata_values = KeyValueTableInput()

        self._redact = QCheckBox("Redact patterns")
        self._redact_patterns = MultiLineInput()
        self._redact_regex = QCheckBox("Use regex")
        self._redact_case_sensitive = QCheckBox("Case sensitive")
        self._redact_pages = QLineEdit()
        self._redact_pages.setPlaceholderText("Optional page range, e.g. 1-3")
        self._redact_label = QLineEdit()
        self._redact_label.setPlaceholderText("Optional label")

        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(14)
        root_layout.addWidget(self._build_general_section())
        root_layout.addWidget(self._build_feature_section())
        root_layout.addStretch(1)
        self._source_mode.currentIndexChanged.connect(self._update_source_mode)
        self._update_source_mode()

        for widget in (
            self._input_dir,
            self._input_files,
            self._output_dir,
            self._report_path,
            self._metadata_values,
            self._redact_patterns,
        ):
            widget.changed.connect(self.changed.emit)
        for widget in (
            self._source_mode,
            self._pattern,
            self._recursive,
            self._fail_fast,
            self._job_name,
            self._compress,
            self._extract_text,
            self._extract_llm,
            self._llm_chunk_size,
            self._llm_overlap,
            self._llm_include_page_markers,
            self._llm_include_metadata,
            self._analyze_llm,
            self._llm_preset,
            self._llm_question,
            self._llm_model,
            self._render,
            self._render_dpi,
            self._render_format,
            self._ocr,
            self._ocr_language,
            self._ocr_skip_text,
            self._ocr_force,
            self._tables,
            self._tables_format,
            self._tables_ocr_first,
            self._metadata,
            self._redact,
            self._redact_regex,
            self._redact_case_sensitive,
            self._redact_pages,
            self._redact_label,
        ):
            if hasattr(widget, "textChanged"):
                widget.textChanged.connect(lambda *_: self.changed.emit())
            elif hasattr(widget, "stateChanged"):
                widget.stateChanged.connect(lambda *_: self.changed.emit())
            elif hasattr(widget, "valueChanged"):
                widget.valueChanged.connect(lambda *_: self.changed.emit())
            elif hasattr(widget, "currentIndexChanged"):
                widget.currentIndexChanged.connect(lambda *_: self.changed.emit())

    def _build_general_section(self) -> QWidget:
        section = QFrame()
        section.setObjectName("InnerSection")
        layout = QFormLayout(section)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setHorizontalSpacing(16)
        layout.setVerticalSpacing(12)
        self._general_layout = layout
        self._folder_row_label = QLabel("Input Folder")
        self._files_row_label = QLabel("Selected PDF Files")
        self._pattern_row_label = QLabel("File Pattern")
        self._recursive_row_label = QLabel("Folder Options")
        layout.addRow("Source", self._source_mode)
        layout.addRow(self._folder_row_label, self._input_dir)
        layout.addRow(self._files_row_label, self._input_files)
        layout.addRow("Output Folder", self._output_dir)
        layout.addRow("Batch Report", self._report_path)
        layout.addRow(self._pattern_row_label, self._pattern)
        layout.addRow("Job Name", self._job_name)
        layout.addRow(self._recursive_row_label, self._recursive)
        layout.addRow("", self._fail_fast)
        return section

    def _build_feature_section(self) -> QWidget:
        section = QFrame()
        section.setObjectName("InnerSection")
        layout = QVBoxLayout(section)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)
        title = QLabel("Features To Run On Each PDF")
        title.setObjectName("FieldLabel")
        layout.addWidget(title)

        layout.addWidget(self._compress)
        layout.addWidget(self._extract_text)

        llm_form = QFormLayout()
        llm_form.addRow(self._extract_llm, QWidget())
        llm_form.addRow("Chunk Size", self._llm_chunk_size)
        llm_form.addRow("Chunk Overlap", self._llm_overlap)
        llm_form.addRow("", self._llm_include_page_markers)
        llm_form.addRow("", self._llm_include_metadata)
        llm_form.addRow(self._analyze_llm, QWidget())
        llm_form.addRow("Analysis Preset", self._llm_preset)
        llm_form.addRow("Question", self._llm_question)
        llm_form.addRow("Model", self._llm_model)
        layout.addLayout(llm_form)

        render_form = QFormLayout()
        render_form.addRow(self._render, QWidget())
        render_form.addRow("Render DPI", self._render_dpi)
        render_form.addRow("Image Format", self._render_format)
        layout.addLayout(render_form)

        ocr_form = QFormLayout()
        ocr_form.addRow(self._ocr, QWidget())
        ocr_form.addRow("OCR Language", self._ocr_language)
        ocr_form.addRow("", self._ocr_skip_text)
        ocr_form.addRow("", self._ocr_force)
        layout.addLayout(ocr_form)

        tables_form = QFormLayout()
        tables_form.addRow(self._tables, QWidget())
        tables_form.addRow("Table Format", self._tables_format)
        tables_form.addRow("", self._tables_ocr_first)
        layout.addLayout(tables_form)

        metadata_form = QFormLayout()
        metadata_form.addRow(self._metadata, QWidget())
        metadata_form.addRow("Metadata Values", self._metadata_values)
        layout.addLayout(metadata_form)

        redact_form = QFormLayout()
        redact_form.addRow(self._redact, QWidget())
        redact_form.addRow("Patterns", self._redact_patterns)
        redact_form.addRow("", self._redact_regex)
        redact_form.addRow("", self._redact_case_sensitive)
        redact_form.addRow("Pages", self._redact_pages)
        redact_form.addRow("Label", self._redact_label)
        layout.addLayout(redact_form)
        return section

    def _update_source_mode(self) -> None:
        folder_mode = self._source_mode.currentData() == "folder"
        self._folder_row_label.setVisible(folder_mode)
        self._input_dir.setVisible(folder_mode)
        self._pattern_row_label.setVisible(folder_mode)
        self._pattern.setVisible(folder_mode)
        self._recursive_row_label.setVisible(folder_mode)
        self._recursive.setVisible(folder_mode)
        self._files_row_label.setVisible(not folder_mode)
        self._input_files.setVisible(not folder_mode)
        self.changed.emit()

    def value(self) -> dict[str, object]:
        steps: list[dict[str, object]] = []
        if self._compress.isChecked():
            steps.append({"action": "compress"})
        if self._extract_text.isChecked():
            steps.append({"action": "extract_text"})
        if self._extract_llm.isChecked():
            steps.append(
                {
                    "action": "extract_llm",
                    "chunk_size": self._llm_chunk_size.value(),
                    "overlap": self._llm_overlap.value(),
                    "include_page_markers": self._llm_include_page_markers.isChecked(),
                    "include_metadata": self._llm_include_metadata.isChecked(),
                }
            )
        if self._analyze_llm.isChecked():
            steps.append(
                {
                    "action": "analyze_llm",
                    "preset": self._llm_preset.currentData(),
                    "question": self._llm_question.text().strip() or None,
                    "model": self._llm_model.text().strip() or "gpt-5-mini",
                }
            )
        if self._render.isChecked():
            steps.append(
                {
                    "action": "render",
                    "dpi": self._render_dpi.value(),
                    "image_format": self._render_format.currentData(),
                }
            )
        if self._ocr.isChecked():
            steps.append(
                {
                    "action": "ocr",
                    "language": self._ocr_language.text().strip() or "eng",
                    "skip_existing_text": self._ocr_skip_text.isChecked(),
                    "force": self._ocr_force.isChecked(),
                }
            )
        if self._tables.isChecked():
            steps.append(
                {
                    "action": "tables_extract",
                    "format": self._tables_format.currentData(),
                    "ocr_first": self._tables_ocr_first.isChecked(),
                }
            )
        if self._metadata.isChecked():
            steps.append(
                {
                    "action": "set_metadata",
                    "values": self._metadata_values.value(),
                }
            )
        if self._redact.isChecked():
            steps.append(
                {
                    "action": "redact",
                    "patterns": self._redact_patterns.value(),
                    "regex": self._redact_regex.isChecked(),
                    "case_sensitive": self._redact_case_sensitive.isChecked(),
                    "pages": self._redact_pages.text().strip() or None,
                    "label": self._redact_label.text().strip() or None,
                }
            )

        return {
            "source_mode": self._source_mode.currentData(),
            "input_root": self._input_dir.value(),
            "input_files": self._input_files.value(),
            "output_root": self._output_dir.value(),
            "report_path": self._report_path.value(),
            "file_patterns": [self._pattern.text().strip() or "*.pdf"],
            "recursive_inputs": self._recursive.isChecked(),
            "fail_fast": self._fail_fast.isChecked(),
            "job_name": self._job_name.text().strip() or "folder-batch",
            "steps": steps,
        }

    def set_value(self, values: dict[str, object] | None) -> None:
        payload = dict(values or {})
        mode = str(payload.get("source_mode", "folder"))
        mode_index = self._source_mode.findData(mode)
        self._source_mode.setCurrentIndex(mode_index if mode_index >= 0 else 0)
        self._input_dir.set_value(payload.get("input_root"))
        self._input_files.set_value(payload.get("input_files", []))
        self._output_dir.set_value(payload.get("output_root"))
        self._report_path.set_value(payload.get("report_path"))
        file_patterns = payload.get("file_patterns", ["*.pdf"])
        first_pattern = file_patterns[0] if isinstance(file_patterns, list) and file_patterns else "*.pdf"
        self._pattern.setText(str(first_pattern or "*.pdf"))
        self._recursive.setChecked(bool(payload.get("recursive_inputs", True)))
        self._fail_fast.setChecked(bool(payload.get("fail_fast", False)))
        self._job_name.setText(str(payload.get("job_name", "folder-batch")))

        self._compress.setChecked(False)
        self._extract_text.setChecked(False)
        self._extract_llm.setChecked(False)
        self._llm_chunk_size.setValue(1200)
        self._llm_overlap.setValue(200)
        self._llm_include_page_markers.setChecked(True)
        self._llm_include_metadata.setChecked(True)
        self._analyze_llm.setChecked(False)
        self._llm_preset.setCurrentIndex(max(self._llm_preset.findData("summary"), 0))
        self._llm_question.clear()
        self._llm_model.setText("gpt-5-mini")
        self._render.setChecked(False)
        self._render_dpi.setValue(150)
        self._render_format.setCurrentIndex(max(self._render_format.findData("png"), 0))
        self._ocr.setChecked(False)
        self._ocr_language.setText("eng")
        self._ocr_skip_text.setChecked(True)
        self._ocr_force.setChecked(False)
        self._tables.setChecked(False)
        self._tables_format.setCurrentIndex(max(self._tables_format.findData("csv"), 0))
        self._tables_ocr_first.setChecked(False)
        self._metadata.setChecked(False)
        self._metadata_values.set_value({})
        self._redact.setChecked(False)
        self._redact_patterns.set_value([])
        self._redact_regex.setChecked(False)
        self._redact_case_sensitive.setChecked(False)
        self._redact_pages.clear()
        self._redact_label.clear()

        for step in list(payload.get("steps", [])):
            if not isinstance(step, dict):
                continue
            action = str(step.get("action", ""))
            if action == "compress":
                self._compress.setChecked(True)
            elif action == "extract_text":
                self._extract_text.setChecked(True)
            elif action == "extract_llm":
                self._extract_llm.setChecked(True)
                self._llm_chunk_size.setValue(int(step.get("chunk_size", 1200)))
                self._llm_overlap.setValue(int(step.get("overlap", 200)))
                self._llm_include_page_markers.setChecked(bool(step.get("include_page_markers", True)))
                self._llm_include_metadata.setChecked(bool(step.get("include_metadata", True)))
            elif action == "analyze_llm":
                self._analyze_llm.setChecked(True)
                preset_index = self._llm_preset.findData(step.get("preset", "summary"))
                self._llm_preset.setCurrentIndex(preset_index if preset_index >= 0 else 0)
                self._llm_question.setText(str(step.get("question", "") or ""))
                self._llm_model.setText(str(step.get("model", "gpt-5-mini")))
            elif action == "render":
                self._render.setChecked(True)
                self._render_dpi.setValue(int(step.get("dpi", 150)))
                render_index = self._render_format.findData(step.get("image_format", "png"))
                self._render_format.setCurrentIndex(render_index if render_index >= 0 else 0)
            elif action == "ocr":
                self._ocr.setChecked(True)
                self._ocr_language.setText(str(step.get("language", "eng")))
                self._ocr_skip_text.setChecked(bool(step.get("skip_existing_text", True)))
                self._ocr_force.setChecked(bool(step.get("force", False)))
            elif action == "tables_extract":
                self._tables.setChecked(True)
                table_index = self._tables_format.findData(step.get("format", "csv"))
                self._tables_format.setCurrentIndex(table_index if table_index >= 0 else 0)
                self._tables_ocr_first.setChecked(bool(step.get("ocr_first", False)))
            elif action == "set_metadata":
                self._metadata.setChecked(True)
                values_map = step.get("values", {})
                self._metadata_values.set_value(values_map if isinstance(values_map, dict) else {})
            elif action == "redact":
                self._redact.setChecked(True)
                patterns = step.get("patterns", [])
                self._redact_patterns.set_value(patterns if isinstance(patterns, list) else [])
                self._redact_regex.setChecked(bool(step.get("regex", False)))
                self._redact_case_sensitive.setChecked(bool(step.get("case_sensitive", False)))
                self._redact_pages.setText(str(step.get("pages", "") or ""))
                self._redact_label.setText(str(step.get("label", "") or ""))


class RedactionPreviewLabel(QLabel):
    box_drawn = Signal(int, float, float, float, float)

    def __init__(self) -> None:
        super().__init__()
        self.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self._pixmap: QPixmap | None = None
        self._start: QPoint | None = None
        self._current: QRect | None = None
        self._page_number = 1
        self._page_rect: fitz.Rect | None = None

    def set_page(self, pixmap: QPixmap | None, page_number: int, page_rect: fitz.Rect | None) -> None:
        self._pixmap = pixmap
        self._page_number = page_number
        self._page_rect = page_rect
        self._current = None
        self.setPixmap(pixmap or QPixmap())

    def mousePressEvent(self, event) -> None:  # pragma: no cover - GUI interaction
        if self._pixmap is None or event.button() != Qt.MouseButton.LeftButton:
            return
        self._start = event.position().toPoint()
        self._current = QRect(self._start, self._start)
        self.update()

    def mouseMoveEvent(self, event) -> None:  # pragma: no cover - GUI interaction
        if self._start is None:
            return
        self._current = QRect(self._start, event.position().toPoint()).normalized()
        self.update()

    def mouseReleaseEvent(self, event) -> None:  # pragma: no cover - GUI interaction
        if self._start is None or self._current is None or self._pixmap is None or self._page_rect is None:
            return
        rect = self._current.normalized()
        if rect.width() < 4 or rect.height() < 4:
            self._start = None
            self._current = None
            self.update()
            return
        scale_x = self._page_rect.width / max(self._pixmap.width(), 1)
        scale_y = self._page_rect.height / max(self._pixmap.height(), 1)
        x1 = rect.left() * scale_x
        y1 = rect.top() * scale_y
        x2 = rect.right() * scale_x
        y2 = rect.bottom() * scale_y
        self.box_drawn.emit(self._page_number, x1, y1, x2, y2)
        self._start = None
        self._current = None
        self.update()

    def paintEvent(self, event) -> None:  # pragma: no cover - GUI interaction
        super().paintEvent(event)
        if self._current is None:
            return
        painter = QPainter(self)
        painter.setPen(QPen(QColor("#d33f49"), 2))
        painter.drawRect(self._current)


class RedactionBoxEditor(QWidget):
    changed = Signal()

    def __init__(self) -> None:
        super().__init__()
        self._source_path: Path | None = None
        self._doc: fitz.Document | None = None
        self._preview = RedactionPreviewLabel()
        self._preview.box_drawn.connect(self._append_box)
        self._page_picker = QSpinBox()
        self._page_picker.setMinimum(1)
        self._page_picker.valueChanged.connect(self._render_page)
        self._table = QTableWidget(0, 5)
        self._table.setHorizontalHeaderLabels(["Page", "x1", "y1", "x2", "y2"])
        self._table.horizontalHeader().setStretchLastSection(True)
        self._table.itemChanged.connect(lambda *_: self.changed.emit())
        remove_button = QPushButton("Remove Selected")
        remove_button.clicked.connect(self._remove_selected)
        controls = QHBoxLayout()
        controls.addWidget(QLabel("Page"))
        controls.addWidget(self._page_picker)
        controls.addWidget(remove_button)
        controls.addStretch(1)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addLayout(controls)
        layout.addWidget(self._preview, 1)
        layout.addWidget(self._table)

    def set_source_pdf(self, path: str | Path | None) -> None:
        self._source_path = Path(path) if path else None
        if self._doc is not None:
            self._doc.close()
            self._doc = None
        if self._source_path and self._source_path.exists():
            try:
                self._doc = fitz.open(self._source_path)
            except Exception:
                self._doc = None
        page_count = self._doc.page_count if self._doc is not None else 1
        self._page_picker.setMaximum(max(page_count, 1))
        self._page_picker.setValue(1)
        self._render_page()

    def value(self) -> list[str]:
        values: list[str] = []
        for row in range(self._table.rowCount()):
            parts = []
            for column in range(5):
                item = self._table.item(row, column)
                parts.append(item.text() if item else "")
            values.append(",".join(parts))
        return values

    def set_value(self, values: list[str] | None) -> None:
        self._table.setRowCount(0)
        for value in values or []:
            parts = [part.strip() for part in value.split(",")]
            if len(parts) == 5:
                self._append_row(parts)

    def _render_page(self) -> None:
        if self._doc is None:
            self._preview.set_page(None, self._page_picker.value(), None)
            return
        page_index = self._page_picker.value() - 1
        pixmap = _pixmap_from_fitz(self._doc, page_index, 1.25)
        page = self._doc.load_page(page_index)
        self._preview.set_page(pixmap, page_index + 1, page.rect)

    def _append_box(self, page_number: int, x1: float, y1: float, x2: float, y2: float) -> None:
        self._append_row([str(page_number), f"{x1:.2f}", f"{y1:.2f}", f"{x2:.2f}", f"{y2:.2f}"])
        self.changed.emit()

    def _append_row(self, values: list[str]) -> None:
        row = self._table.rowCount()
        self._table.insertRow(row)
        for column, value in enumerate(values):
            self._table.setItem(row, column, QTableWidgetItem(value))

    def _remove_selected(self) -> None:
        rows = {item.row() for item in self._table.selectedItems()}
        for row in sorted(rows, reverse=True):
            self._table.removeRow(row)
        self.changed.emit()


class PdfPreviewWidget(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._doc: fitz.Document | None = None
        self._thumbs = QListWidget()
        self._thumbs.setIconSize(QSize(96, 128))
        self._thumbs.currentRowChanged.connect(self._show_page)
        self._canvas = QLabel("No preview")
        self._canvas.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._canvas.setFrameShape(QFrame.Shape.StyledPanel)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self._canvas)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._thumbs, 0)
        layout.addWidget(scroll, 1)

    def clear(self) -> None:
        if self._doc is not None:
            self._doc.close()
            self._doc = None
        self._thumbs.clear()
        self._canvas.setText("No preview")
        self._canvas.setPixmap(QPixmap())

    def load_pdf(self, path: Path | None) -> None:
        self.clear()
        if path is None or not path.exists():
            return
        try:
            self._doc = fitz.open(path)
        except Exception:
            self._canvas.setText(f"Unable to preview {path.name}")
            return
        for index in range(self._doc.page_count):
            icon_pixmap = _pixmap_from_fitz(self._doc, index, 0.2)
            item = QListWidgetItem(f"Page {index + 1}")
            item.setIcon(QIcon(icon_pixmap))
            self._thumbs.addItem(item)
        if self._thumbs.count():
            self._thumbs.setCurrentRow(0)

    def _show_page(self, row: int) -> None:
        if self._doc is None or row < 0 or row >= self._doc.page_count:
            return
        pixmap = _pixmap_from_fitz(self._doc, row, 1.2)
        self._canvas.setPixmap(pixmap)
        self._canvas.resize(pixmap.size())


class StructuredDetailsWidget(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self._summary = QLabel("No results yet")
        self._summary.setObjectName("SummaryPill")
        self._table = QTableWidget()
        self._table.horizontalHeader().setStretchLastSection(True)
        self._text = QPlainTextEdit()
        self._text.setReadOnly(True)
        self._tabs = QTabWidget()
        self._tabs.addTab(self._table, "Table")
        self._tabs.addTab(self._text, "JSON")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._summary)
        layout.addWidget(self._tabs)

    def set_payload(self, payload: dict[str, object]) -> None:
        table_data = self._extract_table(payload)
        self._summary.setText(self._build_summary(payload))
        if table_data is None:
            self._table.setRowCount(0)
            self._table.setColumnCount(0)
        else:
            headers, rows = table_data
            self._table.setColumnCount(len(headers))
            self._table.setHorizontalHeaderLabels(headers)
            self._table.setRowCount(len(rows))
            for row_index, row in enumerate(rows):
                for column_index, value in enumerate(row):
                    self._table.setItem(row_index, column_index, QTableWidgetItem(str(value)))
        self._text.setPlainText(_pretty_json(payload))

    def _extract_table(self, payload: dict[str, object]) -> tuple[list[str], list[list[object]]] | None:
        for key in ("statuses", "pages", "attachments", "fields", "jobs", "tables", "groups"):
            value = payload.get(key)
            if isinstance(value, list) and value and isinstance(value[0], dict):
                headers = list(value[0].keys())
                return headers, [[row.get(header, "") for header in headers] for row in value]
        if isinstance(payload.get("bookmarks"), list):
            return ["bookmark"], [[item] for item in payload["bookmarks"]]
        return None

    def _build_summary(self, payload: dict[str, object]) -> str:
        if not payload:
            return "No results yet"
        if "duplicate_group_count" in payload:
            return (
                f"Scanned {payload.get('scanned_file_count', 0)} PDF(s) · "
                f"Found {payload.get('duplicate_file_count', 0)} duplicate file(s) in "
                f"{payload.get('duplicate_group_count', 0)} group(s) · "
                f"Removed {payload.get('removed_count', 0)}"
            )
        if "page_count" in payload and "summary" in payload:
            return f"{payload.get('page_count', 0)} page(s) · Summary: {payload.get('summary', '')}"
        if "jobs" in payload:
            return f"{len(payload.get('jobs', []))} batch job result(s)"
        if "attachments" in payload:
            return f"{len(payload.get('attachments', []))} attachment(s)"
        if "fields" in payload:
            return f"{len(payload.get('fields', []))} form field(s)"
        if "bookmarks" in payload:
            return f"{len(payload.get('bookmarks', []))} bookmark(s)"
        if "statuses" in payload:
            return f"{len(payload.get('statuses', []))} dependency check(s)"
        return "Operation completed"


class DiagnosticsWidget(QTableWidget):
    def __init__(self) -> None:
        super().__init__(0, 5)
        self.setHorizontalHeaderLabels(["Name", "Type", "Level", "Status", "Remediation"])
        self.horizontalHeader().setStretchLastSection(True)
        self.setAlternatingRowColors(True)
        self.verticalHeader().setVisible(False)

    def refresh(self) -> None:
        statuses = collect_doctor_status("all")
        self.setRowCount(len(statuses))
        for row, status in enumerate(statuses):
            items = [
                QTableWidgetItem(status.name),
                QTableWidgetItem(status.category),
                QTableWidgetItem("core" if status.required else "optional"),
                QTableWidgetItem("ready" if status.available else ("optional add-on" if not status.required else "missing")),
                QTableWidgetItem("" if status.available else status.remediation),
            ]
            if status.available:
                status_color = QColor("#7cff6b")
                row_background = QColor("#142632" if row % 2 == 0 else "#18303d")
            elif status.required:
                status_color = QColor("#ff9b71")
                row_background = QColor("#2a1820" if row % 2 == 0 else "#341d26")
            else:
                status_color = QColor("#ffd166")
                row_background = QColor("#302715" if row % 2 == 0 else "#3a301b")
            for column, item in enumerate(items):
                item.setBackground(row_background)
                item.setForeground(QColor("#e8f6ff"))
                self.setItem(row, column, item)
            items[3].setForeground(status_color)


class StartHerePanel(QFrame):
    workflow_requested = Signal(str)
    template_requested = Signal(str)
    pinned_requested = Signal(str, str)
    template_pin_requested = Signal(str)
    repeat_requested = Signal(bool)
    intake_paths_dropped = Signal(list)

    def __init__(self, templates: list[WorkflowTemplate], settings: QSettings) -> None:
        super().__init__()
        self.setObjectName("WelcomePanel")
        self._settings = settings
        self._templates = templates
        self._template_map = {template.id: template for template in templates}
        self._pinned_items: list[dict[str, str]] = []
        self._current_template_id: str | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 22, 22, 22)
        layout.setSpacing(14)

        eyebrow = QLabel("Start Here")
        eyebrow.setObjectName("WelcomeEyebrow")
        title = QLabel("Get a real result in under a minute")
        title.setObjectName("WelcomeTitle")
        self.summary = QLabel(WELCOME_COPY)
        self.summary.setWordWrap(True)
        self.summary.setObjectName("MutedLabel")
        self._readiness = QLabel("Ready now")
        self._readiness.setObjectName("ResultPill")
        self._readiness_detail = QLabel("Core PDF tools are ready. Optional add-ons can be installed later.")
        self._readiness_detail.setWordWrap(True)
        self._readiness_detail.setObjectName("MutedLabel")

        layout.addWidget(eyebrow)
        layout.addWidget(title)
        layout.addWidget(self.summary)
        layout.addWidget(self._readiness)
        layout.addWidget(self._readiness_detail)

        self._intake_zone = IntakeDropZone(
            settings,
            "Quick Intake",
            "Drop one or more PDFs here to get a suggested task, default outputs, and a fast starting point.",
        )
        self._intake_zone.files_dropped.connect(self.intake_paths_dropped.emit)
        layout.addWidget(self._intake_zone)

        preflight_frame = QFrame()
        preflight_frame.setObjectName("InnerSection")
        preflight_layout = QVBoxLayout(preflight_frame)
        preflight_layout.setContentsMargins(14, 14, 14, 14)
        preflight_layout.setSpacing(8)
        preflight_title = QLabel("Suggested Next Task")
        preflight_title.setObjectName("FieldLabel")
        self._suggested_task = QLabel("Drop PDFs to get a suggested task.")
        self._suggested_task.setObjectName("SummaryPill")
        self._preflight_summary = QLabel("No intake analyzed yet.")
        self._preflight_summary.setWordWrap(True)
        self._preflight_summary.setObjectName("MutedLabel")
        self._preflight_warning = QLabel("")
        self._preflight_warning.setWordWrap(True)
        self._preflight_warning.setObjectName("WelcomeNote")
        self._preflight_warning.setVisible(False)
        self._use_suggested_button = QPushButton("Use Suggested Task")
        self._use_suggested_button.clicked.connect(self._emit_suggested_task)
        preflight_layout.addWidget(preflight_title)
        preflight_layout.addWidget(self._suggested_task)
        preflight_layout.addWidget(self._preflight_summary)
        preflight_layout.addWidget(self._preflight_warning)
        preflight_layout.addWidget(self._use_suggested_button, 0, Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(preflight_frame)

        quick_row = QHBoxLayout()
        quick_row.setSpacing(8)
        for label, operation_id in (
            ("Combine PDFs", "merge"),
            ("Split Pages", "split"),
            ("Redact PII", "redact"),
            ("Export Text", "extract-text"),
            ("Export Tables", "tables-extract"),
            ("Folder Workflow", "batch-run"),
        ):
            button = QPushButton(label)
            button.setProperty("secondary", True)
            button.clicked.connect(lambda _checked=False, op=operation_id: self.workflow_requested.emit(op))
            quick_row.addWidget(button)
        quick_row.addStretch(1)
        layout.addLayout(quick_row)

        template_frame = QFrame()
        template_frame.setObjectName("InnerSection")
        template_layout = QHBoxLayout(template_frame)
        template_layout.setContentsMargins(14, 14, 14, 14)
        template_layout.setSpacing(14)
        self._template_list = QListWidget()
        self._template_list.setMaximumWidth(250)
        self._template_list.currentRowChanged.connect(self._update_template_details)
        for template in self._templates:
            self._template_list.addItem(template.label)
        template_layout.addWidget(self._template_list)

        detail_column = QVBoxLayout()
        detail_column.setContentsMargins(0, 0, 0, 0)
        detail_column.setSpacing(8)
        self._template_eyebrow = QLabel("Workflow Template")
        self._template_eyebrow.setObjectName("CardEyebrow")
        self._template_title = QLabel()
        self._template_title.setObjectName("SectionTitle")
        self._template_description = QLabel()
        self._template_description.setWordWrap(True)
        self._template_description.setObjectName("MutedLabel")
        self._template_hint = QLabel()
        self._template_hint.setWordWrap(True)
        self._template_hint.setObjectName("MetricPill")
        self._template_note = QLabel()
        self._template_note.setWordWrap(True)
        self._template_note.setObjectName("WelcomeNote")
        self._template_note.setVisible(False)
        self._template_pin_button = QPushButton("Pin Template")
        self._template_pin_button.setProperty("secondary", True)
        self._template_pin_button.clicked.connect(self._toggle_current_template_pin)
        included_label = QLabel("Included now: merge, split, preview, extraction, tables, metadata, and automation.")
        included_label.setWordWrap(True)
        included_label.setObjectName("MutedLabel")
        optional_label = QLabel("Optional add-ons: OCR tools and OpenAI-powered LLM analysis.")
        optional_label.setWordWrap(True)
        optional_label.setObjectName("MutedLabel")
        self._template_apply_button = QPushButton("Use Template")
        self._template_apply_button.clicked.connect(self._emit_current_template)
        detail_column.addWidget(self._template_eyebrow)
        detail_column.addWidget(self._template_title)
        detail_column.addWidget(self._template_description)
        detail_column.addWidget(self._template_hint)
        detail_column.addWidget(self._template_note)
        detail_column.addWidget(self._template_pin_button)
        detail_column.addWidget(included_label)
        detail_column.addWidget(optional_label)
        detail_column.addWidget(self._template_apply_button)
        detail_column.addStretch(1)
        template_layout.addLayout(detail_column, 1)
        layout.addWidget(template_frame)

        pinned_frame = QFrame()
        pinned_frame.setObjectName("InnerSection")
        pinned_layout = QVBoxLayout(pinned_frame)
        pinned_layout.setContentsMargins(14, 14, 14, 14)
        pinned_layout.setSpacing(8)
        pinned_title = QLabel("Pinned Workflows")
        pinned_title.setObjectName("FieldLabel")
        self._pinned_empty = QLabel("Pin a template or task to keep your most-used workflows one click away.")
        self._pinned_empty.setWordWrap(True)
        self._pinned_empty.setObjectName("MutedLabel")
        self._pinned_buttons = QHBoxLayout()
        self._pinned_buttons.setContentsMargins(0, 0, 0, 0)
        self._pinned_buttons.setSpacing(8)
        pinned_layout.addWidget(pinned_title)
        pinned_layout.addWidget(self._pinned_empty)
        pinned_layout.addLayout(self._pinned_buttons)
        layout.addWidget(pinned_frame)

        recent_frame = QFrame()
        recent_frame.setObjectName("InnerSection")
        recent_layout = QVBoxLayout(recent_frame)
        recent_layout.setContentsMargins(14, 14, 14, 14)
        recent_layout.setSpacing(8)
        recent_title = QLabel("Recent Activity")
        recent_title.setObjectName("FieldLabel")
        self._recent_task = QLabel("No recent task yet. Use a template or choose a workflow to get started.")
        self._recent_task.setWordWrap(True)
        self._recent_task.setObjectName("MutedLabel")
        self._recent_inputs = QLabel("Recent files will appear here after your first successful run.")
        self._recent_inputs.setWordWrap(True)
        self._recent_inputs.setObjectName("MutedLabel")
        repeat_row = QHBoxLayout()
        repeat_row.setContentsMargins(0, 0, 0, 0)
        repeat_row.setSpacing(8)
        self._repeat_same_button = QPushButton("Repeat With Same Inputs")
        self._repeat_same_button.setProperty("secondary", True)
        self._repeat_same_button.setEnabled(False)
        self._repeat_same_button.clicked.connect(lambda: self.repeat_requested.emit(True))
        self._repeat_new_button = QPushButton("Repeat With New Inputs")
        self._repeat_new_button.setProperty("secondary", True)
        self._repeat_new_button.setEnabled(False)
        self._repeat_new_button.clicked.connect(lambda: self.repeat_requested.emit(False))
        repeat_row.addWidget(self._repeat_same_button)
        repeat_row.addWidget(self._repeat_new_button)
        repeat_row.addStretch(1)
        self._recent_output_links = QHBoxLayout()
        self._recent_output_links.setContentsMargins(0, 0, 0, 0)
        self._recent_output_links.setSpacing(8)
        recent_layout.addWidget(recent_title)
        recent_layout.addWidget(self._recent_task)
        recent_layout.addWidget(self._recent_inputs)
        recent_layout.addLayout(repeat_row)
        recent_layout.addLayout(self._recent_output_links)
        layout.addWidget(recent_frame)

        links = QHBoxLayout()
        links.setSpacing(10)
        docs_button = QPushButton("Open Docs")
        docs_button.setProperty("secondary", True)
        docs_button.clicked.connect(lambda: _open_url(DOCS_URL))
        releases_button = QPushButton("Windows Download")
        releases_button.clicked.connect(lambda: _open_url(RELEASES_URL))
        links.addWidget(docs_button)
        links.addWidget(releases_button)
        links.addStretch(1)
        layout.addLayout(links)

        if self._templates:
            self._template_list.setCurrentRow(0)

    def _emit_suggested_task(self) -> None:
        operation_id = self._use_suggested_button.property("operationId")
        if operation_id:
            self.workflow_requested.emit(str(operation_id))

    def _emit_current_template(self) -> None:
        current_row = self._template_list.currentRow()
        if current_row < 0 or current_row >= len(self._templates):
            return
        self.template_requested.emit(self._templates[current_row].id)

    def _toggle_current_template_pin(self) -> None:
        if self._current_template_id:
            self.template_pin_requested.emit(self._current_template_id)

    def _update_template_details(self, row: int) -> None:
        if row < 0 or row >= len(self._templates):
            self._current_template_id = None
            self._template_title.clear()
            self._template_description.clear()
            self._template_hint.clear()
            self._template_note.clear()
            self._template_note.setVisible(False)
            self._template_apply_button.setEnabled(False)
            return
        template = self._templates[row]
        self._current_template_id = template.id
        self._template_title.setText(template.label)
        self._template_description.setText(template.description)
        self._template_hint.setText(template.output_hint or "Template fills practical defaults for this task.")
        if template.dependency_note:
            self._template_note.setText(template.dependency_note)
            self._template_note.setVisible(True)
        else:
            self._template_note.clear()
            self._template_note.setVisible(False)
        self._template_apply_button.setEnabled(True)
        self._template_apply_button.setText(f"Use {template.label}")
        is_pinned = any(item.get("kind") == "template" and item.get("id") == template.id for item in self._pinned_items)
        self._template_pin_button.setText("Unpin Template" if is_pinned else "Pin Template")

    def select_template(self, template_id: str) -> None:
        for index, template in enumerate(self._templates):
            if template.id == template_id:
                self._template_list.setCurrentRow(index)
                return

    def set_readiness(self, title: str, detail: str) -> None:
        self._readiness.setText(title)
        self._readiness_detail.setText(detail)

    def set_preflight(self, payload: dict[str, Any]) -> None:
        suggested_label = str(payload.get("suggested_label", "Suggested task"))
        suggested_operation = str(payload.get("suggested_operation", "merge"))
        file_count = int(payload.get("file_count", 0) or 0)
        page_count = int(payload.get("page_count", 0) or 0)
        encrypted = bool(payload.get("encrypted"))
        scan_summary = payload.get("scan_summary")
        alternatives = [str(item) for item in payload.get("alternatives", []) if str(item).strip()]
        warnings = [str(item) for item in payload.get("warnings", []) if str(item).strip()]

        parts = [f"{file_count} PDF(s)"]
        if page_count:
            parts.append(f"{page_count} page(s)")
        if scan_summary:
            parts.append(f"scan profile: {scan_summary}")
        if encrypted:
            parts.append("encrypted content detected")
        if alternatives:
            parts.append("other likely tasks: " + ", ".join(alternatives))
        self._suggested_task.setText(suggested_label)
        self._preflight_summary.setText(" · ".join(parts) if parts else "No intake analyzed yet.")
        self._use_suggested_button.setProperty("operationId", suggested_operation)
        self._use_suggested_button.style().unpolish(self._use_suggested_button)
        self._use_suggested_button.style().polish(self._use_suggested_button)
        if warnings:
            self._preflight_warning.setText(" ".join(warnings))
            self._preflight_warning.setVisible(True)
        else:
            self._preflight_warning.clear()
            self._preflight_warning.setVisible(False)

    def set_pinned_items(self, items: list[dict[str, str]]) -> None:
        self._pinned_items = items
        _clear_layout(self._pinned_buttons)
        self._pinned_empty.setVisible(not items)
        for item in items:
            label = str(item.get("label", item.get("id", "Workflow")))
            button = QPushButton(label)
            button.setProperty("secondary", True)
            button.clicked.connect(
                lambda _checked=False, kind=str(item.get("kind", "operation")), target=str(item.get("id", "")): self.pinned_requested.emit(kind, target)
            )
            self._pinned_buttons.addWidget(button)
        self._pinned_buttons.addStretch(1)
        if self._current_template_id:
            current = self._template_map.get(self._current_template_id)
            if current is not None:
                self._template_pin_button.setText(
                    "Unpin Template" if any(item.get("kind") == "template" and item.get("id") == current.id for item in items) else "Pin Template"
                )

    def set_recent_activity(self, latest_run: dict[str, Any] | None, latest_success: dict[str, Any] | None) -> None:
        can_repeat = bool(latest_success)
        self._repeat_same_button.setEnabled(can_repeat)
        self._repeat_new_button.setEnabled(can_repeat)
        _clear_layout(self._recent_output_links)
        if not latest_run:
            self._recent_task.setText("No recent task yet. Use a template or choose a workflow to get started.")
            self._recent_inputs.setText("Recent files will appear here after your first successful run.")
            return

        label = str(latest_run.get("label", "Workflow"))
        status = str(latest_run.get("status", "success"))
        state = "completed" if status == "success" else "needs attention"
        self._recent_task.setText(f"Latest run: {label} {state}.")
        input_paths = [str(path) for path in list(latest_run.get("input_paths", []))[:3]]
        if input_paths:
            input_names = ", ".join(Path(path).name for path in input_paths)
            self._recent_inputs.setText(f"Recent files: {input_names}")
        else:
            self._recent_inputs.setText("No recent input files were captured for the last successful task.")

        output_dirs = latest_success.get("output_dirs", []) if latest_success else latest_run.get("output_dirs", [])
        for folder in list(output_dirs)[:3]:
            folder_path = Path(str(folder))
            button = QPushButton(folder_path.name or str(folder_path))
            button.setProperty("secondary", True)
            button.clicked.connect(lambda _checked=False, target=folder_path: _open_path(target))
            self._recent_output_links.addWidget(button)
        self._recent_output_links.addStretch(1)


class ResultsPanel(QWidget):
    repeat_requested = Signal(bool)

    def __init__(self) -> None:
        super().__init__()
        self._headline = QLabel("No run yet")
        self._headline.setObjectName("ResultTitle")
        self._meta = QLabel("Choose a template or workflow, then run it to populate outputs, diagnostics, and logs.")
        self._meta.setObjectName("MutedLabel")
        self._next_step = QLabel("When a run finishes, the best next action shows here.")
        self._next_step.setWordWrap(True)
        self._next_step.setObjectName("MutedLabel")
        self._status_pill = QLabel("Idle")
        self._status_pill.setObjectName("ResultPill")
        self._count_pill = QLabel("0 outputs")
        self._count_pill.setObjectName("MetricPill")
        self._time_pill = QLabel("0 ms")
        self._time_pill.setObjectName("MetricPill")
        self._primary_output: Path | None = None
        self._output_folder: Path | None = None
        self._report_path: Path | None = None
        self.preview = PdfPreviewWidget()
        self.details = StructuredDetailsWidget()
        self.outputs = QListWidget()
        self.outputs.itemDoubleClicked.connect(lambda item: _open_path(Path(item.text())))
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.diagnostics = DiagnosticsWidget()
        self.tabs = QTabWidget()
        self.tabs.addTab(self.preview, "Preview")
        self.tabs.addTab(self.details, "Details")
        self.tabs.addTab(self.outputs, "Outputs")
        self.tabs.addTab(self.log, "Log")
        self.tabs.addTab(self.diagnostics, "Diagnostics")
        self.outputs.setAlternatingRowColors(True)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        header = QFrame()
        header.setObjectName("ResultsHero")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(18, 18, 18, 18)
        header_layout.setSpacing(10)
        pill_row = QHBoxLayout()
        pill_row.setContentsMargins(0, 0, 0, 0)
        pill_row.setSpacing(8)
        pill_row.addWidget(self._status_pill)
        pill_row.addWidget(self._count_pill)
        pill_row.addWidget(self._time_pill)
        pill_row.addStretch(1)
        action_row = QHBoxLayout()
        action_row.setContentsMargins(0, 0, 0, 0)
        action_row.setSpacing(8)
        self._open_primary_button = QPushButton("Open Primary Output")
        self._open_primary_button.setProperty("secondary", True)
        self._open_primary_button.setEnabled(False)
        self._open_primary_button.clicked.connect(self._open_primary_output)
        self._open_folder_button = QPushButton("Open Output Folder")
        self._open_folder_button.setProperty("secondary", True)
        self._open_folder_button.setEnabled(False)
        self._open_folder_button.clicked.connect(self._open_output_folder)
        self._open_report_button = QPushButton("Open Report")
        self._open_report_button.setProperty("secondary", True)
        self._open_report_button.setEnabled(False)
        self._open_report_button.clicked.connect(self._open_report)
        self._review_diagnostics_button = QPushButton("Review Diagnostics")
        self._review_diagnostics_button.setProperty("secondary", True)
        self._review_diagnostics_button.clicked.connect(lambda: self.tabs.setCurrentWidget(self.diagnostics))
        self._repeat_same_button = QPushButton("Run Another Like This")
        self._repeat_same_button.setProperty("secondary", True)
        self._repeat_same_button.setEnabled(False)
        self._repeat_same_button.clicked.connect(lambda: self.repeat_requested.emit(True))
        self._repeat_new_button = QPushButton("Run With New Inputs")
        self._repeat_new_button.setProperty("secondary", True)
        self._repeat_new_button.setEnabled(False)
        self._repeat_new_button.clicked.connect(lambda: self.repeat_requested.emit(False))
        action_row.addWidget(self._open_primary_button)
        action_row.addWidget(self._open_folder_button)
        action_row.addWidget(self._open_report_button)
        action_row.addWidget(self._review_diagnostics_button)
        action_row.addWidget(self._repeat_same_button)
        action_row.addWidget(self._repeat_new_button)
        action_row.addStretch(1)
        self._quick_links = QHBoxLayout()
        self._quick_links.setContentsMargins(0, 0, 0, 0)
        self._quick_links.setSpacing(8)
        header_layout.addWidget(self._headline)
        header_layout.addWidget(self._meta)
        header_layout.addWidget(self._next_step)
        header_layout.addLayout(pill_row)
        header_layout.addLayout(action_row)
        header_layout.addLayout(self._quick_links)
        layout.addWidget(header)
        layout.addWidget(self.tabs)
        self.diagnostics.refresh()

    def append_log(self, message: str) -> None:
        self.log.appendPlainText(message)

    def _open_primary_output(self) -> None:
        if self._primary_output is not None:
            _open_path(self._primary_output)

    def _open_output_folder(self) -> None:
        if self._output_folder is not None:
            _open_path(self._output_folder)

    def _open_report(self) -> None:
        if self._report_path is not None:
            _open_path(self._report_path)

    def set_result(self, result: JobResult, report_path: Path | None = None, *, display_label: str | None = None) -> None:
        self.details.set_payload(result.details)
        self.outputs.clear()
        _clear_layout(self._quick_links)
        for output in result.outputs:
            self.outputs.addItem(str(output))
        if report_path is not None:
            self.outputs.addItem(str(report_path))
        self._headline.setText(display_label or result.operation_id.replace("-", " ").title())
        self._primary_output = result.outputs[0] if result.outputs else report_path
        self._report_path = report_path
        if result.outputs:
            self._output_folder = result.outputs[0].parent
        elif report_path is not None:
            self._output_folder = report_path.parent
        else:
            self._output_folder = None

        quick_targets = list(result.outputs[:3])
        if report_path is not None and len(quick_targets) < 3:
            quick_targets.append(report_path)
        for target in quick_targets:
            button = QPushButton(Path(target).name)
            button.setProperty("secondary", True)
            button.clicked.connect(lambda _checked=False, path=Path(target): _open_path(path))
            self._quick_links.addWidget(button)
        self._quick_links.addStretch(1)

        if result.status == "success":
            destination = str(self._output_folder) if self._output_folder is not None else "the selected output location"
            self._meta.setText(
                f"Task completed successfully. {len(result.outputs)} output(s) saved"
                + (" plus a report file " if report_path is not None else " ")
                + f"to {destination}."
            )
            self._next_step.setText("Next: open the output folder, inspect the preview/details, or rerun the task with the same setup.")
        else:
            self._meta.setText(result.error or "The last run needs attention.")
            if result.operation_id == "ocr":
                self._next_step.setText("Likely next step: install OCR tools or switch to a core PDF workflow that does not require OCR.")
            elif result.operation_id == "analyze-llm":
                self._next_step.setText("Likely next step: set OPENAI_API_KEY and install the optional llm extras, or keep using local extraction only.")
            else:
                self._next_step.setText("Likely next step: review diagnostics and details, then rerun after fixing the setup.")
        self._status_pill.setText("Completed" if result.status == "success" else "Needs Attention")
        self._count_pill.setText(f"{len(result.outputs)} outputs")
        self._time_pill.setText(f"{result.duration_ms} ms")
        self._open_primary_button.setEnabled(self._primary_output is not None)
        self._open_folder_button.setEnabled(self._output_folder is not None)
        self._open_report_button.setEnabled(report_path is not None)
        self._repeat_same_button.setEnabled(result.status == "success")
        self._repeat_new_button.setEnabled(result.status == "success")
        pdf_output = next((path for path in result.outputs if path.suffix.lower() == ".pdf"), None)
        if pdf_output is not None:
            self.preview.load_pdf(pdf_output)
            self.tabs.setCurrentWidget(self.preview)
        elif result.operation_id == "doctor":
            self.preview.clear()
            self.tabs.setCurrentWidget(self.details)
        elif result.status != "success":
            self.tabs.setCurrentWidget(self.diagnostics)
        elif result.outputs or report_path is not None:
            self.tabs.setCurrentWidget(self.outputs)
        else:
            self.tabs.setCurrentWidget(self.details)
        self.append_log(f"{result.operation_id}: {result.status} ({result.duration_ms} ms)")
        if result.error:
            self.append_log(result.error)


class WorkerSignals(QObject):
    finished = Signal(object, object)


class JobWorker(QRunnable):
    def __init__(self, operation_id: str, values: dict[str, object], *, report_path: str | None, overwrite: bool) -> None:
        super().__init__()
        self.signals = WorkerSignals()
        self._operation_id = operation_id
        self._values = values
        self._report_path = report_path
        self._overwrite = overwrite

    def run(self) -> None:  # pragma: no cover - executed in Qt thread pool
        try:
            request = prepare_request(self._operation_id, self._values, report_path=self._report_path, overwrite=self._overwrite)
            result = execute_job(request)
            self.signals.finished.emit(result, request.report_path)
        except Exception as exc:
            result = JobResult(
                operation_id=self._operation_id,
                status="error",
                outputs=[],
                warnings=[],
                details={},
                error=str(exc),
                duration_ms=0,
            )
            self.signals.finished.emit(result, None)


class WatchController(QObject):
    log = Signal(str)
    changed = Signal(bool)

    def __init__(self) -> None:
        super().__init__()
        self._observer: Observer | None = None

    def active(self) -> bool:
        return self._observer is not None

    def start(self, input_dir: Path, manifest_path: Path, recursive: bool, overwrite: bool) -> None:
        if self._observer is not None:
            return
        config = load_config(Path.cwd())
        self._observer = Observer()
        handler = WatchFolderHandler(
            input_dir,
            manifest_path,
            config,
            overwrite,
            callback=lambda message: self.log.emit(message),
        )
        self._observer.schedule(handler, str(input_dir), recursive=recursive)
        self._observer.start()
        self.log.emit(f"Watching {input_dir}")
        self.changed.emit(True)

    def stop(self) -> None:
        if self._observer is None:
            return
        self._observer.stop()
        self._observer.join(timeout=5)
        self._observer = None
        self.log.emit("Watch stopped")
        self.changed.emit(False)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} {APP_VERSION}")
        self.setWindowIcon(_app_icon())
        self.resize(1680, 980)
        self._settings = QSettings(ORGANIZATION_NAME, "PDFToolkit")
        self._thread_pool = QThreadPool.globalInstance()
        self._watch_controller = WatchController()
        self._watch_controller.log.connect(self._handle_log)
        self._watch_controller.changed.connect(self._sync_watch_state)
        self._definitions = get_operation_definitions()
        self._definition_map = {definition.id: definition for definition in self._definitions}
        self._templates = get_workflow_templates()
        self._operation_items: dict[str, QTreeWidgetItem] = {}
        self._category_items: dict[str, QTreeWidgetItem] = {}
        self._current_definition: OperationDefinition | None = None
        self._field_widgets: dict[str, object] = {}
        self._field_labels: dict[str, QLabel] = {}
        self._watch_active = False
        self._last_request_context: dict[str, Any] | None = None
        self._current_intake_paths: list[str] = []
        self._pinned_workflows = self._load_pinned_workflows()
        self._last_used_values = self._load_last_used_values()
        self._latest_output_folder: Path | None = None
        self._auto_filling = False
        self._recent_runs = self._load_recent_runs()
        self._build_ui()
        self._populate_operation_tree()
        self._select_initial_operation()
        self._refresh_start_here()
        self._update_readiness()

    def _build_ui(self) -> None:
        shell = QWidget()
        shell_layout = QVBoxLayout(shell)
        shell_layout.setContentsMargins(18, 18, 18, 18)
        shell_layout.setSpacing(16)
        shell_layout.addWidget(self._build_shell_header())
        splitter = QSplitter()
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(10)
        splitter.addWidget(self._build_left_pane())
        splitter.addWidget(self._build_center_pane())
        splitter.addWidget(self._build_right_pane())
        splitter.setSizes([320, 560, 800])
        shell_layout.addWidget(splitter, 1)
        self.setCentralWidget(shell)
        self._build_menu()
        refresh_action = QAction("Refresh Diagnostics", self)
        refresh_action.triggered.connect(self._results.diagnostics.refresh)
        self.menuBar().addAction(refresh_action)
        self.statusBar().showMessage(f"{APP_NAME} {APP_VERSION} ready")

    def _build_shell_header(self) -> QWidget:
        header = QFrame()
        header.setObjectName("ShellHeader")
        layout = QHBoxLayout(header)
        layout.setContentsMargins(26, 22, 26, 22)
        layout.setSpacing(22)

        copy_column = QVBoxLayout()
        copy_column.setContentsMargins(0, 0, 0, 0)
        copy_column.setSpacing(8)
        eyebrow = QLabel("Office-grade PDF command center")
        eyebrow.setObjectName("ShellEyebrow")
        title = QLabel(APP_NAME)
        title.setObjectName("ShellTitle")
        tagline = QLabel(APP_TAGLINE)
        tagline.setObjectName("MutedLabel")
        metrics = QHBoxLayout()
        metrics.setContentsMargins(0, 0, 0, 0)
        metrics.setSpacing(8)
        self._header_operation_metric = QLabel("0 workflows")
        self._header_operation_metric.setObjectName("MetricPill")
        self._header_optional_metric = QLabel("OCR + LLM optional")
        self._header_optional_metric.setObjectName("MetricPill")
        self._header_status_metric = QLabel("Desktop ready")
        self._header_status_metric.setObjectName("StatusBadge")
        metrics.addWidget(self._header_operation_metric)
        metrics.addWidget(self._header_optional_metric)
        metrics.addWidget(self._header_status_metric)
        metrics.addStretch(1)
        copy_column.addWidget(eyebrow)
        copy_column.addWidget(title)
        copy_column.addWidget(tagline)
        copy_column.addLayout(metrics)

        action_column = QVBoxLayout()
        action_column.setContentsMargins(0, 0, 0, 0)
        action_column.setSpacing(10)
        links = QHBoxLayout()
        links.setContentsMargins(0, 0, 0, 0)
        links.setSpacing(8)
        for label, callback in (
            ("Documentation", lambda: _open_url(DOCS_URL)),
            ("Windows Releases", lambda: _open_url(RELEASES_URL)),
            ("Project Page", lambda: _open_url(PROJECT_URL)),
        ):
            button = QPushButton(label)
            button.setProperty("secondary", True)
            button.clicked.connect(callback)
            links.addWidget(button)
        shortcuts = QHBoxLayout()
        shortcuts.setContentsMargins(0, 0, 0, 0)
        shortcuts.setSpacing(8)
        for label, operation_id in (
            ("Combine PDFs", "merge"),
            ("Split Pages", "split"),
            ("Export Text", "extract-text"),
            ("Export Tables", "tables-extract"),
            ("Watch Folder", "batch-run"),
        ):
            button = QPushButton(label)
            button.clicked.connect(lambda _checked=False, op=operation_id: self._select_operation(op))
            shortcuts.addWidget(button)
        action_column.addLayout(links)
        action_column.addLayout(shortcuts)

        layout.addLayout(copy_column, 1)
        layout.addLayout(action_column)
        return header

    def _build_menu(self) -> None:
        help_menu = self.menuBar().addMenu("&Help")
        self._docs_action = QAction("Documentation", self)
        self._docs_action.triggered.connect(lambda: _open_url(DOCS_URL))
        self._releases_action = QAction("Windows Releases", self)
        self._releases_action.triggered.connect(lambda: _open_url(RELEASES_URL))
        self._project_action = QAction("Project Page", self)
        self._project_action.triggered.connect(lambda: _open_url(PROJECT_URL))
        self._about_action = QAction("About PDF Toolkit", self)
        self._about_action.triggered.connect(self._show_about_dialog)
        help_menu.addAction(self._docs_action)
        help_menu.addAction(self._releases_action)
        help_menu.addAction(self._project_action)
        help_menu.addSeparator()
        help_menu.addAction(self._about_action)

    def _build_left_pane(self) -> QWidget:
        self._operation_search = QLineEdit()
        self._operation_search.setPlaceholderText("Search tasks and workflows")
        self._operation_search.textChanged.connect(self._filter_operations)
        self._tree = QTreeWidget()
        self._tree.setHeaderHidden(True)
        self._tree.setRootIsDecorated(True)
        self._tree.setUniformRowHeights(True)
        self._tree.setIndentation(18)
        self._tree.itemSelectionChanged.connect(self._handle_operation_selection)
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        helper = QLabel("Browse every task by category, or start from a bundled template to prefill the next job.")
        helper.setWordWrap(True)
        helper.setObjectName("MutedLabel")
        layout.addWidget(helper)
        layout.addWidget(self._operation_search)
        layout.addWidget(self._tree)
        self._operation_count = QLabel()
        self._operation_count.setObjectName("MutedLabel")
        layout.addWidget(self._operation_count)
        return _panel_widget("Operations", body, accent="#d46a3a")

    def _build_center_pane(self) -> QWidget:
        self._start_here_panel = StartHerePanel(self._templates, self._settings)
        self._start_here_panel.workflow_requested.connect(self._select_operation)
        self._start_here_panel.template_requested.connect(self._apply_template)
        self._start_here_panel.repeat_requested.connect(self._repeat_last_task)
        self._start_here_panel.template_pin_requested.connect(self._toggle_template_pin)
        self._start_here_panel.pinned_requested.connect(self._open_pinned_workflow)
        self._start_here_panel.intake_paths_dropped.connect(self._ingest_paths)
        self._welcome_panel = self._start_here_panel
        self._operation_eyebrow = QLabel("Selected Workflow")
        self._operation_eyebrow.setObjectName("CardEyebrow")
        self._title = QLabel()
        self._title.setObjectName("HeroTitle")
        self._description = QLabel()
        self._description.setWordWrap(True)
        self._description.setObjectName("MutedLabel")
        self._summary = QLabel("Choose a task or template, confirm the destination, then run it.")
        self._summary.setObjectName("SummaryPill")
        self._mode_chip = QLabel("Writes files")
        self._mode_chip.setObjectName("MetricPill")
        self._support_chip = QLabel("Preview ready")
        self._support_chip.setObjectName("MetricPill")
        self._report_chip = QLabel("Report export")
        self._report_chip.setObjectName("MetricPill")
        self._pin_operation_button = QPushButton("Pin Task")
        self._pin_operation_button.setProperty("secondary", True)
        self._pin_operation_button.clicked.connect(self._toggle_current_operation_pin)
        self._form_host = QWidget()
        self._form_host.setObjectName("FormHost")
        self._form_layout = QFormLayout(self._form_host)
        self._form_layout.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        self._form_layout.setContentsMargins(12, 12, 12, 12)
        self._form_layout.setHorizontalSpacing(18)
        self._form_layout.setVerticalSpacing(14)
        self._report_label = QLabel("Report Path")
        self._report_input = SinglePathInput(self._settings, save_mode=True)
        self._report_input.changed.connect(self._update_ready_summary)
        self._overwrite = QCheckBox("Overwrite existing outputs")
        self._overwrite.stateChanged.connect(lambda *_: self._update_ready_summary())
        self._validate_button = QPushButton("Validate Setup")
        self._validate_button.setProperty("secondary", True)
        self._validate_button.clicked.connect(lambda: self._update_ready_summary(notify=True))
        self._run_button = QPushButton("Run")
        self._run_button.clicked.connect(self._run_current_operation)
        self._stop_watch_button = QPushButton("Stop Watch")
        self._stop_watch_button.clicked.connect(self._watch_controller.stop)
        self._stop_watch_button.setEnabled(False)
        self._footer_open_output_button = QPushButton("Open Output Folder")
        self._footer_open_output_button.setProperty("secondary", True)
        self._footer_open_output_button.setEnabled(False)
        self._footer_open_output_button.clicked.connect(self._open_latest_output_folder)
        controls = QHBoxLayout()
        controls.addWidget(self._overwrite)
        controls.addStretch(1)
        controls.addWidget(self._footer_open_output_button)
        controls.addWidget(self._validate_button)
        controls.addWidget(self._run_button)
        controls.addWidget(self._stop_watch_button)
        scroll_content = QWidget()
        scroll_content.setObjectName("WorkspaceScrollContent")
        scroll_inner = QVBoxLayout(scroll_content)
        scroll_inner.setContentsMargins(0, 0, 0, 0)
        scroll_inner.setSpacing(14)
        operation_card = QFrame()
        operation_card.setObjectName("OperationHero")
        operation_layout = QVBoxLayout(operation_card)
        operation_layout.setContentsMargins(20, 20, 20, 20)
        operation_layout.setSpacing(10)
        chip_row = QHBoxLayout()
        chip_row.setContentsMargins(0, 0, 0, 0)
        chip_row.setSpacing(8)
        chip_row.addWidget(self._mode_chip)
        chip_row.addWidget(self._support_chip)
        chip_row.addWidget(self._report_chip)
        chip_row.addWidget(self._pin_operation_button)
        chip_row.addStretch(1)
        operation_layout.addWidget(self._operation_eyebrow)
        operation_layout.addWidget(self._title)
        operation_layout.addWidget(self._description)
        operation_layout.addLayout(chip_row)
        operation_layout.addWidget(self._summary)
        self._workspace_intake_zone = IntakeDropZone(
            self._settings,
            "Workspace Intake",
            "Drop PDFs here at any time to prefill the current task or switch to the best next workflow.",
            button_label="Choose Inputs",
        )
        self._workspace_intake_zone.files_dropped.connect(self._ingest_paths)
        self._preflight_card = QFrame()
        self._preflight_card.setObjectName("InnerSection")
        preflight_layout = QVBoxLayout(self._preflight_card)
        preflight_layout.setContentsMargins(14, 14, 14, 14)
        preflight_layout.setSpacing(8)
        preflight_title = QLabel("Intake Preflight")
        preflight_title.setObjectName("FieldLabel")
        self._preflight_metric = QLabel("No intake analyzed yet.")
        self._preflight_metric.setObjectName("MetricPill")
        self._preflight_detail = QLabel("Drop files to see page count, scan profile, encryption hints, and the best suggested task.")
        self._preflight_detail.setWordWrap(True)
        self._preflight_detail.setObjectName("MutedLabel")
        preflight_layout.addWidget(preflight_title)
        preflight_layout.addWidget(self._preflight_metric)
        preflight_layout.addWidget(self._preflight_detail)
        self._parameter_header = QLabel("Operation Parameters")
        self._parameter_header.setObjectName("SectionTitle")
        self._parameter_header.setProperty("workspaceRole", "parameter-header")
        self._parameter_caption = QLabel("Confirm the source files, output location, and practical defaults before you run the task.")
        self._parameter_caption.setObjectName("MutedLabel")
        self._parameter_caption.setProperty("workspaceRole", "parameter-caption")
        scroll_inner.addWidget(self._start_here_panel)
        scroll_inner.addWidget(operation_card)
        scroll_inner.addWidget(self._workspace_intake_zone)
        scroll_inner.addWidget(self._preflight_card)
        scroll_inner.addWidget(self._parameter_header)
        scroll_inner.addWidget(self._parameter_caption)
        scroll_inner.addWidget(self._form_host)
        scroll_inner.addStretch(1)
        scroll = QScrollArea()
        scroll.setObjectName("FormScroll")
        scroll.setWidgetResizable(True)
        scroll.setWidget(scroll_content)
        footer_card = QFrame()
        footer_card.setObjectName("ActionBarCard")
        footer_layout = QVBoxLayout(footer_card)
        footer_layout.setContentsMargins(18, 18, 18, 18)
        footer_layout.setSpacing(10)
        ready_title = QLabel("Ready To Run")
        ready_title.setObjectName("FieldLabel")
        self._ready_summary = QLabel("Choose a task or template, confirm the destination, then run it.")
        self._ready_summary.setObjectName("SummaryPill")
        self._ready_warning = QLabel("Validation and optional add-on guidance will appear here.")
        self._ready_warning.setWordWrap(True)
        self._ready_warning.setObjectName("MutedLabel")
        footer_layout.addWidget(ready_title)
        footer_layout.addWidget(self._ready_summary)
        footer_layout.addWidget(self._ready_warning)
        footer_layout.addWidget(self._report_label)
        footer_layout.addWidget(self._report_input)
        footer_layout.addLayout(controls)
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)
        layout.addWidget(scroll, 1)
        layout.addWidget(footer_card)
        return _panel_widget("Workspace", body, accent="#59b7a7")

    def _build_right_pane(self) -> QWidget:
        self._results = ResultsPanel()
        self._results.repeat_requested.connect(self._repeat_last_task)
        return _panel_widget("Results Deck", self._results, accent="#e08b49")

    def _populate_operation_tree(self) -> None:
        for definition in self._definitions:
            parent = self._category_items.get(definition.category)
            if parent is None:
                parent = QTreeWidgetItem([definition.category])
                parent.setFlags(parent.flags() & ~Qt.ItemFlag.ItemIsSelectable)
                category_font = QFont("Segoe UI", 10)
                category_font.setBold(True)
                parent.setFont(0, category_font)
                self._category_items[definition.category] = parent
                self._tree.addTopLevelItem(parent)
            child = QTreeWidgetItem([definition.label])
            child.setData(0, Qt.ItemDataRole.UserRole, definition.id)
            parent.addChild(child)
            self._operation_items[definition.id] = child
        self._tree.expandAll()
        self._update_operation_count()

    def _select_initial_operation(self) -> None:
        if not self._definitions:
            return
        first = self._operation_items.get(self._definitions[0].id)
        if first is not None:
            self._tree.setCurrentItem(first)
            self._set_current_operation(self._definitions[0].id)

    def _select_operation(self, operation_id: str, *, restore_saved_values: bool = True) -> None:
        item = self._operation_items.get(operation_id)
        if item is None:
            return
        self._tree.setCurrentItem(item)
        self._set_current_operation(operation_id, restore_saved_values=restore_saved_values)

    def _handle_operation_selection(self) -> None:
        item = self._tree.currentItem()
        if item is None:
            return
        operation_id = item.data(0, Qt.ItemDataRole.UserRole)
        if operation_id:
            self._set_current_operation(str(operation_id))

    def _set_current_operation(self, operation_id: str, *, restore_saved_values: bool = True) -> None:
        definition = self._definition_map[operation_id]
        self._current_definition = definition
        self._operation_eyebrow.setText("Selected Task")
        self._title.setText(definition.label)
        self._description.setText(definition.description)
        mode_text = "Creates output files" if definition.mutating else "Read-only inspection"
        optional_text = ""
        if definition.id == "ocr":
            optional_text = " OCR add-on optional."
        elif definition.id == "analyze-llm":
            optional_text = " OpenAI API key optional."
        self._summary.setText(f"{definition.category} task. {mode_text}.{optional_text}")
        self._mode_chip.setText(mode_text)
        self._support_chip.setText("Preview ready" if definition.supports_preview else "No preview")
        self._report_chip.setText("Report export" if definition.supports_report else "No report file")
        self._run_button.setText("Start Watch" if definition.id == "watch-folder" else "Run Task")
        show_global_report = definition.supports_report and definition.id != "batch-run"
        self._report_label.setVisible(show_global_report)
        self._report_input.setVisible(show_global_report)
        if not show_global_report:
            self._report_input.set_value(None)
        self._rebuild_form(definition)
        self._update_parameter_copy(definition)
        self._update_pin_buttons()
        if self._current_intake_paths:
            self._apply_intake_to_current_operation(self._current_intake_paths)
        elif restore_saved_values:
            self._restore_last_used_values(definition.id)
        self._results.preview.clear()
        self._results.details.set_payload({})
        self._update_ready_summary()
        self.statusBar().showMessage(f"Selected {definition.label}")

    def _rebuild_form(self, definition: OperationDefinition) -> None:
        while self._form_layout.rowCount():
            self._form_layout.removeRow(0)
        self._field_widgets.clear()
        self._field_labels.clear()
        if definition.id == "batch-run":
            builder = FolderBatchBuilder(self._settings)
            builder.changed.connect(self._refresh_preview)
            self._field_widgets["manifest_path"] = builder
            label = QLabel("Template Or Batch Setup")
            label.setObjectName("FieldLabel")
            self._field_labels["manifest_path"] = label
            self._form_layout.addRow(label, builder)
            return
        for field_def in definition.fields:
            widget = self._build_field_widget(field_def)
            self._field_widgets[field_def.name] = widget
            label = QLabel(field_def.label + (" *" if field_def.required else ""))
            label.setObjectName("FieldLabel")
            self._field_labels[field_def.name] = label
            if field_def.help:
                label.setToolTip(field_def.help)
                widget.setToolTip(field_def.help)
            self._form_layout.addRow(label, widget)

    def _build_field_widget(self, field_def: OperationField) -> QWidget:
        if field_def.kind == "file":
            widget = MultiPathInput(self._settings) if field_def.multiple else SinglePathInput(self._settings, save_mode=field_def.path_role == "output")
        elif field_def.kind == "directory":
            widget = SinglePathInput(self._settings, directory=True)
        elif field_def.kind in {"text", "page_spec", "password"} and field_def.multiple:
            widget = MultiLineInput()
        elif field_def.kind == "key_value_list":
            widget = KeyValueTableInput()
        elif field_def.kind == "redaction_boxes":
            widget = RedactionBoxEditor()
        elif field_def.kind == "choice":
            combo = QComboBox()
            for choice in field_def.choices:
                combo.addItem(choice.label, choice.value)
            if field_def.default is not None:
                index = combo.findData(field_def.default)
                combo.setCurrentIndex(max(index, 0))
            combo.currentIndexChanged.connect(self._refresh_preview)
            widget = combo
        elif field_def.kind == "checkbox":
            check = QCheckBox()
            check.setChecked(bool(field_def.default))
            check.stateChanged.connect(self._refresh_preview)
            widget = check
        elif field_def.kind == "number":
            if field_def.number_mode == "float":
                spin = QDoubleSpinBox()
                spin.setDecimals(2)
                spin.setRange(-999999.0, 999999.0)
                spin.setValue(float(field_def.default or 0.0))
            else:
                spin = QSpinBox()
                spin.setRange(-999999, 999999)
                spin.setValue(int(field_def.default or 0))
            spin.valueChanged.connect(self._refresh_preview)
            widget = spin
        else:
            line = QLineEdit()
            if field_def.default is not None:
                line.setText(str(field_def.default))
            if field_def.placeholder:
                line.setPlaceholderText(field_def.placeholder)
            if field_def.kind == "password":
                line.setEchoMode(QLineEdit.EchoMode.Password)
            if field_def.kind == "page_spec" and PAGE_SPEC_VALIDATOR is not None:
                line.setValidator(PAGE_SPEC_VALIDATOR)
            line.textChanged.connect(self._refresh_preview)
            widget = line

        if isinstance(widget, (SinglePathInput, MultiPathInput, MultiLineInput, KeyValueTableInput, RedactionBoxEditor)):
            widget.changed.connect(self._refresh_preview)
        return widget

    def _set_widget_value(self, widget: object, value: Any) -> None:
        if isinstance(widget, (SinglePathInput, MultiPathInput, MultiLineInput, KeyValueTableInput, RedactionBoxEditor, FolderBatchBuilder)):
            widget.set_value(value)
        elif isinstance(widget, QComboBox):
            index = widget.findData(value)
            if index < 0 and value is not None:
                index = widget.findText(str(value))
            if index >= 0:
                widget.setCurrentIndex(index)
        elif isinstance(widget, QCheckBox):
            widget.setChecked(bool(value))
        elif isinstance(widget, (QSpinBox, QDoubleSpinBox)):
            if value not in (None, ""):
                widget.setValue(value)
        elif isinstance(widget, QLineEdit):
            widget.setText("" if value in (None, []) else str(value))

    def _apply_values_to_current_form(self, values: dict[str, Any], *, report_path: str | None = None, refresh: bool = True) -> None:
        for field_name, value in values.items():
            widget = self._field_widgets.get(field_name)
            if widget is None:
                continue
            self._set_widget_value(widget, value)
        if report_path is not None and self._report_input.isVisible():
            self._report_input.set_value(report_path)
        if refresh:
            self._refresh_preview()

    def _apply_template(self, template_id: str) -> None:
        template = get_workflow_template(template_id)
        self._start_here_panel.select_template(template_id)
        self._select_operation(template.operation_id, restore_saved_values=False)
        if template.target == "batch":
            self._apply_values_to_current_form({"manifest_path": template.values})
        else:
            self._apply_values_to_current_form(template.values)
        self._current_intake_paths = []
        self._summary.setText(f"Template loaded: {template.label}. {template.description}")
        self._persist_last_used_snapshot(template.operation_id, self._collect_values(), self._report_input.value())
        self._update_ready_summary()
        self.statusBar().showMessage(f"Loaded template: {template.label}")

    def _filter_operations(self, text: str) -> None:
        query = text.strip().lower()
        for operation_id, item in self._operation_items.items():
            definition = self._definition_map[operation_id]
            visible = not query or query in definition.label.lower() or query in definition.category.lower() or query in definition.description.lower()
            item.setHidden(not visible)
        for category, item in self._category_items.items():
            has_visible_child = any(not item.child(index).isHidden() for index in range(item.childCount()))
            item.setHidden(not has_visible_child)
            if has_visible_child:
                item.setExpanded(True)
        self._update_operation_count()

    def _update_operation_count(self) -> None:
        visible_count = sum(1 for item in self._operation_items.values() if not item.isHidden())
        self._operation_count.setText(f"{visible_count} task{'s' if visible_count != 1 else ''}")
        if hasattr(self, "_header_operation_metric"):
            self._header_operation_metric.setText(f"{visible_count} workflows")

    def _collect_values(self) -> dict[str, object]:
        if self._current_definition is None:
            return {}
        values: dict[str, object] = {}
        for field_def in self._current_definition.fields:
            widget = self._field_widgets[field_def.name]
            if isinstance(widget, (SinglePathInput, MultiPathInput, MultiLineInput, KeyValueTableInput, RedactionBoxEditor, FolderBatchBuilder)):
                values[field_def.name] = widget.value()
            elif isinstance(widget, QComboBox):
                values[field_def.name] = widget.currentData()
            elif isinstance(widget, QCheckBox):
                values[field_def.name] = widget.isChecked()
            elif isinstance(widget, (QSpinBox, QDoubleSpinBox)):
                values[field_def.name] = widget.value()
            elif isinstance(widget, QLineEdit):
                values[field_def.name] = widget.text().strip()
            else:
                values[field_def.name] = None
        return values

    def _refresh_preview(self) -> None:
        definition = self._current_definition
        if not self._auto_filling and self._apply_smart_defaults_from_current_inputs():
            return
        self._update_ready_summary()
        if definition is None or not definition.supports_preview:
            return
        values = self._collect_values()
        preview_value = values.get(definition.preview_field or "")
        preview_path: Path | None = None
        if isinstance(preview_value, list) and preview_value:
            preview_path = Path(preview_value[0])
        elif isinstance(preview_value, str) and preview_value:
            preview_path = Path(preview_value)
        if preview_path is not None and preview_path.exists() and preview_path.suffix.lower() == ".pdf":
            self._results.preview.load_pdf(preview_path)
        if definition.id == "redact":
            editor = self._field_widgets.get("box")
            input_widget = self._field_widgets.get("input_path")
            if isinstance(editor, RedactionBoxEditor) and isinstance(input_widget, SinglePathInput):
                editor.set_source_pdf(input_widget.value())

    def _run_current_operation(self) -> None:
        if self._current_definition is None:
            return
        if not self._update_ready_summary(notify=True):
            return
        values = self._collect_values()
        context_values = _json_safe(values)
        report_path: str | None = self._report_input.value() if self._current_definition.supports_report and self._current_definition.id != "batch-run" else None
        if self._current_definition.id == "batch-run":
            try:
                builder_values = values.get("manifest_path", {})
                assert isinstance(builder_values, dict)
                output_root = builder_values.get("output_root")
                if not output_root:
                    raise ValidationError("Output Folder is required.")
                report_path_value = Path(str(builder_values["report_path"])) if builder_values.get("report_path") else None
                if builder_values.get("source_mode") == "files":
                    input_files = [Path(str(path)) for path in builder_values.get("input_files", []) if str(path).strip()]
                    manifest = build_file_batch_manifest(
                        input_files,
                        Path(str(output_root)),
                        steps=list(builder_values.get("steps", [])),
                        report_path=report_path_value,
                        fail_fast=bool(builder_values.get("fail_fast", False)),
                        job_name=str(builder_values.get("job_name", "selected-files-batch")),
                    )
                else:
                    input_root = builder_values.get("input_root")
                    if not input_root:
                        raise ValidationError("Input Folder is required.")
                    manifest = build_folder_batch_manifest(
                        Path(str(input_root)),
                        Path(str(output_root)),
                        steps=list(builder_values.get("steps", [])),
                        recursive_inputs=bool(builder_values.get("recursive_inputs", True)),
                        file_patterns=list(builder_values.get("file_patterns", ["*.pdf"])),
                        report_path=report_path_value,
                        fail_fast=bool(builder_values.get("fail_fast", False)),
                        job_name=str(builder_values.get("job_name", "folder-batch")),
                    )
                manifest_dir = Path(str(output_root)) / "_pdf_toolkit"
                manifest_path = write_manifest(manifest_dir / "folder-batch.yaml", manifest)
                values = {"manifest_path": manifest_path}
            except Exception as exc:
                self._show_inline_failure("Batch setup needs attention", str(exc), operation_id="batch-run")
                return
        if self._current_definition.id == "watch-folder" and not values.get("once"):
            try:
                self._last_request_context = {
                    "operation_id": "watch-folder",
                    "label": self._current_definition.label,
                    "values": _json_safe(values),
                    "report_path": report_path,
                }
                request = prepare_request("watch-folder", values, overwrite=self._overwrite.isChecked())
                self._watch_controller.start(request.values["input_dir"], request.values["manifest_path"], bool(request.values["recursive"]), request.overwrite)
            except Exception as exc:
                self._show_inline_failure("Watch setup needs attention", str(exc), operation_id="watch-folder")
            return
        self._last_request_context = {
            "operation_id": self._current_definition.id,
            "label": self._current_definition.label,
            "values": context_values,
            "report_path": report_path,
        }
        self._run_button.setEnabled(False)
        self.statusBar().showMessage(f"Running {self._current_definition.label}...")
        worker = JobWorker(
            self._current_definition.id,
            values,
            report_path=report_path,
            overwrite=self._overwrite.isChecked(),
        )
        worker.signals.finished.connect(self._handle_result)
        self._thread_pool.start(worker)
        self._results.append_log(f"Running {self._current_definition.id}")

    def _handle_result(self, result: JobResult, report_path: Path | None) -> None:
        self._run_button.setEnabled(True)
        display_label = self._current_definition.label if self._current_definition is not None else None
        self._results.set_result(result, report_path, display_label=display_label)
        self._latest_output_folder = next((path.parent for path in result.outputs), report_path.parent if report_path is not None else None)
        self._footer_open_output_button.setEnabled(self._latest_output_folder is not None)
        self._remember_run(result, report_path)
        if result.status == "error":
            self.statusBar().showMessage(f"{result.operation_id} failed")
            self._update_readiness(last_run_failed=True)
            self._update_ready_summary()
            return
        self._update_readiness(last_run_failed=False)
        preview_path = next((path for path in result.outputs if path.suffix.lower() == ".pdf"), None)
        if preview_path is None:
            self._refresh_preview()
        self.statusBar().showMessage(f"{result.operation_id} completed in {result.duration_ms} ms")

    def _handle_log(self, message: str) -> None:
        self._results.append_log(message)
        self.statusBar().showMessage(message)

    def _sync_watch_state(self, active: bool) -> None:
        self._watch_active = active
        self._stop_watch_button.setEnabled(active)
        self._run_button.setEnabled(not active)

    def _load_recent_runs(self) -> list[dict[str, Any]]:
        raw = self._settings.value("history/recent_runs", "[]")
        try:
            data = json.loads(str(raw))
        except Exception:
            return []
        return [item for item in data if isinstance(item, dict)]

    def _load_pinned_workflows(self) -> list[dict[str, str]]:
        raw = self._settings.value("history/pinned_workflows", "[]")
        try:
            data = json.loads(str(raw))
        except Exception:
            return []
        return [item for item in data if isinstance(item, dict)]

    def _persist_pinned_workflows(self) -> None:
        self._settings.setValue("history/pinned_workflows", json.dumps(self._pinned_workflows))

    def _load_last_used_values(self) -> dict[str, dict[str, Any]]:
        raw = self._settings.value("history/last_values", "{}")
        try:
            data = json.loads(str(raw))
        except Exception:
            return {}
        return data if isinstance(data, dict) else {}

    def _persist_last_used_values(self) -> None:
        self._settings.setValue("history/last_values", json.dumps(self._last_used_values))

    def _persist_recent_runs(self) -> None:
        self._settings.setValue("history/recent_runs", json.dumps(self._recent_runs))

    def _refresh_start_here(self) -> None:
        latest = self._recent_runs[0] if self._recent_runs else None
        latest_success = next((item for item in self._recent_runs if item.get("status") == "success"), None)
        self._start_here_panel.set_recent_activity(latest, latest_success)
        self._start_here_panel.set_pinned_items(self._pinned_workflows)

    def _remember_run(self, result: JobResult, report_path: Path | None) -> None:
        if not self._last_request_context:
            return
        context = dict(self._last_request_context)
        values = dict(context.get("values", {}))
        input_paths: list[str] = []
        operation_id = str(context.get("operation_id", "") or "")
        if operation_id == "batch-run":
            source_mode = str(values.get("source_mode", "folder") or "folder")
            if source_mode == "files":
                input_paths.extend(str(path) for path in values.get("input_files", []) if str(path).strip())
            else:
                input_root = values.get("input_root")
                if input_root and str(input_root).strip():
                    input_paths.append(str(input_root))
        elif self._current_definition is not None:
            for field_def in self._current_definition.fields:
                if field_def.kind not in {"file", "directory"}:
                    continue
                raw = values.get(field_def.name)
                if raw in (None, "", [], {}):
                    continue
                items = raw if isinstance(raw, list) else [raw]
                if field_def.path_role == "input":
                    input_paths.extend(str(item) for item in items if str(item).strip())
        output_dirs = sorted({str(path.parent) for path in [*result.outputs, *([report_path] if report_path is not None else [])]})
        entry = {
            "operation_id": context.get("operation_id"),
            "label": context.get("label"),
            "status": result.status,
            "values": values,
            "report_path": context.get("report_path"),
            "input_paths": input_paths,
            "output_dirs": output_dirs,
            "primary_output": str(result.outputs[0]) if result.outputs else None,
            "error": result.error,
        }
        filtered = [item for item in self._recent_runs if item.get("operation_id") != entry["operation_id"] or item.get("values") != entry["values"]]
        self._recent_runs = [entry, *filtered][:10]
        self._persist_recent_runs()
        self._persist_last_used_snapshot(str(entry["operation_id"] or ""), values, context.get("report_path"))
        self._refresh_start_here()

    def _repeat_last_task(self, use_same_inputs: bool = True) -> None:
        latest = next((item for item in self._recent_runs if item.get("status") == "success"), None)
        if not latest:
            return
        operation_id = str(latest.get("operation_id", "") or "")
        if not operation_id:
            return
        self._select_operation(operation_id, restore_saved_values=False)
        values = latest.get("values", {})
        if isinstance(values, dict):
            values_to_apply = dict(values)
            if not use_same_inputs:
                values_to_apply = self._clear_input_values(operation_id, values_to_apply)
            if operation_id == "batch-run":
                self._apply_values_to_current_form({"manifest_path": values_to_apply}, report_path=latest.get("report_path"))
            else:
                self._apply_values_to_current_form(values_to_apply, report_path=latest.get("report_path"))
        message = "Repeated last task setup" if use_same_inputs else "Repeated last task with fresh inputs"
        self.statusBar().showMessage(f"{message}: {latest.get('label', operation_id)}")

    def _update_readiness(self, *, last_run_failed: bool | None = None) -> None:
        statuses = collect_doctor_status("all")
        required_missing = [status for status in statuses if status.required and not status.available]
        optional_missing = [status for status in statuses if not status.required and not status.available]
        if last_run_failed:
            title = "Last run needs attention"
            detail = "The previous task failed. Review the results panel, then rerun when ready."
        elif required_missing:
            title = "Setup required"
            detail = "Some core dependencies are missing. Review diagnostics before running core workflows."
        elif optional_missing:
            title = "Ready now"
            detail = "Core PDF workflows are ready. Optional add-ons like OCR and LLM analysis can be installed later."
        else:
            title = "Ready now"
            detail = "Everything in the current build is available, including optional add-ons."
        self._header_status_metric.setText(title)
        self._start_here_panel.set_readiness(title, detail)

    def _update_parameter_copy(self, definition: OperationDefinition) -> None:
        if definition.id == "merge":
            self._parameter_header.setText("Files To Combine")
            self._parameter_caption.setText("Confirm the source PDFs and the export file, then run the merge.")
        elif definition.id == "tables-extract":
            self._parameter_header.setText("Export Folder And Table Options")
            self._parameter_caption.setText("Choose the source PDF, export folder, and table format before you run the export.")
        elif definition.id == "batch-run":
            self._parameter_header.setText("Folder Workflow Setup")
            self._parameter_caption.setText("Point the workflow at files or a folder, then confirm the output root and processing steps.")
        else:
            self._parameter_header.setText("Operation Parameters")
            self._parameter_caption.setText("Confirm the source files, output location, and practical defaults before you run the task.")

    def _update_pin_buttons(self) -> None:
        if self._current_definition is None:
            self._pin_operation_button.setText("Pin Task")
            return
        is_pinned = any(
            item.get("kind") == "operation" and item.get("id") == self._current_definition.id for item in self._pinned_workflows
        )
        self._pin_operation_button.setText("Unpin Task" if is_pinned else "Pin Task")

    def _toggle_current_operation_pin(self) -> None:
        if self._current_definition is None:
            return
        self._toggle_pinned_item("operation", self._current_definition.id, self._current_definition.label)

    def _toggle_template_pin(self, template_id: str) -> None:
        template = get_workflow_template(template_id)
        self._toggle_pinned_item("template", template.id, template.label)

    def _toggle_pinned_item(self, kind: str, item_id: str, label: str) -> None:
        existing = next((item for item in self._pinned_workflows if item.get("kind") == kind and item.get("id") == item_id), None)
        if existing is not None:
            self._pinned_workflows = [item for item in self._pinned_workflows if item is not existing]
        else:
            self._pinned_workflows = [{"kind": kind, "id": item_id, "label": label}, *self._pinned_workflows][:8]
        self._persist_pinned_workflows()
        self._refresh_start_here()
        self._update_pin_buttons()

    def _open_pinned_workflow(self, kind: str, item_id: str) -> None:
        if kind == "template":
            self._apply_template(item_id)
        else:
            self._select_operation(item_id)

    def _persist_last_used_snapshot(self, operation_id: str, values: dict[str, Any], report_path: str | None) -> None:
        if not operation_id:
            return
        self._last_used_values[operation_id] = {
            "values": _json_safe(values),
            "report_path": report_path,
        }
        self._persist_last_used_values()

    def _restore_last_used_values(self, operation_id: str) -> None:
        snapshot = self._last_used_values.get(operation_id)
        if not isinstance(snapshot, dict):
            return
        values = snapshot.get("values")
        if isinstance(values, dict):
            if operation_id == "batch-run":
                self._apply_values_to_current_form({"manifest_path": values}, report_path=snapshot.get("report_path"))
            else:
                self._apply_values_to_current_form(values, report_path=snapshot.get("report_path"))

    def _clear_input_values(self, operation_id: str, values: dict[str, Any]) -> dict[str, Any]:
        cleared = dict(values)
        if operation_id == "batch-run":
            cleared["input_root"] = ""
            cleared["input_files"] = []
            return cleared
        definition = self._definition_map.get(operation_id)
        if definition is None:
            return cleared
        for field_def in definition.fields:
            if field_def.path_role != "input":
                continue
            if field_def.multiple:
                cleared[field_def.name] = []
            else:
                cleared[field_def.name] = ""
        return cleared

    def _open_latest_output_folder(self) -> None:
        if self._latest_output_folder is not None:
            _open_path(self._latest_output_folder)

    def _ingest_paths(self, paths: list[str] | list[Path]) -> None:
        normalized = [str(path) for path in _normalize_intake_paths([Path(str(item)) for item in paths]) if Path(str(path)).suffix.lower() == ".pdf"]
        if not normalized:
            self._start_here_panel.set_preflight(_suggest_workflow_for_intake([]))
            return
        self._current_intake_paths = normalized
        preflight = _suggest_workflow_for_intake(normalized)
        self._start_here_panel.set_preflight(preflight)
        self._preflight_metric.setText(
            f"{preflight['file_count']} PDF(s) · {preflight['page_count']} page(s)"
            + (f" · {preflight['scan_summary']}" if preflight.get("scan_summary") else "")
        )
        detail_parts = [f"Suggested task: {preflight['suggested_label']}"]
        if preflight.get("alternatives"):
            detail_parts.append("Also useful: " + ", ".join(preflight["alternatives"]))
        if preflight.get("warnings"):
            detail_parts.append(" ".join(preflight["warnings"]))
        self._preflight_detail.setText(" ".join(detail_parts))
        self._select_operation(str(preflight["suggested_operation"]), restore_saved_values=False)
        self.statusBar().showMessage(f"Intake analyzed: {preflight['suggested_label']}")

    def _apply_intake_to_current_operation(self, paths: list[str]) -> None:
        if self._current_definition is None:
            return
        definition = self._current_definition
        values: dict[str, Any] = {}
        if definition.id == "batch-run":
            current = self._collect_values().get("manifest_path", {})
            payload = dict(current) if isinstance(current, dict) else {}
            payload.setdefault("source_mode", "files")
            payload["source_mode"] = "files"
            payload["input_files"] = list(paths)
            payload.setdefault("output_root", _suggest_output_values("batch-run", paths)["output_root"])
            payload.setdefault("report_path", _suggest_output_values("batch-run", paths)["report_path"])
            values["manifest_path"] = payload
        else:
            for field_def in definition.fields:
                if field_def.path_role != "input":
                    continue
                if field_def.multiple:
                    values[field_def.name] = list(paths)
                else:
                    values[field_def.name] = paths[0]
                break
            suggested_outputs = _suggest_output_values(definition.id, paths)
            current_values = self._collect_values()
            for key, value in suggested_outputs.items():
                if current_values.get(key) in (None, "", [], {}):
                    values[key] = value
        if values:
            self._apply_values_to_current_form(values)

    def _apply_smart_defaults_from_current_inputs(self) -> bool:
        if self._current_definition is None:
            return False
        definition = self._current_definition
        values = self._collect_values()
        updates: dict[str, Any] = {}
        input_paths: list[str] = []

        if definition.id == "batch-run":
            payload = values.get("manifest_path", {})
            if not isinstance(payload, dict):
                return False
            source_mode = str(payload.get("source_mode", "folder") or "folder")
            if source_mode == "files":
                input_paths = [str(path) for path in payload.get("input_files", []) if str(path).strip()]
            elif payload.get("input_root"):
                input_paths = [str(payload.get("input_root"))]
            if not input_paths:
                return False
            suggestions = _suggest_output_values("batch-run", input_paths)
            next_payload = dict(payload)
            changed = False
            for key, value in suggestions.items():
                if next_payload.get(key) in (None, "", [], {}):
                    next_payload[key] = value
                    changed = True
            if changed:
                updates["manifest_path"] = next_payload
        else:
            for field_def in definition.fields:
                if field_def.path_role != "input":
                    continue
                raw = values.get(field_def.name)
                if raw in (None, "", [], {}):
                    continue
                if isinstance(raw, list):
                    input_paths = [str(item) for item in raw if str(item).strip()]
                else:
                    input_paths = [str(raw)]
                if input_paths:
                    break
            if not input_paths:
                return False
            for key, value in _suggest_output_values(definition.id, input_paths).items():
                if values.get(key) in (None, "", [], {}):
                    updates[key] = value

        if not updates:
            return False
        self._auto_filling = True
        try:
            self._apply_values_to_current_form(updates, refresh=False)
        finally:
            self._auto_filling = False
        self._refresh_preview()
        return True

    def _set_validation_state(self, widget: object, state: str | None, message: str | None = None) -> None:
        if not isinstance(widget, QWidget):
            return
        widget.setProperty("validationState", state or "")
        widget.style().unpolish(widget)
        widget.style().polish(widget)
        if message is not None:
            widget.setToolTip(message)

    def _apply_widget_validation(self, field_errors: dict[str, str]) -> None:
        for field_name, widget in self._field_widgets.items():
            message = field_errors.get(field_name)
            self._set_validation_state(widget, "error" if message else None, message)
            label = self._field_labels.get(field_name)
            if label is not None:
                label.setProperty("validationState", "error" if message else "")
                label.style().unpolish(label)
                label.style().polish(label)
                if message is not None:
                    label.setToolTip(message)

    def _update_ready_summary(self, *, notify: bool = False) -> bool:
        if self._current_definition is None:
            self._ready_summary.setText("Choose a task to see setup guidance.")
            self._ready_warning.setText("Validation and optional add-on guidance will appear here.")
            return False
        definition = self._current_definition
        values = self._collect_values()
        errors: list[str] = []
        warnings: list[str] = []
        field_errors: dict[str, str] = {}

        for field_def in definition.fields:
            value = values.get(field_def.name)
            if field_def.required and value in (None, "", [], {}):
                message = f"{field_def.label} is required."
                errors.append(message)
                field_errors[field_def.name] = message

        if definition.id == "split" and not values.get("every_page") and not values.get("ranges"):
            message = "Add page ranges or enable Every Page before running Split Pages."
            errors.append(message)
            field_errors["ranges"] = message
        if definition.id == "analyze-llm" and values.get("preset") == "qa" and not str(values.get("question", "")).strip():
            message = "Question is required when the LLM preset is Q&A."
            errors.append(message)
            field_errors["question"] = message

        statuses = collect_doctor_status("all")
        status_map = {status.name: status for status in statuses}
        if definition.id == "ocr":
            missing = [name for name in ("ocrmypdf", "tesseract", "gswin64c") if name in status_map and not status_map[name].available]
            if missing:
                errors.append("OCR tools are not ready yet. Install OCRmyPDF, Tesseract, and Ghostscript to run OCR Scans.")
        if definition.id == "analyze-llm":
            missing = [name for name in ("openai", "pydantic", "OPENAI_API_KEY") if name in status_map and not status_map[name].available]
            if missing:
                errors.append("OpenAI analysis is optional, but this task needs the llm extras and OPENAI_API_KEY.")
        if definition.id == "tables-extract" and values.get("ocr_first"):
            missing = [name for name in ("ocrmypdf", "tesseract", "gswin64c") if name in status_map and not status_map[name].available]
            if missing:
                warnings.append("Table extraction can still run, but OCR First needs the optional OCR tools.")
        if definition.id == "batch-run":
            payload = values.get("manifest_path", {})
            if isinstance(payload, dict):
                source_mode = str(payload.get("source_mode", "folder") or "folder")
                if source_mode == "files" and not payload.get("input_files"):
                    message = "Choose one or more input files for the folder workflow."
                    errors.append(message)
                    field_errors["manifest_path"] = message
                if source_mode != "files" and not str(payload.get("input_root", "")).strip():
                    message = "Choose an input folder for the folder workflow."
                    errors.append(message)
                    field_errors["manifest_path"] = message
                if not str(payload.get("output_root", "")).strip():
                    message = "Choose an output folder for the folder workflow."
                    errors.append(message)
                    field_errors["manifest_path"] = message

        inputs = [field_def.label for field_def in definition.fields if field_def.path_role == "input" and values.get(field_def.name) not in (None, "", [], {})]
        outputs = [field_def.label for field_def in definition.fields if field_def.path_role == "output" and values.get(field_def.name) not in (None, "", [], {})]
        if self._report_input.isVisible() and self._report_input.value():
            outputs.append("Report Path")
        input_text = ", ".join(inputs) if inputs else "No input selected yet"
        output_text = ", ".join(outputs) if outputs else "No destination chosen yet"
        self._ready_summary.setText(f"Task: {definition.label} · Inputs: {input_text} · Destination: {output_text}")
        self._apply_widget_validation(field_errors)

        detail_parts: list[str] = []
        if warnings:
            detail_parts.append("Warnings: " + " ".join(warnings))
        if errors:
            detail_parts.append("Fix before running: " + " ".join(errors))
        elif definition.id == "ocr":
            detail_parts.append("OCR is optional. This task only works after the OCR add-ons are installed.")
        elif definition.id == "analyze-llm":
            detail_parts.append("OpenAI analysis is optional. Local extraction remains available even without it.")
        else:
            detail_parts.append("Setup looks ready. Validate or run when you are comfortable with the output path.")
        self._ready_warning.setText(" ".join(detail_parts))
        if notify:
            self.statusBar().showMessage("Setup ready" if not errors else "Setup needs attention")
        return not errors

    def _show_inline_failure(self, title: str, message: str, *, operation_id: str) -> None:
        self._results.set_result(
            JobResult(
                operation_id=operation_id,
                status="error",
                outputs=[],
                warnings=[],
                details={"title": title},
                error=message,
                duration_ms=0,
            ),
            None,
            display_label=title,
        )
        self._update_readiness(last_run_failed=True)
        self._ready_warning.setText(message)
        self.statusBar().showMessage(title)

    def closeEvent(self, event) -> None:  # pragma: no cover - Qt close handling
        self._watch_controller.stop()
        super().closeEvent(event)

    def _show_about_dialog(self) -> None:
        QMessageBox.about(
            self,
            f"About {APP_NAME}",
            "\n".join(
                [
                    f"{APP_NAME} {APP_VERSION}",
                    APP_TAGLINE,
                    "",
                    "Built for fast office document cleanup, extraction, redaction, and repeatable desktop workflows.",
                    "OCR is supported as an optional add-on for this first public release.",
                    "",
                    PROJECT_URL,
                ]
            ),
        )


def create_app() -> QApplication:
    if QApplication is None:
        raise RuntimeError(f"PySide6 is required for the GUI: {_PYSIDE_IMPORT_ERROR}")
    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyle(QStyleFactory.create("Fusion"))
    app.setOrganizationName(ORGANIZATION_NAME)
    app.setOrganizationDomain(ORGANIZATION_DOMAIN)
    app.setApplicationName("PDFToolkit")
    app.setWindowIcon(_app_icon())
    app.setStyleSheet(
        """
        QMainWindow {
            background: #071017;
        }
        QMenuBar, QStatusBar {
            background: #0a161e;
            color: #f4efe4;
            border: none;
        }
        QLabel {
            color: #f4efe4;
        }
        QLabel#ShellEyebrow, QLabel#PanelEyebrow, QLabel#CardEyebrow, QLabel#WelcomeEyebrow {
            color: #f0b96a;
            font-size: 11px;
            font-weight: 700;
            letter-spacing: 0.18em;
            text-transform: uppercase;
        }
        QLabel#ShellTitle {
            font-family: "Bahnschrift";
            font-size: 34px;
            font-weight: 700;
            color: #fff7eb;
        }
        QLabel#HeroTitle {
            font-family: "Bahnschrift";
            font-size: 30px;
            font-weight: 700;
            color: #fff7eb;
        }
        QLabel#MutedLabel {
            color: #97a8ab;
        }
        QLabel#SummaryPill {
            background: rgba(240, 185, 106, 0.12);
            color: #ffdca8;
            border: 1px solid rgba(240, 185, 106, 0.35);
            border-radius: 16px;
            padding: 10px 14px;
            font-weight: 600;
        }
        QLabel#FieldLabel {
            color: #f4efe4;
            font-weight: 600;
        }
        QLabel#FieldLabel[validationState="error"] {
            color: #ff9b71;
        }
        QLabel#WelcomeTitle {
            font-family: "Bahnschrift";
            color: #fff7eb;
            font-size: 26px;
            font-weight: 700;
        }
        QLabel#WelcomeNote {
            color: #ffcf8d;
            font-weight: 600;
        }
        QLabel#PanelTitle {
            font-family: "Bahnschrift";
            font-size: 18px;
            font-weight: 700;
            color: #fff7eb;
        }
        QLabel#SectionTitle, QLabel#ResultTitle {
            font-family: "Bahnschrift";
            font-size: 18px;
            font-weight: 700;
            color: #fff7eb;
        }
        QLabel#MetricPill {
            background: rgba(255, 255, 255, 0.05);
            color: #f4efe4;
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 15px;
            padding: 7px 12px;
            font-weight: 600;
        }
        QLabel#StatusBadge, QLabel#ResultPill {
            background: rgba(89, 183, 167, 0.18);
            color: #9ff3e8;
            border: 1px solid rgba(89, 183, 167, 0.42);
            border-radius: 15px;
            padding: 7px 12px;
            font-weight: 700;
        }
        QFrame#PanelCard {
            background: #0d1820;
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 24px;
        }
        QFrame#PanelHeader {
            background: transparent;
            border: none;
        }
        QFrame#PanelAccent {
            background: #f0b96a;
            border-radius: 3px;
        }
        QFrame#ShellHeader {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #122430, stop:0.55 #0f1c25, stop:1 #1f2d22);
            border: 1px solid rgba(240, 185, 106, 0.22);
            border-radius: 24px;
        }
        QFrame#OperationHero, QFrame#ResultsHero, QFrame#ActionBarCard {
            background: #101c24;
            border: 1px solid rgba(255, 255, 255, 0.07);
            border-radius: 20px;
        }
        QFrame#WelcomePanel {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #1f3544, stop:0.45 #16222a, stop:1 #263224);
            border: 1px solid rgba(240, 185, 106, 0.18);
            border-radius: 22px;
        }
        QFrame#InnerSection {
            background: #13202a;
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 12px;
        }
        QWidget#FormHost, QScrollArea#FormScroll, QScrollArea#FormScroll > QWidget > QWidget {
            background: transparent;
        }
        QPushButton {
            background: #f0b96a;
            color: #15110a;
            padding: 10px 16px;
            border: 1px solid rgba(255, 255, 255, 0.06);
            border-radius: 12px;
            font-weight: 700;
        }
        QPushButton:hover {
            background: #ffd095;
        }
        QPushButton[secondary="true"] {
            background: #17242d;
            color: #f4efe4;
            border: 1px solid rgba(255, 255, 255, 0.09);
        }
        QPushButton[secondary="true"]:hover {
            background: #213441;
        }
        QPushButton:disabled {
            background: #2a3640;
            color: #748287;
        }
        QTreeWidget, QListWidget, QTableWidget, QPlainTextEdit, QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
            background: #14212a;
            color: #f4efe4;
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 12px;
            padding: 6px;
            alternate-background-color: #192833;
            gridline-color: rgba(255, 255, 255, 0.08);
            selection-background-color: #de7c4b;
            selection-color: #fff7eb;
        }
        QWidget[validationState="error"], QListWidget[validationState="error"], QPlainTextEdit[validationState="error"], QTableWidget[validationState="error"], QComboBox[validationState="error"], QSpinBox[validationState="error"], QDoubleSpinBox[validationState="error"], QLineEdit[validationState="error"] {
            border: 1px solid rgba(255, 155, 113, 0.9);
            background: #2a1820;
        }
        QTableWidget::item {
            padding: 4px;
        }
        QHeaderView::section {
            background: #1b2b35;
            color: #f0b96a;
            border: none;
            border-bottom: 1px solid rgba(255, 255, 255, 0.08);
            padding: 8px;
            font-weight: 600;
        }
        QTabWidget::pane {
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 16px;
            background: #101c24;
        }
        QTabBar::tab {
            background: #15232c;
            color: #97a8ab;
            padding: 10px 16px;
            margin-right: 6px;
            border-top-left-radius: 10px;
            border-top-right-radius: 10px;
        }
        QTabBar::tab:selected {
            background: #f0b96a;
            color: #15110a;
        }
        QCheckBox {
            color: #f4efe4;
            spacing: 8px;
        }
        QCheckBox::indicator {
            width: 16px;
            height: 16px;
            border-radius: 4px;
            border: 1px solid rgba(255, 255, 255, 0.15);
            background: #14212a;
        }
        QCheckBox::indicator:checked {
            background: #59b7a7;
            border: 1px solid #59b7a7;
        }
        QScrollBar:vertical {
            background: #0c151b;
            width: 12px;
            border-radius: 6px;
        }
        QScrollBar::handle:vertical {
            background: #f0b96a;
            border-radius: 6px;
            min-height: 30px;
        }
        QSplitter::handle {
            background: #132029;
            border-radius: 4px;
            width: 10px;
        }
        """
    )
    return app


def main() -> None:
    app = create_app()
    try:
        window = MainWindow()
        window.show()
        app.exec()
    except Exception as exc:  # pragma: no cover - runtime safety for packaged app
        QMessageBox.critical(
            None,
            f"{APP_NAME} failed to start",
            f"The app could not finish launching.\n\n{exc}\n\nIf the problem persists, check the packaged docs or open an issue.",
        )
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
