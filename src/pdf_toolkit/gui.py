from __future__ import annotations

from pathlib import Path
import json
import sys

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
    layout.setContentsMargins(16, 16, 16, 16)
    layout.setSpacing(12)
    title_label = QLabel(title)
    title_label.setObjectName("PanelTitle")
    title_label.setProperty("accentColor", accent)
    layout.addWidget(title_label)
    layout.addWidget(body, 1)
    return container


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
        self._line_edit.textChanged.connect(self.changed.emit)
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
        self._editor.textChanged.connect(self.changed.emit)
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
                widget.textChanged.connect(self.changed.emit)
            elif hasattr(widget, "stateChanged"):
                widget.stateChanged.connect(self.changed.emit)
            elif hasattr(widget, "valueChanged"):
                widget.valueChanged.connect(self.changed.emit)
            elif hasattr(widget, "currentIndexChanged"):
                widget.currentIndexChanged.connect(self.changed.emit)

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


class WelcomePanel(QFrame):
    workflow_requested = Signal(str)

    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("WelcomePanel")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(10)

        eyebrow = QLabel("Windows-first desktop workflow kit")
        eyebrow.setObjectName("WelcomeEyebrow")
        title = QLabel(APP_NAME)
        title.setObjectName("WelcomeTitle")
        self.summary = QLabel(WELCOME_COPY)
        self.summary.setWordWrap(True)
        self.summary.setObjectName("MutedLabel")
        note = QLabel(OCR_NOTE)
        note.setWordWrap(True)
        note.setObjectName("WelcomeNote")

        layout.addWidget(eyebrow)
        layout.addWidget(title)
        layout.addWidget(self.summary)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        shortcuts = [
            ("Merge files", "merge"),
            ("Split pages", "split"),
            ("Extract text", "extract-text"),
            ("OCR add-on", "ocr"),
        ]
        for label, operation_id in shortcuts:
            button = QPushButton(label)
            button.setProperty("secondary", True)
            button.clicked.connect(lambda _checked=False, op=operation_id: self.workflow_requested.emit(op))
            actions.addWidget(button)
        actions.addStretch(1)

        links = QHBoxLayout()
        links.setSpacing(10)
        docs_button = QPushButton("Open docs")
        docs_button.setProperty("secondary", True)
        docs_button.clicked.connect(lambda: _open_url(DOCS_URL))
        releases_button = QPushButton("Windows download")
        releases_button.clicked.connect(lambda: _open_url(RELEASES_URL))
        links.addWidget(docs_button)
        links.addWidget(releases_button)
        links.addStretch(1)

        layout.addLayout(actions)
        layout.addWidget(note)
        layout.addLayout(links)


class ResultsPanel(QWidget):
    def __init__(self) -> None:
        super().__init__()
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
        layout.addWidget(self.tabs)
        self.diagnostics.refresh()

    def append_log(self, message: str) -> None:
        self.log.appendPlainText(message)

    def set_result(self, result: JobResult, report_path: Path | None = None) -> None:
        self.details.set_payload(result.details)
        self.outputs.clear()
        for output in result.outputs:
            self.outputs.addItem(str(output))
        if report_path is not None:
            self.outputs.addItem(str(report_path))
        pdf_output = next((path for path in result.outputs if path.suffix.lower() == ".pdf"), None)
        if pdf_output is not None:
            self.preview.load_pdf(pdf_output)
            self.tabs.setCurrentWidget(self.preview)
        elif result.operation_id == "doctor":
            self.preview.clear()
            self.tabs.setCurrentWidget(self.details)
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
        self._operation_items: dict[str, QTreeWidgetItem] = {}
        self._category_items: dict[str, QTreeWidgetItem] = {}
        self._current_definition: OperationDefinition | None = None
        self._field_widgets: dict[str, object] = {}
        self._watch_active = False
        self._build_ui()
        self._populate_operation_tree()
        self._select_initial_operation()

    def _build_ui(self) -> None:
        splitter = QSplitter()
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self._build_left_pane())
        splitter.addWidget(self._build_center_pane())
        splitter.addWidget(self._build_right_pane())
        splitter.setSizes([320, 560, 800])
        self.setCentralWidget(splitter)
        self._build_menu()
        refresh_action = QAction("Refresh Diagnostics", self)
        refresh_action.triggered.connect(self._results.diagnostics.refresh)
        self.menuBar().addAction(refresh_action)
        self.statusBar().showMessage(f"{APP_NAME} {APP_VERSION} ready")

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
        self._operation_search.setPlaceholderText("Search operations")
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
        layout.addWidget(self._operation_search)
        layout.addWidget(self._tree)
        self._operation_count = QLabel()
        self._operation_count.setObjectName("MutedLabel")
        layout.addWidget(self._operation_count)
        return _panel_widget("Operations", body, accent="#d46a3a")

    def _build_center_pane(self) -> QWidget:
        self._welcome_panel = WelcomePanel()
        self._welcome_panel.workflow_requested.connect(self._select_operation)
        self._title = QLabel()
        self._title.setObjectName("HeroTitle")
        self._description = QLabel()
        self._description.setWordWrap(True)
        self._description.setObjectName("MutedLabel")
        self._summary = QLabel("Choose an operation, fill the fields, then run it.")
        self._summary.setObjectName("SummaryPill")
        self._form_host = QWidget()
        self._form_host.setObjectName("FormHost")
        self._form_layout = QFormLayout(self._form_host)
        self._form_layout.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
        self._form_layout.setContentsMargins(12, 12, 12, 12)
        self._form_layout.setHorizontalSpacing(18)
        self._form_layout.setVerticalSpacing(14)
        self._report_label = QLabel("Report Path")
        self._report_input = SinglePathInput(self._settings, save_mode=True)
        self._overwrite = QCheckBox("Overwrite existing outputs")
        self._run_button = QPushButton("Run")
        self._run_button.clicked.connect(self._run_current_operation)
        self._stop_watch_button = QPushButton("Stop Watch")
        self._stop_watch_button.clicked.connect(self._watch_controller.stop)
        self._stop_watch_button.setEnabled(False)
        controls = QHBoxLayout()
        controls.addWidget(self._overwrite)
        controls.addStretch(1)
        controls.addWidget(self._run_button)
        controls.addWidget(self._stop_watch_button)
        scroll = QScrollArea()
        scroll.setObjectName("FormScroll")
        scroll.setWidgetResizable(True)
        scroll.setWidget(self._form_host)
        body = QWidget()
        layout = QVBoxLayout(body)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        layout.addWidget(self._welcome_panel)
        layout.addWidget(self._title)
        layout.addWidget(self._description)
        layout.addWidget(self._summary)
        layout.addWidget(QLabel("Operation Parameters"))
        layout.addWidget(scroll, 1)
        layout.addWidget(self._report_label)
        layout.addWidget(self._report_input)
        layout.addLayout(controls)
        return _panel_widget("Workspace", body)

    def _build_right_pane(self) -> QWidget:
        self._results = ResultsPanel()
        return _panel_widget("Results", self._results, accent="#2f6b57")

    def _populate_operation_tree(self) -> None:
        for definition in self._definitions:
            parent = self._category_items.get(definition.category)
            if parent is None:
                parent = QTreeWidgetItem([definition.category])
                parent.setFlags(parent.flags() & ~Qt.ItemFlag.ItemIsSelectable)
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

    def _select_operation(self, operation_id: str) -> None:
        item = self._operation_items.get(operation_id)
        if item is None:
            return
        self._tree.setCurrentItem(item)
        self._set_current_operation(operation_id)

    def _handle_operation_selection(self) -> None:
        item = self._tree.currentItem()
        if item is None:
            return
        operation_id = item.data(0, Qt.ItemDataRole.UserRole)
        if operation_id:
            self._set_current_operation(str(operation_id))

    def _set_current_operation(self, operation_id: str) -> None:
        definition = self._definition_map[operation_id]
        self._current_definition = definition
        self._title.setText(definition.label)
        self._description.setText(definition.description)
        mode_text = "Writes files" if definition.mutating else "Read only"
        optional_text = " - OCR add-on required" if definition.id == "ocr" else ""
        self._summary.setText(f"{definition.category} - {mode_text}{optional_text}")
        self._run_button.setText("Start Watch" if definition.id == "watch-folder" else "Run")
        show_global_report = definition.supports_report and definition.id != "batch-run"
        self._report_label.setVisible(show_global_report)
        self._report_input.setVisible(show_global_report)
        if not show_global_report:
            self._report_input.set_value(None)
        self._rebuild_form(definition)
        self._results.preview.clear()
        self._results.details.set_payload({})
        self.statusBar().showMessage(f"Selected {definition.label}")

    def _rebuild_form(self, definition: OperationDefinition) -> None:
        while self._form_layout.rowCount():
            self._form_layout.removeRow(0)
        self._field_widgets.clear()
        if definition.id == "batch-run":
            builder = FolderBatchBuilder(self._settings)
            self._field_widgets["manifest_path"] = builder
            label = QLabel("Folder Batch Setup")
            label.setObjectName("FieldLabel")
            self._form_layout.addRow(label, builder)
            return
        for field_def in definition.fields:
            widget = self._build_field_widget(field_def)
            self._field_widgets[field_def.name] = widget
            label = QLabel(field_def.label + (" *" if field_def.required else ""))
            label.setObjectName("FieldLabel")
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
        self._operation_count.setText(f"{visible_count} operation{'s' if visible_count != 1 else ''}")

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
        values = self._collect_values()
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
                QMessageBox.critical(self, "Batch Setup Error", str(exc))
                return
        if self._current_definition.id == "watch-folder" and not values.get("once"):
            try:
                request = prepare_request("watch-folder", values, overwrite=self._overwrite.isChecked())
                self._watch_controller.start(request.values["input_dir"], request.values["manifest_path"], bool(request.values["recursive"]), request.overwrite)
            except Exception as exc:
                QMessageBox.critical(self, "Watch Error", str(exc))
            return
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
        self._results.set_result(result, report_path)
        if result.status == "error":
            self.statusBar().showMessage(f"{result.operation_id} failed")
            message = result.error or "Unknown error"
            if result.operation_id == "ocr":
                message = f"{message}\n\nOCR is optional. You can keep using the rest of the toolkit without installing OCR tools."
            QMessageBox.critical(self, "Operation Failed", message)
            return
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
            background: #09131a;
        }
        QMenuBar, QStatusBar {
            background: #0d1a22;
            color: #d7ecff;
        }
        QLabel {
            color: #d7ecff;
        }
        QLabel#HeroTitle {
            font-size: 28px;
            font-weight: 700;
            color: #f4fbff;
        }
        QLabel#MutedLabel {
            color: #8ca8bc;
        }
        QLabel#SummaryPill {
            background: #132a33;
            color: #66f0c9;
            border: 1px solid #1d4d5b;
            border-radius: 14px;
            padding: 8px 12px;
            font-weight: 600;
        }
        QLabel#FieldLabel {
            color: #bfe4ff;
            font-weight: 600;
        }
        QLabel#WelcomeEyebrow {
            color: #72e0a0;
            font-size: 12px;
            font-weight: 700;
            text-transform: uppercase;
        }
        QLabel#WelcomeTitle {
            color: #f4fbff;
            font-size: 24px;
            font-weight: 700;
        }
        QLabel#WelcomeNote {
            color: #ffd166;
            font-weight: 600;
        }
        QFrame#PanelCard {
            background: #0f1d26;
            border: 1px solid #1d3442;
            border-radius: 18px;
        }
        QFrame#WelcomePanel {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #143242, stop:1 #0d1a22);
            border: 1px solid #2d5a6c;
            border-radius: 16px;
        }
        QFrame#InnerSection {
            background: #12222d;
            border: 1px solid #214153;
            border-radius: 12px;
        }
        QLabel#PanelTitle {
            font-size: 16px;
            font-weight: 700;
            color: #ffd166;
        }
        QWidget#FormHost, QScrollArea#FormScroll, QScrollArea#FormScroll > QWidget > QWidget {
            background: #0f1d26;
        }
        QPushButton {
            background: #00c2ff;
            color: #05131a;
            padding: 8px 14px;
            border: none;
            border-radius: 8px;
            font-weight: 600;
        }
        QPushButton:hover {
            background: #5ad7ff;
        }
        QPushButton[secondary="true"] {
            background: #19303d;
            color: #d7ecff;
            border: 1px solid #2d5a6c;
        }
        QPushButton[secondary="true"]:hover {
            background: #234252;
        }
        QPushButton:disabled {
            background: #385160;
            color: #90a8b8;
        }
        QTreeWidget, QListWidget, QTableWidget, QPlainTextEdit, QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
            background: #142632;
            color: #e8f6ff;
            border: 1px solid #26485b;
            border-radius: 10px;
            padding: 4px;
            alternate-background-color: #18303d;
            gridline-color: #26485b;
            selection-background-color: #ff4d8d;
            selection-color: #ffffff;
        }
        QTableWidget::item {
            padding: 4px;
        }
        QHeaderView::section {
            background: #19303d;
            color: #8cf0ff;
            border: none;
            border-bottom: 1px solid #26485b;
            padding: 6px;
            font-weight: 600;
        }
        QTabWidget::pane {
            border: 1px solid #26485b;
            border-radius: 12px;
            background: #12222d;
        }
        QTabBar::tab {
            background: #18313f;
            color: #8ca8bc;
            padding: 8px 14px;
            margin-right: 4px;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
        }
        QTabBar::tab:selected {
            background: #ff4d8d;
            color: #ffffff;
        }
        QCheckBox {
            color: #d7ecff;
            spacing: 8px;
        }
        QCheckBox::indicator {
            width: 16px;
            height: 16px;
            border-radius: 4px;
            border: 1px solid #3a6278;
            background: #142632;
        }
        QCheckBox::indicator:checked {
            background: #7cff6b;
            border: 1px solid #7cff6b;
        }
        QScrollBar:vertical {
            background: #10202a;
            width: 12px;
            border-radius: 6px;
        }
        QScrollBar::handle:vertical {
            background: #ffb703;
            border-radius: 6px;
            min-height: 30px;
        }
        QSplitter::handle {
            background: #1e3949;
            width: 8px;
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
