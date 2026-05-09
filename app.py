"""PySide6 desktop UI for the PWA Data Extractor (v2 wizard layout)."""
from __future__ import annotations

import sys
import traceback
from pathlib import Path
from typing import Optional

import pandas as pd
from PySide6.QtCore import (
    QEvent,
    QObject,
    QPoint,
    QPointF,
    Qt,
    QThread,
    QTimer,
    QUrl,
    Signal,
)
from PySide6.QtGui import (
    QColor,
    QCursor,
    QDesktopServices,
    QDragEnterEvent,
    QDropEvent,
    QIcon,
    QPalette,
)
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWidgets import (
    QAbstractItemView,
    QAbstractSpinBox,
    QApplication,
    QButtonGroup,
    QComboBox,
    QDialog,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QToolButton,
    QToolTip,
    QVBoxLayout,
    QWidget,
)

try:
    from .backend import (
        ANALYSIS_MODE,
        APP_ICON_PATH,
        APP_PUBLISHER,
        APP_TITLE,
        APP_VERSION,
        CONTACT_EMAIL,
        COLUMNS,
        EXTRA_COLUMNS,
        GROUP_MODE_SUBJECT,
        GROUP_MODE_SUBJECT_TIMEPOINT,
        GROUP_MODE_SUBJECT_VISIT,
        GROUP_MODE_SUBJECT_VISIT_TIMEPOINT,
        REPORT_MODE_CLINICAL,
        REPORT_MODE_DETAILED,
        REPOSITORY_URL,
        REVIEW_REASON_BOTH,
        REVIEW_REASON_MULTI_ENTRY,
        REVIEW_REASON_PAIR_ALERT,
        UI_CONTEXT_COLUMNS,
        AnalysisBundle,
        build_analysis,
        calculate_pair_differences,
        filter_columns_for_report_mode,
        default_output_path,
        display_dataframe,
        format_value,
        initial_manual_pairs,
        load_readme_text,
        pair_alert_triggered,
        patient_rows,
        process_pdf,
        record_status,
        save_to_excel,
    )
except ImportError:
    APP_DIR = Path(__file__).resolve().parent
    if str(APP_DIR) not in sys.path:
        sys.path.insert(0, str(APP_DIR))
    from backend import (
        ANALYSIS_MODE,
        APP_ICON_PATH,
        APP_PUBLISHER,
        APP_TITLE,
        APP_VERSION,
        CONTACT_EMAIL,
        COLUMNS,
        EXTRA_COLUMNS,
        GROUP_MODE_SUBJECT,
        GROUP_MODE_SUBJECT_TIMEPOINT,
        GROUP_MODE_SUBJECT_VISIT,
        GROUP_MODE_SUBJECT_VISIT_TIMEPOINT,
        REPORT_MODE_CLINICAL,
        REPORT_MODE_DETAILED,
        REPOSITORY_URL,
        REVIEW_REASON_BOTH,
        REVIEW_REASON_MULTI_ENTRY,
        REVIEW_REASON_PAIR_ALERT,
        UI_CONTEXT_COLUMNS,
        AnalysisBundle,
        build_analysis,
        calculate_pair_differences,
        filter_columns_for_report_mode,
        default_output_path,
        display_dataframe,
        format_value,
        initial_manual_pairs,
        load_readme_text,
        pair_alert_triggered,
        patient_rows,
        process_pdf,
        record_status,
        save_to_excel,
    )


GROUPING_MODE_LABELS = {
    GROUP_MODE_SUBJECT: "Subject only",
    GROUP_MODE_SUBJECT_TIMEPOINT: "Subject + timepoint",
    GROUP_MODE_SUBJECT_VISIT: "Subject + visit",
    GROUP_MODE_SUBJECT_VISIT_TIMEPOINT: "Subject + visit + timepoint",
}

REPORT_MODE_HELP = {
    REPORT_MODE_DETAILED: "Full set of measurements (~40 fields). Use for PWA Detailed Reports.",
    REPORT_MODE_CLINICAL: "Summary report with basic vitals (~16 fields). Use for PWA Clinical Reports.",
}

REVIEW_REASON_LABELS = {
    REVIEW_REASON_MULTI_ENTRY: ("multi-entry", "warn"),
    REVIEW_REASON_PAIR_ALERT: ("pair alert", "danger"),
    REVIEW_REASON_BOTH: ("multi + alert", "danger"),
}

REVIEW_REASON_DESCRIPTIONS = {
    REVIEW_REASON_MULTI_ENTRY: (
        "This patient has more than 2 entries. Pick which two should be averaged "
        "in the export."
    ),
    REVIEW_REASON_PAIR_ALERT: (
        "This patient only has 2 entries, but the difference between them exceeds "
        "your alert threshold. Confirm the pair is acceptable, or adjust the threshold."
    ),
    REVIEW_REASON_BOTH: (
        "This patient has more than 2 entries AND the auto-paired choice exceeds "
        "your alert threshold. Pick a different pair or confirm if the current pair "
        "is acceptable."
    ),
}


# =====================================================================
#   Background worker
# =====================================================================
class ProcessingWorker(QObject):
    progress = Signal(int, int, str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(
        self,
        pdf_paths: list[Path],
        report_mode: str,
        grouping_mode: str,
        filename_pattern: str | None = None,
    ):
        super().__init__()
        self.pdf_paths = pdf_paths
        self.report_mode = report_mode
        self.grouping_mode = grouping_mode
        self.filename_pattern = filename_pattern

    def run(self) -> None:
        records: list[dict[str, object]] = []
        total_files = len(self.pdf_paths)

        try:
            for index, pdf_path in enumerate(self.pdf_paths, start=1):
                self.progress.emit(index - 1, total_files, f"Reading {pdf_path.name}")
                records.append(
                    process_pdf(
                        pdf_path,
                        report_mode=self.report_mode,
                        grouping_mode=self.grouping_mode,
                        filename_pattern=self.filename_pattern,
                    )
                )
                self.progress.emit(index, total_files, f"Processed {pdf_path.name}")
        except Exception:
            self.failed.emit(traceback.format_exc())
            return

        self.finished.emit(records)


# =====================================================================
#   Helper widgets
# =====================================================================
class HelpButton(QToolButton):
    """Circular `?` icon. Tooltip on hover, popup on click for discoverability."""

    def __init__(self, tip: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("helpButton")
        self.setText("?")
        self.setToolTip(tip)
        self.setCursor(Qt.CursorShape.WhatsThisCursor)
        self.setFixedSize(18, 18)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.clicked.connect(self._show_popup)

    def _show_popup(self) -> None:
        tip = self.toolTip()
        if not tip:
            return
        # Anchor the popup just below the icon, but keep it on screen.
        global_pos = self.mapToGlobal(QPoint(self.width() // 2, self.height() + 4))
        QToolTip.showText(global_pos, tip, self, self.rect(), 12000)


class DropZone(QFrame):
    """A drag-and-drop target that accepts PDF files and emits paths."""

    files_dropped = Signal(list)
    clicked = Signal()

    def __init__(self, parent: Optional[QWidget] = None, compact: bool = False):
        super().__init__(parent)
        self.setObjectName("dropZone")
        self.setAcceptDrops(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self._compact = compact
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setFixedHeight(78 if compact else 150)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18 if not compact else 12, 20, 18 if not compact else 12)
        layout.setSpacing(4)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.icon_label = QLabel("↑")  # up arrow
        self.icon_label.setObjectName("dropZoneIcon")
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.title_label = QLabel(
            "Drop more PDFs to add" if compact else "Drop PDFs here"
        )
        self.title_label.setObjectName("dropZoneTitle")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.subtitle_label = QLabel("or click to browse files")
        self.subtitle_label.setObjectName("dropZoneSubtitle")
        self.subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout.addWidget(self.icon_label)
        layout.addWidget(self.title_label)
        if not compact:
            layout.addWidget(self.subtitle_label)

    def set_compact(self, compact: bool) -> None:
        if self._compact == compact:
            return
        self._compact = compact
        self.title_label.setText(
            "Drop more PDFs to add" if compact else "Drop PDFs here"
        )
        self.subtitle_label.setVisible(not compact)
        margin = 12 if compact else 18
        self.setFixedHeight(78 if compact else 150)
        self.layout().setContentsMargins(20, margin, 20, margin)
        self.style().unpolish(self)
        self.style().polish(self)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls() and any(
            url.toLocalFile().lower().endswith(".pdf") for url in event.mimeData().urls()
        ):
            event.acceptProposedAction()
            self.setProperty("dragActive", True)
            self.style().unpolish(self)
            self.style().polish(self)
        else:
            super().dragEnterEvent(event)

    def dragLeaveEvent(self, event: QEvent) -> None:
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)
        urls = event.mimeData().urls()
        paths = [Path(u.toLocalFile()) for u in urls if u.toLocalFile().lower().endswith(".pdf")]
        if paths:
            self.files_dropped.emit(paths)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
            event.accept()
        else:
            super().mousePressEvent(event)


class SegmentedControl(QFrame):
    """Two-option toggle (Detailed / Clinical)."""

    changed = Signal(str)

    def __init__(self, options: list[tuple[str, str]], default: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("segmentedControl")
        self._buttons: dict[str, QPushButton] = {}
        self._group = QButtonGroup(self)
        self._group.setExclusive(True)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)

        for value, label in options:
            btn = QPushButton(label)
            btn.setCheckable(True)
            btn.setObjectName("segmentedButton")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            if value == default:
                btn.setChecked(True)
            btn.clicked.connect(lambda _checked, v=value: self._on_clicked(v))
            self._group.addButton(btn)
            self._buttons[value] = btn
            layout.addWidget(btn, 1)

    def _on_clicked(self, value: str) -> None:
        self.changed.emit(value)

    def value(self) -> str:
        for value, btn in self._buttons.items():
            if btn.isChecked():
                return value
        return ""

    def set_value(self, value: str) -> None:
        if value in self._buttons:
            self._buttons[value].setChecked(True)


class NoFocusItemDelegate(QStyledItemDelegate):
    """Paint table cells without Qt's current-cell focus ring."""

    def paint(self, painter, option, index) -> None:
        option_without_focus = QStyleOptionViewItem(option)
        option_without_focus.state &= ~QStyle.StateFlag.State_HasFocus
        super().paint(painter, option_without_focus, index)


# =====================================================================
#   Dialogs
# =====================================================================
class ReadmeDialog(QDialog):
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle(f"About {APP_TITLE} {APP_VERSION}")
        self.resize(820, 620)
        if APP_ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(APP_ICON_PATH)))

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        title = QLabel(APP_TITLE)
        title.setObjectName("dialogTitle")
        subtitle = QLabel(f"Version {APP_VERSION} · {APP_PUBLISHER}")
        subtitle.setObjectName("dialogSubtitle")

        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setMarkdown(load_readme_text())

        actions = QHBoxLayout()
        repo_button = QPushButton("Open repository")
        email_button = QPushButton("Email support")
        close_button = QPushButton("Close")
        close_button.setObjectName("primaryButton")
        actions.addWidget(repo_button)
        actions.addWidget(email_button)
        actions.addStretch(1)
        actions.addWidget(close_button)

        repo_button.clicked.connect(
            lambda: QDesktopServices.openUrl(QUrl(REPOSITORY_URL))
        )
        email_button.clicked.connect(
            lambda: QDesktopServices.openUrl(QUrl(f"mailto:{CONTACT_EMAIL}"))
        )
        close_button.clicked.connect(self.accept)

        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(browser, 1)
        layout.addLayout(actions)


class PdfViewerDialog(QDialog):
    def __init__(self, pdf_path: Path, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.document = QPdfDocument(self)

        self.setWindowTitle(f"PDF Viewer — {pdf_path.name}")
        self.resize(960, 760)
        if APP_ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(APP_ICON_PATH)))

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        controls = QHBoxLayout()
        self.page_label = QLabel(pdf_path.name)
        self.prev_button = QPushButton("Previous")
        self.next_button = QPushButton("Next")
        self.page_spin = QSpinBox()
        self.page_spin.setMinimum(1)
        self.page_spin.setPrefix("Page ")
        self.page_spin.setMinimumWidth(110)
        self.page_count_label = QLabel("loading...")
        self.zoom_out_button = QPushButton("Zoom −")
        self.fit_width_button = QPushButton("Fit width")
        self.fit_page_button = QPushButton("Fit page")
        self.zoom_in_button = QPushButton("Zoom +")

        controls.addWidget(self.page_label, 1)
        controls.addWidget(self.prev_button)
        controls.addWidget(self.next_button)
        controls.addWidget(self.page_spin)
        controls.addWidget(self.page_count_label)
        controls.addSpacing(16)
        controls.addWidget(self.zoom_out_button)
        controls.addWidget(self.fit_width_button)
        controls.addWidget(self.fit_page_button)
        controls.addWidget(self.zoom_in_button)
        layout.addLayout(controls)

        self.pdf_view = QPdfView()
        self.pdf_view.setPageMode(QPdfView.PageMode.SinglePage)
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        layout.addWidget(self.pdf_view, 1)

        self.prev_button.clicked.connect(self.go_to_previous_page)
        self.next_button.clicked.connect(self.go_to_next_page)
        self.page_spin.valueChanged.connect(self.page_spin_changed)
        self.zoom_in_button.clicked.connect(lambda: self.adjust_zoom(1.2))
        self.zoom_out_button.clicked.connect(lambda: self.adjust_zoom(1 / 1.2))
        self.fit_width_button.clicked.connect(
            lambda: self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        )
        self.fit_page_button.clicked.connect(
            lambda: self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitInView)
        )

        navigator = self.pdf_view.pageNavigator()
        navigator.currentPageChanged.connect(self.sync_page_controls)

        for widget in (
            self.prev_button,
            self.next_button,
            self.page_spin,
            self.zoom_out_button,
            self.fit_width_button,
            self.fit_page_button,
            self.zoom_in_button,
        ):
            widget.setEnabled(False)

        QTimer.singleShot(0, self.load_document)

    def load_document(self) -> None:
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            error = self.document.load(str(self.pdf_path))
        finally:
            QApplication.restoreOverrideCursor()

        if error != QPdfDocument.Error.None_:
            QMessageBox.critical(
                self,
                "Could not open PDF",
                f"Could not load PDF: {self.pdf_path.name}",
            )
            self.close()
            return

        self.pdf_view.setDocument(self.document)
        self.page_spin.setMaximum(max(1, self.document.pageCount()))
        self.page_count_label.setText(f"of {self.document.pageCount()}")
        for widget in (
            self.page_spin,
            self.zoom_out_button,
            self.fit_width_button,
            self.fit_page_button,
            self.zoom_in_button,
        ):
            widget.setEnabled(True)
        self.sync_page_controls(0)

    def sync_page_controls(self, current_page: int) -> None:
        self.page_spin.blockSignals(True)
        self.page_spin.setValue(current_page + 1)
        self.page_spin.blockSignals(False)
        page_count = self.document.pageCount()
        self.prev_button.setEnabled(current_page > 0)
        self.next_button.setEnabled(0 <= current_page < page_count - 1)

    def page_spin_changed(self, page_number: int) -> None:
        self.jump_to_page(page_number - 1)

    def jump_to_page(self, page_index: int) -> None:
        if page_index < 0 or page_index >= self.document.pageCount():
            return
        self.pdf_view.pageNavigator().jump(
            page_index,
            QPointF(0, 0),
            self.pdf_view.zoomFactor(),
        )

    def go_to_previous_page(self) -> None:
        self.jump_to_page(self.pdf_view.pageNavigator().currentPage() - 1)

    def go_to_next_page(self) -> None:
        self.jump_to_page(self.pdf_view.pageNavigator().currentPage() + 1)

    def adjust_zoom(self, multiplier: float) -> None:
        current_zoom = self.pdf_view.zoomFactor()
        if current_zoom <= 0:
            current_zoom = 1.0
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.Custom)
        self.pdf_view.setZoomFactor(max(0.25, min(current_zoom * multiplier, 5.0)))


# =====================================================================
#   Main window
# =====================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # State ----------------------------------------------------------
        self.pdf_paths: list[Path] = []
        self.output_path: Path = default_output_path()
        self.report_mode: str = REPORT_MODE_DETAILED
        self.grouping_mode: str = GROUP_MODE_SUBJECT_TIMEPOINT
        self.filename_pattern: str = ""
        self.diff_green_max: float = 3.0
        self.pair_alert_threshold: float = 6.0

        self.records: list[dict[str, object]] = []
        self.bundle: Optional[AnalysisBundle] = None
        self.auto_pairs: dict[str, tuple[int, int]] = {}
        self.manual_pairs: dict[str, list[int]] = {}
        self.confirmed_patients: set[str] = set()
        self.last_export_path: Optional[Path] = None

        self.thread: Optional[QThread] = None
        self.worker: Optional[ProcessingWorker] = None
        self.pdf_viewers: list[PdfViewerDialog] = []
        self.readme_dialog: Optional[ReadmeDialog] = None
        self._updating_pair_table = False

        # Window ---------------------------------------------------------
        self.setWindowTitle(APP_TITLE)
        self.resize(1320, 880)
        if APP_ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(APP_ICON_PATH)))

        self._build_ui()
        self._apply_styles()
        self._show_import_screen()

    # -- UI construction ----------------------------------------------
    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("root")
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        root_layout.addWidget(self._build_titlebar())

        self.stack = QStackedWidget()
        self.stack.addWidget(self._build_import_screen())  # index 0
        self.stack.addWidget(self._build_review_screen())  # index 1
        root_layout.addWidget(self.stack, 1)

        root_layout.addWidget(self._build_statusbar())

        self.setCentralWidget(root)

    def _build_titlebar(self) -> QWidget:
        bar = QFrame()
        bar.setObjectName("titlebar")
        bar.setFixedHeight(48)
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(16, 0, 16, 0)
        layout.setSpacing(10)

        title = QLabel(APP_TITLE)
        title.setObjectName("brandTitle")

        version = QLabel(f"v{APP_VERSION}")
        version.setObjectName("brandVersion")

        layout.addWidget(title)
        layout.addWidget(version)
        layout.addStretch(1)

        self.help_button = QToolButton()
        self.help_button.setObjectName("titlebarIcon")
        self.help_button.setText("?")
        self.help_button.setToolTip("Help & about")
        self.help_button.setFixedSize(28, 28)
        self.help_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.help_button.clicked.connect(self.show_about_dialog)
        layout.addWidget(self.help_button)

        return bar

    def _build_statusbar(self) -> QWidget:
        bar = QFrame()
        bar.setObjectName("statusBar")
        bar.setFixedHeight(28)
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(16, 0, 16, 0)
        layout.setSpacing(8)

        dot = QLabel()
        dot.setObjectName("statusDot")
        dot.setFixedSize(7, 7)

        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("statusText")

        version_label = QLabel(f"v{APP_VERSION} · local processing")
        version_label.setObjectName("statusVersion")

        layout.addWidget(dot)
        layout.addWidget(self.status_label)
        layout.addStretch(1)
        layout.addWidget(version_label)

        return bar

    # ---- Import screen ----
    def _build_import_screen(self) -> QWidget:
        screen = QFrame()
        screen.setObjectName("importScreen")
        layout = QVBoxLayout(screen)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Wizard header
        header = QFrame()
        header.setObjectName("wizardHead")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(28, 16, 28, 14)
        header_layout.setSpacing(20)

        titles = QVBoxLayout()
        titles.setSpacing(2)
        h1 = QLabel("Import & process")
        h1.setObjectName("wizardTitle")
        sub = QLabel(
            "Add your PWA reports, choose how they should be parsed, then process locally."
        )
        sub.setObjectName("wizardSub")
        sub.setWordWrap(True)
        titles.addWidget(h1)
        titles.addWidget(sub)
        header_layout.addLayout(titles, 1)
        header_layout.addWidget(self._build_stepper(active="import"))

        layout.addWidget(header)

        # Body
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        body = QWidget()
        body.setObjectName("importBody")
        scroll.setWidget(body)

        body_layout = QHBoxLayout(body)
        body_layout.setContentsMargins(28, 18, 28, 18)
        body_layout.setSpacing(0)

        wrap = QWidget()
        wrap.setObjectName("importWrap")
        wrap.setMaximumWidth(1240)
        wrap_layout = QHBoxLayout(wrap)
        wrap_layout.setContentsMargins(0, 0, 0, 0)
        wrap_layout.setSpacing(18)

        source_card = self._build_source_card()
        settings_card = self._build_settings_card()
        source_card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        settings_card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        wrap_layout.addWidget(source_card, 5)
        wrap_layout.addWidget(settings_card, 7)

        # Center the wrap horizontally
        body_layout.addStretch(1)
        body_layout.addWidget(wrap, 10)
        body_layout.addStretch(1)

        layout.addWidget(scroll, 1)

        # Sticky footer with primary action
        footer = QFrame()
        footer.setObjectName("importActions")
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(28, 12, 28, 12)
        footer_layout.setSpacing(12)

        self.import_meta_label = QLabel("Add at least one PDF to continue")
        self.import_meta_label.setObjectName("importMeta")
        footer_layout.addWidget(self.import_meta_label)
        footer_layout.addStretch(1)

        self.process_button = QPushButton("Process PDFs  →")
        self.process_button.setObjectName("primaryButton")
        self.process_button.setMinimumHeight(36)
        self.process_button.clicked.connect(self.process_files)
        self.process_button.setEnabled(False)
        footer_layout.addWidget(self.process_button)

        self.progress = QProgressBar()
        self.progress.setObjectName("progressBar")
        self.progress.setMaximumWidth(180)
        self.progress.setVisible(False)
        footer_layout.addWidget(self.progress)

        layout.addWidget(footer)

        return screen

    def _build_stepper(self, active: str) -> QWidget:
        container = QFrame()
        container.setObjectName("stepperFrame")
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        def make_step(num: str, label: str, state: str) -> QWidget:
            step = QFrame()
            step.setObjectName("stepPill")
            step.setProperty("state", state)
            step_layout = QHBoxLayout(step)
            step_layout.setContentsMargins(11, 5, 11, 5)
            step_layout.setSpacing(8)

            num_lbl = QLabel(num)
            num_lbl.setObjectName("stepNum")
            num_lbl.setProperty("state", state)
            num_lbl.setFixedSize(18, 18)
            num_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)

            text_lbl = QLabel(label)
            text_lbl.setObjectName("stepText")
            text_lbl.setProperty("state", state)

            step_layout.addWidget(num_lbl)
            step_layout.addWidget(text_lbl)
            return step

        if active == "import":
            layout.addWidget(make_step("1", "Import & process", "active"))
        else:
            layout.addWidget(make_step("1", "Import", "done"))

        arrow = QLabel(">")
        arrow.setObjectName("stepArrow")
        layout.addWidget(arrow)

        if active == "review":
            layout.addWidget(make_step("2", "Review & export", "active"))
        else:
            layout.addWidget(make_step("2", "Review & export", "pending"))

        return container

    def _build_source_card(self) -> QWidget:
        card = QFrame()
        card.setObjectName("importCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(0)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        title = QLabel("SOURCE PDFs")
        title.setObjectName("cardEyebrow")
        sub = QLabel(
            "PWA Detailed or Clinical reports in PDF format. Drag and drop or click to browse."
        )
        sub.setObjectName("cardSub")
        sub.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(sub)
        layout.addSpacing(14)

        # Drop zone (with two visual modes)
        self.drop_zone = DropZone(compact=False)
        self.drop_zone.files_dropped.connect(self._on_files_dropped)
        self.drop_zone.clicked.connect(self.add_pdf_files)
        layout.addWidget(self.drop_zone)

        # File list
        self.file_list = QListWidget()
        self.file_list.setObjectName("fileList")
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.file_list.setFrameShape(QFrame.Shape.NoFrame)
        self.file_list.setMinimumHeight(0)
        self.file_list.setMaximumHeight(220)
        self.file_list.setVisible(False)
        self.file_list.itemSelectionChanged.connect(self._sync_remove_file_button)
        layout.addSpacing(8)
        layout.addWidget(self.file_list)

        # Meta line + actions
        self.file_meta = QFrame()
        self.file_meta.setObjectName("fileMeta")
        meta_layout = QHBoxLayout(self.file_meta)
        meta_layout.setContentsMargins(0, 8, 0, 0)
        meta_layout.setSpacing(0)
        self.file_count_label = QLabel("")
        self.file_count_label.setObjectName("fileMetaText")
        meta_layout.addWidget(self.file_count_label, 1)

        self.remove_selected_files_link = QPushButton("Remove selected")
        self.remove_selected_files_link.setObjectName("linkButton")
        self.remove_selected_files_link.setCursor(Qt.CursorShape.PointingHandCursor)
        self.remove_selected_files_link.setFlat(True)
        self.remove_selected_files_link.setEnabled(False)
        self.remove_selected_files_link.clicked.connect(self.remove_selected_pdf_files)
        meta_layout.addWidget(self.remove_selected_files_link)
        meta_layout.addSpacing(14)

        self.clear_files_link = QPushButton("Clear all")
        self.clear_files_link.setObjectName("linkButton")
        self.clear_files_link.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_files_link.setFlat(True)
        self.clear_files_link.clicked.connect(self.clear_pdf_files)
        meta_layout.addWidget(self.clear_files_link)
        self.file_meta.setVisible(False)
        layout.addWidget(self.file_meta)

        return card

    def _build_settings_card(self) -> QWidget:
        card = QFrame()
        card.setObjectName("importCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(0)

        title = QLabel("HOW TO PARSE THEM")
        title.setObjectName("cardEyebrow")
        sub = QLabel(
            "These options decide how filenames map to patients and which fields are extracted."
        )
        sub.setObjectName("cardSub")
        sub.setWordWrap(True)
        layout.addWidget(title)
        layout.addWidget(sub)
        layout.addSpacing(14)

        # Two-column field grid
        grid = QGridLayout()
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(18)
        grid.setVerticalSpacing(10)

        # Report type
        rt_label_row = self._field_label_row(
            "Report type",
            "PWA reports come in two formats. Detailed reports include the full set of "
            "pulse-wave measurements; Clinical reports are a shorter summary. Pick whichever "
            "matches the PDFs you're loading.",
        )
        self.report_mode_seg = SegmentedControl(
            options=[
                (REPORT_MODE_DETAILED, "Detailed"),
                (REPORT_MODE_CLINICAL, "Clinical"),
            ],
            default=self.report_mode,
        )
        self.report_mode_seg.changed.connect(self._on_report_mode_changed)
        self.report_mode_help = QLabel(REPORT_MODE_HELP[self.report_mode])
        self.report_mode_help.setObjectName("fieldHelp")
        self.report_mode_help.setWordWrap(True)

        rt_col = QVBoxLayout()
        rt_col.setContentsMargins(0, 0, 0, 0)
        rt_col.setSpacing(6)
        rt_col.addLayout(rt_label_row)
        rt_col.addWidget(self.report_mode_seg)
        rt_col.addWidget(self.report_mode_help)

        # Grouping
        gp_label_row = self._field_label_row(
            "Group files by",
            "Decides what counts as 'the same patient' when grouping rows. "
            "The grouping key is built from segments parsed out of each filename. "
            "For example: S01_000a.pdf with Subject + timepoint becomes patient '01 T000'.",
        )
        self.grouping_combo = QComboBox()
        self.grouping_combo.setObjectName("groupingCombo")
        for value, label in GROUPING_MODE_LABELS.items():
            self.grouping_combo.addItem(label, value)
        index = self.grouping_combo.findData(self.grouping_mode)
        if index >= 0:
            self.grouping_combo.setCurrentIndex(index)
        self.grouping_combo.currentIndexChanged.connect(self._on_grouping_changed)
        gp_help = QLabel(
            "Files with the same key are treated as repeated measurements of one patient."
        )
        gp_help.setObjectName("fieldHelp")
        gp_help.setWordWrap(True)

        gp_col = QVBoxLayout()
        gp_col.setContentsMargins(0, 0, 0, 0)
        gp_col.setSpacing(6)
        gp_col.addLayout(gp_label_row)
        gp_col.addWidget(self.grouping_combo)
        gp_col.addWidget(gp_help)

        grid.addLayout(rt_col, 0, 0)
        grid.addLayout(gp_col, 0, 1)

        regex_label_row = self._field_label_row(
            "Custom filename regex (optional)",
            "Use the copied AI prompt to generate a regex for your filename format.",
        )
        self.filename_pattern_line = QLineEdit(self.filename_pattern)
        self.filename_pattern_line.setObjectName("pathInput")
        self.filename_pattern_line.setPlaceholderText(
            r"(?P<subject>IAS\d+)(?:[_\s-]T(?P<timepoint>\d+))?"
        )
        self.filename_pattern_line.textChanged.connect(
            self._filename_pattern_changed
        )
        regex_help = QLabel(
            "Leave blank for the built-in parser, or paste an AI-generated regex here."
        )
        regex_help.setObjectName("fieldHelp")
        regex_help.setWordWrap(True)
        regex_ai_help = QLabel(
            "Copy the prompt, paste it into an AI with your example filenames, then paste the returned regex above."
        )
        regex_ai_help.setObjectName("fieldHelp")
        regex_ai_help.setWordWrap(True)
        self.copy_regex_prompt_button = QPushButton("Copy regex prompt")
        self.copy_regex_prompt_button.setObjectName("linkButton")
        self.copy_regex_prompt_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.copy_regex_prompt_button.setFlat(True)
        self.copy_regex_prompt_button.clicked.connect(self.copy_filename_regex_prompt)

        regex_col = QVBoxLayout()
        regex_col.setContentsMargins(0, 0, 0, 0)
        regex_col.setSpacing(6)
        regex_col.addLayout(regex_label_row)
        regex_col.addWidget(self.filename_pattern_line)
        regex_col.addWidget(regex_help)
        regex_col.addWidget(regex_ai_help)
        regex_col.addWidget(
            self.copy_regex_prompt_button,
            0,
            Qt.AlignmentFlag.AlignLeft,
        )

        grid.addLayout(regex_col, 1, 0, 1, 2)

        # Output path (full width)
        out_label_row = self._field_label_row(
            "Output workbook",
            "Where the final Excel workbook will be saved. The workbook contains All Data, "
            "Kept Data, and Averaged Data sheets.",
        )
        out_row = QHBoxLayout()
        out_row.setContentsMargins(0, 0, 0, 0)
        out_row.setSpacing(6)
        self.output_line = QLineEdit(str(self.output_path))
        self.output_line.setObjectName("pathInput")
        self.output_line.textChanged.connect(self._output_path_changed)
        self.browse_button = QPushButton("Browse…")
        self.browse_button.clicked.connect(self.browse_output_path)
        out_row.addWidget(self.output_line, 1)
        out_row.addWidget(self.browse_button)

        out_col = QVBoxLayout()
        out_col.setContentsMargins(0, 0, 0, 0)
        out_col.setSpacing(6)
        out_col.addLayout(out_label_row)
        out_col.addLayout(out_row)

        grid.addLayout(out_col, 2, 0, 1, 2)
        layout.addLayout(grid)

        # Advanced thresholds (always shown)
        layout.addSpacing(14)
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setObjectName("hairline")
        layout.addWidget(sep)
        layout.addSpacing(12)

        thresholds_label_row = self._field_label_row(
            "Pair-difference thresholds",
            "Green is at or below the first value. Yellow is between the two values. Red and review alerts are above the alert value.",
        )
        layout.addLayout(thresholds_label_row)
        layout.addSpacing(6)

        self.advanced_panel = QFrame()
        self.advanced_panel.setObjectName("advancedPanel")
        adv_layout = QGridLayout(self.advanced_panel)
        adv_layout.setContentsMargins(0, 0, 0, 0)
        adv_layout.setHorizontalSpacing(14)
        adv_layout.setVerticalSpacing(6)

        self.green_max_spin = self._make_threshold_spin(
            self.diff_green_max,
            tip="Pair differences at or below this value are highlighted green during review.",
        )
        self.alert_spin = self._make_threshold_spin(
            self.pair_alert_threshold,
            tip="Pair differences above this value are red and flagged for review, even if there are only two entries.",
        )
        self.green_max_spin.valueChanged.connect(self._settings_changed)
        self.alert_spin.valueChanged.connect(self._settings_changed)

        adv_layout.addWidget(
            self._threshold_label("Green up to", "#15803d", self.green_max_spin.toolTip()), 0, 0
        )
        adv_layout.addWidget(self._threshold_value_control(self.green_max_spin), 1, 0)
        adv_layout.addWidget(
            self._threshold_label("Alert above", "#b91c1c", self.alert_spin.toolTip()), 0, 1
        )
        adv_layout.addWidget(self._threshold_value_control(self.alert_spin), 1, 1)

        layout.addWidget(self.advanced_panel)

        return card

    def _field_label_row(self, text: str, help_tip: str) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(6)
        label = QLabel(text)
        label.setObjectName("fieldLabel")
        row.addWidget(label)
        row.addWidget(HelpButton(help_tip))
        row.addStretch(1)
        return row

    def _threshold_label(self, text: str, swatch_color: str, tip: str) -> QWidget:
        wrap = QFrame()
        layout = QHBoxLayout(wrap)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        swatch = QLabel()
        swatch.setFixedSize(9, 9)
        swatch.setStyleSheet(f"background: {swatch_color}; border-radius: 4px;")

        label = QLabel(text)
        label.setObjectName("thresholdLabel")

        layout.addWidget(swatch)
        layout.addWidget(label)
        layout.addWidget(HelpButton(tip))
        layout.addStretch(1)
        return wrap

    def _make_threshold_spin(self, value: float, tip: str) -> QDoubleSpinBox:
        spin = QDoubleSpinBox()
        spin.setObjectName("thresholdValueSpin")
        spin.setRange(0, 100)
        spin.setDecimals(1)
        spin.setSingleStep(0.5)
        spin.setSuffix(" mmHg")
        spin.setValue(value)
        spin.setFixedHeight(38)
        spin.setMinimumWidth(145)
        spin.setToolTip(tip)
        spin.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        spin.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.NoButtons)
        return spin

    def _threshold_value_control(self, spin: QDoubleSpinBox) -> QWidget:
        container = QWidget()
        container.setObjectName("thresholdStepper")
        container.setMinimumWidth(215)
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        down_button = QPushButton("-")
        down_button.setObjectName("stepperButton")
        down_button.setFixedWidth(30)
        down_button.setFixedHeight(38)
        down_button.setToolTip("Decrease threshold")
        down_button.setCursor(Qt.CursorShape.PointingHandCursor)
        down_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        down_button.clicked.connect(spin.stepDown)

        up_button = QPushButton("+")
        up_button.setObjectName("stepperButton")
        up_button.setFixedWidth(30)
        up_button.setFixedHeight(38)
        up_button.setToolTip("Increase threshold")
        up_button.setCursor(Qt.CursorShape.PointingHandCursor)
        up_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        up_button.clicked.connect(spin.stepUp)

        layout.addWidget(spin, 1)
        layout.addWidget(down_button)
        layout.addWidget(up_button)
        return container

    # ---- Review screen ----
    def _build_review_screen(self) -> QWidget:
        screen = QFrame()
        screen.setObjectName("reviewScreen")
        layout = QVBoxLayout(screen)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Run summary header
        header = QFrame()
        header.setObjectName("runSummary")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(28, 12, 28, 12)
        header_layout.setSpacing(18)

        header_layout.addWidget(self._build_stepper(active="review"))
        header_layout.addStretch(1)

        self.edit_setup_button = QPushButton("✎  Edit setup")
        self.edit_setup_button.setObjectName("linkButton")
        self.edit_setup_button.setFlat(True)
        self.edit_setup_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.edit_setup_button.clicked.connect(self._show_import_screen)
        header_layout.addWidget(self.edit_setup_button)

        layout.addWidget(header)

        # Tab bar with action buttons
        tab_row = QFrame()
        tab_row.setObjectName("tabRow")
        tab_layout = QHBoxLayout(tab_row)
        tab_layout.setContentsMargins(20, 0, 20, 0)
        tab_layout.setSpacing(0)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("mainTabs")
        self.tabs.setDocumentMode(True)
        self.tabs.addTab(self._build_overview_tab(), "Overview")
        self.tabs.addTab(self._build_review_tab(), "Review")
        self.tabs.addTab(self._build_all_data_tab(), "All data")
        self.tabs.addTab(self._build_averaged_tab(), "Averaged")

        tab_layout.addWidget(self.tabs, 1)

        # Right-side actions overlay
        action_holder = QWidget()
        action_layout = QHBoxLayout(action_holder)
        action_layout.setContentsMargins(8, 2, 0, 2)
        action_layout.setSpacing(8)

        self.export_button = QPushButton("Export workbook")
        self.export_button.setObjectName("exportButton")
        self.export_button.setMinimumWidth(180)
        self.export_button.setFixedHeight(34)
        self.export_button.clicked.connect(self.export_excel)
        action_layout.addWidget(self.export_button)

        self.open_file_button = QPushButton("Open file")
        self.open_file_button.clicked.connect(self.open_export_file)
        self.open_file_button.setVisible(False)
        action_layout.addWidget(self.open_file_button)

        self.open_folder_button = QPushButton("Open folder")
        self.open_folder_button.clicked.connect(self.open_export_folder)
        self.open_folder_button.setVisible(False)
        action_layout.addWidget(self.open_folder_button)

        self.tabs.setCornerWidget(action_holder, Qt.Corner.TopRightCorner)

        layout.addWidget(tab_row, 1)

        return screen

    def _build_overview_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(28, 22, 28, 22)
        layout.setSpacing(16)

        self.overview_banner = self._make_banner("", "neutral")
        self.overview_banner.setVisible(False)
        layout.addWidget(self.overview_banner)

        # Stat cards
        stats_row = QHBoxLayout()
        stats_row.setSpacing(12)

        self.stat_files = self._make_stat_card("Files processed", "")
        self.stat_averaged = self._make_stat_card("Patients averaged", "accent")
        self.stat_reviewed = self._make_stat_card(
            "Confirmed pairs",
            tooltip="Patients that needed manual review (multi-entry or pair-alert) and were confirmed.",
        )
        self.stat_special = self._make_stat_card(
            "Wrong-type / unrecognized",
            "warn",
            tooltip="PDFs that didn't match the selected report type, or weren't recognized at all.",
        )
        for card in (self.stat_files, self.stat_averaged, self.stat_reviewed, self.stat_special):
            stats_row.addWidget(card, 1)
        layout.addLayout(stats_row)

        # Section heading
        head_row = QHBoxLayout()
        head_label = QLabel("All processed files")
        head_label.setObjectName("sectionTitle")
        head_sub = QLabel("Double-click a row to open the source PDF")
        head_sub.setObjectName("sectionSub")
        head_row.addWidget(head_label)
        head_row.addStretch(1)
        head_row.addWidget(head_sub)
        layout.addLayout(head_row)

        self.overview_table = QTableWidget(0, 8)
        self.overview_table.setHorizontalHeaderLabels(
            [
                "Source File",
                "Patient ID",
                "Subject",
                "Visit",
                "Timepoint",
                "Report type",
                "Scan date",
                "Status",
            ]
        )
        self._configure_table(self.overview_table)
        self.overview_table.itemDoubleClicked.connect(self._open_overview_pdf)
        layout.addWidget(self.overview_table, 1)

        return tab

    def _build_review_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(28, 22, 28, 22)
        layout.setSpacing(14)

        self.unconfirmed_banner = self._make_banner(
            "All patients confirmed.", "success"
        )
        layout.addWidget(self.unconfirmed_banner)

        # Two-column split: queue on left, detail on right
        split = QHBoxLayout()
        split.setContentsMargins(0, 0, 0, 0)
        split.setSpacing(16)

        # Queue
        queue_card = QFrame()
        queue_card.setObjectName("queueCard")
        queue_card.setMinimumWidth(260)
        queue_card.setMaximumWidth(320)
        queue_layout = QVBoxLayout(queue_card)
        queue_layout.setContentsMargins(0, 0, 0, 0)
        queue_layout.setSpacing(0)

        queue_head = QFrame()
        queue_head.setObjectName("queueHead")
        qh_layout = QHBoxLayout(queue_head)
        qh_layout.setContentsMargins(12, 10, 12, 10)
        qh_layout.setSpacing(8)
        qh_label = QLabel("Review queue")
        qh_label.setObjectName("queueHeadTitle")
        self.queue_count_pill = QLabel("0 of 0")
        self.queue_count_pill.setObjectName("pillWarn")
        qh_layout.addWidget(qh_label, 1)
        qh_layout.addWidget(self.queue_count_pill)
        queue_layout.addWidget(queue_head)

        self.queue_list = QListWidget()
        self.queue_list.setObjectName("queueList")
        self.queue_list.setFrameShape(QFrame.Shape.NoFrame)
        self.queue_list.currentRowChanged.connect(self._on_queue_changed)
        queue_layout.addWidget(self.queue_list, 1)

        split.addWidget(queue_card)

        # Detail
        self.detail_scroll = QScrollArea()
        self.detail_scroll.setObjectName("reviewDetailScroll")
        self.detail_scroll.setWidgetResizable(True)
        self.detail_scroll.setFrameShape(QFrame.Shape.NoFrame)

        detail_widget = QWidget()
        detail_widget.setObjectName("reviewDetailContent")
        detail_widget.setAutoFillBackground(True)
        palette = detail_widget.palette()
        palette.setColor(QPalette.ColorRole.Window, QColor("#ffffff"))
        detail_widget.setPalette(palette)
        self.detail_scroll.setWidget(detail_widget)
        self.detail_layout = QVBoxLayout(detail_widget)
        self.detail_layout.setContentsMargins(2, 2, 2, 2)
        self.detail_layout.setSpacing(14)

        # Header row
        self.detail_header = QFrame()
        self.detail_header.setObjectName("detailHeader")
        dh_layout = QHBoxLayout(self.detail_header)
        dh_layout.setContentsMargins(0, 0, 0, 0)
        dh_layout.setSpacing(16)
        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        self.patient_title = QLabel("No patient selected")
        self.patient_title.setObjectName("patientTitle")
        self.patient_subtitle = QLabel("")
        self.patient_subtitle.setObjectName("patientSubtitle")
        title_box.addWidget(self.patient_title)
        title_box.addWidget(self.patient_subtitle)
        dh_layout.addLayout(title_box, 1)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        self.reset_auto_button = QPushButton("↻ Reset to auto")
        self.reset_auto_button.setToolTip("Restore the auto-selected pair")
        self.reset_auto_button.clicked.connect(self.reset_current_to_auto)
        self.view_pdf_button = QPushButton("👁 View PDF")
        self.view_pdf_button.clicked.connect(self.open_current_pair_pdf)
        actions.addWidget(self.reset_auto_button)
        actions.addWidget(self.view_pdf_button)
        dh_layout.addLayout(actions)

        self.detail_layout.addWidget(self.detail_header)

        # Reason callout + Confirm button (same row)
        reason_row = QHBoxLayout()
        reason_row.setContentsMargins(0, 0, 0, 0)
        reason_row.setSpacing(10)
        self.reason_callout = QFrame()
        self.reason_callout.setObjectName("reasonCallout")
        rc_layout = QHBoxLayout(self.reason_callout)
        rc_layout.setContentsMargins(12, 9, 12, 9)
        rc_layout.setSpacing(8)

        self.reason_icon = QLabel("⚠")
        self.reason_icon.setObjectName("reasonIcon")
        self.reason_text = QLabel("Pair alert")
        self.reason_text.setObjectName("reasonText")
        self.reason_help = HelpButton("")
        self.reason_help.setObjectName("reasonHelp")
        rc_layout.addWidget(self.reason_icon)
        rc_layout.addWidget(self.reason_text)
        rc_layout.addWidget(self.reason_help)
        reason_row.addWidget(self.reason_callout, 0, Qt.AlignmentFlag.AlignLeft)
        reason_row.addStretch(1)

        self.confirm_button = QPushButton("Confirm pair")
        self.confirm_button.setObjectName("confirmPairButton")
        self.confirm_button.setCheckable(True)
        self.confirm_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.confirm_button.setToolTip(
            "Tick once you've reviewed the selected rows. Unticked patients still "
            "export using the auto-pair."
        )
        self.confirm_button.toggled.connect(self._on_confirm_toggled)
        reason_row.addWidget(self.confirm_button, 0, Qt.AlignmentFlag.AlignRight)
        self.detail_layout.addLayout(reason_row)

        # Diff section header + strip
        diff_head = QHBoxLayout()
        diff_title = QLabel("Selected pair · absolute differences")
        diff_title.setObjectName("subSectionTitle")
        diff_head.addWidget(diff_title)
        diff_head.addWidget(HelpButton(
            "Difference between the two rows you've kept. Green is at or below the green threshold, "
            "yellow is between green and alert, and red exceeds the alert threshold."
        ))
        diff_head.addStretch(1)
        self.diff_threshold_label = QLabel("")
        self.diff_threshold_label.setObjectName("sectionSub")
        diff_head.addWidget(self.diff_threshold_label)
        self.detail_layout.addLayout(diff_head)

        self.diff_strip = QFrame()
        self.diff_strip.setObjectName("diffStrip")
        diff_layout = QHBoxLayout(self.diff_strip)
        diff_layout.setContentsMargins(0, 0, 0, 0)
        diff_layout.setSpacing(0)
        self.diff_cells: list[tuple[QLabel, QLabel, QFrame]] = []
        for label_text in ("SYS", "DIA", "MAP", "Aortic SYS", "Aortic DIA"):
            cell = QFrame()
            cell.setObjectName("diffCell")
            cell_layout = QVBoxLayout(cell)
            cell_layout.setContentsMargins(14, 10, 14, 10)
            cell_layout.setSpacing(2)
            lbl = QLabel(label_text)
            lbl.setObjectName("diffLabel")
            val = QLabel("—")
            val.setObjectName("diffValue")
            cell_layout.addWidget(lbl)
            cell_layout.addWidget(val)
            self.diff_cells.append((lbl, val, cell))
            diff_layout.addWidget(cell, 1)
        self.detail_layout.addWidget(self.diff_strip)

        # Pair table
        pair_card = QFrame()
        pair_card.setObjectName("pairCard")
        pair_layout = QVBoxLayout(pair_card)
        pair_layout.setContentsMargins(0, 0, 0, 0)
        pair_layout.setSpacing(0)

        pair_head = QFrame()
        pair_head.setObjectName("pairCardHead")
        ph_layout = QHBoxLayout(pair_head)
        ph_layout.setContentsMargins(14, 11, 14, 11)
        ph_layout.setSpacing(8)
        ph_title = QLabel("Measurements")
        ph_title.setObjectName("pairCardTitle")
        ph_help_row = QHBoxLayout()
        ph_help_row.setContentsMargins(0, 0, 0, 0)
        ph_help_row.setSpacing(6)
        ph_help_text = QLabel("Tick the two rows to keep")
        ph_help_text.setObjectName("pairCardHelpText")
        ph_help_row.addWidget(ph_help_text)
        ph_help_row.addWidget(HelpButton(
            "The two checked rows will be averaged in the final export. "
            "The auto-paired choice is preselected."
        ))
        ph_layout.addWidget(ph_title, 1)
        ph_layout.addLayout(ph_help_row)
        pair_layout.addWidget(pair_head)

        self.pair_table = QTableWidget(0, 8)
        self.pair_table.setHorizontalHeaderLabels(
            ["Keep", "SYS", "DIA", "MAP", "Ao SYS", "Ao DIA", "Source", "Pair method"]
        )
        self._configure_table(self.pair_table)
        self.pair_table.verticalHeader().setVisible(False)
        self.pair_table.verticalHeader().setDefaultSectionSize(44)
        self.pair_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.pair_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.pair_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.pair_table.itemDoubleClicked.connect(self.open_current_pair_pdf)
        pair_layout.addWidget(self.pair_table)

        self.detail_layout.addWidget(pair_card)

        self.detail_layout.addStretch(1)

        split.addWidget(self.detail_scroll, 1)
        layout.addLayout(split, 1)

        return tab

    def _build_all_data_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(28, 22, 28, 22)
        layout.setSpacing(10)

        helper = QLabel("Preview of every row that will populate the workbook's All Data sheet.")
        helper.setObjectName("sectionSub")
        layout.addWidget(helper)

        self.all_data_table = QTableWidget()
        self._configure_table(self.all_data_table)
        self.all_data_table.itemDoubleClicked.connect(self._open_all_data_pdf)
        layout.addWidget(self.all_data_table, 1)

        return tab

    def _build_averaged_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(28, 22, 28, 22)
        layout.setSpacing(10)

        helper = QLabel("One averaged row per patient — what the Averaged Data sheet will contain.")
        helper.setObjectName("sectionSub")
        layout.addWidget(helper)

        self.averaged_table = QTableWidget()
        self._configure_table(self.averaged_table)
        layout.addWidget(self.averaged_table, 1)

        return tab

    def _configure_table(self, table: QTableWidget) -> None:
        table.setObjectName("dataTable")
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        table.horizontalHeader().setMinimumSectionSize(60)
        table.horizontalHeader().setStretchLastSection(False)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        table.setItemDelegate(NoFocusItemDelegate(table))
        table.setAlternatingRowColors(False)
        table.setShowGrid(False)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

    def _make_stat_card(
        self, label_text: str, value_class: str = "", tooltip: Optional[str] = None
    ) -> QFrame:
        card = QFrame()
        card.setObjectName("statCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(2)

        value_label = QLabel("0")
        value_label.setObjectName("statValue")
        if value_class:
            value_label.setProperty("variant", value_class)

        label_row = QHBoxLayout()
        label_row.setContentsMargins(0, 0, 0, 0)
        label_row.setSpacing(5)
        label = QLabel(label_text)
        label.setObjectName("statLabel")
        label_row.addWidget(label)
        if tooltip:
            label_row.addWidget(HelpButton(tooltip))
        label_row.addStretch(1)

        layout.addWidget(value_label)
        layout.addLayout(label_row)
        card._value_label = value_label  # type: ignore[attr-defined]
        return card

    def _make_banner(self, text: str, variant: str) -> QFrame:
        banner = QFrame()
        banner.setObjectName("banner")
        banner.setProperty("variant", variant)
        layout = QHBoxLayout(banner)
        layout.setContentsMargins(14, 11, 14, 11)
        layout.setSpacing(10)

        icon = QLabel("ⓘ")
        icon.setObjectName("bannerIcon")
        body = QLabel(text)
        body.setObjectName("bannerBody")
        body.setWordWrap(True)
        layout.addWidget(icon, 0, Qt.AlignmentFlag.AlignTop)
        layout.addWidget(body, 1)
        banner._body_label = body  # type: ignore[attr-defined]
        return banner

    # =================================================================
    #   State transitions
    # =================================================================
    def _show_import_screen(self) -> None:
        self.stack.setCurrentIndex(0)
        self._update_import_meta()

    def _show_review_screen(self) -> None:
        self.stack.setCurrentIndex(1)

    # =================================================================
    #   File handling
    # =================================================================
    def add_pdf_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self, "Add PWA report PDFs", "", "PDF files (*.pdf)"
        )
        if not files:
            return
        self._on_files_dropped([Path(f) for f in files])

    def _on_files_dropped(self, paths: list[Path]) -> None:
        existing = set(self.pdf_paths)
        added = 0
        for p in paths:
            if p not in existing:
                self.pdf_paths.append(p)
                existing.add(p)
                added += 1
        if added:
            self._refresh_file_list()
            self._update_status(f"Added {added} file{'s' if added != 1 else ''}")

    def _refresh_file_list(self) -> None:
        self.file_list.clear()
        has_files = bool(self.pdf_paths)
        self.file_list.setVisible(has_files)
        self.file_meta.setVisible(has_files)
        self.drop_zone.set_compact(False)

        for path in self.pdf_paths:
            item = QListWidgetItem(path.name)
            item.setData(Qt.ItemDataRole.UserRole, str(path))
            item.setToolTip(str(path))
            self.file_list.addItem(item)

        count = len(self.pdf_paths)
        self.file_count_label.setText(
            f"<span style='color:#0f172a;font-weight:600;'>{count}</span> "
            f"file{'s' if count != 1 else ''} selected"
        )
        self._sync_remove_file_button()
        self._update_import_meta()

    def _sync_remove_file_button(self) -> None:
        if hasattr(self, "remove_selected_files_link"):
            self.remove_selected_files_link.setEnabled(bool(self.file_list.selectedItems()))

    def remove_selected_pdf_files(self) -> None:
        selected_paths = {
            Path(item.data(Qt.ItemDataRole.UserRole))
            for item in self.file_list.selectedItems()
        }
        if not selected_paths:
            return

        before_count = len(self.pdf_paths)
        self.pdf_paths = [
            path for path in self.pdf_paths if path not in selected_paths
        ]
        removed_count = before_count - len(self.pdf_paths)
        if removed_count:
            self._refresh_file_list()
            self._update_status(
                f"Removed {removed_count} file{'s' if removed_count != 1 else ''}"
            )

    def clear_pdf_files(self) -> None:
        if not self.pdf_paths:
            return
        self.pdf_paths.clear()
        self._refresh_file_list()
        self._update_status("Cleared file list")

    def _update_import_meta(self) -> None:
        count = len(self.pdf_paths)
        if count == 0:
            self.import_meta_label.setText("Add at least one PDF to continue")
            self.process_button.setEnabled(False)
            self.process_button.setText("Process PDFs  →")
            return
        report_label = "Detailed" if self.report_mode == REPORT_MODE_DETAILED else "Clinical"
        grouping_label = GROUPING_MODE_LABELS[self.grouping_mode]
        self.import_meta_label.setText(
            f"<span style='color:#0f172a;font-weight:600;'>{count}</span> file{'s' if count != 1 else ''} · "
            f"{report_label} · {grouping_label}"
        )
        self.process_button.setEnabled(True)
        self.process_button.setText(
            f"Process {count} PDF{'s' if count != 1 else ''}  →"
        )

    def _on_report_mode_changed(self, value: str) -> None:
        self.report_mode = value
        self.report_mode_help.setText(REPORT_MODE_HELP.get(value, ""))
        self._update_import_meta()

    def _on_grouping_changed(self, _index: int) -> None:
        value = self.grouping_combo.currentData()
        if value:
            self.grouping_mode = value
            self._update_import_meta()

    def _filename_pattern_changed(self, text: str) -> None:
        self.filename_pattern = text.strip()
        self._update_import_meta()

    def _filename_regex_prompt(self) -> str:
        grouping_label = GROUPING_MODE_LABELS.get(self.grouping_mode, self.grouping_mode)
        filenames = [path.name for path in self.pdf_paths[:20]]
        if filenames:
            filename_lines = "\n".join(f"- {name}" for name in filenames)
        else:
            filename_lines = (
                "- IAS003 PWA1.pdf\n"
                "- IAS003_T2 PWA1.pdf\n"
                "- Replace these with my actual filenames before answering."
            )

        return (
            "Help me determine how these PWA PDF filenames should be parsed into "
            "subject, visit, and timepoint values, then write one Python-compatible "
            "regular expression for the PWA Data Extractor app.\n\n"
            "First, ask me what output I expect for each filename if it is not already "
            "clear. Ask follow-up questions until you can confidently determine which "
            "part of the filename is the subject, which part is the visit if any, and "
            "which part is the timepoint if any. Do not provide the final regex until "
            "you have enough information.\n\n"
            "The app will run the regex against the filename stem only, meaning the "
            ".pdf extension is removed before matching.\n\n"
            "When you have enough information, provide the final answer as only the "
            "regex pattern. Do not include quotes, explanation, or a code block.\n\n"
            "Regex requirements:\n"
            "- Use Python named capture groups.\n"
            "- Include a required named group: (?P<subject>...).\n"
            "- Optionally include named groups: (?P<visit>...) and "
            "(?P<timepoint>...).\n"
            "- The subject should include any study prefix before the patient number. "
            "For example, IAS003 should remain IAS003, not 003.\n"
            "- Do not treat PWA1, PWA2, Report1, Run1, or similar measurement suffixes "
            "as visits or timepoints unless the filename has an explicit visit or "
            "timepoint token.\n"
            "- For IAS003 PWA1, expected subject is IAS003 with no visit and no "
            "timepoint.\n"
            "- For IAS003_T2 PWA1, expected subject is IAS003 and expected timepoint "
            "is 2.\n\n"
            f"Current grouping mode in the app: {grouping_label}\n\n"
            "Example filenames:\n"
            f"{filename_lines}\n"
        )

    def copy_filename_regex_prompt(self) -> None:
        QApplication.clipboard().setText(self._filename_regex_prompt())
        self.copy_regex_prompt_button.setText("Copied")
        self.copy_regex_prompt_button.setProperty("copied", True)
        self.copy_regex_prompt_button.style().unpolish(self.copy_regex_prompt_button)
        self.copy_regex_prompt_button.style().polish(self.copy_regex_prompt_button)
        QTimer.singleShot(1800, self._reset_regex_prompt_button)
        self._update_status("Copied regex prompt to clipboard")

    def _reset_regex_prompt_button(self) -> None:
        if not hasattr(self, "copy_regex_prompt_button"):
            return
        self.copy_regex_prompt_button.setText("Copy regex prompt")
        self.copy_regex_prompt_button.setProperty("copied", False)
        self.copy_regex_prompt_button.style().unpolish(self.copy_regex_prompt_button)
        self.copy_regex_prompt_button.style().polish(self.copy_regex_prompt_button)

    def _output_path_changed(self, text: str) -> None:
        self.output_path = Path(text)

    def browse_output_path(self) -> None:
        suggested = str(self.output_path)
        target, _ = QFileDialog.getSaveFileName(
            self, "Choose Excel output", suggested, "Excel files (*.xlsx)"
        )
        if target:
            self.output_line.setText(target)

    def _settings_changed(self) -> None:
        self.diff_green_max = self.green_max_spin.value()
        self.pair_alert_threshold = self.alert_spin.value()
        if self.pair_alert_threshold < self.diff_green_max:
            self.alert_spin.blockSignals(True)
            self.alert_spin.setValue(self.diff_green_max)
            self.alert_spin.blockSignals(False)
            self.pair_alert_threshold = self.diff_green_max
        if self.bundle is not None:
            self._rebuild_analysis()

    # =================================================================
    #   Processing
    # =================================================================
    def process_files(self) -> None:
        if not self.pdf_paths:
            return
        self.process_button.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setRange(0, len(self.pdf_paths))
        self.progress.setValue(0)
        self._update_status("Processing PDFs…")

        self.thread = QThread(self)
        self.worker = ProcessingWorker(
            list(self.pdf_paths),
            self.report_mode,
            self.grouping_mode,
            self.filename_pattern or None,
        )
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_processing_progress)
        self.worker.finished.connect(self._on_processing_finished)
        self.worker.failed.connect(self._on_processing_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

    def _on_processing_progress(self, current: int, total: int, message: str) -> None:
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        self._update_status(message)

    def _on_processing_finished(self, records: list[dict[str, object]]) -> None:
        self.records = records
        self.confirmed_patients.clear()
        self.manual_pairs = {}
        self._rebuild_analysis()
        self.progress.setVisible(False)
        self.process_button.setEnabled(True)
        self._show_review_screen()

    def _on_processing_failed(self, traceback_text: str) -> None:
        self.progress.setVisible(False)
        self.process_button.setEnabled(True)
        self._update_status("Processing failed")
        QMessageBox.critical(
            self,
            "Processing failed",
            f"An error occurred while processing the PDFs:\n\n{traceback_text}",
        )

    def _rebuild_analysis(self) -> None:
        if not self.records:
            self.bundle = None
            return
        manual_tuples = {
            pid: tuple(rows[:2])  # type: ignore[arg-type]
            for pid, rows in self.manual_pairs.items()
            if len(rows) == 2
        }
        self.bundle = build_analysis(
            self.records,
            manual_pairs=manual_tuples,
            mode=ANALYSIS_MODE,
            pair_alert_threshold=self.pair_alert_threshold,
        )
        self.auto_pairs = dict(self.bundle.auto_pairs)
        review_patient_ids = [item.patient_id for item in self.bundle.review_items]
        self.manual_pairs = initial_manual_pairs(
            self.bundle.dataframe, self.bundle.used_pairs, review_patient_ids
        )
        # Drop confirmations for patients no longer in review.
        self.confirmed_patients &= set(review_patient_ids)

        self._refresh_overview_tab()
        self._refresh_review_queue()
        self._refresh_all_data_tab()
        self._refresh_averaged_tab()
        self._refresh_unconfirmed_banner()
        self._refresh_tab_titles()

    # =================================================================
    #   Overview tab
    # =================================================================
    def _refresh_overview_tab(self) -> None:
        if self.bundle is None:
            self.overview_table.setRowCount(0)
            return
        df = display_dataframe(self.bundle)
        self.stat_files.findChild(QLabel, "statValue").setText(str(len(df)))  # type: ignore[union-attr]
        self.stat_averaged.findChild(QLabel, "statValue").setText(str(len(self.bundle.analyzed_df)))  # type: ignore[union-attr]
        self.stat_reviewed.findChild(QLabel, "statValue").setText(str(len(self.confirmed_patients)))  # type: ignore[union-attr]

        special_count = int(self.bundle.special_row_mask.sum())
        self.stat_special.findChild(QLabel, "statValue").setText(str(special_count))  # type: ignore[union-attr]

        # Banner
        if self.last_export_path is not None:
            self.overview_banner.setVisible(True)
            self.overview_banner.setProperty("variant", "success")
            self.overview_banner._body_label.setText(  # type: ignore[attr-defined]
                f"<b>Workbook exported.</b> {self.last_export_path.name} — "
                f"{len(self.bundle.analyzed_df)} averaged patients, {len(df)} rows."
            )
            self.overview_banner.style().unpolish(self.overview_banner)
            self.overview_banner.style().polish(self.overview_banner)
        else:
            self.overview_banner.setVisible(False)

        # Table
        self.overview_table.setRowCount(len(df))
        for row_index, (_, row) in enumerate(df.iterrows()):
            is_special = bool(row.get("Special Row", False))
            cells = [
                str(row.get("Source File") or ""),
                "" if is_special else str(row.get("Patient ID") or ""),
                "" if is_special else str(row.get("Subject ID") or ""),
                "" if is_special else (str(row.get("Visit") or "—")),
                "" if is_special else (str(row.get("Timepoint") or "")),
                "" if is_special else str(row.get("Report Type") or ""),
                "" if is_special else str(row.get("Scan Date") or ""),
                record_status(row.get("Patient ID")) if is_special else (
                    "Kept" if row.get("Analyed") == "Yes" else "Extra"
                ),
            ]
            for col_index, text in enumerate(cells):
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.overview_table.setItem(row_index, col_index, item)
        self.overview_table.resizeColumnsToContents()

    def _open_overview_pdf(self, item: QTableWidgetItem) -> None:
        row = item.row()
        source_item = self.overview_table.item(row, 0)
        if source_item is None:
            return
        filename = source_item.text()
        for path in self.pdf_paths:
            if path.name == filename:
                self._open_pdf(path)
                return

    # =================================================================
    #   Review tab
    # =================================================================
    def _refresh_review_queue(self) -> None:
        self.queue_list.blockSignals(True)
        self.queue_list.clear()
        if self.bundle is None or not self.bundle.review_items:
            self.queue_count_pill.setText("0 of 0")
            self.queue_count_pill.setProperty("variant", "ok")
            self.queue_list.blockSignals(False)
            self._update_detail_view(None)
            return

        for review_item in self.bundle.review_items:
            label_text = self._queue_item_text(review_item.patient_id, review_item.reason)
            list_item = QListWidgetItem(label_text)
            list_item.setData(Qt.ItemDataRole.UserRole, review_item.patient_id)
            if review_item.patient_id in self.confirmed_patients:
                list_item.setForeground(QColor("#94a3b8"))
            self.queue_list.addItem(list_item)

        total = len(self.bundle.review_items)
        left = total - len(self.confirmed_patients)
        if left == 0:
            self.queue_count_pill.setText("all confirmed")
            self.queue_count_pill.setProperty("variant", "ok")
        else:
            self.queue_count_pill.setText(f"{left} of {total} left")
            self.queue_count_pill.setProperty("variant", "warn")
        self.queue_count_pill.style().unpolish(self.queue_count_pill)
        self.queue_count_pill.style().polish(self.queue_count_pill)

        self.queue_list.setCurrentRow(0)
        self.queue_list.blockSignals(False)
        self._on_queue_changed(0)

    def _queue_item_text(self, patient_id: str, reason: str) -> str:
        reason_label, _ = REVIEW_REASON_LABELS.get(reason, ("review", "warn"))
        confirmed = "✓ " if patient_id in self.confirmed_patients else ""
        return f"{confirmed}{patient_id}    [{reason_label}]"

    def _on_queue_changed(self, row: int) -> None:
        if self.bundle is None or row < 0 or row >= self.queue_list.count():
            self._update_detail_view(None)
            return
        item = self.queue_list.item(row)
        patient_id = item.data(Qt.ItemDataRole.UserRole) if item else None
        self._update_detail_view(patient_id)

    def _update_detail_view(self, patient_id: Optional[str]) -> None:
        if patient_id is None or self.bundle is None:
            self.patient_title.setText("No patient selected")
            self.patient_subtitle.setText("")
            self.reason_callout.setVisible(False)
            self._set_confirm_button_state(False, visible=False)
            self.pair_table.setRowCount(0)
            for _, val_label, cell in self.diff_cells:
                val_label.setText("—")
                cell.setProperty("variant", "")
                cell.style().unpolish(cell)
                cell.style().polish(cell)
            self.diff_strip.setProperty("variant", "")
            self.diff_strip.style().unpolish(self.diff_strip)
            self.diff_strip.style().polish(self.diff_strip)
            return

        df = self.bundle.dataframe
        rows = patient_rows(df, patient_id)
        first_row = rows.iloc[0] if not rows.empty else None
        self.patient_title.setText(f"Patient {patient_id}")
        if first_row is not None:
            subject = first_row.get("Subject ID") or "—"
            visit = first_row.get("Visit") or "—"
            timepoint = first_row.get("Timepoint") or "—"
            entry_count = len(rows)
            self.patient_subtitle.setText(
                f"Subject {subject} · Visit {visit} · Timepoint {timepoint} · "
                f"{entry_count} entries"
            )
        else:
            self.patient_subtitle.setText("")

        # Reason callout
        reason = next(
            (item.reason for item in self.bundle.review_items if item.patient_id == patient_id),
            None,
        )
        if reason:
            label, variant = REVIEW_REASON_LABELS.get(reason, ("review", "warn"))
            self.reason_callout.setProperty("variant", variant)
            self.reason_text.setText(label.title())
            self.reason_help.setToolTip(REVIEW_REASON_DESCRIPTIONS.get(reason, ""))
            self.reason_callout.style().unpolish(self.reason_callout)
            self.reason_callout.style().polish(self.reason_callout)
            self.reason_callout.setVisible(True)
        else:
            self.reason_callout.setVisible(False)

        # Diff threshold label
        self.diff_threshold_label.setText(
            f"Threshold: green <= {self.diff_green_max:g} | "
            f"yellow <= {self.pair_alert_threshold:g} | "
            f"alert > {self.pair_alert_threshold:g}"
        )

        # Populate pair table
        self._populate_pair_table(patient_id, rows)

        # Confirm button state
        self._set_confirm_button_state(
            patient_id in self.confirmed_patients, visible=True
        )

    def _populate_pair_table(self, patient_id: str, rows: pd.DataFrame) -> None:
        self._updating_pair_table = True
        self.pair_table.setRowCount(len(rows))
        kept_indices = set(self.manual_pairs.get(patient_id, []))
        auto_pair = set(self.auto_pairs.get(patient_id, ()))

        for visual_row, (df_index, data_row) in enumerate(rows.iterrows()):
            # Keep checkbox
            keep_widget = QPushButton("✓" if df_index in kept_indices else "")
            keep_widget.setObjectName("keepToggle")
            keep_widget.setCheckable(True)
            keep_widget.setFixedSize(22, 22)
            keep_widget.setChecked(df_index in kept_indices)
            keep_widget.toggled.connect(
                lambda checked, pid=patient_id, idx=df_index, button=keep_widget: self._on_keep_toggled(
                    pid, idx, checked, button
                )
            )
            container = QWidget()
            cl = QHBoxLayout(container)
            cl.setContentsMargins(0, 0, 0, 0)
            cl.addWidget(keep_widget)
            cl.setAlignment(keep_widget, Qt.AlignmentFlag.AlignCenter)
            self.pair_table.setCellWidget(visual_row, 0, container)

            for col, col_name in enumerate(
                [
                    "Peripheral Systolic Pressure (mmHg)",
                    "Peripheral Diastolic Pressure (mmHg)",
                    "Peripheral Mean Pressure (mmHg)",
                    "Aortic Systolic Pressure (mmHg)",
                    "Aortic Diastolic Pressure (mmHg)",
                ],
                start=1,
            ):
                item = QTableWidgetItem(format_value(data_row.get(col_name)))
                item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                self.pair_table.setItem(visual_row, col, item)

            source_item = QTableWidgetItem(str(data_row.get("Source File") or ""))
            self.pair_table.setItem(visual_row, 6, source_item)
            method_item = QTableWidgetItem("Auto" if df_index in auto_pair else "Manual")
            method_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.pair_table.setItem(visual_row, 7, method_item)

        self.pair_table.resizeColumnsToContents()
        self._updating_pair_table = False
        self._refresh_diff_strip(patient_id)

    def _on_keep_toggled(
        self,
        patient_id: str,
        df_index: int,
        checked: bool,
        button: Optional[QPushButton] = None,
    ) -> None:
        if self._updating_pair_table:
            return
        kept = list(self.manual_pairs.get(patient_id, []))
        if checked:
            if df_index not in kept and len(kept) >= 2:
                self._updating_pair_table = True
                target_button = button if button is not None else self.sender()
                if isinstance(target_button, QPushButton):
                    target_button.setChecked(False)
                    target_button.setText("")
                self._updating_pair_table = False
                self._update_status("Select exactly two measurements to keep.")
                return
            if df_index not in kept:
                kept.append(df_index)
            if button is not None:
                button.setText("✓")
        else:
            kept = [i for i in kept if i != df_index]
            if button is not None:
                button.setText("")
        self.manual_pairs[patient_id] = kept
        if len(kept) == 2:
            # Recompute analysis to update averaged data
            self._rebuild_analysis()
            # Re-select the same patient
            self._select_queue_patient(patient_id)
        else:
            self._refresh_diff_strip(patient_id)

    def _refresh_diff_strip(self, patient_id: str) -> None:
        if self.bundle is None:
            return
        kept_indices = self.manual_pairs.get(patient_id, [])
        if len(kept_indices) != 2:
            for _, val_label, cell in self.diff_cells:
                val_label.setText("—")
                cell.setProperty("variant", "")
                cell.style().unpolish(cell)
                cell.style().polish(cell)
            return

        pair_df = self.bundle.dataframe.loc[list(kept_indices)]
        diffs = calculate_pair_differences(pair_df)
        diff_keys = [
            "Pair Diff Peripheral Systolic (mmHg)",
            "Pair Diff Peripheral Diastolic (mmHg)",
            "Pair Diff Peripheral Mean (mmHg)",
            "Pair Diff Aortic Systolic (mmHg)",
            "Pair Diff Aortic Diastolic (mmHg)",
        ]
        for (label, val_label, cell), key in zip(self.diff_cells, diff_keys):
            value = diffs.get(key)
            if value is None:
                val_label.setText("—")
                variant = ""
            else:
                val_label.setText(f"{value:.1f}")
                if value <= self.diff_green_max:
                    variant = "green"
                elif value <= self.pair_alert_threshold:
                    variant = "yellow"
                else:
                    variant = "red"
            cell.setProperty("variant", variant)
            cell.style().unpolish(cell)
            cell.style().polish(cell)
        variants = [
            str(cell.property("variant") or "")
            for _, _, cell in self.diff_cells
        ]
        active_variants = [variant for variant in variants if variant]
        strip_variant = (
            active_variants[0]
            if active_variants and len(set(active_variants)) == 1
            else ""
        )
        self.diff_strip.setProperty("variant", strip_variant)
        self.diff_strip.style().unpolish(self.diff_strip)
        self.diff_strip.style().polish(self.diff_strip)

    def _set_confirm_button_state(self, checked: bool, visible: bool = True) -> None:
        self.confirm_button.setVisible(visible)
        self.confirm_button.blockSignals(True)
        self.confirm_button.setChecked(checked)
        self.confirm_button.blockSignals(False)
        self.confirm_button.setText("✓  Confirmed" if checked else "Confirm pair")

    def _on_confirm_toggled(self, checked: bool) -> None:
        item = self.queue_list.currentItem()
        if item is None:
            return
        patient_id = item.data(Qt.ItemDataRole.UserRole)
        if checked:
            self.confirmed_patients.add(patient_id)
        else:
            self.confirmed_patients.discard(patient_id)
        # Update button label, queue row, and counts.
        self.confirm_button.setText("✓  Confirmed" if checked else "Confirm pair")
        item.setText(self._queue_item_text(patient_id, self._reason_for(patient_id)))
        item.setForeground(
            QColor("#94a3b8") if checked else QColor("#0f172a")
        )
        self._refresh_queue_count()
        self._refresh_unconfirmed_banner()
        self._refresh_tab_titles()
        self._refresh_overview_tab()

    def _reason_for(self, patient_id: str) -> str:
        if self.bundle is None:
            return REVIEW_REASON_MULTI_ENTRY
        for item in self.bundle.review_items:
            if item.patient_id == patient_id:
                return item.reason
        return REVIEW_REASON_MULTI_ENTRY

    def _refresh_queue_count(self) -> None:
        if self.bundle is None:
            return
        total = len(self.bundle.review_items)
        left = total - len(self.confirmed_patients)
        if left == 0:
            self.queue_count_pill.setText("all confirmed")
            self.queue_count_pill.setProperty("variant", "ok")
        else:
            self.queue_count_pill.setText(f"{left} of {total} left")
            self.queue_count_pill.setProperty("variant", "warn")
        self.queue_count_pill.style().unpolish(self.queue_count_pill)
        self.queue_count_pill.style().polish(self.queue_count_pill)

    def _refresh_unconfirmed_banner(self) -> None:
        if self.bundle is None or not self.bundle.review_items:
            self.unconfirmed_banner.setProperty("variant", "neutral")
            self.unconfirmed_banner._body_label.setText(  # type: ignore[attr-defined]
                "No patients require review."
            )
            self.unconfirmed_banner.style().unpolish(self.unconfirmed_banner)
            self.unconfirmed_banner.style().polish(self.unconfirmed_banner)
            return
        total = len(self.bundle.review_items)
        left = total - len(self.confirmed_patients)
        if left == 0:
            self.unconfirmed_banner.setProperty("variant", "success")
            self.unconfirmed_banner._body_label.setText(  # type: ignore[attr-defined]
                "<b>All patients confirmed.</b> You're good to export the workbook."
            )
        else:
            self.unconfirmed_banner.setProperty("variant", "attention")
            self.unconfirmed_banner._body_label.setText(  # type: ignore[attr-defined]
                f"<b>{left} of {total} patients aren't confirmed yet.</b> "
                "You can still export — unconfirmed patients use the auto-paired choice — "
                "but it's safer to confirm each one first."
            )
        self.unconfirmed_banner.style().unpolish(self.unconfirmed_banner)
        self.unconfirmed_banner.style().polish(self.unconfirmed_banner)

    def _refresh_tab_titles(self) -> None:
        if self.bundle is None:
            self.tabs.setTabText(1, "Review")
            self.tabs.setTabText(2, "All data")
            self.tabs.setTabText(3, "Averaged")
            return
        review_count = len(self.bundle.review_items)
        all_data_count = len(self.bundle.dataframe)
        averaged_count = len(self.bundle.analyzed_df)
        self.tabs.setTabText(1, f"Review  ({review_count})" if review_count else "Review")
        self.tabs.setTabText(2, f"All data  ({all_data_count})")
        self.tabs.setTabText(3, f"Averaged  ({averaged_count})")

    def _select_queue_patient(self, patient_id: str) -> None:
        for index in range(self.queue_list.count()):
            item = self.queue_list.item(index)
            if item and item.data(Qt.ItemDataRole.UserRole) == patient_id:
                self.queue_list.setCurrentRow(index)
                return

    def reset_current_to_auto(self) -> None:
        item = self.queue_list.currentItem()
        if item is None:
            return
        patient_id = item.data(Qt.ItemDataRole.UserRole)
        auto_pair = self.auto_pairs.get(patient_id)
        if auto_pair is None:
            return
        self.manual_pairs[patient_id] = list(auto_pair)
        self._rebuild_analysis()
        self._select_queue_patient(patient_id)

    def open_current_pair_pdf(self) -> None:
        if self.pair_table.rowCount() == 0:
            return
        row = self.pair_table.currentRow()
        if row < 0:
            row = 0
        source_item = self.pair_table.item(row, 6)
        if source_item is None:
            return
        filename = source_item.text()
        for path in self.pdf_paths:
            if path.name == filename:
                self._open_pdf(path)
                return

    # =================================================================
    #   All data / Averaged tabs
    # =================================================================
    def _refresh_all_data_tab(self) -> None:
        if self.bundle is None:
            self.all_data_table.setRowCount(0)
            self.all_data_table.setColumnCount(0)
            return
        df = display_dataframe(self.bundle).drop(
            columns=["Special Row", *EXTRA_COLUMNS, *UI_CONTEXT_COLUMNS],
            errors="ignore",
        )
        df = filter_columns_for_report_mode(df, self.report_mode)
        self._populate_data_table(self.all_data_table, df)

    def _refresh_averaged_tab(self) -> None:
        if self.bundle is None:
            self.averaged_table.setRowCount(0)
            self.averaged_table.setColumnCount(0)
            return
        df = self.bundle.analyzed_df.drop(
            columns=[*EXTRA_COLUMNS, *UI_CONTEXT_COLUMNS],
            errors="ignore",
        ).copy()
        df = filter_columns_for_report_mode(df, self.report_mode)
        self._populate_data_table(self.averaged_table, df)

    def _populate_data_table(self, table: QTableWidget, df: pd.DataFrame) -> None:
        columns = list(df.columns)
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)
        table.setRowCount(len(df))
        for row_index, (_, row) in enumerate(df.iterrows()):
            for col_index, col_name in enumerate(columns):
                value = row.get(col_name)
                text = format_value(value)
                item = QTableWidgetItem(text)
                if col_name == "Source File":
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
                    )
                else:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                table.setItem(row_index, col_index, item)
        table.resizeColumnsToContents()

    def _open_all_data_pdf(self, item: QTableWidgetItem) -> None:
        row = item.row()
        # First column is Source File for the All Data dataframe
        source_item = self.all_data_table.item(row, 0)
        if source_item is None:
            return
        filename = source_item.text()
        for path in self.pdf_paths:
            if path.name == filename:
                self._open_pdf(path)
                return

    # =================================================================
    #   Export
    # =================================================================
    def export_excel(self) -> None:
        if self.bundle is None or not self.records:
            QMessageBox.information(
                self, "Nothing to export", "Process some PDFs first."
            )
            return

        # Soft warning if some patients are unconfirmed
        if self.bundle.review_items:
            unconfirmed = [
                item.patient_id
                for item in self.bundle.review_items
                if item.patient_id not in self.confirmed_patients
            ]
            if unconfirmed:
                names = ", ".join(unconfirmed[:5])
                if len(unconfirmed) > 5:
                    names += f", and {len(unconfirmed) - 5} more"
                response = QMessageBox.warning(
                    self,
                    "Export with unconfirmed patients?",
                    f"<b>{len(unconfirmed)} patient{'s' if len(unconfirmed) != 1 else ''} "
                    f"haven't been confirmed yet.</b><br><br>"
                    f"Unconfirmed patients ({names}) will export using the auto-paired choice. "
                    "It's safer to confirm each one first.<br><br>"
                    "Export anyway?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
                    QMessageBox.StandardButton.Cancel,
                )
                if response != QMessageBox.StandardButton.Yes:
                    return

        try:
            manual_tuples = {
                pid: tuple(rows[:2])  # type: ignore[arg-type]
                for pid, rows in self.manual_pairs.items()
                if len(rows) == 2
            }
            count = save_to_excel(
                self.records,
                self.output_path,
                manual_pairs=manual_tuples,
                mode=ANALYSIS_MODE,
                pair_alert_threshold=self.pair_alert_threshold,
                report_mode=self.report_mode,
            )
        except Exception:
            QMessageBox.critical(
                self, "Export failed", traceback.format_exc()
            )
            return

        self.last_export_path = self.output_path
        self.open_file_button.setVisible(True)
        self.open_folder_button.setVisible(True)
        self._refresh_overview_tab()
        self._update_status(f"Exported {count} rows to {self.output_path.name}")
        self.tabs.setCurrentIndex(0)

    def open_export_file(self) -> None:
        if self.last_export_path and self.last_export_path.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.last_export_path)))

    def open_export_folder(self) -> None:
        if self.last_export_path and self.last_export_path.parent.exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.last_export_path.parent)))

    # =================================================================
    #   PDF viewer / About
    # =================================================================
    def _open_pdf(self, pdf_path: Path) -> None:
        for viewer in list(self.pdf_viewers):
            if viewer.pdf_path == pdf_path:
                viewer.show()
                viewer.raise_()
                viewer.activateWindow()
                return
        try:
            viewer = PdfViewerDialog(pdf_path, self)
        except Exception:
            QMessageBox.critical(
                self, "Could not open PDF", traceback.format_exc()
            )
            return
        self.pdf_viewers.append(viewer)
        viewer.finished.connect(
            lambda _: self.pdf_viewers.remove(viewer)
            if viewer in self.pdf_viewers
            else None
        )
        viewer.show()

    def show_about_dialog(self) -> None:
        if self.readme_dialog is None:
            self.readme_dialog = ReadmeDialog(self)
        self.readme_dialog.show()
        self.readme_dialog.raise_()
        self.readme_dialog.activateWindow()

    def _update_status(self, text: str) -> None:
        self.status_label.setText(text)

    # =================================================================
    #   Stylesheet
    # =================================================================
    def _apply_styles(self) -> None:
        self.setStyleSheet(STYLESHEET)


# =====================================================================
#   Stylesheet (clean clinical, single accent)
# =====================================================================
STYLESHEET = """
* {
    font-family: "Segoe UI Variable", "Segoe UI", "Inter", system-ui;
    font-size: 13px;
    color: #0f172a;
}

QMainWindow, QWidget#root {
    background: #ffffff;
}

QFrame#titlebar {
    background: #ffffff;
    border-bottom: 1px solid #e3e8ee;
}
QLabel#brandTitle { font-weight: 600; font-size: 14px; color: #0f172a; }
QLabel#brandVersion { color: #94a3b8; font-size: 12px; }
QToolButton#titlebarIcon {
    border: 1px solid transparent;
    border-radius: 6px;
    background: transparent;
    color: #64748b;
    font-weight: 600;
    font-size: 14px;
}
QToolButton#titlebarIcon:hover {
    background: #f7f9fb;
    border-color: #e3e8ee;
    color: #0f172a;
}

QToolButton#helpButton {
    background: #f1f4f7;
    color: #64748b;
    border: 1px solid #cfd6de;
    border-radius: 9px;
    font-size: 10px;
    font-weight: 700;
    padding: 0;
}
QToolButton#helpButton:hover {
    background: #ecfdf5;
    color: #0f766e;
    border-color: #0f766e;
}
QToolTip {
    background: #0f172a;
    color: #ffffff;
    border: 1px solid #0f172a;
    padding: 6px 10px;
    font-size: 12px;
}

QFrame#statusBar {
    background: #f7f9fb;
    border-top: 1px solid #e3e8ee;
}
QLabel#statusDot {
    background: #15803d;
    border-radius: 3px;
}
QLabel#statusText { color: #64748b; font-size: 12px; }
QLabel#statusVersion { color: #94a3b8; font-size: 11px; }

/* ---------- Wizard header ---------- */
QFrame#wizardHead {
    background: #ffffff;
    border-bottom: 1px solid #e3e8ee;
}
QLabel#wizardTitle { font-size: 18px; font-weight: 600; }
QLabel#wizardSub { color: #64748b; font-size: 12px; }

QFrame#stepperFrame { background: transparent; }
QFrame#stepPill {
    background: #f7f9fb;
    border: 1px solid #e3e8ee;
    border-radius: 999px;
}
QFrame#stepPill[state="active"] {
    background: #ecfdf5;
    border-color: #99f6e4;
}
QLabel#stepNum {
    background: #f1f4f7;
    color: #94a3b8;
    border-radius: 9px;
    font-size: 11px;
    font-weight: 700;
    qproperty-alignment: AlignCenter;
}
QLabel#stepNum[state="active"] {
    background: #0f766e;
    color: white;
}
QLabel#stepNum[state="done"] {
    background: #15803d;
    color: white;
}
QLabel#stepText { color: #64748b; font-size: 12px; }
QLabel#stepText[state="active"] { color: #115e59; font-weight: 600; }
QLabel#stepText[state="done"] { color: #15803d; }
QLabel#stepArrow { color: #94a3b8; font-size: 14px; }

/* ---------- Run summary header ---------- */
QFrame#runSummary {
    background: #ffffff;
    border-bottom: 1px solid #e3e8ee;
}

/* ---------- Import body / cards ---------- */
QFrame#importScreen, QFrame#reviewScreen, QFrame#importBody, QWidget#importWrap {
    background: #ffffff;
}
QScrollArea { background: #ffffff; border: none; }
QScrollArea > QWidget > QWidget { background: #ffffff; }
QWidget#reviewDetailContent { background: #ffffff; }
QFrame#detailHeader { background: #ffffff; }

QFrame#importCard {
    background: #ffffff;
    border: 1px solid #e3e8ee;
    border-radius: 8px;
}
QLabel#cardEyebrow {
    color: #94a3b8;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 1px;
}
QLabel#cardSub {
    color: #64748b;
    font-size: 12px;
}
QLabel#fieldLabel {
    color: #0f172a;
    font-size: 12px;
    font-weight: 500;
}
QLabel#fieldHelp {
    color: #64748b;
    font-size: 11px;
}

/* Drop zone */
QFrame#dropZone {
    border: 2px dashed #cfd6de;
    border-radius: 8px;
    background: #f7f9fb;
}
QFrame#dropZone:hover {
    border-color: #0f766e;
    background: #ecfdf5;
}
QFrame#dropZone[dragActive="true"] {
    border-color: #0f766e;
    background: #ecfdf5;
}
QLabel#dropZoneIcon { color: #0f766e; font-size: 22px; font-weight: 700; }
QLabel#dropZoneTitle { color: #0f172a; font-weight: 600; font-size: 14px; }
QLabel#dropZoneSubtitle { color: #64748b; font-size: 12px; }

/* File list */
QListWidget#fileList {
    background: #f7f9fb;
    border: 1px solid #e3e8ee;
    border-radius: 6px;
    padding: 4px 0;
}
QListWidget#fileList::item {
    padding: 7px 12px;
    color: #0f172a;
    border-bottom: 1px solid #e3e8ee;
}
QListWidget#fileList::item:last { border-bottom: none; }
QListWidget#fileList::item:selected {
    background: #ecfdf5;
    color: #115e59;
}
QListWidget#fileList::item:hover { background: #f1f4f7; }

QLabel#fileMetaText { color: #64748b; font-size: 12px; }

/* Inputs */
QLineEdit, QComboBox, QDoubleSpinBox, QSpinBox {
    background: #ffffff;
    border: 1px solid #e3e8ee;
    border-radius: 6px;
    padding: 7px 10px;
    selection-background-color: #ecfdf5;
    selection-color: #115e59;
}
QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus, QSpinBox:focus {
    border-color: #0f766e;
}
QDoubleSpinBox#thresholdValueSpin {
    padding: 4px 10px;
}
QLineEdit#pathInput {
    font-family: "Cascadia Mono", "Consolas", monospace;
    font-size: 12px;
}
QComboBox::drop-down {
    border: none;
    width: 20px;
}
QComboBox::down-arrow {
    width: 10px;
    height: 10px;
}

QWidget#thresholdStepper {
    background: transparent;
}
QPushButton#stepperButton {
    background: #ffffff;
    border: 1px solid #e3e8ee;
    border-radius: 6px;
    color: #0f172a;
    font-size: 14px;
    font-weight: 700;
    padding: 0;
}
QPushButton#stepperButton:hover {
    background: #ecfdf5;
    border-color: #0f766e;
    color: #0f766e;
}
QPushButton#stepperButton:pressed {
    background: #d9f4eb;
}

/* Segmented control */
QFrame#segmentedControl {
    background: #f1f4f7;
    border: 1px solid #e3e8ee;
    border-radius: 8px;
}
QPushButton#segmentedButton {
    background: transparent;
    border: none;
    border-radius: 6px;
    padding: 7px 12px;
    color: #64748b;
    font-weight: 500;
}
QPushButton#segmentedButton:hover { color: #0f172a; }
QPushButton#segmentedButton:checked {
    background: #ffffff;
    color: #0f766e;
    font-weight: 600;
}

/* Buttons */
QPushButton {
    background: #ffffff;
    border: 1px solid #e3e8ee;
    border-radius: 6px;
    padding: 7px 14px;
    color: #0f172a;
    font-weight: 500;
}
QPushButton:hover { background: #f7f9fb; }
QPushButton:disabled { color: #94a3b8; background: #f1f4f7; border-color: #e3e8ee; }

QPushButton#primaryButton {
    background: #0f766e;
    border-color: #0f766e;
    color: white;
    font-weight: 600;
    padding: 8px 16px;
}
QPushButton#primaryButton:hover {
    background: #115e59;
    border-color: #115e59;
}
QPushButton#primaryButton:disabled {
    background: #94a3b8;
    border-color: #94a3b8;
}
QPushButton#exportButton {
    background: #0f766e;
    border: 1px solid #0f766e;
    border-radius: 6px;
    color: white;
    font-weight: 600;
    padding: 4px 16px;
}
QPushButton#exportButton:hover {
    background: #115e59;
    border-color: #115e59;
}
QPushButton#exportButton:disabled {
    background: #94a3b8;
    border-color: #94a3b8;
}

QPushButton#linkButton {
    background: transparent;
    border: none;
    color: #0f766e;
    font-weight: 500;
    padding: 4px 8px;
}
QPushButton#linkButton:hover {
    background: #ecfdf5;
    border-radius: 4px;
}
QPushButton#linkButton[copied="true"] {
    background: #0f766e;
    color: #ffffff;
    border-radius: 4px;
}
QPushButton#linkButton:disabled {
    color: #94a3b8;
    background: transparent;
}

QPushButton#disclosureToggle {
    background: transparent;
    border: none;
    color: #64748b;
    font-size: 12px;
    font-weight: 500;
    padding: 4px 0;
    text-align: left;
}
QPushButton#disclosureToggle:hover { color: #0f172a; }

QFrame#hairline {
    background: #e3e8ee;
    max-height: 1px;
}
QFrame#advancedPanel { background: transparent; }
QLabel#thresholdLabel { color: #64748b; font-size: 11px; }

/* Sticky import action bar */
QFrame#importActions {
    background: #ffffff;
    border-top: 1px solid #e3e8ee;
}
QLabel#importMeta { color: #64748b; font-size: 12px; }

QProgressBar#progressBar {
    border: 1px solid #e3e8ee;
    border-radius: 4px;
    background: #f1f4f7;
    text-align: center;
    color: #64748b;
    font-size: 11px;
}
QProgressBar#progressBar::chunk {
    background: #0f766e;
    border-radius: 3px;
}

/* ---------- Tabs ---------- */
QTabWidget::pane {
    border: none;
    background: #ffffff;
}
QTabBar { background: #ffffff; border-bottom: 1px solid #e3e8ee; }
QTabBar::tab {
    background: transparent;
    color: #64748b;
    padding: 10px 16px;
    border-bottom: 2px solid transparent;
    margin-right: 2px;
}
QTabBar::tab:hover { color: #0f172a; }
QTabBar::tab:selected {
    color: #0f766e;
    border-bottom-color: #0f766e;
    font-weight: 600;
}

/* ---------- Stat cards ---------- */
QFrame#statCard {
    background: #ffffff;
    border: 1px solid #e3e8ee;
    border-radius: 8px;
}
QLabel#statValue {
    font-size: 22px;
    font-weight: 600;
    color: #0f172a;
}
QLabel#statValue[variant="accent"] { color: #0f766e; }
QLabel#statValue[variant="warn"] { color: #b45309; }
QLabel#statValue[variant="danger"] { color: #b91c1c; }
QLabel#statLabel { color: #64748b; font-size: 12px; }

/* ---------- Banners ---------- */
QFrame#banner {
    background: #f7f9fb;
    border: 1px solid #e3e8ee;
    border-radius: 8px;
}
QFrame#banner[variant="attention"] {
    background: #fffbeb;
    border-color: #fcd34d;
}
QFrame#banner[variant="success"] {
    background: #ecfdf5;
    border-color: #99f6e4;
}
QFrame#banner[variant="danger"] {
    background: #fef2f2;
    border-color: #fca5a5;
}
QFrame#banner[variant="neutral"] { background: #f7f9fb; }
QLabel#bannerIcon { color: #64748b; font-size: 14px; }
QLabel#bannerBody { color: #0f172a; font-size: 13px; }
QFrame#banner[variant="attention"] QLabel#bannerBody { color: #78350f; }
QFrame#banner[variant="success"] QLabel#bannerBody { color: #115e59; }
QFrame#banner[variant="danger"] QLabel#bannerBody { color: #7f1d1d; }

/* ---------- Section titles ---------- */
QLabel#sectionTitle {
    font-size: 15px;
    font-weight: 600;
    color: #0f172a;
}
QLabel#sectionSub { color: #64748b; font-size: 12px; }
QLabel#subSectionTitle {
    font-size: 13px;
    font-weight: 600;
    color: #0f172a;
}

/* ---------- Tables ---------- */
QTableWidget#dataTable {
    background: #ffffff;
    border: 1px solid #e3e8ee;
    border-radius: 8px;
    gridline-color: #f1f4f7;
}
QTableWidget#dataTable QHeaderView::section {
    background: #f7f9fb;
    color: #64748b;
    border: none;
    border-bottom: 1px solid #e3e8ee;
    border-right: 1px solid #f1f4f7;
    padding: 8px 12px;
    font-size: 11px;
    font-weight: 600;
}
QTableWidget#dataTable::item {
    padding: 8px 12px;
    border-bottom: 1px solid #f1f4f7;
    color: #0f172a;
}
QTableWidget#dataTable::item:selected {
    background: #ecfdf5;
    color: #115e59;
    border: none;
}
QTableWidget#dataTable::item:focus {
    border: none;
    outline: none;
}

/* ---------- Review queue ---------- */
QFrame#queueCard {
    background: #ffffff;
    border: 1px solid #e3e8ee;
    border-radius: 8px;
}
QFrame#queueHead {
    background: #f7f9fb;
    border-bottom: 1px solid #e3e8ee;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
}
QLabel#queueHeadTitle { color: #0f172a; font-size: 12px; font-weight: 600; }
QLabel#pillWarn {
    background: #fffbeb;
    color: #b45309;
    border: 1px solid #fcd34d;
    border-radius: 9px;
    padding: 1px 8px;
    font-size: 11px;
    font-weight: 500;
}
QLabel#pillWarn[variant="ok"] {
    background: #ecfdf5;
    color: #15803d;
    border-color: #bbf7d0;
}

QListWidget#queueList {
    background: #ffffff;
    border: none;
}
QListWidget#queueList::item {
    padding: 10px 12px;
    border-bottom: 1px solid #e3e8ee;
    border-left: 3px solid transparent;
}
QListWidget#queueList::item:hover { background: #f3f7f9; }
QListWidget#queueList::item:selected {
    background: #ecfdf5;
    border-left-color: #0f766e;
    color: #115e59;
}

/* ---------- Review detail ---------- */
QLabel#patientTitle { font-size: 18px; font-weight: 600; color: #0f172a; }
QLabel#patientSubtitle { color: #64748b; font-size: 12px; }

/* Reason callout */
QFrame#reasonCallout {
    border: 1px solid #fcd34d;
    background: #fffbeb;
    border-radius: 8px;
}
QFrame#reasonCallout[variant="danger"] {
    border-color: #fca5a5;
    background: #fef2f2;
}
QLabel#reasonIcon { font-size: 14px; color: #b45309; }
QLabel#reasonText { font-weight: 600; color: #78350f; font-size: 13px; }
QFrame#reasonCallout[variant="danger"] QLabel#reasonIcon { color: #b91c1c; }
QFrame#reasonCallout[variant="danger"] QLabel#reasonText { color: #7f1d1d; }

/* Diff strip */
QFrame#diffStrip {
    background: #e3e8ee;
    border: 1px solid #e3e8ee;
    border-radius: 8px;
}
QFrame#diffStrip[variant="green"] {
    background: #ecfdf5;
    border-color: #bbf7d0;
}
QFrame#diffStrip[variant="yellow"] {
    background: #fffbeb;
    border-color: #fde68a;
}
QFrame#diffStrip[variant="red"] {
    background: #fef2f2;
    border-color: #fecaca;
}
QFrame#diffCell {
    background: #ffffff;
}
QFrame#diffCell[variant="green"] {
    background: #ecfdf5;
    border: 1px solid #bbf7d0;
}
QFrame#diffCell[variant="yellow"] {
    background: #fffbeb;
    border: 1px solid #fde68a;
}
QFrame#diffCell[variant="red"] {
    background: #fef2f2;
    border: 1px solid #fecaca;
}
QLabel#diffLabel {
    color: #94a3b8;
    font-size: 11px;
    font-weight: 500;
    letter-spacing: 0.5px;
}
QLabel#diffValue {
    font-size: 18px;
    font-weight: 600;
    color: #0f172a;
}
QFrame#diffCell[variant="green"] QLabel#diffValue { color: #15803d; }
QFrame#diffCell[variant="yellow"] QLabel#diffValue { color: #b45309; }
QFrame#diffCell[variant="red"] QLabel#diffValue { color: #b91c1c; }
QFrame#diffCell[variant="green"] QLabel#diffLabel { color: #166534; }
QFrame#diffCell[variant="yellow"] QLabel#diffLabel { color: #92400e; }
QFrame#diffCell[variant="red"] QLabel#diffLabel { color: #991b1b; }

/* Pair card */
QFrame#pairCard {
    background: #ffffff;
    border: 1px solid #e3e8ee;
    border-radius: 8px;
}
QFrame#pairCardHead {
    background: #f7f9fb;
    border-bottom: 1px solid #e3e8ee;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
}
QLabel#pairCardTitle { font-weight: 600; color: #0f172a; font-size: 13px; }
QLabel#pairCardHelpText { color: #64748b; font-size: 12px; }

/* Confirm pair button (sits next to reason callout) */
QPushButton#confirmPairButton {
    background: #ffffff;
    border: 1px solid #cfd6de;
    border-radius: 6px;
    padding: 7px 14px;
    color: #0f172a;
    font-weight: 500;
}
QPushButton#confirmPairButton:hover {
    border-color: #0f766e;
    color: #0f766e;
}
QPushButton#confirmPairButton:checked {
    background: #ecfdf5;
    border-color: #0f766e;
    color: #115e59;
    font-weight: 600;
}
QPushButton#confirmPairButton:checked:hover {
    background: #d1fae5;
}

QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border: 1.5px solid #cfd6de;
    border-radius: 4px;
    background: #ffffff;
}
QCheckBox::indicator:hover { border-color: #0f766e; }
QCheckBox::indicator:checked {
    background: #ffffff;
    border-color: #0f766e;
}
QPushButton#keepToggle {
    background: #ffffff;
    border: 1.5px solid #cfd6de;
    border-radius: 4px;
    color: #ffffff;
    font-size: 13px;
    font-weight: 700;
    padding: 0;
    margin: 0;
}
QPushButton#keepToggle:hover {
    border-color: #0f766e;
}
QPushButton#keepToggle:checked {
    background: #0f766e;
    border-color: #0f766e;
    color: #ffffff;
}

/* Dialog */
QDialog { background: #ffffff; }
QLabel#dialogTitle { font-size: 20px; font-weight: 600; color: #0f172a; }
QLabel#dialogSubtitle { color: #64748b; font-size: 12px; }
QTextBrowser {
    background: #ffffff;
    border: 1px solid #e3e8ee;
    border-radius: 6px;
    padding: 8px;
}

/* Scrollbars (thin, rounded, no arrow buttons) */
QScrollBar:vertical {
    background: transparent;
    width: 12px;
    margin: 4px 2px 4px 2px;
    border: none;
}
QScrollBar::handle:vertical {
    background: #cbd5e1;
    border-radius: 4px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover { background: #94a3b8; }
QScrollBar::handle:vertical:pressed { background: #64748b; }
QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical {
    background: transparent;
    border: none;
    height: 0;
    width: 0;
}
QScrollBar::add-page:vertical,
QScrollBar::sub-page:vertical { background: transparent; }
QScrollBar:horizontal {
    background: transparent;
    height: 12px;
    margin: 2px 4px 2px 4px;
    border: none;
}
QScrollBar::handle:horizontal {
    background: #cbd5e1;
    border-radius: 4px;
    min-width: 30px;
}
QScrollBar::handle:horizontal:hover { background: #94a3b8; }
QScrollBar::handle:horizontal:pressed { background: #64748b; }
QScrollBar::add-line:horizontal,
QScrollBar::sub-line:horizontal {
    background: transparent;
    border: none;
    height: 0;
    width: 0;
}
QScrollBar::add-page:horizontal,
QScrollBar::sub-page:horizontal { background: transparent; }
"""


# =====================================================================
def main() -> int:
    app = QApplication(sys.argv)
    if APP_ICON_PATH.exists():
        app.setWindowIcon(QIcon(str(APP_ICON_PATH)))
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
