from __future__ import annotations

import sys
import traceback
from pathlib import Path
from typing import Optional

import pandas as pd
from PySide6.QtCore import QEvent, QObject, QPointF, Qt, QThread, QUrl, Signal
from PySide6.QtGui import QColor, QDesktopServices, QIcon
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCheckBox,
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
    QPushButton,
    QProgressBar,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextBrowser,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

try:
    from .backend import (
        ANALYSIS_MODE,
        APP_ICON_PATH,
        APP_SUBTITLE,
        APP_TITLE,
        CONTACT_EMAIL,
        COLUMNS,
        EXTRA_COLUMNS,
        REPOSITORY_URL,
        AnalysisBundle,
        build_analysis,
        calculate_pair_differences,
        default_output_path,
        display_dataframe,
        format_value,
        initial_manual_pairs,
        load_readme_text,
        pair_alert_triggered,
        patient_rows,
        patient_entry_counts,
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
        APP_SUBTITLE,
        APP_TITLE,
        CONTACT_EMAIL,
        COLUMNS,
        EXTRA_COLUMNS,
        REPOSITORY_URL,
        AnalysisBundle,
        build_analysis,
        calculate_pair_differences,
        default_output_path,
        display_dataframe,
        format_value,
        initial_manual_pairs,
        load_readme_text,
        pair_alert_triggered,
        patient_rows,
        patient_entry_counts,
        process_pdf,
        record_status,
        save_to_excel,
    )


class ProcessingWorker(QObject):
    progress = Signal(int, int, str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, pdf_paths: list[Path]):
        super().__init__()
        self.pdf_paths = pdf_paths

    def run(self) -> None:
        records: list[dict[str, object]] = []
        total_files = len(self.pdf_paths)

        try:
            for index, pdf_path in enumerate(self.pdf_paths, start=1):
                self.progress.emit(index - 1, total_files, f"Reading {pdf_path.name}")
                records.append(process_pdf(pdf_path))
                self.progress.emit(index, total_files, f"Processed {pdf_path.name}")
        except Exception:
            self.failed.emit(traceback.format_exc())
            return

        self.finished.emit(records)


class ReadmeDialog(QDialog):
    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("About PWA Data Extractor")
        self.resize(860, 620)
        if APP_ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(APP_ICON_PATH)))

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        title = QLabel(APP_TITLE)
        title.setObjectName("dialogTitle")
        subtitle = QLabel(
            "Instructions, export notes, repository link, and support contact."
        )
        subtitle.setObjectName("dialogSubtitle")
        subtitle.setWordWrap(True)

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

        self.setWindowTitle(f"PDF Viewer - {pdf_path.name}")
        self.resize(980, 760)
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
        self.page_count_label = QLabel("of 0")
        self.zoom_out_button = QPushButton("Zoom -")
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
        navigator.backAvailableChanged.connect(self._sync_nav_buttons)
        navigator.forwardAvailableChanged.connect(self._sync_nav_buttons)

        error = self.document.load(str(pdf_path))
        if error != QPdfDocument.Error.None_:
            raise RuntimeError(f"Could not load PDF: {pdf_path.name}")

        self.pdf_view.setDocument(self.document)
        self.page_spin.setMaximum(max(1, self.document.pageCount()))
        self.page_count_label.setText(f"of {self.document.pageCount()}")
        self.sync_page_controls(0)

    def sync_page_controls(self, current_page: int) -> None:
        self.page_spin.blockSignals(True)
        self.page_spin.setValue(current_page + 1)
        self.page_spin.blockSignals(False)
        self._sync_nav_buttons()

    def _sync_nav_buttons(self) -> None:
        current_page = self.pdf_view.pageNavigator().currentPage()
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


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.pdf_paths: list[Path] = []
        self.output_path: Path = default_output_path()
        self.records: list[dict[str, object]] = []
        self.bundle: Optional[AnalysisBundle] = None
        self.auto_pairs: dict[str, tuple[int, int]] = {}
        self.manual_pairs: dict[str, list[int]] = {}
        self.pdf_viewers: list[PdfViewerDialog] = []
        self.thread: Optional[QThread] = None
        self.worker: Optional[ProcessingWorker] = None
        self.last_export_path: Optional[Path] = None
        self.updating_pair_table = False
        self.updating_settings = False
        self.readme_dialog: Optional[ReadmeDialog] = None
        self.pair_alert_threshold = 5.0
        self.diff_green_max = 3.0
        self.diff_yellow_max = 6.0
        self.setup_scroll_area: Optional[QScrollArea] = None
        self.review_split: Optional[QSplitter] = None
        self.review_patient_panel: Optional[QFrame] = None
        self.review_detail_panel: Optional[QFrame] = None
        self.review_detail_scroll: Optional[QScrollArea] = None
        self.review_split_compact: Optional[bool] = None

        self.setWindowTitle(APP_TITLE)
        self.resize(1400, 900)
        if APP_ICON_PATH.exists():
            self.setWindowIcon(QIcon(str(APP_ICON_PATH)))

        self._build_ui()
        self._apply_styles()
        self._sync_controls()
        self._refresh_results_views()

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(26, 22, 26, 22)
        root_layout.setSpacing(18)

        root_layout.addWidget(self._build_header())

        workflow = QSplitter(Qt.Orientation.Horizontal)
        workflow.setChildrenCollapsible(False)
        workflow.addWidget(self._build_setup_panel())
        workflow.addWidget(self._build_results_panel())
        workflow.setStretchFactor(0, 0)
        workflow.setStretchFactor(1, 1)
        workflow.setSizes([360, 1200])
        root_layout.addWidget(workflow, 1)

        self.setCentralWidget(root)

    def _build_header(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)

        title_row = QHBoxLayout()
        title = QLabel(APP_TITLE)
        title.setObjectName("title")
        subtitle = QLabel(APP_SUBTITLE)
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)

        self.info_button = QToolButton()
        self.info_button.setObjectName("infoButton")
        self.info_button.setText("?")
        self.info_button.setToolTip("Open instructions and repository details")
        self.info_button.clicked.connect(self.show_info_dialog)

        title_row.addWidget(title)
        title_row.addWidget(self.info_button, 0, Qt.AlignmentFlag.AlignVCenter)
        title_row.addStretch(1)

        layout.addLayout(title_row)
        layout.addWidget(subtitle)
        return container

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._update_review_split_layout()

    def _build_setup_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("panel")
        panel.setMinimumWidth(360)
        panel.setMaximumWidth(460)
        outer_layout = QVBoxLayout(panel)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setup_scroll_area = scroll
        outer_layout.addWidget(scroll)

        content = QWidget()
        scroll.setWidget(content)

        layout = QVBoxLayout(content)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        layout.addWidget(
            self._section_title(
                "1. Input reports",
                "Choose one or more PWA PDF files for processing.",
            )
        )

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.file_list.setMinimumHeight(110)
        self.file_list.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.file_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.file_list.viewport().installEventFilter(self)
        layout.addWidget(self.file_list, 1)

        file_buttons = QHBoxLayout()
        file_buttons.setSpacing(8)
        self.add_files_button = QPushButton("Add PDFs")
        self.remove_files_button = QPushButton("Remove selected")
        self.clear_files_button = QPushButton("Clear")
        file_buttons.addWidget(self.add_files_button)
        file_buttons.addWidget(self.remove_files_button)
        file_buttons.addWidget(self.clear_files_button)
        layout.addLayout(file_buttons)

        layout.addWidget(
            self._section_title(
                "2. Export workbook",
                "Choose where the Excel output should be saved.",
            )
        )

        output_row = QHBoxLayout()
        output_row.setSpacing(8)
        self.output_line = QLineEdit(str(self.output_path))
        self.output_line.setPlaceholderText("Choose an .xlsx output file")
        self.output_line.setCursorPosition(0)
        self.browse_output_button = QPushButton("Browse")
        output_row.addWidget(self.output_line, 1)
        output_row.addWidget(self.browse_output_button)
        layout.addLayout(output_row)

        settings_header = QWidget()
        settings_header_layout = QHBoxLayout(settings_header)
        settings_header_layout.setContentsMargins(0, 0, 0, 0)
        settings_header_layout.setSpacing(6)
        settings_header_layout.addWidget(
            self._section_title(
                "3. Settings",
                "Adjust the default thresholds used for review highlighting and pair alerts.",
            ),
            1,
        )
        self.thresholds_help_button = QToolButton()
        self.thresholds_help_button.setObjectName("infoButton")
        self.thresholds_help_button.setText("?")
        self.thresholds_help_button.setToolTip("Explain threshold settings")
        settings_header_layout.addWidget(
            self.thresholds_help_button,
            0,
            Qt.AlignmentFlag.AlignTop,
        )
        layout.addWidget(settings_header)

        settings_grid = QGridLayout()
        settings_grid.setHorizontalSpacing(8)
        settings_grid.setVerticalSpacing(6)
        settings_grid.setColumnStretch(0, 1)
        settings_grid.setColumnStretch(1, 1)

        self.green_max_spin = QDoubleSpinBox()
        self.green_max_spin.setRange(0, 100)
        self.green_max_spin.setDecimals(1)
        self.green_max_spin.setValue(self.diff_green_max)
        self.green_max_spin.setSuffix(" mmHg")
        self.green_max_spin.setMinimumWidth(160)
        self.green_max_spin.setMinimumHeight(36)
        self.green_max_spin.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.green_max_spin.installEventFilter(self)

        self.yellow_max_spin = QDoubleSpinBox()
        self.yellow_max_spin.setRange(0, 100)
        self.yellow_max_spin.setDecimals(1)
        self.yellow_max_spin.setValue(self.diff_yellow_max)
        self.yellow_max_spin.setSuffix(" mmHg")
        self.yellow_max_spin.setMinimumWidth(160)
        self.yellow_max_spin.setMinimumHeight(36)
        self.yellow_max_spin.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.yellow_max_spin.installEventFilter(self)

        self.pair_alert_spin = QDoubleSpinBox()
        self.pair_alert_spin.setRange(0, 100)
        self.pair_alert_spin.setDecimals(1)
        self.pair_alert_spin.setValue(self.pair_alert_threshold)
        self.pair_alert_spin.setSuffix(" mmHg")
        self.pair_alert_spin.setMinimumWidth(160)
        self.pair_alert_spin.setMinimumHeight(36)
        self.pair_alert_spin.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.pair_alert_spin.installEventFilter(self)

        green_label = QLabel("Green up to")
        green_label.setMinimumHeight(28)
        yellow_label = QLabel("Yellow up to")
        yellow_label.setMinimumHeight(28)
        alert_label = QLabel("Pair alert above")
        alert_label.setMinimumHeight(28)

        settings_grid.addWidget(green_label, 0, 0)
        settings_grid.addWidget(self.green_max_spin, 0, 1)
        settings_grid.addWidget(yellow_label, 1, 0)
        settings_grid.addWidget(self.yellow_max_spin, 1, 1)
        settings_grid.addWidget(alert_label, 2, 0)
        settings_grid.addWidget(self.pair_alert_spin, 2, 1)
        layout.addLayout(settings_grid)

        layout.addWidget(
            self._section_title(
                "4. Process",
                "Analyze the reports locally, then review multi-entry patients before exporting.",
            )
        )

        self.process_button = QPushButton("Process PDFs")
        self.process_button.setObjectName("primaryButton")
        layout.addWidget(self.process_button)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        self.add_files_button.clicked.connect(self.add_pdf_files)
        self.remove_files_button.clicked.connect(self.remove_selected_files)
        self.clear_files_button.clicked.connect(self.clear_pdf_files)
        self.browse_output_button.clicked.connect(self.browse_output_path)
        self.output_line.textChanged.connect(self._output_path_changed)
        self.thresholds_help_button.clicked.connect(self.show_thresholds_help_dialog)
        self.green_max_spin.valueChanged.connect(self._settings_changed)
        self.yellow_max_spin.valueChanged.connect(self._settings_changed)
        self.pair_alert_spin.valueChanged.connect(self._settings_changed)
        self.process_button.clicked.connect(self.process_files)

        return panel

    def _build_results_panel(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("panel")
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        header_card = QFrame()
        header_card.setObjectName("subpanel")
        header_layout = QHBoxLayout(header_card)
        header_layout.setContentsMargins(16, 14, 16, 14)
        header_layout.setSpacing(14)

        header_copy = QWidget()
        header_copy_layout = QVBoxLayout(header_copy)
        header_copy_layout.setContentsMargins(0, 0, 0, 0)
        header_copy_layout.setSpacing(4)

        header_title = QLabel("5. Review and export")
        header_title.setObjectName("sectionTitle")
        header_subtitle = QLabel(
            "Automatic pairing is applied by default. Adjust only the patients that need manual selection."
        )
        header_subtitle.setObjectName("statusLabel")
        header_subtitle.setWordWrap(True)
        header_copy_layout.addWidget(header_title)
        header_copy_layout.addWidget(header_subtitle)
        header_layout.addWidget(header_copy, 1)

        header_actions = QHBoxLayout()
        header_actions.setSpacing(10)
        self.export_button = QPushButton("Export Excel")
        self.export_button.setObjectName("primaryButton")
        self.open_output_button = QPushButton("Open folder")
        self.open_output_button.setEnabled(False)
        header_actions.addWidget(self.export_button)
        header_actions.addWidget(self.open_output_button)
        header_layout.addLayout(header_actions)

        layout.addWidget(header_card)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_overview_tab(), "Overview")
        self.tabs.addTab(self._build_review_tab(), "Multi-entry review")
        self.tabs.addTab(self._build_all_data_tab(), "All data")
        self.tabs.addTab(self._build_averaged_tab(), "Averaged data")
        layout.addWidget(self.tabs, 1)

        self.export_button.clicked.connect(self.export_excel)
        self.open_output_button.clicked.connect(self.open_export_folder)

        return panel

    def eventFilter(self, source: QObject, event: object) -> bool:
        if (
            event is not None
            and isinstance(event, QEvent)
            and event.type() == QEvent.Type.Wheel
        ):
            table_viewports = {
                self.overview_table.viewport(),
                self.pair_table.viewport(),
                self.diff_table.viewport(),
                self.all_data_table.viewport(),
                self.averaged_table.viewport(),
            }
            if (
                event.modifiers() & Qt.KeyboardModifier.ShiftModifier
                and source
                in table_viewports
            ):
                table = source.parent()
                if isinstance(table, QTableWidget):
                    scrollbar = table.horizontalScrollBar()
                    delta = event.angleDelta().y() or event.angleDelta().x()
                    steps = delta / 120
                    scroll_step = 1 if steps > 0 else -1 if steps < 0 else 0
                    scrollbar.setValue(scrollbar.value() - scroll_step)
                    return True
            elif source in table_viewports:
                if (
                    source in {self.pair_table.viewport(), self.diff_table.viewport()}
                    and self.review_detail_scroll is not None
                ):
                    scrollbar = self.review_detail_scroll.verticalScrollBar()
                    delta = event.angleDelta().y()
                    steps = delta / 120
                    base_step = max(1, scrollbar.singleStep())
                    scroll_amount = max(1, int(round(abs(steps) * base_step)))
                    if steps > 0:
                        scrollbar.setValue(scrollbar.value() - scroll_amount)
                    elif steps < 0:
                        scrollbar.setValue(scrollbar.value() + scroll_amount)
                    return True
                else:
                    table = source.parent()
                    if not isinstance(table, QTableWidget):
                        table = None
                    scrollbar = table.verticalScrollBar() if table is not None else None

                if scrollbar is not None:
                    delta = event.angleDelta().y()
                    steps = delta / 120
                    base_step = max(1, scrollbar.singleStep())
                    scroll_amount = max(1, int(round(abs(steps) * base_step * 3 / 25)))
                    if steps > 0:
                        scrollbar.setValue(scrollbar.value() - scroll_amount)
                    elif steps < 0:
                        scrollbar.setValue(scrollbar.value() + scroll_amount)
                    return True

        if (
            event is not None
            and isinstance(event, QEvent)
            and event.type() == QEvent.Type.Wheel
            and self.setup_scroll_area is not None
        ):
            focus_widget = QApplication.focusWidget()
            should_redirect = False

            if source is self.file_list.viewport():
                should_redirect = focus_widget is not self.file_list
            elif source in {
                self.green_max_spin,
                self.yellow_max_spin,
                self.pair_alert_spin,
            }:
                should_redirect = focus_widget is not source

            if should_redirect:
                scrollbar = self.setup_scroll_area.verticalScrollBar()
                scrollbar.setValue(scrollbar.value() - event.angleDelta().y())
                return True

        return super().eventFilter(source, event)

    def _build_overview_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        self.summary_label = QLabel("Add PDFs and process them to see results.")
        self.summary_label.setObjectName("summaryLabel")
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)

        stats_grid = QGridLayout()
        stats_grid.setHorizontalSpacing(12)
        stats_grid.setVerticalSpacing(12)
        self.processed_value, processed_card = self._stat_card("Processed files")
        self.review_value, review_card = self._stat_card("Manual review patients")
        self.averaged_value, averaged_card = self._stat_card("Averaged patients")
        self.special_value, special_card = self._stat_card("Special rows")
        stats_grid.addWidget(processed_card, 0, 0)
        stats_grid.addWidget(review_card, 0, 1)
        stats_grid.addWidget(averaged_card, 0, 2)
        stats_grid.addWidget(special_card, 0, 3)
        layout.addLayout(stats_grid)

        self.overview_table = QTableWidget(0, 10)
        self.overview_table.setHorizontalHeaderLabels(
            [
                "Source File",
                "Patient ID",
                "Record #",
                "Scan Date",
                "Scan Time",
                "SYS",
                "DIA",
                "MAP",
                "Pairing",
                "Status",
            ]
        )
        self._apply_overview_column_widths()
        self.overview_table.verticalHeader().setVisible(False)
        self.overview_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.overview_table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self.overview_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.overview_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.overview_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.overview_table.viewport().installEventFilter(self)
        self.overview_table.itemDoubleClicked.connect(self.open_selected_overview_pdf)
        self.overview_table.itemSelectionChanged.connect(self._sync_controls)
        layout.addWidget(self.overview_table, 1)

        overview_buttons = QHBoxLayout()
        self.open_selected_overview_pdf_button = QPushButton("View selected PDF")
        overview_buttons.addWidget(self.open_selected_overview_pdf_button)
        overview_buttons.addStretch(1)
        layout.addLayout(overview_buttons)

        self.open_selected_overview_pdf_button.clicked.connect(
            self.open_selected_overview_pdf
        )
        return tab

    def _apply_overview_column_widths(self) -> None:
        header = self.overview_table.horizontalHeader()
        header.setMinimumSectionSize(54)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(9, QHeaderView.ResizeMode.Fixed)
        self.overview_table.setColumnWidth(0, 136)
        self.overview_table.setColumnWidth(1, 96)
        self.overview_table.setColumnWidth(2, 86)
        self.overview_table.setColumnWidth(3, 96)
        self.overview_table.setColumnWidth(4, 86)
        self.overview_table.setColumnWidth(5, 58)
        self.overview_table.setColumnWidth(6, 58)
        self.overview_table.setColumnWidth(7, 64)
        self.overview_table.setColumnWidth(8, 80)
        self.overview_table.setColumnWidth(9, 150)
        self.overview_table.horizontalScrollBar().setValue(0)

    def _apply_pair_table_layout(self) -> None:
        header = self.pair_table.horizontalHeader()
        header.setMinimumSectionSize(54)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(7, QHeaderView.ResizeMode.Fixed)
        header.setStretchLastSection(False)
        self.pair_table.setColumnWidth(0, 72)
        self.pair_table.setColumnWidth(1, 78)
        self.pair_table.setColumnWidth(2, 78)
        self.pair_table.setColumnWidth(3, 68)
        self.pair_table.setColumnWidth(4, 96)
        self.pair_table.setColumnWidth(5, 96)
        self.pair_table.setColumnWidth(6, 320)
        self.pair_table.setColumnWidth(7, 118)
        self.pair_table.horizontalScrollBar().setValue(0)

    def _update_pair_table_height(self) -> None:
        header_height = self.pair_table.horizontalHeader().height()
        row_height = self.pair_table.verticalHeader().defaultSectionSize()
        frame = self.pair_table.frameWidth() * 2
        scrollbar_height = self.pair_table.horizontalScrollBar().sizeHint().height()
        row_count = max(self.pair_table.rowCount(), 1)
        total_height = header_height + (row_count * row_height) + frame + scrollbar_height + 4
        self.pair_table.setFixedHeight(total_height)
        self.pair_table.verticalScrollBar().setValue(0)

    def _apply_diff_table_layout(self) -> None:
        header = self.diff_table.horizontalHeader()
        header.setMinimumSectionSize(108)
        for column_index in range(self.diff_table.columnCount()):
            header.setSectionResizeMode(column_index, QHeaderView.ResizeMode.Stretch)
        self.diff_table.horizontalScrollBar().setValue(0)
        header_height = self.diff_table.horizontalHeader().height()
        row_height = self.diff_table.verticalHeader().defaultSectionSize()
        frame = self.diff_table.frameWidth() * 2
        total_height = header_height + row_height + frame + 2
        self.diff_table.setFixedHeight(total_height)
        self.diff_table.verticalScrollBar().setValue(0)

    def _update_review_split_layout(self) -> None:
        if (
            self.review_split is None
            or self.review_patient_panel is None
            or self.review_detail_panel is None
        ):
            return

        if self.review_split.orientation() != Qt.Orientation.Horizontal:
            self.review_split.setOrientation(Qt.Orientation.Horizontal)
            self.review_split.setSizes([220, 760])

        self.review_patient_panel.setMinimumWidth(220)
        self.review_patient_panel.setMaximumWidth(320)
        self.review_patient_panel.setMinimumHeight(0)
        self.review_patient_panel.setMaximumHeight(16777215)
        self.review_split_compact = False

    def _apply_data_table_widths(
        self,
        table: QTableWidget,
        columns: list[str],
    ) -> None:
        header = table.horizontalHeader()
        header.setMinimumSectionSize(56)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

        width_map = {
            "Source File": 185,
            "Patient ID": 90,
            "Scanned ID": 96,
            "Scan Date": 96,
            "Scan Time": 84,
            "Record #": 84,
            "Analyed": 70,
            "Date of Birth": 96,
            "Age": 56,
            "Gender": 72,
            "Height (m)": 78,
            "# of Pulses": 80,
            "Pulse Height": 84,
            "Source Path": 210,
        }

        for column_index, column_name in enumerate(columns):
            width = width_map.get(column_name)
            if width is None:
                if any(
                    token in column_name
                    for token in ["(mmHg)", "(%)", "(ms)", "(bpm)", "(m/s)"]
                ):
                    width = 92
                elif "Variation" in column_name or "Pressure" in column_name:
                    width = 98
                else:
                    width = min(max(table.columnWidth(column_index), 72), 120)
            width = max(
                width,
                header.fontMetrics().horizontalAdvance(column_name) + 28,
            )
            table.setColumnWidth(column_index, width)

        table.horizontalScrollBar().setValue(0)

    def _build_review_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        review_banner = QFrame()
        review_banner.setObjectName("subpanel")
        review_banner_layout = QHBoxLayout(review_banner)
        review_banner_layout.setContentsMargins(14, 12, 14, 12)
        review_banner_layout.setSpacing(12)

        self.review_count_badge = QLabel("0")
        self.review_count_badge.setObjectName("reviewBadge")
        self.review_count_badge.setAlignment(
            Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter
        )
        review_banner_layout.addWidget(self.review_count_badge, 0)

        self.review_status_label = QLabel(
            "Patients with more than two entries will appear here after processing."
        )
        self.review_status_label.setObjectName("summaryLabel")
        self.review_status_label.setWordWrap(True)
        review_banner_layout.addWidget(self.review_status_label, 1)
        layout.addWidget(review_banner)

        review_split = QSplitter(Qt.Orientation.Horizontal)
        review_split.setChildrenCollapsible(False)
        self.review_split = review_split

        patient_panel = QFrame()
        patient_panel.setObjectName("subpanel")
        patient_panel.setMinimumWidth(220)
        patient_panel.setMaximumWidth(320)
        self.review_patient_panel = patient_panel
        patient_layout = QVBoxLayout(patient_panel)
        patient_layout.setContentsMargins(14, 14, 14, 14)
        patient_layout.setSpacing(10)
        patient_header = QHBoxLayout()
        patient_header.setSpacing(8)
        patient_header.addWidget(self._micro_title("Review queue"), 1)
        self.review_queue_badge = QLabel("0")
        self.review_queue_badge.setObjectName("reviewBadge")
        self.review_queue_badge.setAlignment(
            Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter
        )
        patient_header.addWidget(self.review_queue_badge, 0)
        patient_layout.addLayout(patient_header)
        self.patient_list = QListWidget()
        self.patient_list.setAlternatingRowColors(True)
        self.patient_list.currentRowChanged.connect(self._manual_patient_changed)
        patient_layout.addWidget(self.patient_list, 1)
        self.review_hint_label = QLabel(
            "Each patient starts with the automatic pair already selected."
        )
        self.review_hint_label.setObjectName("helperSmall")
        self.review_hint_label.setWordWrap(True)
        patient_layout.addWidget(self.review_hint_label)

        detail_panel = QFrame()
        detail_panel.setObjectName("subpanel")
        self.review_detail_panel = detail_panel
        detail_panel_layout = QVBoxLayout(detail_panel)
        detail_panel_layout.setContentsMargins(0, 0, 0, 0)
        detail_panel_layout.setSpacing(0)

        detail_scroll = QScrollArea()
        detail_scroll.setWidgetResizable(True)
        detail_scroll.setFrameShape(QFrame.Shape.NoFrame)
        detail_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        detail_scroll.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOn
        )
        self.review_detail_scroll = detail_scroll
        detail_panel_layout.addWidget(detail_scroll)

        detail_content = QWidget()
        detail_scroll.setWidget(detail_content)

        detail_layout = QVBoxLayout(detail_content)
        detail_layout.setContentsMargins(14, 14, 14, 14)
        detail_layout.setSpacing(12)

        patient_context_card = QFrame()
        patient_context_card.setObjectName("subpanel")
        patient_context_layout = QVBoxLayout(patient_context_card)
        patient_context_layout.setContentsMargins(14, 12, 14, 12)
        patient_context_layout.setSpacing(8)

        patient_context_header = QHBoxLayout()
        patient_context_header.setSpacing(10)
        self.current_patient_label = QLabel("No patient selected")
        self.current_patient_label.setObjectName("reviewPatientTitle")
        self.current_patient_label.setWordWrap(True)
        patient_context_header.addWidget(self.current_patient_label, 1)
        self.review_selection_badge = QLabel("")
        self.review_selection_badge.setObjectName("neutralPill")
        self.review_selection_badge.setAlignment(
            Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignVCenter
        )
        patient_context_header.addWidget(self.review_selection_badge, 0)

        self.selection_label = QLabel("")
        self.selection_label.setObjectName("helperSmall")
        self.selection_label.setWordWrap(True)
        self.selected_files_label = QLabel("")
        self.selected_files_label.setObjectName("fileSummary")
        self.selected_files_label.setWordWrap(True)
        self.review_warning_label = QLabel("")
        self.review_warning_label.setObjectName("warningLabel")
        self.review_warning_label.setWordWrap(True)
        self.review_warning_label.hide()
        patient_context_layout.addLayout(patient_context_header)
        patient_context_layout.addWidget(self.selection_label)
        patient_context_layout.addWidget(self.selected_files_label)
        detail_layout.addWidget(patient_context_card)

        action_row = QHBoxLayout()
        action_row.setSpacing(10)
        self.reset_auto_button = QPushButton("Reset to auto pair")
        self.view_pair_pdf_button = QPushButton("View selected PDF")
        self.reset_auto_button.setMinimumWidth(190)
        self.view_pair_pdf_button.setMinimumWidth(190)
        action_row.addWidget(self.reset_auto_button)
        action_row.addWidget(self.view_pair_pdf_button)
        action_row.addStretch(1)
        detail_layout.addLayout(action_row)

        pair_card = QFrame()
        pair_card.setObjectName("subpanel")
        pair_card_layout = QVBoxLayout(pair_card)
        pair_card_layout.setContentsMargins(14, 12, 14, 14)
        pair_card_layout.setSpacing(10)
        pair_card_layout.addWidget(self._micro_title("Selected measurements"))

        self.pair_table = QTableWidget(0, 8)
        self.pair_table.setHorizontalHeaderLabels(
            [
                "Keep",
                "SYS",
                "DIA",
                "MAP",
                "Aortic SYS",
                "Aortic DIA",
                "Source File",
                "Pair method",
            ]
        )
        header = self.pair_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.pair_table.verticalHeader().setVisible(False)
        self.pair_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.pair_table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self.pair_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.pair_table.setWordWrap(False)
        self.pair_table.verticalHeader().setDefaultSectionSize(34)
        self.pair_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.pair_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.pair_table.viewport().installEventFilter(self)
        self.pair_table.itemDoubleClicked.connect(self.open_selected_pair_pdf)
        self.pair_table.itemSelectionChanged.connect(self._sync_controls)
        pair_card_layout.addWidget(self.pair_table)
        detail_layout.addWidget(pair_card)

        diff_card = QFrame()
        diff_card.setObjectName("subpanel")
        diff_card_layout = QVBoxLayout(diff_card)
        diff_card_layout.setContentsMargins(14, 12, 14, 12)
        diff_card_layout.setSpacing(8)
        diff_card_layout.addWidget(self._micro_title("Selected pair absolute differences"))

        self.diff_table = QTableWidget(1, 5)
        self.diff_table.setHorizontalHeaderLabels(
            [
                "Peripheral SYS",
                "Peripheral DIA",
                "MAP",
                "Aortic SYS",
                "Aortic DIA",
            ]
        )
        self.diff_table.verticalHeader().setVisible(False)
        self.diff_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.diff_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.diff_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.diff_table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.diff_table.setFixedHeight(92)
        self.diff_table.verticalHeader().setDefaultSectionSize(34)
        self.diff_table.viewport().installEventFilter(self)
        self._apply_diff_table_layout()
        diff_card_layout.addWidget(self.diff_table)

        self.diff_status_label = QLabel(
            "Select exactly two rows to see pair differences."
        )
        self.diff_status_label.setObjectName("helperSmall")
        self.diff_status_label.setWordWrap(True)
        diff_card_layout.addWidget(self.diff_status_label)
        detail_layout.addWidget(diff_card)

        self.reset_auto_button.clicked.connect(self.reset_current_patient_to_auto)
        self.view_pair_pdf_button.clicked.connect(self.open_selected_pair_pdf)

        review_split.addWidget(patient_panel)
        review_split.addWidget(detail_panel)
        review_split.setStretchFactor(0, 0)
        review_split.setStretchFactor(1, 1)
        review_split.setSizes([220, 760])
        self._update_review_split_layout()
        self._apply_pair_table_layout()
        layout.addWidget(review_split, 1)

        return tab

    def _build_all_data_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        helper = QLabel(
            "Preview the rows that will populate the workbook's All Data sheet."
        )
        helper.setObjectName("statusLabel")
        helper.setWordWrap(True)
        layout.addWidget(helper)

        buttons = QHBoxLayout()
        self.open_selected_all_data_pdf_button = QPushButton("View selected PDF")
        buttons.addWidget(self.open_selected_all_data_pdf_button)
        buttons.addStretch(1)
        layout.addLayout(buttons)

        self.all_data_table = QTableWidget()
        self.all_data_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Interactive
        )
        self.all_data_table.verticalHeader().setVisible(False)
        self.all_data_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.all_data_table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self.all_data_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.all_data_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.all_data_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.all_data_table.viewport().installEventFilter(self)
        self.all_data_table.itemDoubleClicked.connect(self.open_selected_all_data_pdf)
        self.all_data_table.itemSelectionChanged.connect(self._sync_controls)
        layout.addWidget(self.all_data_table, 1)

        self.open_selected_all_data_pdf_button.clicked.connect(
            self.open_selected_all_data_pdf
        )
        return tab

    def _build_averaged_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        helper = QLabel(
            "Preview the per-patient rows that will populate the Averaged Data sheet."
        )
        helper.setObjectName("statusLabel")
        helper.setWordWrap(True)
        layout.addWidget(helper)

        self.averaged_table = QTableWidget()
        self.averaged_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Interactive
        )
        self.averaged_table.verticalHeader().setVisible(False)
        self.averaged_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.averaged_table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self.averaged_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.averaged_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.averaged_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.averaged_table.viewport().installEventFilter(self)
        layout.addWidget(self.averaged_table, 1)

        return tab

    def _section_title(self, title: str, description: str) -> QWidget:
        container = QWidget()
        container.setMinimumHeight(64)
        container.setSizePolicy(
            QSizePolicy.Policy.Preferred,
            QSizePolicy.Policy.Minimum,
        )
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)

        title_label = QLabel(title)
        title_label.setObjectName("sectionTitle")
        title_label.setMinimumHeight(22)
        description_label = QLabel(description)
        description_label.setObjectName("sectionDescription")
        description_label.setWordWrap(True)
        description_label.setMinimumHeight(32)
        description_label.setSizePolicy(
            QSizePolicy.Policy.Preferred,
            QSizePolicy.Policy.Minimum,
        )

        layout.addWidget(title_label)
        layout.addWidget(description_label)
        return container

    def _micro_title(self, title: str) -> QLabel:
        label = QLabel(title)
        label.setObjectName("summaryLabel")
        return label

    def _stat_card(self, label_text: str) -> tuple[QLabel, QWidget]:
        card = QFrame()
        card.setObjectName("statCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(2)

        value_label = QLabel("0")
        value_label.setObjectName("cardValue")
        caption = QLabel(label_text)
        caption.setObjectName("cardLabel")
        caption.setWordWrap(True)

        layout.addWidget(value_label)
        layout.addWidget(caption)
        return value_label, card

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QWidget {
                background: #f3f1eb;
                color: #1e2a2f;
                font-family: "Segoe UI";
                font-size: 10pt;
            }
            QFrame#panel {
                background: #fffdf8;
                border: 1px solid #d8d0c3;
                border-radius: 12px;
            }
            QFrame#subpanel,
            QFrame#statCard {
                background: #f9f6ef;
                border: 1px solid #e3dbce;
                border-radius: 10px;
            }
            QLabel#title {
                font-size: 24pt;
                font-weight: 700;
                color: #17313a;
            }
            QLabel#subtitle,
            QLabel#sectionDescription,
            QLabel#statusLabel,
            QLabel#cardLabel,
            QLabel#dialogSubtitle {
                color: #66757b;
            }
            QLabel#sectionTitle,
            QLabel#summaryLabel,
            QLabel#dialogTitle,
            QLabel#reviewPatientTitle {
                font-size: 11pt;
                font-weight: 650;
                color: #17313a;
            }
            QLabel#reviewPatientTitle {
                font-size: 14pt;
                font-weight: 700;
            }
            QLabel#dialogTitle {
                font-size: 20pt;
            }
            QLabel#cardValue {
                font-size: 20pt;
                font-weight: 700;
                color: #0e6d69;
            }
            QLabel#warningLabel {
                color: #9a4d00;
                font-weight: 600;
            }
            QLabel#helperSmall,
            QLabel#fileSummary {
                color: #5f6d73;
            }
            QLabel#fileSummary {
                background: #f3f7f5;
                border: 1px solid #d7e2dc;
                border-radius: 8px;
                padding: 8px 10px;
            }
            QLabel#reviewBadge,
            QLabel#successPill,
            QLabel#neutralPill {
                font-weight: 650;
                border-radius: 10px;
                padding: 4px 10px;
                min-height: 20px;
            }
            QLabel#reviewBadge {
                background: #e7efe9;
                color: #0e6d69;
                border: 1px solid #bad0c7;
                min-width: 26px;
            }
            QLabel#successPill {
                background: #dff1ea;
                color: #0e6d69;
                border: 1px solid #badfd1;
            }
            QLabel#neutralPill {
                background: #efe9dc;
                color: #6a5635;
                border: 1px solid #ddd1bc;
            }
            QLabel#banner {
                background: #ebe5d8;
                border: 1px solid #d8d0c3;
                border-radius: 12px;
            }
            QToolButton#infoButton {
                background: #e7efe9;
                border: 1px solid #bad0c7;
                border-radius: 12px;
                min-width: 24px;
                max-width: 24px;
                min-height: 24px;
                max-height: 24px;
                font-size: 10pt;
                font-weight: 700;
                color: #0e6d69;
                padding: 0px;
            }
            QToolButton#infoButton:hover {
                background: #dbe7e1;
            }
            QPushButton {
                background: #eff2ee;
                border: 1px solid #c0ccc6;
                border-radius: 7px;
                padding: 8px 12px;
            }
            QPushButton:hover {
                background: #e1e8e3;
            }
            QPushButton:disabled {
                color: #8f989c;
                background: #f2f2f2;
                border-color: #dadada;
            }
            QPushButton#primaryButton {
                background: #0e6d69;
                color: #ffffff;
                border: 1px solid #0b5955;
                font-weight: 650;
            }
            QPushButton#primaryButton:hover {
                background: #0c5e5a;
            }
            QListWidget,
            QLineEdit,
            QTextBrowser,
            QTableWidget {
                background: #ffffff;
                border: 1px solid #d1d7d4;
                border-radius: 8px;
                selection-background-color: #cde7e0;
                selection-color: #17313a;
            }
            QListWidget::item {
                padding: 8px 10px;
                border-radius: 6px;
                margin: 2px 0;
            }
            QListWidget::item:selected {
                background: #dff1ea;
                color: #17313a;
            }
            QLineEdit {
                padding: 8px;
            }
            QTableWidget {
                gridline-color: #e5e7e7;
                alternate-background-color: #fbfcfb;
            }
            QTableWidget::item:selected {
                border: 1px solid #0e6d69;
                color: #17313a;
            }
            QHeaderView::section {
                background: #eff2ee;
                border: 0;
                border-right: 1px solid #dde2df;
                border-bottom: 1px solid #dde2df;
                padding: 8px 10px;
                font-weight: 600;
            }
            QProgressBar {
                background: #f0f0ee;
                border: 1px solid #d6d1c8;
                border-radius: 7px;
                text-align: center;
            }
            QProgressBar::chunk {
                background: #0e6d69;
                border-radius: 6px;
            }
            QTabWidget::pane {
                border: 0;
            }
            QTabBar::tab {
                background: #ece7dd;
                border: 1px solid #ddd4c7;
                border-bottom: 0;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                padding: 9px 14px;
                margin-right: 4px;
            }
            QTabBar::tab:selected {
                background: #fffdf8;
                color: #17313a;
            }
            """
        )

    def show_info_dialog(self) -> None:
        self.readme_dialog = ReadmeDialog(self)
        self.readme_dialog.exec()

    def show_thresholds_help_dialog(self) -> None:
        dialog = QMessageBox(self)
        dialog.setWindowTitle("Threshold settings")
        dialog.setIcon(QMessageBox.Icon.Information)
        dialog.setTextFormat(Qt.TextFormat.RichText)
        dialog.setText(
            """
            <div style="min-width: 460px; line-height: 1.45;">
              <p style="margin: 0 0 14px 0;">
                These settings control how the difference row is highlighted during
                multi-entry review.
              </p>
              <p style="margin: 0 0 4px 0;"><b>Green up to</b></p>
              <p style="margin: 0 0 14px 0;">
                Differences at or below this value are shown in green.<br>
                Use this range for pairs that look closely matched.
              </p>
              <p style="margin: 0 0 4px 0;"><b>Yellow up to</b></p>
              <p style="margin: 0 0 14px 0;">
                Differences above the green range and up to this value are shown in yellow.<br>
                Use this range for pairs that may still be acceptable, but should be looked at more carefully.
              </p>
              <p style="margin: 0 0 4px 0;"><b>Pair alert above</b></p>
              <p style="margin: 0 0 14px 0;">
                If the selected pair differs by more than this value in peripheral systolic
                or peripheral diastolic pressure, the pair is flagged in the review tab and
                the overview tab.
              </p>
              <p style="margin: 0 0 4px 0;"><b>Notes</b></p>
              <table style="margin: 0; border-collapse: collapse;">
                <tr>
                  <td style="padding: 0 8px 6px 0; vertical-align: top;">&bull;</td>
                  <td style="padding: 0 0 6px 0;">Differences above the yellow range are shown in red.</td>
                </tr>
                <tr>
                  <td style="padding: 0 8px 0 0; vertical-align: top;">&bull;</td>
                  <td style="padding: 0;">Changing these settings updates the review colors and alert rules for the current session.</td>
                </tr>
              </table>
            </div>
            """
        )
        dialog.setStandardButtons(QMessageBox.StandardButton.Ok)
        dialog.exec()

    def add_pdf_files(self) -> None:
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PWA PDF files",
            "",
            "PDF Files (*.pdf)",
        )
        if not file_paths:
            return

        existing = {path.resolve() for path in self.pdf_paths}
        for file_path in file_paths:
            candidate = Path(file_path).resolve()
            if candidate not in existing:
                self.pdf_paths.append(candidate)
                existing.add(candidate)

        self.pdf_paths.sort(key=lambda path: path.name.lower())
        self._refresh_file_list()
        self._sync_controls()

    def remove_selected_files(self) -> None:
        selected_rows = sorted(
            {self.file_list.row(item) for item in self.file_list.selectedItems()},
            reverse=True,
        )
        if not selected_rows:
            return

        for row in selected_rows:
            del self.pdf_paths[row]

        self._refresh_file_list()
        self._sync_controls()

    def clear_pdf_files(self) -> None:
        if not self.pdf_paths:
            return
        self.pdf_paths = []
        self._refresh_file_list()
        self._sync_controls()

    def _refresh_file_list(self) -> None:
        self.file_list.clear()
        for path in self.pdf_paths:
            self.file_list.addItem(path.name)

    def browse_output_path(self) -> None:
        current = (
            str(self.output_path)
            if str(self.output_path)
            else str(default_output_path())
        )
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Excel file as",
            current,
            "Excel Workbook (*.xlsx)",
        )
        if not output_path:
            return

        selected_path = Path(output_path)
        if selected_path.suffix.lower() != ".xlsx":
            selected_path = selected_path.with_suffix(".xlsx")
        self.output_line.setText(str(selected_path))

    def _output_path_changed(self, value: str) -> None:
        self.output_path = Path(value.strip()) if value.strip() else Path()
        self.output_line.setCursorPosition(0)
        self._sync_controls()

    def process_files(self) -> None:
        if not self.pdf_paths:
            QMessageBox.warning(self, "No PDFs selected", "Add one or more PWA PDF files.")
            return
        if not str(self.output_path):
            QMessageBox.warning(
                self,
                "No export file",
                "Choose an Excel export location.",
            )
            return

        self.records = []
        self.bundle = None
        self.auto_pairs = {}
        self.manual_pairs = {}
        self.last_export_path = None
        self.open_output_button.setEnabled(False)
        self.progress.setMaximum(len(self.pdf_paths))
        self.progress.setValue(0)
        self.status_label.setText("Starting PWA processing...")
        self._refresh_results_views()
        self._set_processing_state(True)

        self.thread = QThread()
        self.worker = ProcessingWorker(self.pdf_paths)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.processing_progress)
        self.worker.finished.connect(self.processing_finished)
        self.worker.failed.connect(self.processing_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(self._clear_worker_refs)
        self.thread.start()

    def processing_progress(self, current: int, total: int, message: str) -> None:
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        self.status_label.setText(message)

    def processing_finished(self, records: object) -> None:
        self.records = list(records)
        self._rebuild_analysis(seed_manual=True)
        self.status_label.setText(
            f"Processed {len(self.records)} file(s). Review multi-entry patients before exporting."
        )
        self._set_processing_state(False)
        if self.bundle and self.bundle.manual_patients:
            self.tabs.setCurrentIndex(1)
        else:
            self.tabs.setCurrentIndex(0)

    def processing_failed(self, details: str) -> None:
        self.status_label.setText("Processing failed.")
        self._set_processing_state(False)
        QMessageBox.critical(
            self,
            "Processing failed",
            "The extractor could not finish processing the selected PDFs.\n\n"
            f"{details}",
        )
        self._refresh_results_views()

    def _clear_worker_refs(self) -> None:
        self.thread = None
        self.worker = None
        self._sync_controls()

    def _rebuild_analysis(
        self,
        seed_manual: bool = False,
        preserve_patient_id: Optional[str] = None,
    ) -> None:
        if not self.records:
            self.bundle = None
            self.auto_pairs = {}
            self.manual_pairs = {}
            self._refresh_results_views()
            return

        auto_bundle = build_analysis(
            self.records,
            pair_alert_threshold=self.pair_alert_threshold,
        )
        self.auto_pairs = auto_bundle.used_pairs
        if seed_manual or not self.manual_pairs:
            self.manual_pairs = initial_manual_pairs(
                auto_bundle.dataframe,
                auto_bundle.used_pairs,
                auto_bundle.manual_patients,
            )
        else:
            valid_patients = set(auto_bundle.manual_patients)
            self.manual_pairs = {
                patient_id: selection
                for patient_id, selection in self.manual_pairs.items()
                if patient_id in valid_patients
            }
            for patient_id in auto_bundle.manual_patients:
                if patient_id not in self.manual_pairs:
                    fallback = initial_manual_pairs(
                        auto_bundle.dataframe,
                        auto_bundle.used_pairs,
                        [patient_id],
                    )
                    self.manual_pairs[patient_id] = fallback.get(patient_id, [])

        manual_pair_tuples = {
            patient_id: tuple(selection)
            for patient_id, selection in self.manual_pairs.items()
            if len(selection) == 2
        }
        self.bundle = build_analysis(
            self.records,
            manual_pairs=manual_pair_tuples,
            pair_alert_threshold=self.pair_alert_threshold,
        )
        self._refresh_results_views(preserve_patient_id=preserve_patient_id)

    def _refresh_results_views(self, preserve_patient_id: Optional[str] = None) -> None:
        self._refresh_overview()
        self._refresh_review_panel(preserve_patient_id=preserve_patient_id)
        self._refresh_all_data_table()
        self._refresh_averaged_table()
        self._sync_controls()

    def _refresh_overview(self) -> None:
        if not self.bundle:
            self.summary_label.setText("Add PDFs and process them to see results.")
            self.processed_value.setText("0")
            self.review_value.setText("0")
            self.averaged_value.setText("0")
            self.special_value.setText("0")
            self.overview_table.setRowCount(0)
            return

        overview_df = display_dataframe(self.bundle)
        entry_counts = patient_entry_counts(self.bundle.dataframe)
        detailed_count = int((overview_df["Special Row"] != True).sum())
        special_count = int((overview_df["Special Row"] == True).sum())
        single_entry_count = sum(1 for count in entry_counts.values() if count == 1)
        pair_alert_count = 0
        for patient_id, pair in self.bundle.used_pairs.items():
            pair_df = self.bundle.dataframe.loc[list(pair)]
            if pair_alert_triggered(pair_df, self.pair_alert_threshold):
                pair_alert_count += 1
        self.processed_value.setText(str(len(overview_df)))
        self.review_value.setText(str(len(self.bundle.manual_patients)))
        self.averaged_value.setText(str(len(self.bundle.analyzed_df)))
        self.special_value.setText(str(special_count))

        self.summary_label.setText(
            f"{len(overview_df)} file(s) processed. "
            f"{detailed_count} detailed report row(s), "
            f"{special_count} special row(s), "
            f"{single_entry_count} single-entry patient(s), and "
            f"{pair_alert_count} selected pair alert(s) above {self.pair_alert_threshold:.1f} mmHg."
        )

        self.overview_table.setRowCount(len(overview_df))
        for table_row, (frame_index, row) in enumerate(overview_df.iterrows()):
            status = record_status(row.get("Patient ID"))
            pairing = row.get("Analyed", "")
            if row.get("Special Row") == True:
                pairing = ""
            else:
                patient_id = str(row.get("Patient ID") or "")
                patient_count = entry_counts.get(patient_id, 0)
                pair = self.bundle.used_pairs.get(patient_id)
                pair_alert = False
                if pair is not None:
                    pair_df = self.bundle.dataframe.loc[list(pair)]
                    pair_alert = pair_alert_triggered(pair_df, self.pair_alert_threshold)

                if patient_count == 1:
                    status = "Single entry"
                elif pair_alert:
                    status = f"Pair alert > {self.pair_alert_threshold:.1f}"
                elif patient_count > 2:
                    status = "Multi-entry review"
                elif patient_count == 2:
                    status = "Two-entry pair"

            values = [
                format_value(row.get("Source File")),
                format_value(row.get("Patient ID")),
                format_value(row.get("Record #")),
                format_value(row.get("Scan Date")),
                format_value(row.get("Scan Time")),
                format_value(row.get("Peripheral Systolic Pressure (mmHg)"))
                if row.get("Special Row") != True
                else "",
                format_value(row.get("Peripheral Diastolic Pressure (mmHg)"))
                if row.get("Special Row") != True
                else "",
                format_value(row.get("Peripheral Mean Pressure (mmHg)"))
                if row.get("Special Row") != True
                else "",
                pairing,
                status,
            ]

            for column_index, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column_index == 0:
                    item.setData(Qt.ItemDataRole.UserRole, frame_index)
                if column_index in {2, 3, 4, 5, 6, 7, 8}:
                    item.setTextAlignment(
                        int(
                            Qt.AlignmentFlag.AlignCenter
                            | Qt.AlignmentFlag.AlignVCenter
                        )
                    )
                if status.startswith("Pair alert"):
                    item.setBackground(QColor("#f6d7d4"))
                elif status == "Single entry":
                    item.setBackground(QColor("#fff2bf"))
                elif pairing == "Yes":
                    item.setBackground(QColor("#dff1ea"))
                self.overview_table.setItem(table_row, column_index, item)

        self._apply_overview_column_widths()
        if self.overview_table.rowCount() > 0:
            self.overview_table.selectRow(0)
            self.overview_table.horizontalScrollBar().setValue(0)

    def _refresh_review_panel(self, preserve_patient_id: Optional[str] = None) -> None:
        self.patient_list.blockSignals(True)
        self.patient_list.clear()

        if not self.bundle or not self.bundle.manual_patients:
            self.review_status_label.setText(
                "No multi-entry patients need manual review. Automatic pairing is ready."
                if self.records
                else "Patients with more than two entries will appear here after processing."
            )
            self.review_count_badge.setText("0")
            self.review_queue_badge.setText("0")
            self.current_patient_label.setText("No patient selected")
            self.selection_label.setText("")
            self.selected_files_label.setText("")
            self.review_selection_badge.setText("")
            self.review_warning_label.setText("")
            self.pair_table.setRowCount(0)
            self._refresh_difference_table(None)
            self.patient_list.blockSignals(False)
            return

        current_patient = preserve_patient_id or self.current_manual_patient_id()
        patient_count = len(self.bundle.manual_patients)
        self.review_count_badge.setText(str(patient_count))
        self.review_queue_badge.setText(str(patient_count))
        self.review_status_label.setText(
            "Automatic pairs are preselected. Review only the patients that need changes."
        )

        selected_row = 0
        for row_index, patient_id in enumerate(self.bundle.manual_patients):
            selection = self.manual_pairs.get(patient_id, [])
            label = f"{patient_id} ({len(selection)}/2 selected)"
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, patient_id)
            self.patient_list.addItem(item)
            if patient_id == current_patient:
                selected_row = row_index

        self.patient_list.setCurrentRow(selected_row)
        self.patient_list.blockSignals(False)
        self._render_current_patient()

    def current_manual_patient_id(self) -> Optional[str]:
        item = self.patient_list.currentItem()
        if item is None:
            return None
        return item.data(Qt.ItemDataRole.UserRole)

    def _manual_patient_changed(self, row: int) -> None:
        patient_id = None
        if row >= 0:
            item = self.patient_list.item(row)
            if item is not None:
                patient_id = item.data(Qt.ItemDataRole.UserRole)
        self._render_current_patient(patient_id_override=patient_id)
        self._sync_controls()

    def _render_current_patient(self, patient_id_override: Optional[str] = None) -> None:
        if not self.bundle:
            self.current_patient_label.setText("No patient selected")
            self.selection_label.setText("")
            self.selected_files_label.setText("")
            self.review_selection_badge.setText("")
            self.review_warning_label.setText("")
            self.pair_table.setRowCount(0)
            self._refresh_difference_table(None)
            return

        patient_id = patient_id_override or self.current_manual_patient_id()
        if not patient_id:
            self.current_patient_label.setText("No patient selected")
            self.selection_label.setText("")
            self.selected_files_label.setText("")
            self.review_selection_badge.setText("")
            self.review_warning_label.setText("")
            self.pair_table.setRowCount(0)
            self._refresh_difference_table(None)
            return

        rows = patient_rows(self.bundle.dataframe, patient_id)
        selection = self.manual_pairs.get(patient_id, [])
        auto_pair = set(self.auto_pairs.get(patient_id, ()))

        self.current_patient_label.setText(patient_id)
        self.selection_label.setText("Choose two rows to average for export.")
        selected_files = []
        for frame_index in selection:
            if frame_index in rows.index:
                file_name = rows.loc[frame_index].get("Source File")
                if file_name:
                    selected_files.append(str(file_name))
        if len(selection) == 2:
            self.review_selection_badge.setObjectName("successPill")
            self.review_selection_badge.setText("Ready for export")
            self.selected_files_label.setText(
                "Selected files: " + " | ".join(selected_files)
                if selected_files
                else "Selected files: automatic pair"
            )
            self.review_warning_label.setText("")
        else:
            self.review_selection_badge.setObjectName("neutralPill")
            self.review_selection_badge.setText(f"{len(selection)}/2 selected")
            self.selected_files_label.setText(
                "Selected files: " + " | ".join(selected_files)
                if selected_files
                else "Selected files: none selected yet"
            )
            self.review_warning_label.setText("")
        self.review_selection_badge.style().unpolish(self.review_selection_badge)
        self.review_selection_badge.style().polish(self.review_selection_badge)

        self.updating_pair_table = True
        self.pair_table.setRowCount(len(rows))
        for table_row, (frame_index, row) in enumerate(rows.iterrows()):
            self.pair_table.setCellWidget(
                table_row,
                0,
                self._create_keep_checkbox(
                    patient_id,
                    frame_index,
                    frame_index in selection,
                ),
            )

            values = [
                format_value(row.get("Peripheral Systolic Pressure (mmHg)")),
                format_value(row.get("Peripheral Diastolic Pressure (mmHg)")),
                format_value(row.get("Peripheral Mean Pressure (mmHg)")),
                format_value(row.get("Aortic Systolic Pressure (mmHg)")),
                format_value(row.get("Aortic Diastolic Pressure (mmHg)")),
                format_value(row.get("Source File")),
                "Auto" if frame_index in auto_pair else "",
            ]
            for column_offset, value in enumerate(values, start=1):
                item = QTableWidgetItem(value)
                if column_offset == 6:
                    item.setData(Qt.ItemDataRole.UserRole, frame_index)
                if column_offset != 6:
                    item.setTextAlignment(
                        int(
                            Qt.AlignmentFlag.AlignCenter
                            | Qt.AlignmentFlag.AlignVCenter
                        )
                    )
                if frame_index in selection:
                    item.setBackground(QColor("#dff1ea"))
                elif frame_index in auto_pair:
                    item.setBackground(QColor("#eef8f4"))
                self.pair_table.setItem(table_row, column_offset, item)

        self.updating_pair_table = False
        self._apply_pair_table_layout()
        self._update_pair_table_height()
        if self.pair_table.rowCount() > 0:
            self.pair_table.setCurrentCell(0, 1)
            self.pair_table.selectRow(0)
            self.pair_table.horizontalScrollBar().setValue(0)
        self._refresh_difference_table(patient_id)

    def reset_current_patient_to_auto(self) -> None:
        if not self.bundle:
            return
        patient_id = self.current_manual_patient_id()
        if not patient_id:
            return

        fallback = list(patient_rows(self.bundle.dataframe, patient_id).index[:2])
        auto_selection = list(self.auto_pairs.get(patient_id, ()))
        self.manual_pairs[patient_id] = (
            auto_selection[:2] if len(auto_selection) == 2 else fallback
        )
        self._render_current_patient()
        self._refresh_difference_table(patient_id)
        self._rebuild_analysis(preserve_patient_id=patient_id)
        if self.current_manual_patient_id() == patient_id:
            self._refresh_difference_table(patient_id)

    def _settings_changed(self) -> None:
        if self.updating_settings:
            return

        self.diff_green_max = float(self.green_max_spin.value())
        self.diff_yellow_max = max(float(self.yellow_max_spin.value()), self.diff_green_max)
        self.pair_alert_threshold = float(self.pair_alert_spin.value())

        if self.yellow_max_spin.value() != self.diff_yellow_max:
            self.updating_settings = True
            self.yellow_max_spin.setValue(self.diff_yellow_max)
            self.updating_settings = False

        if self.records:
            self._rebuild_analysis(
                preserve_patient_id=self.current_manual_patient_id()
            )

    def _diff_background(self, value: float | None) -> QColor:
        if value is None:
            return QColor("#f3f4f4")
        if value <= self.diff_green_max:
            return QColor("#dff1ea")
        if value <= self.diff_yellow_max:
            return QColor("#fff2bf")
        return QColor("#f6d7d4")

    def _pair_checkbox_toggled(
        self,
        patient_id: str,
        frame_index: int,
        checked: bool,
        checkbox: Optional[QCheckBox] = None,
    ) -> None:
        if self.updating_pair_table:
            return

        selection = list(self.manual_pairs.get(patient_id, []))
        if checked and frame_index not in selection:
            if len(selection) >= 2:
                if isinstance(checkbox, QCheckBox):
                    self.updating_pair_table = True
                    checkbox.setChecked(False)
                    self.updating_pair_table = False
                else:
                    self._render_current_patient()
                QMessageBox.warning(
                    self,
                    "Selection limit",
                    "You can only select two rows for each patient.",
                )
                self._refresh_difference_table(patient_id)
                return
            selection.append(frame_index)
        elif not checked and frame_index in selection:
            selection.remove(frame_index)

        self.manual_pairs[patient_id] = selection
        self._refresh_difference_table(patient_id)
        self._rebuild_analysis(preserve_patient_id=patient_id)

    def _create_keep_checkbox(self, patient_id: str, frame_index: int, checked: bool) -> QWidget:
        holder = QWidget()
        holder_layout = QHBoxLayout(holder)
        holder_layout.setContentsMargins(0, 0, 0, 0)
        holder_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        checkbox = QCheckBox()
        checkbox.setChecked(checked)
        checkbox.stateChanged.connect(
            lambda state, pid=patient_id, idx=frame_index, cb=checkbox: self._pair_checkbox_toggled(
                pid,
                idx,
                state == Qt.CheckState.Checked.value,
                cb,
            )
        )
        holder_layout.addWidget(checkbox)
        return holder

    def _refresh_difference_table(self, patient_id: Optional[str]) -> None:
        self.diff_table.setRowCount(1)
        if not self.bundle or not patient_id:
            for column_index in range(self.diff_table.columnCount()):
                self.diff_table.setItem(0, column_index, QTableWidgetItem(""))
            self._apply_diff_table_layout()
            self.diff_status_label.setText(
                "Select exactly two rows to see pair differences."
            )
            return

        selection = self.manual_pairs.get(patient_id, [])
        if len(selection) != 2:
            for column_index in range(self.diff_table.columnCount()):
                self.diff_table.setItem(0, column_index, QTableWidgetItem(""))
            self._apply_diff_table_layout()
            self.diff_status_label.setText(
                "Select exactly two rows to see pair differences."
            )
            return

        pair_df = self.bundle.dataframe.loc[selection]
        differences = calculate_pair_differences(pair_df)
        columns = [
            "Pair Diff Peripheral Systolic (mmHg)",
            "Pair Diff Peripheral Diastolic (mmHg)",
            "Pair Diff Peripheral Mean (mmHg)",
            "Pair Diff Aortic Systolic (mmHg)",
            "Pair Diff Aortic Diastolic (mmHg)",
        ]

        for column_index, column_name in enumerate(columns):
            raw_value = differences.get(column_name)
            item = QTableWidgetItem(format_value(raw_value))
            item.setBackground(self._diff_background(raw_value))
            item.setTextAlignment(int(Qt.AlignmentFlag.AlignCenter))
            self.diff_table.setItem(0, column_index, item)

        self._apply_diff_table_layout()
        if pair_alert_triggered(pair_df, self.pair_alert_threshold):
            self.diff_status_label.setText(
                f"Alert: the selected pair differs by more than {self.pair_alert_threshold:.1f} mmHg in peripheral systolic or diastolic pressure."
            )
            self.diff_status_label.setStyleSheet("color: #9a4d00; font-weight: 600;")
        else:
            self.diff_status_label.setText(
                "Selected pair is within the current peripheral systolic/diastolic alert threshold."
            )
            self.diff_status_label.setStyleSheet("")

    def _refresh_all_data_table(self) -> None:
        if not self.bundle:
            self.all_data_table.setRowCount(0)
            self.all_data_table.setColumnCount(0)
            return

        frame = display_dataframe(self.bundle).drop(columns=["Special Row"], errors="ignore")
        columns = [*COLUMNS, *EXTRA_COLUMNS]
        self._populate_dataframe_table(
            self.all_data_table,
            frame,
            columns,
            index_role_column=0,
        )

    def _refresh_averaged_table(self) -> None:
        if not self.bundle or self.bundle.analyzed_df.empty:
            self.averaged_table.setRowCount(0)
            self.averaged_table.setColumnCount(0)
            return

        frame = self.bundle.analyzed_df.drop(
            columns=["Record #", "Source Path", "Special Row"],
            errors="ignore",
        )
        self._populate_dataframe_table(
            self.averaged_table,
            frame,
            list(frame.columns),
        )

    def _populate_dataframe_table(
        self,
        table: QTableWidget,
        frame: pd.DataFrame,
        columns: list[str],
        index_role_column: Optional[int] = None,
    ) -> None:
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)
        table.setRowCount(len(frame))

        for table_row, (frame_index, row) in enumerate(frame.iterrows()):
            for column_index, column_name in enumerate(columns):
                item = QTableWidgetItem(format_value(row.get(column_name)))
                if index_role_column is not None and column_index == index_role_column:
                    item.setData(Qt.ItemDataRole.UserRole, frame_index)
                if column_name == "Analyed" and row.get(column_name) == "Yes":
                    item.setBackground(QColor("#dff1ea"))
                table.setItem(table_row, column_index, item)

        self._apply_data_table_widths(table, columns)
        if table.rowCount() > 0:
            table.selectRow(0)

    def selected_frame_index(self, table: QTableWidget) -> Optional[int]:
        row = table.currentRow()
        if row < 0 or table.columnCount() == 0:
            return None
        item = table.item(row, 0)
        if item is None:
            return None
        return item.data(Qt.ItemDataRole.UserRole)

    def frame_pdf_path(self, frame_index: int) -> Optional[Path]:
        if not self.bundle or frame_index not in self.bundle.dataframe.index:
            return None
        source_path = self.bundle.dataframe.loc[frame_index].get("Source Path")
        if not source_path:
            return None
        path = Path(str(source_path))
        return path if path.exists() else None

    def open_pdf(self, pdf_path: Path) -> None:
        try:
            viewer = PdfViewerDialog(pdf_path, self)
        except Exception as exc:
            QMessageBox.critical(
                self,
                "PDF preview failed",
                f"Could not open {pdf_path.name}.\n\n{exc}",
            )
            return

        self.pdf_viewers.append(viewer)
        viewer.finished.connect(
            lambda _result=0, dlg=viewer: self._release_pdf_viewer(dlg)
        )
        viewer.show()
        viewer.raise_()
        viewer.activateWindow()

    def _release_pdf_viewer(self, viewer: PdfViewerDialog) -> None:
        if viewer in self.pdf_viewers:
            self.pdf_viewers.remove(viewer)

    def open_selected_overview_pdf(self) -> None:
        frame_index = self.selected_frame_index(self.overview_table)
        if frame_index is None:
            return
        pdf_path = self.frame_pdf_path(frame_index)
        if pdf_path is not None:
            self.open_pdf(pdf_path)

    def open_selected_pair_pdf(self) -> None:
        row = self.pair_table.currentRow()
        if row < 0:
            return
        item = self.pair_table.item(row, 6)
        if item is None:
            return
        frame_index = item.data(Qt.ItemDataRole.UserRole)
        if frame_index is None:
            return
        pdf_path = self.frame_pdf_path(frame_index)
        if pdf_path is not None:
            self.open_pdf(pdf_path)

    def open_selected_all_data_pdf(self) -> None:
        frame_index = self.selected_frame_index(self.all_data_table)
        if frame_index is None:
            return
        pdf_path = self.frame_pdf_path(frame_index)
        if pdf_path is not None:
            self.open_pdf(pdf_path)

    def manual_review_complete(self) -> bool:
        if not self.bundle:
            return False
        for patient_id in self.bundle.manual_patients:
            if len(self.manual_pairs.get(patient_id, [])) != 2:
                return False
        return True

    def export_excel(self) -> None:
        if not self.records or not self.bundle:
            QMessageBox.warning(
                self,
                "Nothing to export",
                "Process PDFs before exporting.",
            )
            return
        if not self.manual_review_complete():
            QMessageBox.warning(
                self,
                "Review incomplete",
                "Each patient in the multi-entry review tab must have exactly two selected rows before export.",
            )
            self.tabs.setCurrentIndex(1)
            return

        manual_pair_tuples = {
            patient_id: tuple(selection)
            for patient_id, selection in self.manual_pairs.items()
            if len(selection) == 2
        }

        try:
            self.output_path.parent.mkdir(parents=True, exist_ok=True)
            exported_count = save_to_excel(
                self.records,
                self.output_path,
                manual_pairs=manual_pair_tuples,
                pair_alert_threshold=self.pair_alert_threshold,
            )
        except Exception as exc:
            QMessageBox.critical(
                self,
                "Export failed",
                f"Could not save the Excel workbook.\n\n{exc}",
            )
            return

        self.last_export_path = self.output_path
        self.open_output_button.setEnabled(True)
        self.status_label.setText(
            f"Saved {exported_count} row(s) to {self.output_path}"
        )
        QMessageBox.information(
            self,
            "Export complete",
            f"Saved {exported_count} row(s) to:\n{self.output_path}",
        )

    def open_export_folder(self) -> None:
        if not self.last_export_path:
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.last_export_path.parent)))

    def _set_processing_state(self, processing: bool) -> None:
        for widget in [
            self.add_files_button,
            self.remove_files_button,
            self.clear_files_button,
            self.browse_output_button,
            self.output_line,
            self.process_button,
            self.export_button,
            self.reset_auto_button,
            self.view_pair_pdf_button,
            self.green_max_spin,
            self.yellow_max_spin,
            self.pair_alert_spin,
        ]:
            widget.setEnabled(not processing)
        if processing:
            return
        self._sync_controls()

    def _sync_controls(self) -> None:
        has_files = bool(self.pdf_paths)
        has_output = bool(self.output_line.text().strip())
        is_processing = self.thread is not None
        has_records = bool(self.records)
        has_manual_patients = bool(self.bundle and self.bundle.manual_patients)
        manual_complete = self.manual_review_complete() if has_records else False

        self.remove_files_button.setEnabled(has_files and not is_processing)
        self.clear_files_button.setEnabled(has_files and not is_processing)
        self.process_button.setEnabled(has_files and has_output and not is_processing)
        self.export_button.setEnabled(
            has_records and has_output and manual_complete and not is_processing
        )
        self.open_output_button.setEnabled(
            self.last_export_path is not None and not is_processing
        )
        self.open_selected_overview_pdf_button.setEnabled(
            has_records and self.overview_table.currentRow() >= 0 and not is_processing
        )
        self.open_selected_all_data_pdf_button.setEnabled(
            has_records and self.all_data_table.currentRow() >= 0 and not is_processing
        )
        self.view_pair_pdf_button.setEnabled(
            has_manual_patients and self.pair_table.currentRow() >= 0 and not is_processing
        )
        self.reset_auto_button.setEnabled(has_manual_patients and not is_processing)

    def closeEvent(self, event) -> None:  # noqa: N802
        if self.thread is not None:
            self.thread.quit()
            self.thread.wait(3000)
        super().closeEvent(event)


def main() -> int:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
