from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple
from uuid import uuid4
import logging
import json

from PySide6.QtCore import QDate, QEventLoop, QTime, QTimer, Qt
from PySide6.QtGui import QColor, QPixmap, QCursor
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QProgressBar,
    QRadioButton,
    QSpinBox,
    QStackedWidget,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QTimeEdit,
    QVBoxLayout,
    QWidget,
    QFrame,
    QGraphicsDropShadowEffect,
    QHeaderView,
    QScrollArea,
    QToolTip,
    QWizard,
    QWizardPage,
)

from ..core.config import AppConfig, load_config, save_config, token_path
from ..core.points import MODIFIERS, PointsStore
from ..core.sanitise import (
    build_sanity_check,
    preprocess_lines,
    sanitize_line,
    slice_by_date,
    validate_lines,
    MULTI_NOT_MARKER,
)
from ..core.aliases import add_boss_alias, add_points_value
from ..core.workflow import (
    CalculationResult,
    Resolution,
    calculate_points,
    estimate_unknown_count,
)
from ..core.sheets import get_names_from_sheets
from ..core.runs import build_run_meta, iter_active_events, iso_to_dt, normalize_event, save_run


@dataclass
class WizardContext:
    base_dir: Path
    config: AppConfig
    timers_path: Path
    credentials_path: Path
    spreadsheet_id: str
    range_name: str
    use_all_entries: bool
    start_datetime: Optional[datetime]
    end_datetime: Optional[datetime]
    sanity_text: str = ""
    errors_text: str = ""
    calculation: Optional[CalculationResult] = None


@dataclass
class ErrorItem:
    kind: str
    line_index: int
    boss: Optional[str] = None


class SetupPage(QWizardPage):
    def __init__(self, context: WizardContext) -> None:
        super().__init__()
        self.context = context
        self.setTitle("")

        layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        self.run_tab = QWidget()
        run_layout = QVBoxLayout(self.run_tab)
        form = QFormLayout()

        self.timers_input = QLineEdit()
        self.timers_input.setText(context.config.last_timers_path)
        timers_button = QPushButton("Browse")
        timers_button.clicked.connect(self._browse_timers)
        timers_row = QHBoxLayout()
        timers_row.addWidget(self.timers_input)
        timers_row.addWidget(timers_button)
        form.addRow("Band export (timers.txt)", timers_row)

        self.credentials_input = QLineEdit()
        self.credentials_input.setText(context.config.last_credentials_path)
        cred_button = QPushButton("Browse")
        cred_button.clicked.connect(self._browse_credentials)
        cred_row = QHBoxLayout()
        cred_row.addWidget(self.credentials_input)
        cred_row.addWidget(cred_button)
        form.addRow("Google credentials.json", cred_row)

        self.sheet_input = QLineEdit()
        self.sheet_input.setText(context.config.spreadsheet_id)
        form.addRow("Spreadsheet ID", self.sheet_input)

        self.range_input = QLineEdit()
        self.range_input.setText(context.config.range_name)
        form.addRow("Range", self.range_input)

        self.date_input = QDateEdit()
        self.date_input.setCalendarPopup(True)
        self.time_input = QTimeEdit()
        self.time_input.setDisplayFormat("HH:mm")

        self.end_date_input = QDateEdit()
        self.end_date_input.setCalendarPopup(True)
        self.end_time_input = QTimeEdit()
        self.end_time_input.setDisplayFormat("HH:mm")

        if context.config.start_date_iso:
            try:
                start_dt = datetime.fromisoformat(context.config.start_date_iso)
                if start_dt.tzinfo:
                    start_dt = start_dt.astimezone(timezone.utc)
                self.date_input.setDate(QDate(start_dt.year, start_dt.month, start_dt.day))
                self.time_input.setTime(QTime(start_dt.hour, start_dt.minute))
            except ValueError:
                pass
        else:
            now_utc = datetime.now(timezone.utc)
            self.date_input.setDate(QDate(now_utc.year, now_utc.month, now_utc.day))
            self.time_input.setTime(QTime(0, 0))

        if context.config.end_date_iso:
            try:
                end_dt = datetime.fromisoformat(context.config.end_date_iso)
                if end_dt.tzinfo:
                    end_dt = end_dt.astimezone(timezone.utc)
                self.end_date_input.setDate(QDate(end_dt.year, end_dt.month, end_dt.day))
                self.end_time_input.setTime(QTime(end_dt.hour, end_dt.minute))
            except ValueError:
                pass
        else:
            default_end = datetime(
                self.date_input.date().year(),
                self.date_input.date().month(),
                self.date_input.date().day(),
                self.time_input.time().hour(),
                self.time_input.time().minute(),
                tzinfo=timezone.utc,
            ) + timedelta(days=6)
            self.end_date_input.setDate(
                QDate(default_end.year, default_end.month, default_end.day)
            )
            self.end_time_input.setTime(QTime(23, 59))

        date_row = QHBoxLayout()
        date_row.addWidget(self.date_input)
        date_row.addWidget(self.time_input)
        form.addRow("Start date (UTC)", date_row)

        end_date_row = QHBoxLayout()
        end_date_row.addWidget(self.end_date_input)
        end_date_row.addWidget(self.end_time_input)
        form.addRow("End date (UTC)", end_date_row)

        run_layout.addLayout(form)

        self.test_button = QPushButton("Test Google Sheets connection")
        self.test_button.clicked.connect(self._test_connection)
        self.test_spinner = QProgressBar()
        self.test_spinner.setObjectName("InlineSpinner")
        self.test_spinner.setRange(0, 0)
        self.test_spinner.setTextVisible(False)
        self.test_spinner.setFixedWidth(120)
        self.test_spinner.setVisible(False)
        self.test_status_indicator = QLabel("")
        self.test_status_indicator.setObjectName("StatusIndicator")
        self.test_status_indicator.setProperty("state", "idle")
        self.test_status_indicator.setAlignment(Qt.AlignCenter)
        self.test_status_indicator.setFixedSize(18, 18)
        self.test_status_text = QLabel("")
        self.test_status_text.setObjectName("TestStatusText")
        self.test_status_text.setWordWrap(True)
        test_row = QHBoxLayout()
        test_row.setSpacing(8)
        test_row.addWidget(self.test_button)
        test_row.addWidget(self.test_spinner)
        test_row.addWidget(self.test_status_indicator)
        test_row.addWidget(self.test_status_text, 1)
        run_layout.addLayout(test_row)

        note = QLabel("Date range uses UTC; timers.txt is local time and converted to UTC.")
        note.setStyleSheet("color: gray;")
        run_layout.addWidget(note)
        self.tabs.addTab(self.run_tab, "Setup")

        self.points_tab = QWidget()
        points_tab_layout = QVBoxLayout(self.points_tab)
        points_scroll = QScrollArea()
        points_scroll.setWidgetResizable(True)
        points_scroll.setFrameShape(QFrame.NoFrame)
        points_tab_layout.addWidget(points_scroll)

        points_container = QWidget()
        points_layout = QVBoxLayout(points_container)
        points_scroll.setWidget(points_container)
        self.points_loading = False
        self.points_last_snapshot = None
        self.points_undo_stack: List[dict] = []
        self.points_redo_stack: List[dict] = []

        regular_header = QLabel("Regular entries")
        regular_header.setObjectName("SubtleTitle")
        points_layout.addWidget(regular_header)

        self.points_table = QTableWidget()
        self.points_table.setColumnCount(2)
        self.points_table.setHorizontalHeaderLabels(["Boss Key", "Points"])
        self.points_table.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.points_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.points_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.points_table.verticalHeader().setVisible(False)
        points_header_view = self.points_table.horizontalHeader()
        points_header_view.setSectionResizeMode(0, QHeaderView.Interactive)
        points_header_view.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        points_header_view.setStretchLastSection(False)
        self.points_table.itemChanged.connect(self._on_points_changed)
        points_layout.addWidget(self.points_table)
        self.points_table.setColumnWidth(0, 240)

        regular_buttons = QHBoxLayout()
        self.points_add_button = QPushButton("Add entry")
        self.points_add_button.clicked.connect(self._add_points_row)
        regular_buttons.addWidget(self.points_add_button)
        self.points_remove_button = QPushButton("Remove selected")
        self.points_remove_button.clicked.connect(self._remove_points_row)
        regular_buttons.addWidget(self.points_remove_button)
        regular_buttons.addStretch(1)
        points_layout.addLayout(regular_buttons)

        legacy_header = QLabel("/legacy tiers")
        legacy_header.setObjectName("SubtleTitle")
        points_layout.addWidget(legacy_header)

        self.legacy_table = QTableWidget()
        self.legacy_table.setColumnCount(3)
        self.legacy_table.setHorizontalHeaderLabels(["Level", "5-star", "6-star"])
        self.legacy_table.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.legacy_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.legacy_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.legacy_table.verticalHeader().setVisible(False)
        legacy_header_view = self.legacy_table.horizontalHeader()
        legacy_header_view.setSectionResizeMode(0, QHeaderView.Interactive)
        legacy_header_view.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        legacy_header_view.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        legacy_header_view.setStretchLastSection(False)
        self.legacy_table.itemChanged.connect(self._on_points_changed)
        points_layout.addWidget(self.legacy_table)
        self.legacy_table.setColumnWidth(0, 140)

        legacy_buttons = QHBoxLayout()
        self.legacy_add_button = QPushButton("Add tier")
        self.legacy_add_button.clicked.connect(self._add_legacy_row)
        legacy_buttons.addWidget(self.legacy_add_button)
        self.legacy_remove_button = QPushButton("Remove selected tier")
        self.legacy_remove_button.clicked.connect(self._remove_legacy_row)
        legacy_buttons.addWidget(self.legacy_remove_button)
        legacy_buttons.addStretch(1)
        points_layout.addLayout(legacy_buttons)

        points_divider = QFrame()
        points_divider.setObjectName("Divider")
        points_divider.setFrameShape(QFrame.HLine)
        points_divider.setFrameShadow(QFrame.Sunken)
        points_layout.addWidget(points_divider)

        points_buttons = QHBoxLayout()
        self.points_undo_button = QPushButton("Undo")
        self.points_undo_button.clicked.connect(self._undo_points_change)
        points_buttons.addWidget(self.points_undo_button)
        self.points_redo_button = QPushButton("Redo")
        self.points_redo_button.clicked.connect(self._redo_points_change)
        points_buttons.addWidget(self.points_redo_button)
        self.points_reload_button = QPushButton("Reload points.json")
        self.points_reload_button.clicked.connect(self._load_points_json)
        points_buttons.addWidget(self.points_reload_button)
        self.points_save_button = QPushButton("Apply / Save")
        self.points_save_button.clicked.connect(self._save_points_json)
        points_buttons.addWidget(self.points_save_button)
        points_buttons.addStretch(1)
        points_layout.addLayout(points_buttons)

        self.points_status = QLabel("")
        self.points_status.setObjectName("ProgressLabel")
        points_layout.addWidget(self.points_status)

        points_layout.setStretchFactor(self.points_table, 2)
        points_layout.setStretchFactor(self.legacy_table, 1)
        self.points_table.setMinimumHeight(220)
        self.legacy_table.setMinimumHeight(160)
        self.chart_tab = QWidget()
        chart_layout = QVBoxLayout(self.chart_tab)
        chart_controls = QHBoxLayout()
        self.week_selector = QComboBox()
        self.week_selector.currentIndexChanged.connect(self._render_selected_week)
        chart_controls.addWidget(QLabel("Week"))
        chart_controls.addWidget(self.week_selector)
        chart_controls.addStretch(1)
        chart_controls.addWidget(QLabel("A ≥"))
        self.activity_a_input = QSpinBox()
        self.activity_a_input.setRange(0, 10000)
        self.activity_a_input.setValue(self.context.config.activity_a_threshold)
        self.activity_a_input.setKeyboardTracking(False)
        self.activity_a_input.setMinimumWidth(80)
        self.activity_a_input.editingFinished.connect(self._on_activity_thresholds_committed)
        chart_controls.addWidget(self.activity_a_input)
        chart_controls.addWidget(QLabel("A+ ≥"))
        self.activity_aplus_input = QSpinBox()
        self.activity_aplus_input.setRange(0, 10000)
        self.activity_aplus_input.setValue(self.context.config.activity_aplus_threshold)
        self.activity_aplus_input.setKeyboardTracking(False)
        self.activity_aplus_input.setMinimumWidth(80)
        self.activity_aplus_input.editingFinished.connect(self._on_activity_thresholds_committed)
        chart_controls.addWidget(self.activity_aplus_input)
        self.export_include_streaks = QCheckBox("Include A/A+ in export")
        self.export_include_streaks.setChecked(False)
        self.chart_refresh_button = QPushButton("Refresh")
        self.chart_refresh_button.clicked.connect(self._load_weekly_chart)
        chart_controls.addWidget(self.chart_refresh_button)
        chart_layout.addLayout(chart_controls)

        self.chart_status = QLabel("No saved runs yet.")
        self.chart_status.setObjectName("ProgressLabel")
        chart_layout.addWidget(self.chart_status)
        self.chart_table = QTableWidget()
        self.chart_table.setColumnCount(0)
        self.chart_table.setRowCount(0)
        self.chart_table.setSortingEnabled(False)
        self.chart_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.chart_table.setSelectionMode(QAbstractItemView.NoSelection)
        self.chart_table.verticalHeader().setVisible(False)
        self.chart_table.horizontalHeader().setStretchLastSection(True)
        chart_layout.addWidget(self.chart_table)

        export_row = QHBoxLayout()
        export_txt_button = QPushButton("Export weekly.txt")
        export_txt_button.clicked.connect(self._export_weekly_txt)
        export_row.addWidget(export_txt_button)
        export_csv_button = QPushButton("Export weekly.csv")
        export_csv_button.clicked.connect(self._export_weekly_csv)
        export_row.addWidget(export_csv_button)
        copy_button = QPushButton("Copy weekly to clipboard")
        copy_button.clicked.connect(self._copy_weekly_clipboard)
        export_row.addWidget(copy_button)
        self.export_include_streaks = QCheckBox("Include A/A+ in export")
        self.export_include_streaks.setChecked(False)
        export_row.addWidget(self.export_include_streaks)
        chart_layout.addLayout(export_row)
        self.tabs.addTab(self.chart_tab, "Weekly Chart")
        self.tabs.addTab(self.points_tab, "Points Editor")
        self.tabs.currentChanged.connect(self._on_tab_changed)

    def initializePage(self) -> None:
        wizard = self.wizard()
        if wizard:
            wizard.setButtonText(QWizard.NextButton, "Start DKP Wizard")
        self.tabs.setCurrentIndex(0)
        self._load_weekly_chart()
        self._load_points_json()

    def _browse_timers(self) -> None:
        try:
            start_dir = (
                str(Path(self.timers_input.text()).parent)
                if self.timers_input.text()
                else str(Path.cwd())
            )
            logging.info("Browse timers start dir: %s", start_dir)
            path = self._open_file_dialog(
                title="Select timers.txt",
                start_dir=start_dir,
                filter_text="Text Files (*.txt);;All Files (*)",
            )
            if path:
                self.timers_input.setText(path)
        except Exception as exc:
            QMessageBox.critical(self, "Browse failed", str(exc))

    def _browse_credentials(self) -> None:
        try:
            start_dir = (
                str(Path(self.credentials_input.text()).parent)
                if self.credentials_input.text()
                else str(Path.cwd())
            )
            logging.info("Browse credentials start dir: %s", start_dir)
            path = self._open_file_dialog(
                title="Select credentials.json",
                start_dir=start_dir,
                filter_text="JSON Files (*.json);;All Files (*)",
            )
            if path:
                self.credentials_input.setText(path)
        except Exception as exc:
            QMessageBox.critical(self, "Browse failed", str(exc))

    def _test_connection(self) -> None:
        try:
            self._set_test_status("working", "Testing connection...", show_spinner=True)
            spreadsheet_id = self.sheet_input.text().strip()
            range_name = self.range_input.text().strip()
            credentials_path = Path(self.credentials_input.text().strip())
            if not spreadsheet_id or not range_name or not credentials_path.exists():
                raise ValueError("Spreadsheet ID, range, and credentials.json are required.")

            token_file = token_path()
            names = get_names_from_sheets(
                spreadsheet_id=spreadsheet_id,
                range_name=range_name,
                credentials_path=credentials_path,
                token_path=token_file,
            )

            self._set_test_status(
                "ok",
                f"Loaded {len(names)} names from the sheet.",
                show_spinner=False,
            )
        except Exception as exc:
            self._set_test_status("error", str(exc), show_spinner=False)

    def _set_test_status(self, state: str, message: str, show_spinner: bool) -> None:
        self.test_button.setEnabled(not show_spinner)
        self.test_spinner.setVisible(show_spinner)
        if state == "ok":
            self.test_status_indicator.setText("OK")
        elif state == "error":
            self.test_status_indicator.setText("X")
        else:
            self.test_status_indicator.setText("")
        self.test_status_indicator.setProperty("state", state)
        self.test_status_text.setText(message)
        self.test_status_indicator.style().unpolish(self.test_status_indicator)
        self.test_status_indicator.style().polish(self.test_status_indicator)
        QApplication.processEvents()

    def _on_tab_changed(self, index: int) -> None:
        if self.tabs.widget(index) is self.chart_tab:
            self._load_weekly_chart()
        if self.tabs.widget(index) is self.points_tab:
            self._load_points_json()

    def _points_path(self) -> Path:
        return self.context.base_dir / "points.json"

    def _points_snapshot(self) -> dict:
        regular, rings, legacy = self._collect_points_data(allow_invalid=True)
        return {
            "regular": regular,
            "rings": rings,
            "legacy": legacy,
        }

    def _push_points_snapshot(self, previous: dict, current: dict) -> None:
        if previous == current:
            return
        self.points_undo_stack.append(previous)
        self.points_redo_stack.clear()
        self.points_last_snapshot = current
        self._update_points_undo_buttons()

    def _update_points_undo_buttons(self) -> None:
        self.points_undo_button.setEnabled(bool(self.points_undo_stack))
        self.points_redo_button.setEnabled(bool(self.points_redo_stack))

    def _on_points_changed(self) -> None:
        if self.points_loading:
            return
        previous = self.points_last_snapshot or self._points_snapshot()
        current = self._points_snapshot()
        self._push_points_snapshot(previous, current)

    def _undo_points_change(self) -> None:
        if not self.points_undo_stack:
            return
        current = self.points_last_snapshot or self._points_snapshot()
        snapshot = self.points_undo_stack.pop()
        self.points_redo_stack.append(current)
        self._load_points_snapshot(snapshot)
        self._update_points_undo_buttons()

    def _redo_points_change(self) -> None:
        if not self.points_redo_stack:
            return
        current = self.points_last_snapshot or self._points_snapshot()
        snapshot = self.points_redo_stack.pop()
        self.points_undo_stack.append(current)
        self._load_points_snapshot(snapshot)
        self._update_points_undo_buttons()

    def _load_points_snapshot(self, snapshot: dict) -> None:
        self.points_loading = True
        self._load_regular_points(
            snapshot.get("regular", []), snapshot.get("rings", {"5": 0, "6": 0})
        )
        self._load_legacy_points(snapshot.get("legacy", []))
        self.points_loading = False
        self.points_last_snapshot = self._points_snapshot()

    @staticmethod
    def _ring_row_label(star: int) -> str:
        return f"/rings {star}-star"

    @staticmethod
    def _ring_row_marker(star: int) -> str:
        return f"rings_star_{star}"

    def _load_regular_points(
        self, entries: List[Tuple[str, int]], rings: Dict[str, int]
    ) -> None:
        self.points_table.setRowCount(0)
        for star in (5, 6):
            row = self.points_table.rowCount()
            self.points_table.insertRow(row)
            key_item = QTableWidgetItem(self._ring_row_label(star))
            key_item.setFlags(key_item.flags() & ~Qt.ItemIsEditable)
            key_item.setData(Qt.UserRole, self._ring_row_marker(star))
            self.points_table.setItem(row, 0, key_item)
            value = int(rings.get(str(star), 0))
            value_item = QTableWidgetItem(str(value))
            value_item.setData(Qt.UserRole, self._ring_row_marker(star))
            self.points_table.setItem(row, 1, value_item)
        for boss, value in entries:
            row = self.points_table.rowCount()
            self.points_table.insertRow(row)
            self.points_table.setItem(row, 0, QTableWidgetItem(str(boss)))
            self.points_table.setItem(row, 1, QTableWidgetItem(str(value)))

    def _load_legacy_points(self, tiers: List[dict]) -> None:
        self.legacy_table.setRowCount(0)
        for tier in tiers:
            row = self.legacy_table.rowCount()
            self.legacy_table.insertRow(row)
            self.legacy_table.setItem(row, 0, QTableWidgetItem(str(tier.get("level", 0))))
            self.legacy_table.setItem(row, 1, QTableWidgetItem(str(tier.get("5", 0))))
            self.legacy_table.setItem(row, 2, QTableWidgetItem(str(tier.get("6", 0))))

    def _add_points_row(self) -> None:
        row = self.points_table.rowCount()
        self.points_table.insertRow(row)
        self.points_table.setItem(row, 0, QTableWidgetItem(""))
        self.points_table.setItem(row, 1, QTableWidgetItem("0"))
        self.points_table.setCurrentCell(row, 0)
        self._on_points_changed()

    def _remove_points_row(self) -> None:
        row = self.points_table.currentRow()
        if row >= 0:
            key_item = self.points_table.item(row, 0)
            if key_item and key_item.data(Qt.UserRole) in {
                self._ring_row_marker(5),
                self._ring_row_marker(6),
            }:
                QToolTip.showText(
                    QCursor.pos(),
                    "Ring rows cannot be removed because they control /rings points.",
                )
                self.points_status.setText("Ring rows cannot be removed.")
                return
            self.points_table.removeRow(row)
            self._on_points_changed()

    def _add_legacy_row(self) -> None:
        row = self.legacy_table.rowCount()
        self.legacy_table.insertRow(row)
        self.legacy_table.setItem(row, 0, QTableWidgetItem("0"))
        self.legacy_table.setItem(row, 1, QTableWidgetItem("0"))
        self.legacy_table.setItem(row, 2, QTableWidgetItem("0"))
        self.legacy_table.setCurrentCell(row, 0)
        self._on_points_changed()

    def _remove_legacy_row(self) -> None:
        row = self.legacy_table.currentRow()
        if row >= 0:
            self.legacy_table.removeRow(row)
            self._on_points_changed()

    def _collect_points_data(
        self, allow_invalid: bool = False
    ) -> Tuple[List[Tuple[str, int]], Dict[str, int], List[dict]]:
        regular: List[Tuple[str, int]] = []
        seen: Set[str] = set()
        rings = {"5": 0, "6": 0}
        for row in range(self.points_table.rowCount()):
            boss_item = self.points_table.item(row, 0)
            points_item = self.points_table.item(row, 1)
            boss = boss_item.text().strip() if boss_item else ""
            points_text = points_item.text().strip() if points_item else ""
            marker = boss_item.data(Qt.UserRole) if boss_item else None
            if marker in {self._ring_row_marker(5), self._ring_row_marker(6)}:
                star = "5" if marker == self._ring_row_marker(5) else "6"
                try:
                    value = int(points_text)
                except ValueError:
                    if allow_invalid:
                        continue
                    raise ValueError("Ring points must be an integer.")
                if value < 0 or value > 999:
                    if allow_invalid:
                        continue
                    raise ValueError("Ring points must be 0-999.")
                rings[star] = value
                continue
            if not boss and allow_invalid:
                continue
            if ":" in boss:
                if allow_invalid:
                    continue
                raise ValueError("Boss key cannot contain ':'")
            if boss in seen:
                if allow_invalid:
                    continue
                raise ValueError(f"Duplicate boss key: {boss}")
            seen.add(boss)
            try:
                value = int(points_text)
            except ValueError:
                if allow_invalid:
                    continue
                raise ValueError(f"Points must be an integer for {boss or '(blank)'}")
            if value < 0 or value > 999:
                if allow_invalid:
                    continue
                raise ValueError(f"Points must be 0-999 for {boss or '(blank)'}")
            if boss:
                regular.append((boss, value))

        legacy: List[dict] = []
        legacy_levels: Set[int] = set()
        for row in range(self.legacy_table.rowCount()):
            level_item = self.legacy_table.item(row, 0)
            five_item = self.legacy_table.item(row, 1)
            six_item = self.legacy_table.item(row, 2)
            level_text = level_item.text().strip() if level_item else ""
            five_text = five_item.text().strip() if five_item else ""
            six_text = six_item.text().strip() if six_item else ""
            if allow_invalid and not level_text and not five_text and not six_text:
                continue
            try:
                level = int(level_text or "0")
                five_val = int(five_text or "0")
                six_val = int(six_text or "0")
            except ValueError:
                if allow_invalid:
                    continue
                raise ValueError("Legacy tiers must be integers.")
            if five_val < 0 or five_val > 999 or six_val < 0 or six_val > 999:
                if allow_invalid:
                    continue
                raise ValueError("Legacy tier points must be 0-999.")
            if level < 0:
                if allow_invalid:
                    continue
                raise ValueError("Legacy level must be 0 or greater.")
            if level in legacy_levels:
                if allow_invalid:
                    continue
                raise ValueError(f"Duplicate legacy level: {level}")
            legacy_levels.add(level)
            legacy.append({"level": level, "5": five_val, "6": six_val})

        legacy.sort(key=lambda item: item["level"], reverse=True)
        return regular, rings, legacy

    def _load_points_json(self) -> None:
        path = self._points_path()
        try:
            if not path.exists():
                self.points_loading = True
                self._load_regular_points([], {"5": 0, "6": 0})
                self._load_legacy_points([])
                self.points_loading = False
                self.points_status.setText("points.json not found.")
                self.points_undo_stack.clear()
                self.points_redo_stack.clear()
                self.points_last_snapshot = self._points_snapshot()
                self._update_points_undo_buttons()
                return
            raw = json.loads(path.read_text(encoding="utf-8"))
            regular = []
            rings = {"5": 0, "6": 0}
            legacy: List[dict] = []
            for key, value in raw.items():
                if key == "/rings" and isinstance(value, dict):
                    rings["5"] = int(value.get("5", 0))
                    rings["6"] = int(value.get("6", 0))
                elif key == "/legacy":
                    if isinstance(value, list):
                        legacy = [
                            {"level": int(item.get("level", 0)),
                             "5": int(item.get("5", 0)),
                             "6": int(item.get("6", 0))}
                            for item in value
                        ]
                    elif isinstance(value, int):
                        legacy = [{"level": 0, "5": int(value), "6": int(value)}]
                else:
                    if isinstance(value, int):
                        regular.append((key, int(value)))
            regular.sort(key=lambda item: item[0].lower())
            self.points_loading = True
            self._load_regular_points(regular, rings)
            self._load_legacy_points(legacy)
            self.points_loading = False
            self.points_status.setText(f"Loaded {path.name}.")
            self.points_undo_stack.clear()
            self.points_redo_stack.clear()
            self.points_last_snapshot = self._points_snapshot()
            self._update_points_undo_buttons()
        except Exception as exc:
            self.points_status.setText(f"Failed to load points.json: {exc}")

    def _save_points_json(self) -> None:
        path = self._points_path()
        try:
            regular, rings, legacy = self._collect_points_data()
            payload: Dict[str, object] = {}
            for key, value in sorted(regular, key=lambda item: item[0].lower()):
                payload[key] = value
            payload["/rings"] = {"5": rings["5"], "6": rings["6"]}
            payload["/legacy"] = legacy
            formatted = json.dumps(payload, indent=2, ensure_ascii=True)
            path.write_text(formatted + "\n", encoding="utf-8")
            self.points_status.setText("Saved points.json.")
        except Exception as exc:
            self.points_status.setText(f"Failed to save points.json: {exc}")

    def _on_activity_thresholds_committed(self) -> None:
        a_value = self.activity_a_input.value()
        aplus_value = self.activity_aplus_input.value()
        if aplus_value < a_value:
            self.activity_aplus_input.blockSignals(True)
            self.activity_aplus_input.setValue(a_value)
            self.activity_aplus_input.blockSignals(False)
            aplus_value = a_value
        self.context.config.activity_a_threshold = a_value
        self.context.config.activity_aplus_threshold = aplus_value
        save_config(self.context.config)
        self._load_weekly_chart()

    @staticmethod
    def _week_start_utc(dt: datetime) -> datetime:
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        dt = dt.astimezone(timezone.utc)
        days_since_sunday = (dt.weekday() + 1) % 7
        week_start = dt - timedelta(days=days_since_sunday)
        return week_start.replace(hour=0, minute=0, second=0, microsecond=0)

    def _load_weekly_chart(self) -> None:
        events = iter_active_events(self.context.base_dir)
        if not events:
            self.chart_status.setText("No saved runs yet.")
            self.chart_table.setRowCount(0)
            self.chart_table.setColumnCount(0)
            self.chart_table.setVisible(False)
            self.week_selector.blockSignals(True)
            self.week_selector.clear()
            self.week_selector.blockSignals(False)
            return

        def normalize_boss_key(raw_boss: str) -> str:
            cleaned = raw_boss.strip()
            if "(" in cleaned and cleaned.endswith(")"):
                cleaned = cleaned.split("(", 1)[0]
            if cleaned.startswith("/"):
                cleaned = cleaned[1:]
            return cleaned

        weekly: Dict[datetime, Dict[str, Dict[str, Dict[str, int]]]] = {}
        boss_set: Set[str] = set()

        for event in events:
            event_time_raw = event.get("event_time_utc")
            if not event_time_raw:
                continue
            event_time = iso_to_dt(event_time_raw)
            week_start = self._week_start_utc(event_time)
            bucket = weekly.setdefault(week_start, {})
            boss = event.get("boss", "")
            boss_key = normalize_boss_key(boss) if boss else ""
            has_positive = False
            for entry in event.get("entries", []):
                name = entry.get("name", "")
                delta = int(entry.get("delta", 0))
                if not name:
                    continue
                player = bucket.setdefault(name, {"dkp": 0, "boss_counts": {}})
                player["dkp"] += delta
                if boss_key:
                    boss_counts = player["boss_counts"]
                    current = boss_counts.get(boss_key, 0)
                    if delta > 0:
                        boss_counts[boss_key] = current + 1
                        has_positive = True
                    elif delta < 0 and current > 0:
                        new_value = current - 1
                        if new_value > 0:
                            boss_counts[boss_key] = new_value
                        else:
                            boss_counts.pop(boss_key, None)
            if boss_key and has_positive:
                boss_set.add(boss_key)

        boss_list = sorted(boss_set, key=str.lower)
        weeks = sorted(weekly.keys())
        self._weekly_data = weekly
        self._boss_list = boss_list
        self._weeks = weeks

        current_value = self.week_selector.currentData()
        self.week_selector.blockSignals(True)
        self.week_selector.clear()
        for week_start in weeks:
            week_end = week_start + timedelta(days=6)
            label = f"{week_start.date()} to {week_end.date()} (UTC)"
            self.week_selector.addItem(label, week_start.isoformat())
        self.week_selector.blockSignals(False)
        if current_value and current_value in [w.isoformat() for w in weeks]:
            index = [w.isoformat() for w in weeks].index(current_value)
            self.week_selector.setCurrentIndex(index)
        else:
            self.week_selector.setCurrentIndex(len(weeks) - 1)

        self._render_selected_week()

    def _compute_streaks(self) -> Dict[str, Dict[str, int]]:
        weekly = getattr(self, "_weekly_data", {})
        weeks = getattr(self, "_weeks", [])
        if not weekly or not weeks:
            return {}

        min_week = min(weeks)
        max_week = max(weeks)
        full_weeks: List[datetime] = []
        current = min_week
        while current <= max_week:
            full_weeks.append(current)
            current = current + timedelta(days=7)

        a_threshold = self.activity_a_input.value()
        aplus_threshold = self.activity_aplus_input.value()
        players: Set[str] = set()
        for week_players in weekly.values():
            players.update(week_players.keys())

        streaks: Dict[str, Dict[str, int]] = {}
        for player in players:
            a_streak = 0
            aplus_streak = 0
            for week_start in reversed(full_weeks):
                dkp = weekly.get(week_start, {}).get(player, {}).get("dkp", 0)
                if dkp >= a_threshold:
                    a_streak += 1
                else:
                    break
            for week_start in reversed(full_weeks):
                dkp = weekly.get(week_start, {}).get(player, {}).get("dkp", 0)
                if dkp >= aplus_threshold:
                    aplus_streak += 1
                else:
                    break
            streaks[player] = {"a": a_streak, "aplus": aplus_streak}
        return streaks

    def _render_selected_week(self) -> None:
        weekly = getattr(self, "_weekly_data", {})
        boss_list = getattr(self, "_boss_list", [])
        weeks = getattr(self, "_weeks", [])
        if not weekly or not weeks or self.week_selector.currentIndex() < 0:
            self.chart_status.setText("No saved runs yet.")
            self.chart_table.setRowCount(0)
            self.chart_table.setColumnCount(0)
            self.chart_table.setVisible(False)
            return

        selected_iso = self.week_selector.currentData()
        if not selected_iso:
            return
        selected_week = datetime.fromisoformat(selected_iso)
        all_players: Set[str] = set()
        for week_players in weekly.values():
            all_players.update(week_players.keys())
        if not all_players:
            self.chart_status.setText("No data for this week.")
            self.chart_table.setRowCount(0)
            self.chart_table.setColumnCount(0)
            self.chart_table.setVisible(False)
            return

        streaks = self._compute_streaks()
        columns = ["Player", "Weekly DKP", "A Streak", "A+ Streak"] + boss_list
        rows = []
        selected_players = weekly.get(selected_week, {})
        for player_name in sorted(all_players, key=str.lower):
            player_data = selected_players.get(player_name, {"dkp": 0, "boss_counts": {}})
            rows.append(
                {
                    "player": player_name,
                    "dkp": player_data["dkp"],
                    "boss_counts": player_data["boss_counts"],
                    "a_streak": streaks.get(player_name, {}).get("a", 0),
                    "aplus_streak": streaks.get(player_name, {}).get("aplus", 0),
                }
            )

        week_end = selected_week + timedelta(days=6)
        self.chart_status.setText(
            f"Week: {selected_week.date()} to {week_end.date()} (UTC) | {len(rows)} players"
        )
        self.chart_table.setVisible(True)
        self.chart_table.setSortingEnabled(False)
        self.chart_table.setRowCount(len(rows))
        self.chart_table.setColumnCount(len(columns))
        self.chart_table.setHorizontalHeaderLabels(columns)

        for row_idx, row in enumerate(rows):
            self.chart_table.setItem(row_idx, 0, QTableWidgetItem(str(row["player"])))
            dkp_item = QTableWidgetItem()
            dkp_item.setData(Qt.DisplayRole, int(row["dkp"]))
            self.chart_table.setItem(row_idx, 1, dkp_item)
            a_item = QTableWidgetItem()
            a_item.setData(Qt.DisplayRole, int(row["a_streak"]))
            self.chart_table.setItem(row_idx, 2, a_item)
            aplus_item = QTableWidgetItem()
            aplus_item.setData(Qt.DisplayRole, int(row["aplus_streak"]))
            self.chart_table.setItem(row_idx, 3, aplus_item)
            counts = row["boss_counts"]
            for col_idx, boss in enumerate(boss_list, start=4):
                value = int(counts.get(boss, 0))
                item = QTableWidgetItem()
                item.setData(Qt.DisplayRole, value)
                self.chart_table.setItem(row_idx, col_idx, item)

        self.chart_table.resizeColumnsToContents()
        self.chart_table.resizeRowsToContents()
        self.chart_table.setSortingEnabled(True)

    def _selected_week_rows(self, include_streaks: bool) -> (List[str], List[List[str]]):
        weekly = getattr(self, "_weekly_data", {})
        weeks = getattr(self, "_weeks", [])
        if not weekly or not weeks or self.week_selector.currentIndex() < 0:
            return [], []
        selected_iso = self.week_selector.currentData()
        if not selected_iso:
            return [], []
        selected_week = datetime.fromisoformat(selected_iso)

        all_players: Set[str] = set()
        for week_players in weekly.values():
            all_players.update(week_players.keys())
        if not all_players:
            return [], []

        headers = ["Player", "Weekly DKP"]
        streaks = {}
        if include_streaks:
            headers += ["A Streak", "A+ Streak"]
            streaks = self._compute_streaks()

        rows: List[List[str]] = []
        selected_players = weekly.get(selected_week, {})
        for player_name in sorted(all_players, key=str.lower):
            player_data = selected_players.get(player_name, {"dkp": 0})
            row = [player_name, str(int(player_data.get("dkp", 0)))]
            if include_streaks:
                streak = streaks.get(player_name, {"a": 0, "aplus": 0})
                row.append(str(int(streak.get("a", 0))))
                row.append(str(int(streak.get("aplus", 0))))
            rows.append(row)

        return headers, rows

    def _export_weekly_txt(self) -> None:
        headers, rows = self._selected_week_rows(self.export_include_streaks.isChecked())
        if not rows:
            QMessageBox.information(self, "No data", "No weekly data to export.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save weekly output", "weekly.txt", "Text Files (*.txt)"
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            for row in rows:
                f.write(", ".join(row) + "\n")

    def _export_weekly_csv(self) -> None:
        headers, rows = self._selected_week_rows(self.export_include_streaks.isChecked())
        if not rows:
            QMessageBox.information(self, "No data", "No weekly data to export.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save weekly CSV", "weekly.csv", "CSV Files (*.csv)"
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write(",".join(headers) + "\n")
            for row in rows:
                f.write(",".join(row) + "\n")

    def _copy_weekly_clipboard(self) -> None:
        headers, rows = self._selected_week_rows(self.export_include_streaks.isChecked())
        if not rows:
            QMessageBox.information(self, "No data", "No weekly data to copy.")
            return
        lines = [", ".join(headers)] + [", ".join(row) for row in rows]
        QApplication.clipboard().setText("\n".join(lines))

    def validatePage(self) -> bool:
        timers_path = Path(self.timers_input.text().strip())
        credentials_path = Path(self.credentials_input.text().strip())
        spreadsheet_id = self.sheet_input.text().strip()
        range_name = self.range_input.text().strip()

        if not timers_path.exists():
            QMessageBox.critical(self, "Missing file", "Please select a valid timers.txt file.")
            return False
        if not credentials_path.exists():
            QMessageBox.critical(
                self, "Missing file", "Please select a valid credentials.json file."
            )
            return False
        if not spreadsheet_id or not range_name:
            QMessageBox.critical(self, "Missing data", "Spreadsheet ID and range are required.")
            return False

        date = self.date_input.date()
        time = self.time_input.time()
        start_dt = datetime(
            date.year(),
            date.month(),
            date.day(),
            time.hour(),
            time.minute(),
            tzinfo=timezone.utc,
        )
        end_date = self.end_date_input.date()
        end_time = self.end_time_input.time()
        end_dt = datetime(
            end_date.year(),
            end_date.month(),
            end_date.day(),
            end_time.hour(),
            end_time.minute(),
            tzinfo=timezone.utc,
        )
        if end_dt < start_dt:
            QMessageBox.critical(
                self, "Invalid date range", "End date must be after the start date."
            )
            return False

        self.context.timers_path = timers_path
        self.context.credentials_path = credentials_path
        self.context.spreadsheet_id = spreadsheet_id
        self.context.range_name = range_name
        self.context.use_all_entries = False
        self.context.start_datetime = start_dt
        self.context.end_datetime = end_dt

        cfg = self.context.config
        cfg.last_timers_path = str(timers_path)
        cfg.last_credentials_path = str(credentials_path)
        cfg.spreadsheet_id = spreadsheet_id
        cfg.range_name = range_name
        cfg.use_all_entries = False
        cfg.start_date_iso = start_dt.isoformat() if start_dt else ""
        cfg.end_date_iso = end_dt.isoformat() if end_dt else ""
        cfg.use_native_dialog = True
        save_config(cfg)

        return True

    def _open_file_dialog(self, title: str, start_dir: str, filter_text: str) -> str:
        parent = self.window()
        logging.info("Dialog parent=%s winId=%s", type(parent).__name__, int(parent.winId()))

        dialog = QFileDialog(parent, title, start_dir, filter_text)
        dialog.setFileMode(QFileDialog.ExistingFile)
        dialog.setOption(QFileDialog.DontUseCustomDirectoryIcons, True)

        dialog.setOption(QFileDialog.DontUseNativeDialog, False)

        if dialog.exec():
            selected = dialog.selectedFiles()
            return selected[0] if selected else ""
        return ""


class SanityCheckPage(QWizardPage):
    def __init__(self, context: WizardContext) -> None:
        super().__init__()
        self.context = context
        self.setTitle("Fix line errors")
        layout = QVBoxLayout(self)

        self.summary = QLabel("Parsing timers...")
        self.summary.setObjectName("SectionTitle")
        layout.addWidget(self.summary)

        self.ok_label = QLabel("All checks passed. Click Next to continue.")
        self.ok_label.setObjectName("ProgressLabel")
        self.ok_label.setVisible(False)
        layout.addWidget(self.ok_label)

        self.progress_label = QLabel("Fixed 0 / 0")
        self.progress_label.setObjectName("ProgressLabel")
        layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)

        self.fix_panel = QFrame()
        self.fix_panel.setObjectName("FixPanel")
        shadow = QGraphicsDropShadowEffect(self.fix_panel)
        shadow.setBlurRadius(18)
        shadow.setOffset(0, 4)
        shadow.setColor(QColor(0, 0, 0, 160))
        self.fix_panel.setGraphicsEffect(shadow)
        fix_layout = QVBoxLayout(self.fix_panel)
        fix_layout.setContentsMargins(12, 12, 12, 12)
        fix_layout.setSpacing(10)

        self.error_header = QLabel("")
        self.error_header.setObjectName("ErrorHeader")
        self.error_header.setWordWrap(True)
        fix_layout.addWidget(self.error_header)

        self.context_view = QTextEdit()
        self.context_view.setObjectName("ContextView")
        self.context_view.setReadOnly(True)
        self.context_view.setMinimumHeight(120)
        fix_layout.addWidget(self.context_view)

        self.options_label = QLabel("Fix Options")
        self.options_label.setObjectName("SectionTitle")
        fix_layout.addWidget(self.options_label)

        self.stack = QStackedWidget()
        fix_layout.addWidget(self.stack)

        self.retype_page = QWidget()
        retype_layout = QVBoxLayout(self.retype_page)
        self.retype_mode_widget = QWidget()
        retype_mode_layout = QHBoxLayout(self.retype_mode_widget)
        retype_mode_layout.setContentsMargins(0, 0, 0, 0)
        self.retype_mode_edit = QRadioButton("Edit entry")
        self.retype_mode_multinot = QRadioButton("Use multi-not")
        self.retype_mode_group = QButtonGroup(self)
        self.retype_mode_group.addButton(self.retype_mode_edit)
        self.retype_mode_group.addButton(self.retype_mode_multinot)
        self.retype_mode_edit.setChecked(True)
        self.retype_mode_widget.setVisible(False)
        self.retype_mode_edit.toggled.connect(self._update_retype_inputs)
        self.retype_mode_multinot.toggled.connect(self._update_retype_inputs)
        retype_mode_layout.addWidget(self.retype_mode_edit)
        retype_mode_layout.addWidget(self.retype_mode_multinot)
        retype_layout.addWidget(self.retype_mode_widget)

        self.retype_label = QLabel("Edit the entry below (date/time preserved)")
        self.retype_label.setObjectName("SectionTitle")
        retype_layout.addWidget(self.retype_label)
        self.retype_input = QPlainTextEdit()
        self.retype_input.setObjectName("RetypeInput")
        self.retype_input.setPlaceholderText("Edit the entry (no date/prefix)...")
        self.retype_input.setMinimumHeight(120)
        retype_layout.addWidget(self.retype_input)
        self.stack.addWidget(self.retype_page)

        self.boss_page = QWidget()
        boss_layout = QVBoxLayout(self.boss_page)
        self.boss_map_radio = QRadioButton("Map to existing boss")
        self.boss_line_only_radio = QRadioButton("Fix this line only (no alias saved)")
        self.boss_add_radio = QRadioButton("Add new boss with points")
        self.boss_map_radio.setChecked(True)
        self.boss_map_radio.toggled.connect(self._update_boss_inputs)
        self.boss_line_only_radio.toggled.connect(self._update_boss_inputs)
        self.boss_add_radio.toggled.connect(self._update_boss_inputs)
        boss_layout.addWidget(self.boss_map_radio)
        boss_layout.addWidget(self.boss_line_only_radio)
        self.boss_combo = QComboBox()
        boss_layout.addWidget(self.boss_combo)
        boss_layout.addWidget(self.boss_add_radio)
        self.boss_token_input = QLineEdit()
        self.boss_token_input.setPlaceholderText("New boss token")
        boss_layout.addWidget(self.boss_token_input)
        self.boss_points = QSpinBox()
        self.boss_points.setRange(-1000, 10000)
        self.boss_points.setValue(0)
        boss_layout.addWidget(self.boss_points)
        self.boss_manual_radio = QRadioButton("Edit line manually")
        self.boss_manual_radio.toggled.connect(self._update_boss_inputs)
        boss_layout.addWidget(self.boss_manual_radio)
        self.boss_manual_input = QPlainTextEdit()
        self.boss_manual_input.setPlaceholderText("Edit the full line here...")
        self.boss_manual_input.setMinimumHeight(90)
        self.boss_manual_input.setVisible(False)
        boss_layout.addWidget(self.boss_manual_input)
        self.stack.addWidget(self.boss_page)

        self.single_page = QWidget()
        single_layout = QVBoxLayout(self.single_page)
        self.single_token_combo = QComboBox()
        single_layout.addWidget(self.single_token_combo)
        self.single_replace_radio = QRadioButton("Replace with name")
        self.single_remove_radio = QRadioButton("Remove token (ignore)")
        self.single_replace_radio.setChecked(True)
        self.single_replace_radio.toggled.connect(self._update_single_inputs)
        self.single_remove_radio.toggled.connect(self._update_single_inputs)
        single_layout.addWidget(self.single_replace_radio)
        self.single_replace_input = QLineEdit()
        self.single_replace_input.setPlaceholderText("Replacement name")
        single_layout.addWidget(self.single_replace_input)
        single_layout.addWidget(self.single_remove_radio)
        self.stack.addWidget(self.single_page)

        self.boss_or_not_controls = QWidget()
        boss_or_not_layout = QHBoxLayout(self.boss_or_not_controls)
        self.choice_boss_radio = QRadioButton("Treat as boss error")
        self.choice_not_radio = QRadioButton("Treat as NOT error")
        self.choice_boss_radio.setChecked(True)
        self.choice_boss_radio.toggled.connect(self._on_boss_or_not_choice_changed)
        self.choice_not_radio.toggled.connect(self._on_boss_or_not_choice_changed)
        boss_or_not_layout.addWidget(self.choice_boss_radio)
        boss_or_not_layout.addWidget(self.choice_not_radio)
        self.boss_or_not_controls.setVisible(False)
        fix_layout.addWidget(self.boss_or_not_controls)

        buttons_row = QHBoxLayout()
        self.apply_button = QPushButton("Apply Fix")
        self.apply_button.clicked.connect(self._apply_fix)
        buttons_row.addWidget(self.apply_button)
        self.skip_button = QPushButton("Skip / Discard Line")
        self.skip_button.clicked.connect(self._skip_line)
        buttons_row.addWidget(self.skip_button)
        fix_layout.addLayout(buttons_row)

        layout.addWidget(self.fix_panel)

        self._line_map: Dict[int, str] = {}
        self._raw_line_map: Dict[int, str] = {}
        self._overrides: Dict[int, Optional[str]] = {}
        self._error_items: List[ErrorItem] = []
        self._current_error: Optional[ErrorItem] = None
        self._initial_total = 0
        self._source_key = None
        self._bosses: List[str] = []
        self._backup_created = False
        self._update_boss_inputs()
        self._update_single_inputs()
        self._update_retype_inputs()

    def reset_state(self) -> None:
        self._line_map = {}
        self._raw_line_map = {}
        self._overrides = {}
        self._error_items = []
        self._current_error = None
        self._initial_total = 0
        self._source_key = None
        self._bosses = []
        self._backup_created = False
        self.complete = False
        self.summary.setText("Parsing timers...")
        self.ok_label.setVisible(False)
        self.fix_panel.setVisible(False)
        self.progress_bar.setMaximum(1)
        self.progress_bar.setValue(0)
        self.progress_label.setText("Fixed 0 / 0")

    def initializePage(self) -> None:
        wizard = self.wizard()
        if wizard:
            wizard.setButtonText(QWizard.NextButton, "Next")

        current_key = (
            str(self.context.timers_path),
            self.context.start_datetime,
            self.context.use_all_entries,
        )

        if self._source_key != current_key:
            self._source_key = current_key
            self._overrides = {}
            self._initial_total = 0

        self._revalidate()

    def isComplete(self) -> bool:
        return getattr(self, "complete", True)

    def _format_errors(self, errors) -> str:
        if not errors.any():
            return "No validation errors found."

        parts = []
        if errors.date_lines:
            parts.append("Cannot read date in lines: " + ", ".join(map(str, errors.date_lines)))
        if errors.boss_lines:
            parts.append("Cannot read boss in lines: " + ", ".join(map(str, errors.boss_lines)))
        if errors.at_lines:
            parts.append("Word 'at' in lines: " + ", ".join(map(str, errors.at_lines)))
        if errors.incorrect_not_lines:
            parts.append(
                "Incorrect use of 'not' in lines: "
                + ", ".join(map(str, errors.incorrect_not_lines))
            )
        if errors.ambiguous_not_boss_lines:
            parts.append(
                "Boss or NOT? Review lines: "
                + ", ".join(map(str, errors.ambiguous_not_boss_lines))
            )
        if errors.single_char_lines:
            parts.append(
                "Single character name in lines: "
                + ", ".join(map(str, errors.single_char_lines))
            )
        if errors.general_lines:
            parts.append("Error at lines: " + ", ".join(map(str, errors.general_lines)))

        return "\n".join(parts)

    def _revalidate(self) -> None:
        points_store = PointsStore(self.context.base_dir)
        self._bosses = sorted(
            k for k, v in points_store.points_map.items() if isinstance(v, int)
        )

        lines, line_map = self._build_lines()
        self._line_map = line_map
        sanity = build_sanity_check(lines)
        _, errors = validate_lines(lines, points_store)

        summary_text = f"Total lines: {sanity.total_lines}"
        self.summary.setText(summary_text)

        errors_text = self._format_errors(errors)

        self.context.sanity_text = summary_text
        self.context.errors_text = errors_text

        self._error_items = self._build_error_queue(errors)
        remaining = len(self._error_items)

        if self._initial_total == 0:
            self._initial_total = remaining
        elif remaining > self._initial_total:
            self._initial_total = remaining

        total = self._initial_total
        if total == 0:
            self.progress_bar.setMaximum(1)
            self.progress_bar.setValue(1)
            self.progress_label.setText("Fixed 0 / 0")
        else:
            fixed = max(total - remaining, 0)
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(fixed)
            self.progress_label.setText(f"Fixed {fixed} / {total}")

        self.complete = remaining == 0
        self.ok_label.setVisible(self.complete)
        self.fix_panel.setVisible(not self.complete)
        if self._error_items:
            self._set_current_error(self._error_items[0])

        self.completeChanged.emit()

    def _build_error_queue(self, errors) -> List[ErrorItem]:
        boss_by_line: Dict[int, str] = {}
        for boss, line_nums in errors.unknown_bosses.items():
            for line_num in line_nums:
                boss_by_line[line_num] = boss

        items: List[ErrorItem] = []
        for line in errors.date_lines:
            items.append(ErrorItem("date", line))
        for line in errors.boss_lines:
            items.append(ErrorItem("boss", line, boss_by_line.get(line)))
        for line in errors.incorrect_not_lines:
            items.append(ErrorItem("not", line))
        for line in errors.ambiguous_not_boss_lines:
            items.append(ErrorItem("boss_or_not", line, boss_by_line.get(line)))
        for line in errors.at_lines:
            items.append(ErrorItem("at", line))
        for line in errors.single_char_lines:
            items.append(ErrorItem("single_char", line))
        for line in errors.general_lines:
            items.append(ErrorItem("general", line))

        items.sort(key=lambda item: item.line_index)
        return items

    def _build_lines(self) -> (List[tuple], Dict[int, str]):
        lines = preprocess_lines(self.context.timers_path, self.context.base_dir)
        line_map = {idx: line for idx, line in lines}
        raw_line_map: Dict[int, str] = {}
        if self.context.timers_path.exists():
            try:
                with self.context.timers_path.open(
                    "r", encoding="utf-8", errors="ignore"
                ) as f:
                    raw_lines = f.read().splitlines()
                raw_line_map = {
                    idx: raw for idx, raw in enumerate(raw_lines, start=1)
                }
            except Exception:
                raw_line_map = {}

        for idx, override in self._overrides.items():
            if override is None:
                line_map.pop(idx, None)
                raw_line_map.pop(idx, None)
            else:
                line_map[idx] = sanitize_line(
                    override, base_dir=self.context.base_dir
                )
                raw_line_map[idx] = override

        self._raw_line_map = raw_line_map

        ordered = [(idx, line_map[idx]) for idx in sorted(line_map.keys())]
        if (
            not self.context.use_all_entries
            and self.context.start_datetime
            and self.context.end_datetime
        ):
            ordered = slice_by_date(ordered, self.context.start_datetime, self.context.end_datetime)
            line_map = {idx: line for idx, line in ordered}

        return ordered, line_map

    def _set_current_error(self, item: ErrorItem) -> None:
        self._current_error = item
        line_text_raw = self._current_line_raw()
        line_text = self._display_line(line_text_raw)
        self._set_line_display(line_text)
        self._configure_retype_mode(False)

        if item.kind == "boss":
            unknown = item.boss or self._extract_boss_token(line_text_raw)
            self.error_header.setText("Boss error")
            self.boss_or_not_controls.setVisible(False)
            self.stack.setCurrentWidget(self.boss_page)
            self.boss_combo.clear()
            self.boss_combo.addItems(self._bosses)
            self.boss_token_input.setText(unknown)
            raw_line = self._raw_line_map.get(item.line_index, line_text_raw)
            self.boss_manual_input.setPlainText(raw_line)
            self.boss_map_radio.setChecked(True)
            self.boss_line_only_radio.setChecked(False)
            self._update_boss_inputs()
            self._set_line_display(line_text)
        elif item.kind == "boss_or_not":
            unknown = item.boss or self._extract_boss_token(line_text_raw)
            self.error_header.setText("Boss or NOT?")
            self.boss_or_not_controls.setVisible(True)
            self.choice_boss_radio.setChecked(True)
            self.stack.setCurrentWidget(self.boss_page)
            self.boss_combo.clear()
            self.boss_combo.addItems(self._bosses)
            self.boss_token_input.setText(unknown)
            raw_line = self._raw_line_map.get(item.line_index, line_text_raw)
            self.boss_manual_input.setPlainText(raw_line)
            self.boss_map_radio.setChecked(True)
            self.boss_line_only_radio.setChecked(False)
            self._update_boss_inputs()
            self._set_line_display(line_text)
        elif item.kind == "single_char":
            self.error_header.setText("Single-character name")
            self.boss_or_not_controls.setVisible(False)
            self.stack.setCurrentWidget(self.single_page)
            self.single_token_combo.clear()
            _, entry = self._split_prefix_entry(line_text_raw)
            tokens = [t for t in entry.split() if len(t) == 1]
            self.single_token_combo.addItems(tokens)
            self.single_replace_input.clear()
            self.single_replace_radio.setChecked(True)
            self._update_single_inputs()
        elif item.kind == "date":
            self.error_header.setText("Date error")
            self.boss_or_not_controls.setVisible(False)
            self.stack.setCurrentWidget(self.retype_page)
            self._set_retype_input(line_text_raw, entry_only=False)
        elif item.kind == "not":
            self.error_header.setText("Incorrect NOT usage")
            self.boss_or_not_controls.setVisible(False)
            self.stack.setCurrentWidget(self.retype_page)
            self._set_retype_input(line_text_raw, entry_only=True)
            self._set_line_display(line_text)
            self._configure_retype_mode(self._has_multi_not(line_text_raw))
        elif item.kind == "at":
            self.error_header.setText("Unexpected 'at'")
            self.boss_or_not_controls.setVisible(False)
            self.stack.setCurrentWidget(self.retype_page)
            self._set_retype_input(line_text_raw, entry_only=True)
        else:
            self.error_header.setText("General error")
            self.boss_or_not_controls.setVisible(False)
            self.stack.setCurrentWidget(self.retype_page)
            entry_only = ":" in line_text_raw
            self._set_retype_input(line_text_raw, entry_only=entry_only)

    def _set_context_view(self, line_index: int) -> None:
        import html

        prev_line = self._display_line(
            self._raw_line_map.get(
                line_index - 1, self._line_map.get(line_index - 1, "(no previous line)")
            )
        )
        next_line = self._display_line(
            self._raw_line_map.get(
                line_index + 1, self._line_map.get(line_index + 1, "(no next line)")
            )
        )
        current_line = self._display_line(
            self._raw_line_map.get(line_index, self._line_map.get(line_index, ""))
        )

        def esc(text: str) -> str:
            return html.escape(text)

        html_content = (
            f"<div style='color:#8d8273'>{esc(prev_line)}</div>"
            f"<div style='color:#f0e6d2; font-weight:bold'>{esc(current_line)}</div>"
            f"<div style='color:#8d8273'>{esc(next_line)}</div>"
        )
        self.context_view.setHtml(html_content)
        self._autosize_text_box(self.context_view, min_height=90, max_height=220)

    def _set_line_display(self, line_text: str) -> None:
        import html

        if self._current_error:
            self._set_context_view(self._current_error.line_index)
            return

        content = (
            f"<div style='color:#f0e6d2; font-weight:bold'>{html.escape(line_text)}</div>"
        )
        self.context_view.setHtml(content)
        self._autosize_text_box(self.context_view, min_height=90, max_height=220)

    @staticmethod
    def _autosize_text_box(
        widget: QPlainTextEdit, min_height: int = 90, max_height: int = 220
    ) -> None:
        doc = widget.document()
        width = widget.viewport().width() or 600
        doc.setTextWidth(width)
        height = int(doc.size().height()) + 14
        height = max(min_height, min(height, max_height))
        widget.setMinimumHeight(height)

    def _configure_retype_mode(self, allow_multinot: bool) -> None:
        self.retype_mode_widget.setVisible(allow_multinot)
        if allow_multinot:
            self.retype_mode_edit.setChecked(True)
        self._update_retype_inputs()

    def _update_retype_inputs(self) -> None:
        show_edit = not self.retype_mode_widget.isVisible() or self.retype_mode_edit.isChecked()
        self.retype_label.setVisible(show_edit)
        self.retype_input.setVisible(show_edit)
        if show_edit and self.retype_input.isVisible():
            QTimer.singleShot(0, self.retype_input.setFocus)

    def _is_multinot_selected(self) -> bool:
        return self.retype_mode_widget.isVisible() and self.retype_mode_multinot.isChecked()

    def _update_boss_inputs(self) -> None:
        is_add = self.boss_add_radio.isChecked()
        is_manual = self.boss_manual_radio.isChecked()
        self.boss_combo.setVisible(not is_add and not is_manual)
        self.boss_token_input.setVisible(is_add)
        self.boss_points.setVisible(is_add)
        self.boss_manual_input.setVisible(is_manual)
        if is_manual:
            QTimer.singleShot(0, self.boss_manual_input.setFocus)
        elif is_add:
            QTimer.singleShot(0, self.boss_token_input.setFocus)
        else:
            QTimer.singleShot(0, self.boss_combo.setFocus)

    def _update_single_inputs(self) -> None:
        self.single_replace_input.setVisible(self.single_replace_radio.isChecked())
        if self.single_replace_input.isVisible():
            QTimer.singleShot(0, self.single_replace_input.setFocus)

    def _on_boss_or_not_choice_changed(self) -> None:
        if not self._current_error or self._current_error.kind != "boss_or_not":
            return
        if self.choice_not_radio.isChecked():
            self.stack.setCurrentWidget(self.retype_page)
            line_text_raw = self._current_line_raw()
            self._set_retype_input(line_text_raw, entry_only=True)
            self._configure_retype_mode(self._has_multi_not(line_text_raw))
        else:
            self.stack.setCurrentWidget(self.boss_page)
            self._configure_retype_mode(False)

    def _current_line_raw(self) -> str:
        if not self._current_error:
            return ""
        return self._line_map.get(self._current_error.line_index, "")

    def _display_line(self, line_text: str) -> str:
        if not line_text:
            return ""
        cleaned = line_text.replace(MULTI_NOT_MARKER, "")
        cleaned = " ".join(cleaned.split())
        return cleaned

    def _extract_boss_token(self, line_text: str) -> str:
        if ":" not in line_text:
            return ""
        segment = line_text.rsplit(":", 1)[1].strip()
        return segment.split()[0] if segment else ""

    def _split_entry_tokens(self, line_text: str) -> List[str]:
        if ":" not in line_text:
            return []
        segment = line_text.rsplit(":", 1)[1].strip()
        tokens = segment.split()
        if len(tokens) < 2:
            return []

        full_line = list(tokens)
        modifier = full_line.pop(1)
        is_valid_modifier = False
        for test in MODIFIERS:
            if modifier == f"({test})":
                is_valid_modifier = True
                break

        if is_valid_modifier:
            full_line[0] = f"{full_line[0]}{modifier}"
        else:
            full_line.insert(1, modifier)

        boss = full_line.pop(0)
        if not boss:
            return []

        return [t for t in full_line if t != MULTI_NOT_MARKER]

    def _split_prefix_entry(self, line_text: str) -> (str, str):
        if ":" not in line_text:
            return "", line_text
        prefix, entry = line_text.rsplit(":", 1)
        return f"{prefix}:", entry.strip()

    def _replace_boss_in_entry(self, line_text: str, selection: str) -> str:
        prefix, entry = self._split_prefix_entry(line_text)
        tokens = entry.split()
        if not tokens:
            return line_text
        tokens[0] = selection
        return f"{prefix}{' '.join(tokens)}"

    def _set_retype_input(self, line_text: str, entry_only: bool) -> None:
        if entry_only:
            _, entry = self._split_prefix_entry(line_text)
            self.retype_label.setText("Edit the entry below (date/time preserved)")
            self.retype_input.setPlaceholderText("Edit the entry (no date/prefix)...")
            self.retype_input.setPlainText(self._display_line(entry))
        else:
            self.retype_label.setText("Edit the full line below (required)")
            self.retype_input.setPlaceholderText("Edit the full line here...")
            self.retype_input.setPlainText(self._display_line(line_text))
        self._update_retype_inputs()

    def _has_multi_not(self, line_text: str) -> bool:
        tokens = self._split_entry_tokens(line_text)
        if "not" not in tokens:
            return False
        not_index = tokens.index("not")
        return len(tokens[not_index + 1 :]) > 1

    def _apply_fix(self) -> None:
        if not self._current_error:
            return

        item = self._current_error
        line_text = self._current_line_raw()

        if item.kind in {"date", "not", "at", "general"}:
            entry_only = item.kind != "date" and ":" in line_text
            use_multinot = item.kind == "not" and self._is_multinot_selected()
            if use_multinot:
                if entry_only:
                    _, entry = self._split_prefix_entry(line_text)
                    new_raw = self._display_line(entry)
                else:
                    new_raw = self._display_line(line_text)
            else:
                new_raw = self.retype_input.toPlainText().strip()
                if not new_raw:
                    QMessageBox.critical(self, "Missing input", "Please edit the line.")
                    return
            if entry_only:
                prefix, _ = self._split_prefix_entry(line_text)
                sanitized_entry = sanitize_line(new_raw, base_dir=self.context.base_dir)
                if not sanitized_entry:
                    QMessageBox.critical(self, "Invalid line", "The line could not be parsed.")
                    return
                sanitized = f"{prefix}{sanitized_entry}"
            else:
                sanitized = sanitize_line(new_raw, base_dir=self.context.base_dir)
                if not sanitized:
                    QMessageBox.critical(self, "Invalid line", "The line could not be parsed.")
                    return
            if use_multinot:
                if MULTI_NOT_MARKER not in sanitized.split():
                    sanitized = f"{sanitized} {MULTI_NOT_MARKER}"
            self._overrides[item.line_index] = sanitized
            self._persist_overrides()
            self._revalidate()
            return

        if item.kind == "boss_or_not":
            if self.choice_not_radio.isChecked():
                entry_only = ":" in line_text
                use_multinot = self._is_multinot_selected()
                if use_multinot:
                    if entry_only:
                        _, entry = self._split_prefix_entry(line_text)
                        new_raw = self._display_line(entry)
                    else:
                        new_raw = self._display_line(line_text)
                else:
                    new_raw = self.retype_input.toPlainText().strip()
                    if not new_raw:
                        QMessageBox.critical(
                            self, "Missing input", "Please edit the line."
                        )
                        return
                if entry_only:
                    prefix, _ = self._split_prefix_entry(line_text)
                    sanitized_entry = sanitize_line(new_raw, base_dir=self.context.base_dir)
                    if not sanitized_entry:
                        QMessageBox.critical(self, "Invalid line", "The line could not be parsed.")
                        return
                    sanitized = f"{prefix}{sanitized_entry}"
                else:
                    sanitized = sanitize_line(new_raw, base_dir=self.context.base_dir)
                    if not sanitized:
                        QMessageBox.critical(self, "Invalid line", "The line could not be parsed.")
                        return
                if use_multinot:
                    if MULTI_NOT_MARKER not in sanitized.split():
                        sanitized = f"{sanitized} {MULTI_NOT_MARKER}"
                self._overrides[item.line_index] = sanitized
                self._persist_overrides()
                self._revalidate()
                return
            # Treat as boss error and apply immediately
            item = ErrorItem("boss", item.line_index, item.boss)

        if item.kind == "boss":
            unknown = item.boss or self._extract_boss_token(line_text)
            if self.boss_manual_radio.isChecked():
                new_raw = self.boss_manual_input.toPlainText().strip()
                if not new_raw:
                    QMessageBox.critical(self, "Missing input", "Please edit the line.")
                    return
                sanitized = sanitize_line(new_raw, base_dir=self.context.base_dir)
                if not sanitized:
                    QMessageBox.critical(self, "Invalid line", "The line could not be parsed.")
                    return
                self._overrides[item.line_index] = new_raw
            elif self.boss_add_radio.isChecked():
                new_token = self.boss_token_input.text().strip().lower()
                if not new_token:
                    QMessageBox.critical(self, "Missing boss", "Enter a boss token.")
                    return
                points = self.boss_points.value()
                add_points_value(self.context.base_dir, new_token, points)
                self._overrides[item.line_index] = self._replace_boss_in_entry(
                    line_text, new_token
                )
            elif self.boss_line_only_radio.isChecked():
                selection = self.boss_combo.currentText().strip()
                if not selection:
                    QMessageBox.critical(self, "Missing boss", "Select a boss.")
                    return
                self._overrides[item.line_index] = self._replace_boss_in_entry(
                    line_text, selection
                )
            else:
                selection = self.boss_combo.currentText().strip()
                if not selection:
                    QMessageBox.critical(self, "Missing boss", "Select a boss.")
                    return
                add_boss_alias(self.context.base_dir, unknown, selection)
                self._overrides[item.line_index] = self._replace_boss_in_entry(
                    line_text, selection
                )

            self._persist_overrides()
            self._revalidate()
            return

        if item.kind == "single_char":
            token = self.single_token_combo.currentText()
            if not token:
                QMessageBox.critical(self, "Missing token", "Select a token to fix.")
                return
            prefix, entry = self._split_prefix_entry(line_text)
            tokens = entry.split()
            if self.single_remove_radio.isChecked():
                removed = False
                new_tokens = []
                for t in tokens:
                    if t == token and not removed:
                        removed = True
                        continue
                    new_tokens.append(t)
                self._overrides[item.line_index] = f"{prefix}{' '.join(new_tokens)}"
            else:
                replacement_raw = self.single_replace_input.text().strip()
                if not replacement_raw:
                    QMessageBox.critical(
                        self, "Missing name", "Enter a replacement name."
                    )
                    return
                replacement_token = replacement_raw.lower()
                replaced = False
                new_tokens = []
                for t in tokens:
                    if t == token and not replaced:
                        new_tokens.append(replacement_token)
                        replaced = True
                    else:
                        new_tokens.append(t)
                self._overrides[item.line_index] = f"{prefix}{' '.join(new_tokens)}"

            self._persist_overrides()
            self._revalidate()
            return

    def _skip_line(self) -> None:
        if not self._current_error:
            return
        self._overrides[self._current_error.line_index] = None
        self._persist_overrides()
        self._revalidate()

    def _persist_overrides(self) -> None:
        if not self._overrides:
            return

        path = self.context.timers_path
        if not path.exists():
            return

        try:
            with path.open("r", encoding="utf-8", errors="ignore") as f:
                lines = f.read().splitlines()

            if not self._backup_created:
                backup_path = path.with_suffix(path.suffix + ".bak")
                if not backup_path.exists():
                    backup_text = "\n".join(lines)
                    if lines:
                        backup_text += "\n"
                    with backup_path.open("w", encoding="utf-8") as backup_file:
                        backup_file.write(backup_text)
                self._backup_created = True

            max_line = len(lines)
            for idx, override in self._overrides.items():
                pos = idx - 1
                if pos < 0 or pos >= max_line:
                    continue
                if override is None:
                    lines[pos] = ""
                else:
                    lines[pos] = override

            new_text = "\n".join(lines)
            if lines:
                new_text += "\n"
            with path.open("w", encoding="utf-8") as f:
                f.write(new_text)

            self._overrides = {}
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", f"Could not write timers.txt: {exc}")


class AutocorrectPage(QWizardPage):
    def __init__(self, context: WizardContext) -> None:
        super().__init__()
        self.context = context
        self.setTitle("Autocorrect")
        layout = QVBoxLayout(self)

        self.status = QLabel("Run autocorrect to continue.")
        layout.addWidget(self.status)

        self.resolve_progress_label = QLabel("Resolved 0 / 0")
        self.resolve_progress_label.setObjectName("ProgressLabel")
        layout.addWidget(self.resolve_progress_label)

        self.resolve_progress_bar = QProgressBar()
        self.resolve_progress_bar.setMaximum(1)
        self.resolve_progress_bar.setValue(0)
        layout.addWidget(self.resolve_progress_bar)

        self.review_label = QLabel("Names queued for review")
        self.review_label.setObjectName("SectionTitle")
        self.review_label.setVisible(False)
        layout.addWidget(self.review_label)

        self.review_box = QPlainTextEdit()
        self.review_box.setReadOnly(True)
        self.review_box.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        self.review_box.setMinimumHeight(80)
        self.review_box.setVisible(False)
        layout.addWidget(self.review_box)

        self.resolve_panel = QFrame()
        self.resolve_panel.setObjectName("FixPanel")
        resolve_layout = QVBoxLayout(self.resolve_panel)
        resolve_layout.setContentsMargins(12, 12, 12, 12)
        resolve_layout.setSpacing(10)

        self.resolve_line_view = QTextEdit()
        self.resolve_line_view.setObjectName("ContextView")
        self.resolve_line_view.setReadOnly(True)
        self.resolve_line_view.setLineWrapMode(QTextEdit.WidgetWidth)
        self.resolve_line_view.setMinimumHeight(60)
        resolve_layout.addWidget(self.resolve_line_view)

        self.resolve_unknown_label = QLabel("")
        self.resolve_unknown_label.setObjectName("ErrorHeader")
        resolve_layout.addWidget(self.resolve_unknown_label)

        self.resolve_options_header = QWidget()
        resolve_header_layout = QHBoxLayout(self.resolve_options_header)
        resolve_header_layout.setContentsMargins(0, 0, 0, 0)
        resolve_header_layout.setSpacing(8)
        self.resolve_options_label = QLabel("Pick a fix")
        self.resolve_options_label.setObjectName("SectionTitle")
        resolve_header_layout.addWidget(self.resolve_options_label)
        resolve_header_layout.addStretch(1)
        self.resolve_persist_checkbox = QCheckBox(
            "Save this mapping to name_aliases.json"
        )
        self.resolve_persist_checkbox.setChecked(True)
        resolve_header_layout.addWidget(self.resolve_persist_checkbox)
        resolve_layout.addWidget(self.resolve_options_header)

        self.resolve_options_container = QWidget()
        self.resolve_options_layout = QVBoxLayout(self.resolve_options_container)
        self.resolve_options_layout.setContentsMargins(0, 0, 0, 0)
        self.resolve_options_layout.setSpacing(6)
        resolve_layout.addWidget(self.resolve_options_container)

        resolve_buttons = QHBoxLayout()
        self.resolve_apply_button = QPushButton("Apply Resolution")
        self.resolve_apply_button.clicked.connect(self._apply_resolution)
        resolve_buttons.addWidget(self.resolve_apply_button)
        self.resolve_discard_button = QPushButton("Discard")
        self.resolve_discard_button.clicked.connect(self._discard_resolution)
        resolve_buttons.addWidget(self.resolve_discard_button)
        resolve_layout.addLayout(resolve_buttons)

        self.resolve_panel.setVisible(False)
        layout.addWidget(self.resolve_panel)

        self._resolve_loop: Optional[QEventLoop] = None
        self._resolve_result: Optional[Resolution] = None
        self._resolve_group: Optional[QButtonGroup] = None
        self._resolve_suggestion_map: Dict[QRadioButton, str] = {}
        self._resolve_custom_input: Optional[QLineEdit] = None
        self._resolve_split_first: Optional[QLineEdit] = None
        self._resolve_split_second: Optional[QLineEdit] = None
        self._resolve_add_input: Optional[QLineEdit] = None
        self._resolve_persist_checkbox: Optional[QCheckBox] = self.resolve_persist_checkbox
        self._resolve_merge_prev_btn: Optional[QRadioButton] = None
        self._resolve_merge_next_btn: Optional[QRadioButton] = None
        self._resolve_merge_prev_value: Optional[str] = None
        self._resolve_merge_next_value: Optional[str] = None
        self._resolve_total = 0
        self._resolve_count = 0
        self._review_queue: Set[str] = set()
        self._resolve_input_map: Dict[QRadioButton, List[QWidget]] = {}
        self._autocorrect_in_progress = False
        self._autocorrect_started = False
        self._autocorrect_source_key = None

    def reset_state(self) -> None:
        self._resolve_loop = None
        self._resolve_result = None
        self._resolve_group = None
        self._resolve_suggestion_map = {}
        self._resolve_custom_input = None
        self._resolve_split_first = None
        self._resolve_split_second = None
        self._resolve_add_input = None
        self._resolve_merge_prev_btn = None
        self._resolve_merge_next_btn = None
        self._resolve_merge_prev_value = None
        self._resolve_merge_next_value = None
        self._resolve_total = 0
        self._resolve_count = 0
        self._review_queue.clear()
        self._resolve_input_map = {}
        self._autocorrect_in_progress = False
        self._autocorrect_started = False
        self._autocorrect_source_key = None
        self.complete = False
        self.status.setText("Run autocorrect to continue.")
        self.resolve_progress_bar.setMaximum(1)
        self.resolve_progress_bar.setValue(0)
        self.resolve_progress_label.setText("Resolved 0 / 0")
        self.review_label.setVisible(False)
        self.review_box.setVisible(False)
        self.resolve_panel.setVisible(False)

    def initializePage(self) -> None:
        self._ensure_wizard_size(min_width=900, min_height=700)
        current_key = (
            str(self.context.timers_path),
            self.context.start_datetime,
            self.context.use_all_entries,
            self.context.spreadsheet_id,
            self.context.range_name,
            str(self.context.credentials_path),
        )
        if self._autocorrect_source_key != current_key:
            self._autocorrect_source_key = current_key
            self._autocorrect_started = False
            self._autocorrect_in_progress = False
        if not self._autocorrect_started:
            self._autocorrect_started = True
            self.status.setText("Running autocorrect...")
            QTimer.singleShot(0, self._run_autocorrect)

    def _run_autocorrect(self) -> None:
        if self._autocorrect_in_progress:
            return
        self._autocorrect_in_progress = True
        try:
            token_file = token_path()
            self._reset_resolve_progress()
            QApplication.processEvents()
            estimated, _sheet_names = estimate_unknown_count(
                timers_path=self.context.timers_path,
                start_date=self.context.start_datetime,
                end_date=self.context.end_datetime,
                use_all_entries=self.context.use_all_entries,
                spreadsheet_id=self.context.spreadsheet_id,
                range_name=self.context.range_name,
                credentials_path=self.context.credentials_path,
                token_path=token_file,
                base_dir=self.context.base_dir,
            )
            self._resolve_total = estimated
            self._update_resolve_progress()

            def resolver(
                name: str,
                suggestions: List[str],
                line_text: str,
                prev_token: str,
                next_token: str,
                prev_line_raw: str,
                next_line_raw: str,
            ) -> Optional[Resolution]:
                if self._resolve_count >= self._resolve_total:
                    self._resolve_total = self._resolve_count + 1
                self._update_resolve_progress()
                result = self._resolve_name_inline(
                    name,
                    suggestions,
                    line_text,
                    prev_token,
                    next_token,
                    prev_line_raw,
                    next_line_raw,
                )
                self._resolve_count += 1
                self._update_resolve_progress()
                return result

            calculation = calculate_points(
                timers_path=self.context.timers_path,
                start_date=self.context.start_datetime,
                end_date=self.context.end_datetime,
                use_all_entries=self.context.use_all_entries,
                spreadsheet_id=self.context.spreadsheet_id,
                range_name=self.context.range_name,
                credentials_path=self.context.credentials_path,
                token_path=token_file,
                base_dir=self.context.base_dir,
                resolve_unknown=resolver,
            )

            if calculation.errors.any():
                QMessageBox.critical(
                    self,
                    "Validation errors",
                    "Please fix the errors shown in the previous step and try again.",
                )
                return

            self.context.calculation = calculation
            self.status.setText(
                f"Autocorrect complete. {len(calculation.totals)} names with points."
            )
            self.complete = True
            self.completeChanged.emit()
        except Exception as exc:
            QMessageBox.critical(self, "Autocorrect failed", str(exc))
        finally:
            self._autocorrect_in_progress = False

    def isComplete(self) -> bool:
        return getattr(self, "complete", False)

    def _resolve_name_inline(
        self,
        name: str,
        suggestions: List[str],
        line_text: str,
        prev_token: str,
        next_token: str,
        prev_line_raw: str,
        next_line_raw: str,
    ) -> Optional[Resolution]:
        self._build_resolution_ui(
            name,
            suggestions,
            line_text,
            prev_token,
            next_token,
            prev_line_raw,
            next_line_raw,
        )
        self.resolve_panel.setVisible(True)
        self.resolve_panel.raise_()
        self._ensure_wizard_size(min_width=900, min_height=720)
        self._resolve_result = None
        self._resolve_loop = QEventLoop()
        self._resolve_loop.exec()
        self._resolve_loop = None
        self.resolve_panel.setVisible(False)
        return self._resolve_result

    def _reset_resolve_progress(self) -> None:
        self._resolve_total = 0
        self._resolve_count = 0
        self._update_resolve_progress()
        self._review_queue.clear()
        self._update_review_queue()

    def _update_resolve_progress(self) -> None:
        total = self._resolve_total
        resolved = min(self._resolve_count, total)
        if total <= 0:
            self.resolve_progress_bar.setMaximum(1)
            self.resolve_progress_bar.setValue(0)
            self.resolve_progress_label.setText("Resolved 0 / 0")
            return
        self.resolve_progress_bar.setMaximum(total)
        self.resolve_progress_bar.setValue(resolved)
        self.resolve_progress_label.setText(f"Resolved {resolved} / {total}")

    @staticmethod
    def _autosize_text_view(
        widget: QPlainTextEdit, min_height: int = 90, max_height: int = 220
    ) -> None:
        doc = widget.document()
        width = widget.viewport().width() or 600
        doc.setTextWidth(width)
        height = int(doc.size().height()) + 14
        height = max(min_height, min(height, max_height))
        widget.setMinimumHeight(height)

    @staticmethod
    def _fit_text_view(
        widget: QTextEdit, min_height: int = 40, max_height: int = 160
    ) -> None:
        doc = widget.document()
        width = widget.viewport().width() or 600
        doc.setTextWidth(width)
        height = int(doc.size().height()) + 12
        height = max(min_height, min(height, max_height))
        widget.setFixedHeight(height)

    @staticmethod
    def _display_line(line_text: str) -> str:
        if not line_text:
            return ""
        cleaned = line_text.replace(MULTI_NOT_MARKER, "")
        cleaned = " ".join(cleaned.split())
        return cleaned

    def _set_resolve_line_context(
        self, prev_line_raw: str, line_text: str, next_line_raw: str
    ) -> None:
        import html

        prev_line = self._display_line(prev_line_raw or "(no previous line)")
        current_line = self._display_line(line_text)
        next_line = self._display_line(next_line_raw or "(no next line)")

        def esc(text: str) -> str:
            return html.escape(text)

        html_content = (
            f"<div style='color:#8d8273'>{esc(prev_line)}</div>"
            f"<div style='color:#f0e6d2; font-weight:bold'>{esc(current_line)}</div>"
            f"<div style='color:#8d8273'>{esc(next_line)}</div>"
        )
        self.resolve_line_view.setHtml(html_content)

    def _update_resolve_inputs(self) -> None:
        for btn, widgets in self._resolve_input_map.items():
            show = btn.isChecked()
            for widget in widgets:
                widget.setVisible(show)
            if show and widgets:
                QTimer.singleShot(0, widgets[0].setFocus)

    def _build_resolution_ui(
        self,
        name: str,
        suggestions: List[str],
        line_text: str,
        prev_token: str,
        next_token: str,
        prev_line_raw: str,
        next_line_raw: str,
    ) -> None:
        self.resolve_unknown_label.setText(f"Unknown name: {name}")
        self._set_resolve_line_context(prev_line_raw, line_text, next_line_raw)
        self._fit_text_view(self.resolve_line_view, min_height=60, max_height=180)

        self.resolve_options_container.setUpdatesEnabled(False)
        self._clear_layout(self.resolve_options_layout)
        self._resolve_group = QButtonGroup(self)
        self._resolve_suggestion_map = {}
        self._resolve_custom_input = None
        self._resolve_split_first = None
        self._resolve_split_second = None
        self._resolve_add_input = None
        self._resolve_merge_prev_btn = None
        self._resolve_merge_next_btn = None
        self._resolve_merge_prev_value = None
        self._resolve_merge_next_value = None

        suggested_label = QLabel("Suggested names")
        suggested_label.setObjectName("SubtleTitle")
        self.resolve_options_layout.addWidget(suggested_label)

        for idx, suggestion in enumerate(suggestions, start=1):
            btn = QRadioButton(f"Use suggestion {idx}: {suggestion}")
            self._resolve_group.addButton(btn)
            self.resolve_options_layout.addWidget(btn)
            self._resolve_suggestion_map[btn] = suggestion

        divider = QFrame()
        divider.setObjectName("Divider")
        divider.setFrameShape(QFrame.HLine)
        divider.setFrameShadow(QFrame.Sunken)
        self.resolve_options_layout.addWidget(divider)

        other_label = QLabel("Other options")
        other_label.setObjectName("SubtleTitle")
        self.resolve_options_layout.addWidget(other_label)

        if prev_token:
            merged = f"{prev_token}{name}"
            self._resolve_merge_prev_value = merged
            btn = QRadioButton(f"Merge with previous token: {merged}")
            self._resolve_group.addButton(btn)
            self.resolve_options_layout.addWidget(btn)
            self._resolve_merge_prev_btn = btn

        if next_token:
            merged = f"{name}{next_token}"
            self._resolve_merge_next_value = merged
            btn = QRadioButton(f"Merge with next token: {merged}")
            self._resolve_group.addButton(btn)
            self.resolve_options_layout.addWidget(btn)
            self._resolve_merge_next_btn = btn

        self._resolve_custom_btn = QRadioButton("Enter a different name")
        self._resolve_group.addButton(self._resolve_custom_btn)
        self.resolve_options_layout.addWidget(self._resolve_custom_btn)
        self._resolve_custom_input = QLineEdit()
        self._resolve_custom_input.setPlaceholderText("Type a name")
        self._resolve_custom_input.setVisible(False)
        self.resolve_options_layout.addWidget(self._resolve_custom_input)

        self._resolve_split_btn = QRadioButton("Split into two names")
        self._resolve_group.addButton(self._resolve_split_btn)
        self.resolve_options_layout.addWidget(self._resolve_split_btn)
        split_row = QHBoxLayout()
        self._resolve_split_first = QLineEdit()
        self._resolve_split_first.setPlaceholderText("First name")
        self._resolve_split_second = QLineEdit()
        self._resolve_split_second.setPlaceholderText("Second name")
        split_row.addWidget(self._resolve_split_first)
        split_row.addWidget(self._resolve_split_second)
        self._resolve_split_first.setVisible(False)
        self._resolve_split_second.setVisible(False)
        self.resolve_options_layout.addLayout(split_row)

        self._resolve_add_btn = QRadioButton("Add as a new name")
        self._resolve_group.addButton(self._resolve_add_btn)
        self.resolve_options_layout.addWidget(self._resolve_add_btn)
        self._resolve_add_input = QLineEdit()
        self._resolve_add_input.setPlaceholderText("New name")
        self._resolve_add_input.setVisible(False)
        self.resolve_options_layout.addWidget(self._resolve_add_input)

        self._resolve_input_map = {
            self._resolve_custom_btn: [self._resolve_custom_input],
            self._resolve_split_btn: [self._resolve_split_first, self._resolve_split_second],
            self._resolve_add_btn: [self._resolve_add_input],
        }
        self._resolve_group.buttonToggled.connect(self._update_resolve_inputs)
        self._update_resolve_inputs()

        buttons = self._resolve_group.buttons()
        if buttons:
            buttons[0].setChecked(True)
            self._update_resolve_inputs()
        self.resolve_options_container.setUpdatesEnabled(True)
        self.resolve_options_container.update()

    def _apply_resolution(self) -> None:
        if not self._resolve_group:
            return
        selected = self._resolve_group.checkedButton()
        if not selected:
            QMessageBox.critical(self, "Missing choice", "Please choose a resolution.")
            return

        resolution = None

        if self._resolve_merge_prev_btn and selected == self._resolve_merge_prev_btn:
            if not self._resolve_merge_prev_value:
                QMessageBox.critical(
                    self, "Missing merge", "No previous token to merge."
                )
                return
            resolution = Resolution(
                names=[self._resolve_merge_prev_value],
                cache_original=False,
                persist_alias=False,
                merge_with_prev=True,
                reprocess=True,
            )

        elif self._resolve_merge_next_btn and selected == self._resolve_merge_next_btn:
            if not self._resolve_merge_next_value:
                QMessageBox.critical(
                    self, "Missing merge", "No next token to merge."
                )
                return
            resolution = Resolution(
                names=[self._resolve_merge_next_value],
                cache_original=False,
                persist_alias=False,
                merge_with_next=True,
                reprocess=True,
            )

        elif selected in self._resolve_suggestion_map:
            suggestion = self._resolve_suggestion_map[selected]
            resolution = Resolution(
                names=[suggestion],
                cache_original=True,
                persist_alias=self._resolve_persist_checkbox.isChecked()
                if self._resolve_persist_checkbox
                else False,
            )

        elif selected == getattr(self, "_resolve_custom_btn", None):
            value = (
                self._resolve_custom_input.text().strip()
                if self._resolve_custom_input
                else ""
            )
            if not value:
                QMessageBox.critical(self, "Missing name", "Enter a name.")
                return
            resolution = Resolution(
                names=[value],
                cache_original=True,
                persist_alias=self._resolve_persist_checkbox.isChecked()
                if self._resolve_persist_checkbox
                else False,
                reprocess=True,
            )

        elif selected == getattr(self, "_resolve_split_btn", None):
            first = (
                self._resolve_split_first.text().strip()
                if self._resolve_split_first
                else ""
            )
            second = (
                self._resolve_split_second.text().strip()
                if self._resolve_split_second
                else ""
            )
            if not first or not second:
                QMessageBox.critical(
                    self, "Missing names", "Enter both split names."
                )
                return
            resolution = Resolution(
                names=[first, second],
                cache_original=False,
                persist_alias=False,
                reprocess=True,
            )

        elif selected == getattr(self, "_resolve_add_btn", None):
            value = (
                self._resolve_add_input.text().strip()
                if self._resolve_add_input
                else ""
            )
            if not value:
                QMessageBox.critical(self, "Missing name", "Enter a new name.")
                return
            resolution = Resolution(
                names=[value],
                cache_original=True,
                add_new=True,
                persist_alias=self._resolve_persist_checkbox.isChecked()
                if self._resolve_persist_checkbox
                else False,
            )

        if resolution is None:
            return

        if resolution.add_new:
            for name in resolution.names:
                if name:
                    self._review_queue.add(name)
            self._update_review_queue()

        self._finish_resolution(resolution)

    def _discard_resolution(self) -> None:
        self._finish_resolution(None)

    def _finish_resolution(self, result: Optional[Resolution]) -> None:
        self._resolve_result = result
        if self._resolve_loop:
            self._resolve_loop.quit()

    def _update_review_queue(self) -> None:
        if not self._review_queue:
            self.review_label.setVisible(False)
            self.review_box.setVisible(False)
            return
        self.review_label.setVisible(True)
        self.review_box.setVisible(True)
        self.review_box.setPlainText("\n".join(sorted(self._review_queue, key=str.lower)))
        self._autosize_text_view(self.review_box, min_height=80, max_height=200)

    def _ensure_wizard_size(self, min_width: int, min_height: int) -> None:
        wizard = self.wizard()
        if not wizard:
            return
        current = wizard.size()
        target_w = max(current.width(), min_width)
        target_h = max(current.height(), min_height)
        if current.width() != target_w or current.height() != target_h:
            wizard.resize(target_w, target_h)

    @staticmethod
    def _clear_layout(layout: QVBoxLayout) -> None:
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.hide()
                widget.setParent(None)
                widget.deleteLater()
            child_layout = item.layout()
            if child_layout:
                AutocorrectPage._clear_layout(child_layout)
                child_layout.setParent(None)
                child_layout.deleteLater()


class ResultsPage(QWizardPage):
    def __init__(self, context: WizardContext) -> None:
        super().__init__()
        self.context = context
        self.setTitle("Results")
        layout = QVBoxLayout(self)

        self.results_box = QPlainTextEdit()
        self.results_box.setReadOnly(True)
        layout.addWidget(self.results_box)

        self.breakdown_label = QLabel("Boss breakdown")
        self.breakdown_label.setObjectName("SectionTitle")
        layout.addWidget(self.breakdown_label)

        self.breakdown_table = QTableWidget()
        self.breakdown_table.setColumnCount(0)
        self.breakdown_table.setRowCount(0)
        self.breakdown_table.setSortingEnabled(False)
        self.breakdown_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.breakdown_table.setSelectionMode(QAbstractItemView.NoSelection)
        self.breakdown_table.verticalHeader().setVisible(False)
        self.breakdown_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.breakdown_table)

        buttons_row = QHBoxLayout()
        export_button = QPushButton("Export to output.txt")
        export_button.clicked.connect(self._export_txt)
        buttons_row.addWidget(export_button)

        export_csv_button = QPushButton("Export CSV")
        export_csv_button.clicked.connect(self._export_csv)
        buttons_row.addWidget(export_csv_button)

        copy_button = QPushButton("Copy to clipboard")
        copy_button.clicked.connect(self._copy_clipboard)
        buttons_row.addWidget(copy_button)

        layout.addLayout(buttons_row)

    def reset_state(self) -> None:
        self.results_box.setPlainText("")
        self.breakdown_table.setRowCount(0)
        self.breakdown_table.setColumnCount(0)

    def initializePage(self) -> None:
        calculation = self.context.calculation
        if not calculation:
            self.results_box.setPlainText("No results available.")
            self.breakdown_table.setRowCount(0)
            self.breakdown_table.setColumnCount(0)
            return

        lines = [f"{name}, {points}" for name, points in calculation.totals]
        self.results_box.setPlainText("\n".join(lines))

        boss_list = calculation.boss_list
        boss_counts = calculation.boss_counts
        total_names = {name for name, _points in calculation.totals}
        players = sorted(total_names | set(boss_counts.keys()), key=str.lower)
        columns = ["Player"] + boss_list
        self.breakdown_table.setColumnCount(len(columns))
        self.breakdown_table.setHorizontalHeaderLabels(columns)
        self.breakdown_table.setRowCount(len(players))

        for row, player in enumerate(players):
            self.breakdown_table.setItem(row, 0, QTableWidgetItem(player))
            counts = boss_counts.get(player, {})
            for col, boss in enumerate(boss_list, start=1):
                value = counts.get(boss, 0)
                item = QTableWidgetItem()
                item.setData(Qt.DisplayRole, int(value))
                self.breakdown_table.setItem(row, col, item)

        self.breakdown_table.resizeColumnsToContents()
        self.breakdown_table.resizeRowsToContents()
        self.breakdown_table.setSortingEnabled(True)

    def _export_txt(self) -> None:
        calculation = self.context.calculation
        if not calculation:
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save output", "output.txt", "Text Files (*.txt)"
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            for name, points in calculation.totals:
                f.write(f"{name}, {points}\n")

    def _export_csv(self) -> None:
        calculation = self.context.calculation
        if not calculation:
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save CSV", "output.csv", "CSV Files (*.csv)"
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write("Name,Points\n")
            for name, points in calculation.totals:
                safe_name = name.replace('"', '""')
                f.write(f"\"{safe_name}\",{points}\n")

    def _copy_clipboard(self) -> None:
        calculation = self.context.calculation
        if not calculation:
            return
        lines = [f"{name}, {points}" for name, points in calculation.totals]
        QApplication.clipboard().setText("\n".join(lines))


class DkpWizard(QWizard):
    def __init__(self, base_dir: Path) -> None:
        super().__init__()
        config = load_config()
        self.context = WizardContext(
            base_dir=base_dir,
            config=config,
            timers_path=Path(config.last_timers_path) if config.last_timers_path else Path(),
            credentials_path=Path(config.last_credentials_path)
            if config.last_credentials_path
            else Path(),
            spreadsheet_id=config.spreadsheet_id,
            range_name=config.range_name,
            use_all_entries=config.use_all_entries,
            start_datetime=None,
            end_datetime=None,
        )

        self.base_dir = base_dir
        self.setWindowTitle("DKP Automator")
        self.setWizardStyle(QWizard.ModernStyle)
        banner = QPixmap(1, 1)
        banner.fill(QColor("#1d1a18"))
        self.setPixmap(QWizard.BannerPixmap, banner)
        self.setup_page = SetupPage(self.context)
        self.sanity_page = SanityCheckPage(self.context)
        self.autocorrect_page = AutocorrectPage(self.context)
        self.results_page = ResultsPage(self.context)
        self.addPage(self.setup_page)
        self.addPage(self.sanity_page)
        self.addPage(self.autocorrect_page)
        self.addPage(self.results_page)

    def closeEvent(self, event) -> None:
        event.accept()

    def accept(self) -> None:
        self._save_current_run()
        self._reset_run_state()

    def reject(self) -> None:
        self._reset_run_state()

    def _reset_run_state(self) -> None:
        self.context.calculation = None
        self.sanity_page.reset_state()
        self.autocorrect_page.reset_state()
        self.results_page.reset_state()
        self.restart()
        self.setCurrentId(0)

    def _save_current_run(self) -> None:
        calculation = self.context.calculation
        start_dt = self.context.start_datetime
        end_dt = self.context.end_datetime
        if not calculation or not calculation.events or not start_dt or not end_dt:
            return
        run_id = uuid4().hex
        created = datetime.now(timezone.utc)
        run_meta = build_run_meta(
            run_id=run_id,
            created_utc=created,
            start_utc=start_dt,
            end_utc=end_dt,
            event_count=len(calculation.events),
            timers_path=str(self.context.timers_path),
        )
        event_payloads = []
        for event in calculation.events:
            entries = [{"name": entry.name, "delta": entry.delta} for entry in event.entries]
            event_payloads.append(
                normalize_event(
                    run_id=run_id,
                    created_utc=created,
                    event_time=event.event_time,
                    boss=event.boss,
                    points=event.points,
                    entries=entries,
                    source_line=event.source_line,
                )
            )
        save_run(self.base_dir, run_meta, event_payloads)
