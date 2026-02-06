import logging
import os
import sys
from pathlib import Path
import shutil

from PySide6.QtCore import QtMsgType, qInstallMessageHandler
from PySide6.QtWidgets import QApplication

from .gui.wizard import DkpWizard


def _setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
    )

    def qt_message_handler(mode, context, message):
        level = logging.INFO
        if mode == QtMsgType.QtWarningMsg:
            level = logging.WARNING
        elif mode == QtMsgType.QtCriticalMsg:
            level = logging.ERROR
        elif mode == QtMsgType.QtFatalMsg:
            level = logging.CRITICAL
        logging.log(level, "Qt: %s", message)

    qInstallMessageHandler(qt_message_handler)


def _seed_config_files(target_dir: Path) -> None:
    source_dir = None
    if hasattr(sys, "_MEIPASS"):
        source_dir = Path(getattr(sys, "_MEIPASS")) / "Config files"
        if not source_dir.exists():
            source_dir = Path(getattr(sys, "_MEIPASS")) / "data"
    if source_dir is None or not source_dir.exists():
        return

    target_dir.mkdir(parents=True, exist_ok=True)
    for name in ("points.json", "name_aliases.json", "boss_aliases.json", "prios.json"):
        src = source_dir / name
        dst = target_dir / name
        if src.exists() and not dst.exists():
            shutil.copy2(src, dst)


def _resolve_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        config_dir = exe_dir / "Config files"
        _seed_config_files(config_dir)
        return config_dir
    return Path(__file__).resolve().parents[1]


def main() -> int:
    base_dir = _resolve_base_dir()

    #os.environ.setdefault("QT_LOGGING_RULES", "qt.qpa.*=true;qt.widgets.*=true")
    _setup_logging()

    app = QApplication(sys.argv)
    logging.info("Python: %s", sys.version.replace("\n", " "))
    logging.info("App base dir: %s", base_dir)
    logging.info("Qt platform: %s", app.platformName())
    app.setStyleSheet(
        """
        QWidget {
            font-family: "Georgia", "Times New Roman", serif;
            color: #f0e6d2;
            background-color: #1d1a18;
        }
        QWizard, QDialog {
            background-color: #1d1a18;
        }
        QWizard::header {
            background-color: #1d1a18;
            border-bottom: 1px solid #3a322b;
        }
        QWizard QFrame#qt_wizard_titlebar,
        QWizard QFrame#qt_wizard_header,
        QWizard QWidget#qt_wizard_header,
        QWizard QLabel#qt_wizard_title,
        QWizard QLabel#qt_wizard_subtitle {
            background-color: #1d1a18;
            color: #e8d9b5;
        }
        QWizard::title {
            color: #e8d9b5;
        }
        QWizard::subTitle {
            color: #d7c49a;
        }
        QWizard QWidget#qt_wizard_header {
            background-color: #1d1a18;
        }
        QWizard QFrame#qt_wizard_buttonbox,
        QWizard QWidget#qt_wizard_buttonbox,
        QWizard QDialogButtonBox {
            background-color: #1d1a18;
            border-top: 1px solid #3a322b;
        }
        QLabel#SectionTitle {
            font-size: 16px;
            font-weight: bold;
            color: #e8d9b5;
        }
        QLabel#ErrorHeader {
            font-size: 14px;
            font-weight: bold;
            color: #f0c36d;
        }
        QLabel#ProgressLabel {
            color: #d7c49a;
        }
        QLabel#SubtleTitle {
            font-size: 12px;
            font-weight: bold;
            color: #cdbb95;
        }
        QLineEdit, QPlainTextEdit, QTextEdit, QComboBox, QDateEdit, QTimeEdit, QSpinBox {
            background-color: #2a2521;
            border: 1px solid #4a4036;
            border-radius: 6px;
            padding: 6px;
            color: #f5efdf;
        }
        QComboBox {
            padding-right: 18px;
        }
        QComboBox QAbstractItemView {
            background-color: #2a2521;
            selection-background-color: #4a3a2b;
            color: #f5efdf;
        }
        QPlainTextEdit, QLineEdit {
            selection-background-color: #4a3a2b;
        }
        QTableWidget {
            background-color: #2a2521;
            gridline-color: #4a4036;
            color: #f5efdf;
            border: 1px solid #4a4036;
        }
        QTableWidget::item:selected {
            background-color: #8a6a3f;
            color: #1d1a18;
        }
        QHeaderView::section {
            background-color: #26211d;
            color: #e8d9b5;
            padding: 4px 6px;
            border: 1px solid #3a322b;
        }
        QTextEdit#ContextView {
            background-color: #211d19;
            border: 1px dashed #5a4a39;
        }
        QPlainTextEdit#LineView, QPlainTextEdit#RetypeInput {
            font-family: "Consolas", "Courier New", monospace;
        }
        QTabWidget::pane {
            border: 1px solid #5a4a39;
            border-radius: 8px;
            background-color: #221e1a;
            padding-top: 4px;
        }
        QTabBar::tab {
            background-color: #2f2721;
            color: #d7c49a;
            border: 1px solid #6a5236;
            border-bottom: none;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            padding: 6px 14px;
            margin-right: 4px;
        }
        QTabBar::tab:selected {
            background-color: #8a6a3f;
            color: #1d1a18;
            border-color: #b4874f;
        }
        QTabBar::tab:hover {
            background-color: #3b2e23;
            color: #f0e6d2;
        }
        QPushButton {
            background-color: #3b2e23;
            border: 1px solid #7a5c3a;
            border-radius: 6px;
            padding: 6px 10px;
            color: #f0e4c8;
        }
        QPushButton:hover {
            background-color: #4a3a2b;
        }
        QPushButton:disabled {
            color: #6f665c;
            background-color: #2a2622;
            border-color: #3a322b;
        }
        QCheckBox, QRadioButton {
            spacing: 8px;
        }
        QProgressBar {
            border: 1px solid #4a4036;
            border-radius: 6px;
            background-color: #2a2521;
            text-align: center;
            color: #f0e6d2;
            height: 16px;
        }
        QProgressBar::chunk {
            background-color: #8a6a3f;
            border-radius: 6px;
        }
        QProgressBar#InlineSpinner {
            height: 10px;
        }
        QFrame#FixPanel {
            background-color: #26211d;
            border: 1px solid #5a4a39;
            border-radius: 10px;
        }
        QFrame#Divider {
            border: 0;
            border-top: 1px solid #3a322b;
            margin-top: 6px;
            margin-bottom: 6px;
        }
        QLabel#StatusIndicator {
            background-color: #3a322b;
            border: 1px solid #5a4a39;
            border-radius: 8px;
            color: #f5efdf;
            font-weight: bold;
        }
        QLabel#StatusIndicator[state="ok"] {
            background-color: #2d7a3f;
            border-color: #3e9b57;
        }
        QLabel#StatusIndicator[state="error"] {
            background-color: #8a3b2f;
            border-color: #b24a39;
        }
        QLabel#StatusIndicator[state="working"] {
            background-color: #7a5c3a;
            border-color: #9a6f45;
        }
        QLabel#TestStatusText {
            color: #d7c49a;
        }
        """
    )
    wizard = DkpWizard(base_dir)
    wizard.resize(960, 720)
    wizard.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
