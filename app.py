import os
import sys
import logging
import logging.handlers
from pathlib import Path
from typing import Dict

from PySide6.QtCore import QThread, Signal, QUrl
from PySide6.QtGui import QDesktopServices, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QLineEdit,
    QProgressBar,
    QMenuBar,
    QMenu,
    QFileDialog,
    QMessageBox,
    QFrame,
    QTabWidget,
    QDialog,
    QPlainTextEdit,
    QFormLayout,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
)

SRC_DIR = Path(__file__).resolve().parent / "src"
if SRC_DIR.exists() and str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

try:
    from src.config_manager import load_config, save_config, ConfigError
    from src.access_bridge import AccessBridge, AccessBridgeError, AuthenticationError, PermissionError
    from src.validator import (
        load_excel,
        validate_dataframe,
        FileNotFoundError as ValidatorFileNotFoundError,
        InvalidFormatError,
    )
    from src.splash_screen import create_splash_screen
except ModuleNotFoundError:
    from config_manager import load_config, save_config, ConfigError # type: ignore
    from access_bridge import AccessBridge, AccessBridgeError, AuthenticationError, PermissionError # type: ignore
    from validator import ( # type: ignore
        load_excel,
        validate_dataframe,
        FileNotFoundError as ValidatorFileNotFoundError,
        InvalidFormatError,
    )
    from splash_screen import create_splash_screen # type: ignore


APP_TITLE = "SharePoint Upload Tool"


try:
    config = load_config()
except ConfigError as exc:
    print(f"Configuration Error: {exc}")
    sys.exit(1)

log_file = config["logging"]["log_file"]
log_level = getattr(logging, config["logging"].get("log_level", "INFO").upper(), logging.INFO)

os.makedirs(os.path.dirname(log_file) if os.path.dirname(log_file) else ".", exist_ok=True)

logger = logging.getLogger()
logger.setLevel(log_level)

handler = logging.handlers.RotatingFileHandler(
    log_file, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
)
formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
handler.setFormatter(formatter)
logger.addHandler(handler)


class UploadWorker(QThread):
    status = Signal(str)
    finished = Signal(bool, str)

    def __init__(self, file_path: str, config_obj, required_columns):
        super().__init__()
        self.file_path = file_path
        self.config_obj = config_obj
        self.required_columns = required_columns

    def run(self) -> None:
        self.status.emit("Starting upload...")
        logger.info("Starting upload process.")

        try:
            self.status.emit("Validating file...")
            df = load_excel(self.file_path)
            ok, msg = validate_dataframe(df, self.required_columns)
            if not ok:
                self.finished.emit(False, msg)
                return

            with AccessBridge(
                db_path=self.config_obj["access"]["db_path"],
                linked_table=self.config_obj["access"]["linked_table"],
                temp_table=self.config_obj["access"]["temp_table"],
                macro_name=self.config_obj["access"]["macro_name"],
            ) as bridge:
                try:
                    self.status.emit("Checking authentication...")
                    bridge.ensure_authenticated(timeout_seconds=60, poll_interval=3)
                except PermissionError as exc:
                    logger.error(f"Permission error: {exc}")
                    self.finished.emit(False, str(exc))
                    return
                except AuthenticationError as exc:
                    logger.error(f"Authentication error: {exc}")
                    self.finished.emit(False, str(exc))
                    return

                try:
                    self.status.emit(f"Importing {len(df)} rows into Access...")
                    bridge.import_excel_to_temp(self.file_path)
                except AccessBridgeError as exc:
                    logger.error(f"Import error: {exc}")
                    self.finished.emit(False, str(exc))
                    return

                try:
                    self.status.emit("Running append macro...")
                    bridge.run_macro()
                except AccessBridgeError as exc:
                    logger.error(f"Macro error: {exc}")
                    try:
                        bridge.clear_temp_table()
                    except Exception:
                        pass
                    self.finished.emit(False, str(exc))
                    return

            self.finished.emit(True, "Upload completed successfully.")
            logger.info("Upload completed successfully.")

        except ValidatorFileNotFoundError as exc:
            self.finished.emit(False, str(exc))
        except InvalidFormatError as exc:
            self.finished.emit(False, str(exc))
        except Exception as exc:
            logger.exception("Upload failed.")
            self.finished.emit(False, f"Upload failed: {exc}")


class ConnectionTestWorker(QThread):
    finished = Signal(bool, str)

    def __init__(self, config_obj):
        super().__init__()
        self.config_obj = config_obj

    def run(self) -> None:
        bridge = AccessBridge(
            db_path=self.config_obj["access"]["db_path"],
            linked_table=self.config_obj["access"]["linked_table"],
            temp_table=self.config_obj["access"]["temp_table"],
            macro_name=self.config_obj["access"]["macro_name"],
        )
        try:
            bridge.test_connection()
            self.finished.emit(True, "Connection and permissions look OK.")
        except PermissionError:
            self.finished.emit(False, "You do not have permission to access the SharePoint list.")
        except AuthenticationError as exc:
            self.finished.emit(False, str(exc))
        except Exception as exc:
            self.finished.emit(False, str(exc))
        finally:
            try:
                bridge.close()
            except Exception:
                pass


class LogViewer(QDialog):
    def __init__(self, parent, log_path: str):
        super().__init__(parent)
        self.setWindowTitle("Log Viewer")
        self.resize(900, 500)
        self.log_path = log_path

        layout = QVBoxLayout()

        self.text = QPlainTextEdit()
        self.text.setReadOnly(True)
        self.text.setStyleSheet("font-family: Consolas; font-size: 10px;")
        layout.addWidget(self.text)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self.load_log)
        btn_layout.addWidget(refresh_btn)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

        self.setLayout(layout)
        self.load_log()

    def load_log(self) -> None:
        self.text.clear()
        if os.path.exists(self.log_path):
            try:
                with open(self.log_path, "r", encoding="utf-8", errors="ignore") as file_handle:
                    self.text.setPlainText(file_handle.read())
            except Exception as exc:
                self.text.setPlainText(f"Error reading log file: {exc}")
        else:
            self.text.setPlainText("Log file not found.")


class ConfigEditor(QDialog):
    def __init__(self, parent, config_obj):
        super().__init__(parent)
        self.setWindowTitle("Config Editor")
        self.resize(650, 400)
        self.config_obj = config_obj
        self.entries: Dict[str, QLineEdit] = {}

        layout = QVBoxLayout()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        form_layout = QFormLayout()

        for section in ("access", "validation", "logging"):
            section_label = QLabel(f"[{section}]")
            section_label.setStyleSheet("font-weight: bold; color: #c41e3a;")
            form_layout.addRow(section_label, QLabel(""))
            for key, value in self.config_obj[section].items():
                entry = QLineEdit(str(value))
                self.entries[f"{section}.{key}"] = entry
                form_layout.addRow(QLabel(key), entry)

        inner.setLayout(form_layout)
        scroll.setWidget(inner)
        layout.addWidget(scroll)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.save)
        btn_layout.addWidget(save_btn)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def save(self) -> None:
        for key, entry in self.entries.items():
            section, option = key.split(".", 1)
            self.config_obj[section][option] = entry.text().strip()
        save_config(self.config_obj)
        QMessageBox.information(self, "Config", "Configuration saved. Restart the app to apply changes.")


class UploadApp(QMainWindow):
    def __init__(self, config_obj):
        super().__init__()
        self.config = config_obj

        self.db_path = config_obj["access"]["db_path"]
        self.macro_name = config_obj["access"]["macro_name"]
        self.linked_table = config_obj["access"]["linked_table"]
        self.temp_table = config_obj["access"]["temp_table"]

        self.required_columns = [
            c.strip() for c in config_obj["validation"]["required_columns"].split(",") if c.strip()
        ]

        self.log_file = config_obj["logging"]["log_file"]
        self.upload_worker = None
        self.connection_worker = None

        self.setWindowTitle(APP_TITLE)
        self.setGeometry(100, 100, 1200, 800)
        self.setMinimumSize(1000, 700)

        icon_path = Path(__file__).parent / "icons" / "scotiabank_logo_icon_170755.png"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self.apply_stylesheet()
        self.create_ui()

    def apply_stylesheet(self) -> None:
        style = """
        QMainWindow {
            background-color: #0f1419;
        }

        QWidget {
            background-color: #0f1419;
            color: #ffffff;
        }

        QTabWidget::pane {
            border: none;
            background-color: #0f1419;
        }

        QTabBar::tab {
            background-color: #1a1f26;
            color: #ffffff;
            padding: 8px 20px;
            border: none;
            border-bottom: 3px solid transparent;
            font-weight: bold;
        }

        QTabBar::tab:selected {
            background-color: #c41e3a;
            border-bottom: 3px solid #ff4455;
            color: #ffffff;
        }

        QTabBar::tab:hover:!selected {
            background-color: #252d35;
            color: #ff9999;
        }

        QPushButton {
            background-color: #c41e3a;
            color: #ffffff;
            border: none;
            border-radius: 6px;
            padding: 8px 16px;
            font-weight: bold;
            font-size: 11px;
        }

        QPushButton:hover {
            background-color: #e6233d;
        }

        QPushButton:pressed {
            background-color: #a01830;
        }

        QPushButton:disabled {
            background-color: #555555;
            color: #999999;
        }

        QLineEdit, QTextEdit, QComboBox, QPlainTextEdit {
            background-color: #1a1f26;
            color: #ffffff;
            border: 1px solid #333333;
            border-radius: 4px;
            padding: 6px;
            selection-background-color: #c41e3a;
        }

        QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QPlainTextEdit:focus {
            border: 2px solid #c41e3a;
            background-color: #262d35;
        }

        QLabel {
            color: #ffffff;
        }

        QLabel#title {
            font-size: 24px;
            font-weight: bold;
            color: #ffffff;
        }

        QLabel#subtitle {
            font-size: 11px;
            color: #999999;
        }

        QLabel#section-title {
            font-size: 13px;
            font-weight: bold;
            color: #c41e3a;
        }

        QProgressBar {
            background-color: #1a1f26;
            border: 1px solid #333333;
            border-radius: 4px;
            height: 24px;
            text-align: center;
            color: #ffffff;
        }

        QProgressBar::chunk {
            background-color: #c41e3a;
            border-radius: 2px;
        }

        QTableWidget {
            background-color: #1a1f26;
            alternate-background-color: #252d35;
            gridline-color: #333333;
            border: 1px solid #333333;
            border-radius: 4px;
        }

        QTableWidget::item:selected {
            background-color: #c41e3a;
            color: #ffffff;
        }

        QHeaderView::section {
            background-color: #0f1419;
            color: #ffffff;
            padding: 6px;
            border: none;
            border-right: 1px solid #333333;
            font-weight: bold;
        }

        QScrollBar:vertical {
            background-color: #1a1f26;
            width: 12px;
            border: none;
        }

        QScrollBar::handle:vertical {
            background-color: #c41e3a;
            border-radius: 6px;
            min-height: 20px;
        }

        QScrollBar::handle:vertical:hover {
            background-color: #e6233d;
        }

        QScrollBar:horizontal {
            background-color: #1a1f26;
            height: 12px;
            border: none;
        }

        QScrollBar::handle:horizontal {
            background-color: #c41e3a;
            border-radius: 6px;
            min-width: 20px;
        }

        QScrollBar::handle:horizontal:hover {
            background-color: #e6233d;
        }

        QFrame#header {
            background-color: #1a1f26;
            border-bottom: 1px solid #333333;
            padding: 16px;
        }

        QFrame#footer {
            background-color: #1a1f26;
            border-top: 1px solid #333333;
            padding: 12px;
        }
        """
        self.setStyleSheet(style)

    def create_ui(self) -> None:
        menubar = QMenuBar(self)
        tools_menu = QMenu("Tools", self)
        tools_menu.addAction("Test Connection", self.test_connection)
        tools_menu.addAction("View Log", self.open_log_viewer)
        tools_menu.addAction("Edit Config", self.open_config_editor)
        menubar.addMenu(tools_menu)
        self.setMenuBar(menubar)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        header = self.create_header()
        main_layout.addWidget(header)

        tabs = QTabWidget()
        tabs.addTab(self.create_upload_tab(), "Upload")
        main_layout.addWidget(tabs, 1)

        footer = self.create_footer()
        main_layout.addWidget(footer)

        central_widget.setLayout(main_layout)

    def create_header(self) -> QFrame:
        header = QFrame()
        header.setObjectName("header")
        layout = QVBoxLayout()
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(6)

        title = QLabel(APP_TITLE)
        title.setObjectName("title")
        layout.addWidget(title)

        subtitle = QLabel(
            "Upload Excel data to SharePoint through the Access bridge, with validation and preview."
        )
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        header.setLayout(layout)
        return header

    def create_footer(self) -> QFrame:
        footer = QFrame()
        footer.setObjectName("footer")
        layout = QHBoxLayout()
        layout.setContentsMargins(16, 8, 16, 8)
        layout.setSpacing(8)

        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("color: #999999; font-size: 10px;")
        layout.addWidget(self.status_label)

        layout.addStretch()

        open_logs_btn = QPushButton("Open Logs Folder")
        open_logs_btn.clicked.connect(self.open_logs_folder)
        layout.addWidget(open_logs_btn)

        about_btn = QPushButton("About")
        about_btn.clicked.connect(self.show_about)
        layout.addWidget(about_btn)

        footer.setLayout(layout)
        return footer

    def create_upload_tab(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        title = QLabel("Step 1: Select Excel File")
        title.setObjectName("section-title")
        layout.addWidget(title)

        file_layout = QHBoxLayout()
        self.file_input = QLineEdit()
        self.file_input.setPlaceholderText("No file selected")
        self.file_input.setReadOnly(True)
        file_layout.addWidget(self.file_input)

        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.pick_file)
        browse_btn.setMaximumWidth(120)
        file_layout.addWidget(browse_btn)
        layout.addLayout(file_layout)

        action_title = QLabel("Step 2: Validate, Preview, Upload")
        action_title.setObjectName("section-title")
        layout.addWidget(action_title)

        action_layout = QHBoxLayout()
        validate_btn = QPushButton("Validate File")
        validate_btn.clicked.connect(self.validate_file)
        action_layout.addWidget(validate_btn)

        preview_btn = QPushButton("Preview Data")
        preview_btn.clicked.connect(self.preview_data)
        action_layout.addWidget(preview_btn)

        upload_btn = QPushButton("Upload to SharePoint")
        upload_btn.clicked.connect(self.start_upload)
        action_layout.addWidget(upload_btn)

        action_layout.addStretch()
        layout.addLayout(action_layout)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        layout.addStretch()
        container.setLayout(layout)
        return container

    def set_status(self, text: str) -> None:
        self.status_label.setText(text)
        logger.debug(f"Status: {text}")

    def pick_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xls)",
        )
        if file_path:
            self.file_input.setText(file_path)
            logger.info(f"Selected file: {file_path}")
            self.set_status(f"Selected: {Path(file_path).name}")

    def _validate_file_path(self, path: str) -> bool:
        if not path:
            QMessageBox.warning(self, "Error", "Please select a file first.")
            return False

        if not Path(path).exists():
            QMessageBox.warning(self, "Error", f"File not found: {path}")
            return False

        return True

    def validate_file(self) -> None:
        path = self.file_input.text().strip()
        if not self._validate_file_path(path):
            return

        try:
            df = load_excel(path)
            ok, msg = validate_dataframe(df, self.required_columns)
            if ok:
                QMessageBox.information(self, "Validation Passed", msg)
            else:
                QMessageBox.warning(self, "Validation Failed", msg)
        except ValidatorFileNotFoundError as exc:
            QMessageBox.warning(self, "File Error", f"Cannot find file:\n{exc}")
        except InvalidFormatError as exc:
            QMessageBox.warning(self, "Format Error", f"Invalid Excel file:\n{exc}")
        except Exception as exc:
            logger.exception("Validation failed.")
            QMessageBox.warning(self, "Error", f"Validation failed:\n{exc}")

    def preview_data(self) -> None:
        path = self.file_input.text().strip()
        if not self._validate_file_path(path):
            return

        try:
            df = load_excel(path)
            ok, msg = validate_dataframe(df, self.required_columns)
            if not ok:
                QMessageBox.warning(self, "Cannot Preview", f"File validation failed:\n{msg}")
                return

            dialog = QDialog(self)
            dialog.setWindowTitle("Data Preview")
            dialog.resize(900, 500)

            layout = QVBoxLayout()

            info = QLabel(f"File: {Path(path).name}  |  Rows: {len(df)}  |  Columns: {len(df.columns)}")
            info.setStyleSheet("color: #999999; font-size: 10px;")
            layout.addWidget(info)

            table = QTableWidget()
            table.setAlternatingRowColors(True)
            preview_df = df.head(100)
            table.setRowCount(len(preview_df))
            table.setColumnCount(len(preview_df.columns))
            table.setHorizontalHeaderLabels([str(col) for col in preview_df.columns])
            for row_offset, (_, row) in enumerate(preview_df.iterrows()):
                for col_idx, col in enumerate(preview_df.columns):
                    item = QTableWidgetItem(str(row[col])[:50])
                    table.setItem(row_offset, col_idx, item)

            table.resizeColumnsToContents()
            layout.addWidget(table)

            btn_layout = QHBoxLayout()
            btn_layout.addStretch()
            close_btn = QPushButton("Close")
            close_btn.clicked.connect(dialog.close)
            btn_layout.addWidget(close_btn)
            layout.addLayout(btn_layout)

            dialog.setLayout(layout)
            dialog.exec()
        except ValidatorFileNotFoundError as exc:
            QMessageBox.warning(self, "File Error", f"Cannot find file:\n{exc}")
        except InvalidFormatError as exc:
            QMessageBox.warning(self, "Format Error", f"Invalid Excel file:\n{exc}")
        except Exception as exc:
            logger.exception("Preview failed.")
            QMessageBox.warning(self, "Error", f"Preview failed:\n{exc}")

    def start_upload(self) -> None:
        path = self.file_input.text().strip()
        if not self._validate_file_path(path):
            self.set_status("Upload cancelled.")
            return

        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.set_status("Starting upload...")

        self.upload_worker = UploadWorker(path, self.config, self.required_columns)
        self.upload_worker.status.connect(self.set_status)
        self.upload_worker.finished.connect(self.on_upload_finished)
        self.upload_worker.start()

    def on_upload_finished(self, success: bool, message: str) -> None:
        self.progress.setVisible(False)
        self.progress.setRange(0, 100)
        if success:
            self.set_status("Completed successfully.")
            QMessageBox.information(self, "Success", message)
        else:
            self.set_status("Upload failed.")
            QMessageBox.warning(self, "Upload Failed", message)

    def test_connection(self) -> None:
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.set_status("Testing connection...")

        self.connection_worker = ConnectionTestWorker(self.config)
        self.connection_worker.finished.connect(self.on_connection_finished)
        self.connection_worker.start()

    def on_connection_finished(self, success: bool, message: str) -> None:
        self.progress.setVisible(False)
        self.progress.setRange(0, 100)
        if success:
            self.set_status("Connection test passed.")
            QMessageBox.information(self, "Connection Test", message)
        else:
            self.set_status("Connection test failed.")
            QMessageBox.warning(self, "Connection Test", message)

    def open_log_viewer(self) -> None:
        viewer = LogViewer(self, self.log_file)
        viewer.exec()

    def open_config_editor(self) -> None:
        editor = ConfigEditor(self, self.config)
        editor.exec()

    def open_logs_folder(self) -> None:
        try:
            log_dir = Path(self.log_file).parent
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(log_dir)))
            self.set_status("Opened logs folder")
        except Exception as exc:
            self.set_status(f"Error: {exc}")

    def show_about(self) -> None:
        about_text = f"""{APP_TITLE}

Version: 1.0

Description:
Upload Excel data to SharePoint via an Access bridge with validation, preview, and logging."""
        QMessageBox.information(self, "About", about_text)


def main() -> None:
    app = QApplication(sys.argv)

    splash = create_splash_screen()

    try:
        window = UploadApp(config)

        def on_splash_finished():
            splash.close()
            window.show()

        splash.finished.connect(on_splash_finished)

        sys.exit(app.exec())
    except Exception as exc:
        logger.exception("Fatal error during application startup.")
        splash.close()
        QMessageBox.critical(None, "Fatal Error", str(exc))
        sys.exit(1)


if __name__ == "__main__":
    main()