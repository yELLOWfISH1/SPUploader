import logging
import getpass
import socket
from pathlib import Path
from typing import Optional

from PySide6.QtCore import QUrl, QDate, QTimer
from PySide6.QtGui import QDesktopServices, QIcon
from PySide6.QtWidgets import (
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
    QTableWidget,
    QTableWidgetItem,
    QCheckBox,
    QPlainTextEdit,
    QDateEdit,
    QListWidget,
    QListWidgetItem,
)
import pandas as pd
from datetime import date, timedelta

from .access_bridge import AccessBridge, AccessBridgeError, AuthenticationError, PermissionError
from .config_manager import save_config
from .ui_workers import UploadWorker, ConnectionTestWorker, AuthenticationWorker, MacroWorker, PreviewWorker
from .ui_dialogs import LogViewer, ConfigEditor, ValidationErrorsDialog
from .validator import (
    load_excel,
    validate_dataframe,
    collect_validation_errors,
    FileNotFoundError as ValidatorFileNotFoundError,
    InvalidFormatError,
)

logger = logging.getLogger()


class UploadApp(QMainWindow):
    def __init__(self, config_obj, audit_logger, log_file: str):
        super().__init__()
        self.config = config_obj
        self.audit_logger = audit_logger

        self.db_path = config_obj["eol"]["db_path"]
        self.macro_name = config_obj["eol"]["macro_name"]
        self.delete_macro_name = config_obj["eol"]["delete_macro_name"]
        self.eol_preview_query = config_obj["eol"]["preview_query_incomplete"]
        self.eol_preview_limit = self._parse_preview_limit(
            config_obj["eol"].get("preview_row_limit", "50")
        )
        self.linked_table = config_obj["eol"]["linked_table"]
        self.temp_table = config_obj["eol"]["temp_table"]

        self.archive_db_path = config_obj["archive"]["db_path"]
        self.archive_macros = {
            "New Starters Transfer": config_obj["archive"]["macro_new_starters"],
            "Leavers Transfer": config_obj["archive"]["macro_leavers"],
            "Mat Leavers Transfer": config_obj["archive"]["macro_mat_leavers"],
            "Transfers Transfer": config_obj["archive"]["macro_transfers"],
        }
        self.archive_preview_queries = {
            "New Starters Transfer": config_obj["archive"]["preview_query_new_starters"],
            "Leavers Transfer": config_obj["archive"]["preview_query_leavers"],
            "Mat Leavers Transfer": config_obj["archive"]["preview_query_mat_leavers"],
            "Transfers Transfer": config_obj["archive"]["preview_query_transfers"],
        }
        self.archive_preview_limit = self._parse_preview_limit(
            config_obj["archive"].get("preview_row_limit", "50")
        )

        self.required_columns = [
            c.strip() for c in config_obj["validation"]["required_columns"].split(",") if c.strip()
        ]

        self.log_file = log_file

        if config_obj.has_section("new_starters"):
            new_starters_cfg = config_obj["new_starters"]
        else:
            new_starters_cfg = {}

        self.new_starters_db_path = new_starters_cfg.get("db_path", self.db_path)
        self.new_starters_macro_name = new_starters_cfg.get("macro_name", "")
        self.new_starters_temp_table = new_starters_cfg.get("temp_table", self.temp_table)
        self.new_starters_linked_table = new_starters_cfg.get("linked_table", self.linked_table)

        self.london_names_saved = new_starters_cfg.get("london_names", "")
        self.dublin_names_saved = new_starters_cfg.get("dublin_names", "")

        self.new_starters_df = None
        self.upload_worker = None
        self.connection_worker = None
        self.macro_worker = None
        self.preview_worker = None
        self.pending_upload_path = ""
        self.pending_upload_row_count = 0

        self.setWindowTitle("SharePoint Upload Tool")
        self.setGeometry(100, 100, 1200, 800)
        self.setMinimumSize(1000, 700)

        icon_path = Path(__file__).resolve().parent.parent / "icons" / "scotiabank_logo_icon_170755.png"
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
        tools_menu.addAction("Refresh Linked Table", self.refresh_linked_table)
        tools_menu.addAction("Recreate Linked Table", self.recreate_linked_table)
        tools_menu.addAction("Open Access for Sign-in", self.open_access_for_signin)
        tools_menu.addAction("View Log", self.open_log_viewer)
        # tools_menu.addAction("Edit Config", self.open_config_editor)
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
        tabs.addTab(self.create_upload_tab(), "End Of Life Uploader")
        tabs.addTab(self.create_archive_tab(), "Archiver")
        tabs.addTab(self.create_new_starters_tab(), "New Starters Schedule")
        main_layout.addWidget(tabs, 1)

        footer = self.create_footer()
        main_layout.addWidget(footer)

        # Initial SharePoint authentication check at startup (non-blocking)
        QTimer.singleShot(100, self.check_sharepoint_authentication)

        central_widget.setLayout(main_layout)

    def create_header(self) -> QFrame:
        header = QFrame()
        header.setObjectName("header")
        layout = QVBoxLayout()
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(6)

        title = QLabel("SharePoint Upload Tool")
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

    def set_status(self, text: str) -> None:
        if hasattr(self, "status_label") and self.status_label is not None:
            self.status_label.setText(text)
        logger.debug(f"Status: {text}")

    def auto_save_new_starters_names(self) -> None:
        if not self.config.has_section("new_starters"):
            self.config.add_section("new_starters")

        london_text = self.london_names_box.toPlainText().strip()
        dublin_text = self.dublin_names_box.toPlainText().strip()

        self.config.set("new_starters", "london_names", london_text)
        self.config.set("new_starters", "dublin_names", dublin_text)

        try:
            save_config(self.config)
            self.set_status("New Starters names saved.")
        except Exception as exc:
            logger.exception("Failed to auto-save New Starters names.")
            self.set_status(f"Error saving names: {exc}")

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

        delete_btn = QPushButton("Delete All")
        delete_btn.clicked.connect(self.confirm_delete_all)
        action_layout.addWidget(delete_btn)

        upload_btn = QPushButton("Upload to SharePoint")
        upload_btn.clicked.connect(self.start_upload)
        action_layout.addWidget(upload_btn)

        action_layout.addStretch()
        layout.addLayout(action_layout)

        self.dry_run_checkbox = QCheckBox("Dry run (validate + preview only)")
        layout.addWidget(self.dry_run_checkbox)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        layout.addStretch()
        container.setLayout(layout)
        return container

    def create_archive_tab(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        title = QLabel("Transfer SharePoint data to the archive database")
        title.setObjectName("section-title")
        layout.addWidget(title)

        info = QLabel(
            "Archive completed items from the Europe Helpdesk SharePoint lists used by the Service Desk. "
            "Each action writes completed items to the local Access archive database and then removes them "
            "from the source list."
        )
        info.setStyleSheet("color: #ffffff; font-size: 11px; font-weight: bold;")
        info.setWordWrap(True)
        layout.addWidget(info)

        sections = [
            (
                "New Starters Transfer",
                "Moves completed New Starters items from the Europe Helpdesk list into the archive database "
                "and removes them from SharePoint.",
            ),
            (
                "Leavers Transfer",
                "Moves completed Leavers items from the Europe Helpdesk list into the archive database and "
                "removes them from SharePoint.",
            ),
            (
                "Mat Leavers Transfer",
                "Moves completed MAT Leavers items from the Europe Helpdesk list into the archive database "
                "and removes them from SharePoint.",
            ),
            (
                "Transfers Transfer",
                "Moves completed Transfers items from the Europe Helpdesk list into the archive database "
                "and removes them from SharePoint.",
            ),
        ]

        for label, description in sections:
            section_layout = QVBoxLayout()

            header = QLabel(label)
            header.setObjectName("section-title")
            section_layout.addWidget(header)

            paragraph = QLabel(description)
            paragraph.setStyleSheet("color: #999999; font-size: 10px;")
            paragraph.setWordWrap(True)
            section_layout.addWidget(paragraph)

            action_layout = QHBoxLayout()
            preview_btn = QPushButton("Preview")
            preview_btn.clicked.connect(
                lambda checked=False, action=label: self.start_preview_action(action)
            )
            action_layout.addWidget(preview_btn)
            btn = QPushButton(label)
            btn.clicked.connect(lambda checked=False, action=label: self.confirm_archive_transfer(action))
            action_layout.addWidget(btn)
            action_layout.addStretch()
            section_layout.addLayout(action_layout)

            layout.addLayout(section_layout)

        self.archive_progress = QProgressBar()
        self.archive_progress.setVisible(False)
        layout.addWidget(self.archive_progress)

        layout.addStretch()
        container.setLayout(layout)
        return container

    def create_new_starters_tab(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        title = QLabel("New Starters Schedule")
        title.setObjectName("section-title")
        layout.addWidget(title)

        info = QLabel(
            "Create a weekly schedule for new starters and upload it to SharePoint via Access."
        )
        info.setStyleSheet("color: #999999; font-size: 11px;")
        info.setWordWrap(True)
        layout.addWidget(info)

        london_label = QLabel("London Names (one per line):")
        layout.addWidget(london_label)
        self.london_names_box = QPlainTextEdit()
        self.london_names_box.setPlaceholderText("Jack\nLouise\nJoe")
        self.london_names_box.setPlainText(self.london_names_saved)
        self.london_names_box.setFixedHeight(100)
        self.london_names_box.textChanged.connect(self.auto_save_new_starters_names)
        layout.addWidget(self.london_names_box)

        dublin_label = QLabel("Dublin Names (one per line):")
        layout.addWidget(dublin_label)
        self.dublin_names_box = QPlainTextEdit()
        self.dublin_names_box.setPlaceholderText("Ann\nJohn")
        self.dublin_names_box.setPlainText(self.dublin_names_saved)
        self.dublin_names_box.setFixedHeight(100)
        self.dublin_names_box.textChanged.connect(self.auto_save_new_starters_names)
        layout.addWidget(self.dublin_names_box)

        dates_layout = QHBoxLayout()
        self.start_date_edit = QDateEdit(QDate.currentDate())
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setMinimumDate(QDate.currentDate())
        dates_layout.addWidget(QLabel("Start Date:"))
        dates_layout.addWidget(self.start_date_edit)

        self.end_date_edit = QDateEdit(QDate.currentDate().addMonths(3))
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setMinimumDate(QDate.currentDate().addMonths(3))
        dates_layout.addWidget(QLabel("End Date:"))
        dates_layout.addWidget(self.end_date_edit)

        layout.addLayout(dates_layout)

        action_layout = QHBoxLayout()
        gen_btn = QPushButton("Generate Schedule")
        gen_btn.clicked.connect(self.generate_new_starters_schedule)
        action_layout.addWidget(gen_btn)

        export_btn = QPushButton("Export to Excel")
        export_btn.clicked.connect(self.export_new_starters_schedule)
        action_layout.addWidget(export_btn)

        upload_btn = QPushButton("Upload to SharePoint")
        upload_btn.clicked.connect(self.upload_new_starters_schedule)
        action_layout.addWidget(upload_btn)

        layout.addLayout(action_layout)

        self.new_starters_preview = QTableWidget()
        self.new_starters_preview.setColumnCount(3)
        self.new_starters_preview.setHorizontalHeaderLabels(["Name", "Week Start", "Week End"])
        self.new_starters_preview.setVisible(False)
        layout.addWidget(self.new_starters_preview)

        layout.addStretch()
        container.setLayout(layout)
        return container

    def _get_new_starter_names(self, box: QPlainTextEdit):
        raw_text = box.toPlainText().strip()
        names = [n.strip() for n in raw_text.splitlines() if n.strip()]
        return list(dict.fromkeys(names))

    def generate_new_starters_schedule(self) -> None:
        london_names = self._get_new_starter_names(self.london_names_box)
        dublin_names = self._get_new_starter_names(self.dublin_names_box)

        if not london_names:
            QMessageBox.warning(self, "Input Error", "Please provide at least one London name.")
            return

        if not dublin_names:
            QMessageBox.warning(self, "Input Error", "Please provide at least one Dublin name.")
            return

        start_qdate = self.start_date_edit.date()
        end_qdate = self.end_date_edit.date()
        start = start_qdate.toPython()
        end = end_qdate.toPython()

        if start < date.today():
            QMessageBox.warning(self, "Date Error", "Start date must be today or in the future.")
            return

        if end < date.today() + timedelta(days=90):
            QMessageBox.warning(self, "Date Error", "End date must be at least 3 months from today.")
            return

        if end < start:
            QMessageBox.warning(self, "Date Error", "End date must be after start date.")
            return

        # Align the first week to Monday -> Sunday
        if start.weekday() != 0:
            offset_days = 7 - start.weekday()
            start = start + timedelta(days=offset_days)

        rows = []
        current_start = start
        idx = 0

        while current_start <= end:
            week_end = current_start + timedelta(days=6)
            if week_end > end:
                week_end = end

            london = london_names[idx % len(london_names)]
            dublin = dublin_names[idx % len(dublin_names)]
            combined_name = f"{london} / {dublin}"

            rows.append(
                {
                    "Name": combined_name,
                    "Week Start": current_start.isoformat(),
                    "Week End": week_end.isoformat(),
                }
            )

            idx += 1
            current_start = week_end + timedelta(days=1)

        self.new_starters_df = pd.DataFrame(rows)
        self._display_new_starters_preview(self.new_starters_df)
        self.set_status(f"Generated schedule: {len(rows)} weeks")

    def _display_new_starters_preview(self, df):
        self.new_starters_preview.setVisible(True)
        self.new_starters_preview.setRowCount(len(df))
        self.new_starters_preview.setColumnCount(len(df.columns))
        self.new_starters_preview.setHorizontalHeaderLabels(list(df.columns))
        for r, row in df.iterrows():
            for c, value in enumerate(row):
                self.new_starters_preview.setItem(r, c, QTableWidgetItem(str(value)))
        self.new_starters_preview.resizeColumnsToContents()

    def export_new_starters_schedule(self) -> None:
        if self.new_starters_df is None or self.new_starters_df.empty:
            QMessageBox.warning(self, "Export Error", "No schedule generated to export.")
            return

        excel_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save New Starters Schedule",
            "new_starters_schedule.xlsx",
            "Excel Files (*.xlsx)",
        )
        if not excel_path:
            return

        try:
            self.new_starters_df.to_excel(excel_path, index=False)
            self.set_status(f"Schedule exported: {excel_path}")
            QMessageBox.information(self, "Exported", f"Schedule exported to {excel_path}")
        except Exception as exc:
            logger.exception("Excel export failed.")
            QMessageBox.warning(self, "Export Error", f"Failed to save Excel file:\n{exc}")

    def upload_new_starters_schedule(self) -> None:
        if self.new_starters_df is None or self.new_starters_df.empty:
            QMessageBox.warning(self, "Upload Error", "Generate the schedule before uploading.")
            return

        default_path = Path.cwd() / "new_starters_schedule.xlsx"
        excel_path = str(default_path)
        try:
            self.new_starters_df.to_excel(excel_path, index=False)
        except Exception as exc:
            logger.exception("Excel temp save failed.")
            QMessageBox.warning(self, "Upload Error", f"Failed to prepare upload file:\n{exc}")
            return

        reply = QMessageBox.question(
            self,
            "Confirm Upload",
            f"Upload generated schedule ({len(self.new_starters_df)} rows)?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        self.set_status("Starting New Starters upload...")
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)

        try:
            with AccessBridge(
                db_path=self.new_starters_db_path,
                linked_table=self.new_starters_linked_table,
                temp_table=self.new_starters_temp_table,
                macro_name=self.new_starters_macro_name,
            ) as bridge:
                try:
                    bridge.ensure_authenticated(timeout_seconds=120, poll_interval=5)
                except PermissionError as exc:
                    QMessageBox.warning(self, "Permission Error", str(exc))
                    raise
                except Exception as exc:
                    QMessageBox.warning(self, "Authentication Error", "Authentication failed. Please sign in via Microsoft login and retry.")
                    raise

                bridge.clear_temp_table()
                bridge.import_excel_to_temp(excel_path)
                bridge.run_macro()

            self.set_status("New Starters upload completed.")
            QMessageBox.information(self, "Success", "New Starters schedule uploaded successfully.")
        except Exception as exc:
            logger.exception("New Starters upload failed.")
            self.set_status("New Starters upload failed.")
            QMessageBox.warning(self, "Upload Error", f"Upload failed:\n{exc}")
        finally:
            self.progress.setVisible(False)
            self.progress.setRange(0, 100)

    def check_sharepoint_authentication(self) -> None:
        self.set_status("Checking SharePoint authentication...")
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)

        self.auth_worker = AuthenticationWorker(self.config, interactive=False)
        self.auth_worker.finished.connect(self.on_auth_check_finished)
        self.auth_worker.start()

    def on_auth_check_finished(self, success: bool, message: str) -> None:
        self.progress.setVisible(False)
        self.progress.setRange(0, 100)

        if success:
            self.set_status("SharePoint authentication OK.")
            QMessageBox.information(self, "SharePoint Auth", "SharePoint authentication succeeded.")
            return

        self.set_status("SharePoint authentication failed.")
        logger.warning(f"Auth check failed: {message}")

        result = QMessageBox.question(
            self,
            "Authentication Required",
            "Could not authenticate silently. Do you want to try interactive authentication with Access visible?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )

        if result == QMessageBox.Yes:
            self.authenticate_interactively()
        else:
            QMessageBox.warning(
                self,
                "Authentication Required",
                "Authentication is required to use the uploader. Please sign in via Access and retry."
            )

    def authenticate_interactively(self) -> None:
        self.set_status("Starting interactive authentication...")
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)

        self.auth_worker = AuthenticationWorker(self.config, interactive=True)
        self.auth_worker.finished.connect(self.on_interactive_auth_finished)
        self.auth_worker.start()

    def on_interactive_auth_finished(self, success: bool, message: str) -> None:
        self.progress.setVisible(False)
        self.progress.setRange(0, 100)

        if success:
            self.set_status("Interactive authentication succeeded.")
            QMessageBox.information(self, "Authenticated", "Interactive authentication completed successfully.")
        else:
            self.set_status("Interactive authentication failed.")
            QMessageBox.warning(self, "Authentication Failed", f"Interactive authentication failed:\n{message}")

    def refresh_linked_table(self) -> None:
        self.set_status("Refreshing SharePoint linked table...")
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)

        self.auth_worker = AuthenticationWorker(self.config, interactive=False)
        self.auth_worker.finished.connect(self.on_refresh_done)
        self.auth_worker.start()

    def on_refresh_done(self, success: bool, message: str) -> None:
        self.progress.setVisible(False)
        self.progress.setRange(0, 100)

        if success:
            self.set_status("Linked table refresh completed.")
            QMessageBox.information(self, "Refresh Complete", "SharePoint linked table has been refreshed silently.")
        else:
            self.set_status("Linked table refresh failed.")
            logger.warning(f"Refresh action failed: {message}")
            QMessageBox.warning(self, "Refresh Failed", f"Could not refresh linked table:\n{message}")

    def open_access_for_signin(self) -> None:
        if hasattr(self, "manual_access_bridge") and self.manual_access_bridge is not None:
            QMessageBox.information(self, "Access Open", "Access is already open for sign-in.")
            return

        try:
            self.manual_access_bridge = AccessBridge(
                db_path=self.db_path,
                linked_table=self.linked_table,
                temp_table=self.temp_table,
                macro_name=self.macro_name,
            )
            self.manual_access_bridge.open()
            self.manual_access_bridge.access.Visible = True
            self.manual_access_bridge.access.UserControl = True
            self.set_status("Access opened for manual sign-in. Complete sign-in and close Access when done.")
            QMessageBox.information(
                self,
                "Manual Authentication",
                "Access window opened. Please sign in manually in the Access UI, then use Refresh or Test Connection."
            )
        except Exception as exc:
            logger.exception("Failed to open Access for manual sign-in.")
            self.set_status("Failed to open Access for manual sign-in.")
            QMessageBox.warning(self, "Error", f"Could not open Access for sign-in: {exc}")

    def recreate_linked_table(self) -> None:
        self.set_status("Recreating SharePoint linked table...")
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)

        try:
            with AccessBridge(
                db_path=self.db_path,
                linked_table=self.linked_table,
                temp_table=self.temp_table,
                macro_name=self.macro_name,
            ) as bridge:
                bridge.recreate_linked_table()
            self.set_status("Linked table recreated.")
            QMessageBox.information(
                self,
                "Recreate Complete",
                "Linked table was recreated. Please run Refresh Linked Table/Test Connection now."
            )
        except Exception as exc:
            logger.exception("Recreate linked table failed.")
            self.set_status("Could not recreate linked table.")
            QMessageBox.warning(self, "Recreate Failed", f"Could not recreate linked table:\n{exc}")

    def closeEvent(self, event):
        if hasattr(self, "manual_access_bridge") and self.manual_access_bridge is not None:
            try:
                self.manual_access_bridge.close()
            except Exception:
                pass
        super().closeEvent(event)

    def closeEvent(self, event):
        if hasattr(self, "manual_access_bridge") and self.manual_access_bridge is not None:
            try:
                self.manual_access_bridge.close()
            except Exception:
                pass
        super().closeEvent(event)

        try:
            with AccessBridge(
                db_path=self.db_path,
                linked_table=self.linked_table,
                temp_table=self.temp_table,
                macro_name=self.macro_name,
            ) as bridge:
                bridge.ensure_authenticated(timeout_seconds=60, poll_interval=2, interactive=False)
                bridge.refresh_linked_table()
            self.set_status("Linked table refresh completed.")
            QMessageBox.information(self, "Refresh Complete", "SharePoint linked table has been refreshed silently.")
        except Exception as exc:
            logger.exception("Linked table refresh failed.")
            self.set_status("Linked table refresh failed.")
            QMessageBox.warning(self, "Refresh Failed", f"Could not refresh linked table:\n{exc}")
        finally:
            self.progress.setVisible(False)
            self.progress.setRange(0, 100)

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
                self._audit("validate_success", f"file={path}")
                QMessageBox.information(self, "Validation Passed", msg)
            else:
                errors = collect_validation_errors(df, self.required_columns)
                self._audit("validate_failed", f"file={path} reason={msg}")
                self._show_validation_failure("Validation Failed", msg, errors, path)
        except ValidatorFileNotFoundError as exc:
            self._audit("validate_failed", f"file={path} error={exc}")
            QMessageBox.warning(self, "File Error", f"Cannot find file:\n{exc}")
        except InvalidFormatError as exc:
            self._audit("validate_failed", f"file={path} error={exc}")
            QMessageBox.warning(self, "Format Error", f"Invalid Excel file:\n{exc}")
        except Exception as exc:
            logger.exception("Validation failed.")
            self._audit("validate_failed", f"file={path} error={exc}")
            QMessageBox.warning(self, "Error", f"Validation failed:\n{exc}")

    def preview_data(self) -> None:
        path = self.file_input.text().strip()
        if not self._validate_file_path(path):
            return

        try:
            df = load_excel(path)
            ok, msg = validate_dataframe(df, self.required_columns)
            if not ok:
                errors = collect_validation_errors(df, self.required_columns)
                self._audit("preview_failed", f"file={path} reason={msg}")
                self._show_validation_failure("Cannot Preview", msg, errors, path)
                return

            self._show_dataframe_preview(df, "Data Preview")
            self._audit("preview_success", f"file={path}")
        except ValidatorFileNotFoundError as exc:
            self._audit("preview_failed", f"file={path} error={exc}")
            QMessageBox.warning(self, "File Error", f"Cannot find file:\n{exc}")
        except InvalidFormatError as exc:
            self._audit("preview_failed", f"file={path} error={exc}")
            QMessageBox.warning(self, "Format Error", f"Invalid Excel file:\n{exc}")
        except Exception as exc:
            logger.exception("Preview failed.")
            self._audit("preview_failed", f"file={path} error={exc}")
            QMessageBox.warning(self, "Error", f"Preview failed:\n{exc}")

    def start_upload(self) -> None:
        path = self.file_input.text().strip()
        if not self._validate_file_path(path):
            self.set_status("Upload cancelled.")
            return

        row_count = self._get_upload_row_count(path)
        if row_count is None:
            return

        if self.dry_run_checkbox.isChecked():
            self._audit("dry_run_start", f"file={path} rows={row_count}")
            self._run_dry_run(path)
            return

        self._audit("upload_prompt", f"file={path} rows={row_count}")

        reply = QMessageBox.question(
            self,
            "Confirm Upload",
            "This will upload new data for this year EOL. Did you want to delete the data in the list first?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            self.pending_upload_path = path
            self.pending_upload_row_count = row_count
            self.confirm_delete_all()
            return

        self._confirm_upload_and_start(path, row_count)

    def _run_dry_run(self, path: str) -> None:
        try:
            df = load_excel(path)
            ok, msg = validate_dataframe(df, self.required_columns)
            if not ok:
                errors = collect_validation_errors(df, self.required_columns)
                self._audit("dry_run_failed", f"file={path} reason={msg}")
                self._show_validation_failure("Dry Run Failed", msg, errors, path)
                return

            self._show_dataframe_preview(df, "Dry Run Preview")
            self._audit("dry_run_success", f"file={path} rows={len(df)}")
        except ValidatorFileNotFoundError as exc:
            self._audit("dry_run_failed", f"file={path} error={exc}")
            QMessageBox.warning(self, "File Error", f"Cannot find file:\n{exc}")
        except InvalidFormatError as exc:
            self._audit("dry_run_failed", f"file={path} error={exc}")
            QMessageBox.warning(self, "Format Error", f"Invalid Excel file:\n{exc}")
        except Exception as exc:
            logger.exception("Dry run failed.")
            self._audit("dry_run_failed", f"file={path} error={exc}")
            QMessageBox.warning(self, "Error", f"Dry run failed:\n{exc}")

    def _show_dataframe_preview(self, df, title: str) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.resize(900, 500)

        layout = QVBoxLayout()

        info = QLabel(
            f"Rows: {len(df)}  |  Columns: {len(df.columns)}  |  Limit: 100"
        )
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

    def _get_upload_row_count(self, path: str):
        try:
            df = load_excel(path)
            return len(df)
        except ValidatorFileNotFoundError as exc:
            QMessageBox.warning(self, "File Error", f"Cannot find file:\n{exc}")
        except InvalidFormatError as exc:
            QMessageBox.warning(self, "Format Error", f"Invalid Excel file:\n{exc}")
        except Exception as exc:
            logger.exception("Failed to read Excel file.")
            QMessageBox.warning(self, "Error", f"Failed to read Excel file:\n{exc}")
        return None

    def _confirm_upload_and_start(self, path: str, row_count: int) -> None:
        reply = QMessageBox.question(
            self,
            "Confirm Upload",
            f"Upload will start with {row_count} rows from the Excel file. Continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            self._audit("upload_cancelled", f"file={path} rows={row_count}")
            self.set_status("Upload cancelled.")
            return

        self._audit("upload_start", f"file={path} rows={row_count}")
        self._begin_upload(path)

    def _begin_upload(self, path: str) -> None:
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.set_status("Starting upload...")

        self.upload_worker = UploadWorker(path, self.config, self.required_columns)
        self.upload_worker.status.connect(self.set_status)
        self.upload_worker.finished.connect(self.on_upload_finished)
        self.upload_worker.start()

    def confirm_delete_all(self) -> None:
        if not self.eol_preview_query:
            QMessageBox.warning(self, "Config Error", "Preview query is missing in config.ini.")
            return

        if not self.db_path or not Path(self.db_path).exists():
            QMessageBox.warning(self, "Config Error", f"Database not found: {self.db_path}")
            return

        self._audit("delete_preview_start", f"query={self.eol_preview_query}")

        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.set_status("Loading incomplete devices...")

        self.preview_worker = PreviewWorker(
            self.db_path,
            self.eol_preview_query,
            self.eol_preview_limit,
            "Delete All",
        )
        self.preview_worker.finished.connect(self.on_delete_preview_finished)
        self.preview_worker.start()

    def on_delete_preview_finished(
        self,
        success: bool,
        message: str,
        columns,
        rows,
    ) -> None:
        self.progress.setVisible(False)
        self.progress.setRange(0, 100)
        if not success:
            self.set_status("Delete preview failed.")
            QMessageBox.warning(self, "Preview Failed", message)
            return

        self.set_status("Delete preview ready.")
        if len(rows) == 0:
            self._audit("delete_preview_empty", "no rows")
            self._proceed_delete_all()
            return

        self.show_delete_preview_dialog(columns, rows)

    def show_delete_preview_dialog(self, columns, rows) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Delete All Preview")
        dialog.resize(900, 520)

        layout = QVBoxLayout()

        warning = QLabel("There are machines on the list that are still not completed yet.")
        warning.setStyleSheet("color: #ffffff; font-size: 11px; font-weight: bold;")
        warning.setWordWrap(True)
        layout.addWidget(warning)

        info = QLabel(
            f"Rows: {len(rows)}  |  Columns: {len(columns)}  |  Limit: {self.eol_preview_limit}"
        )
        info.setStyleSheet("color: #999999; font-size: 10px;")
        layout.addWidget(info)

        if columns:
            table = QTableWidget()
            table.setAlternatingRowColors(True)
            table.setRowCount(len(rows))
            table.setColumnCount(len(columns))
            table.setHorizontalHeaderLabels([str(col) for col in columns])
            for row_idx, row in enumerate(rows):
                for col_idx, value in enumerate(row):
                    item = QTableWidgetItem(str(value)[:80])
                    table.setItem(row_idx, col_idx, item)

            table.resizeColumnsToContents()
            layout.addWidget(table)
        else:
            empty = QLabel("No incomplete devices found.")
            empty.setStyleSheet("color: #999999; font-size: 10px;")
            layout.addWidget(empty)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        proceed_btn = QPushButton("Proceed with Delete")
        proceed_btn.clicked.connect(
            lambda: self._proceed_delete_all(dialog)
        )
        btn_layout.addWidget(proceed_btn)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        dialog.setLayout(layout)
        dialog.exec()

    def _proceed_delete_all(self, dialog: Optional[QDialog] = None) -> None:
        if dialog is not None:
            dialog.accept()
        self._audit("delete_all_start", "")
        self.start_macro_action(
            db_path=self.db_path,
            macro_name=self.delete_macro_name,
            action_label="Delete All",
            progress_bar=self.progress,
        )

    def confirm_archive_transfer(self, action_label: str) -> None:
        reply = QMessageBox.question(
            self,
            "Confirm Transfer",
            "This will archive items and delete them from SharePoint. Continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            self._audit("archive_confirmed", f"action={action_label}")
            macro_name = self.archive_macros.get(action_label, "")
            self.start_macro_action(
                db_path=self.archive_db_path,
                macro_name=macro_name,
                action_label=action_label,
                progress_bar=self.archive_progress,
            )
        else:
            self._audit("archive_cancelled", f"action={action_label}")

    def start_macro_action(
        self,
        db_path: str,
        macro_name: str,
        action_label: str,
        progress_bar: QProgressBar,
    ) -> None:
        if not db_path or not Path(db_path).exists():
            QMessageBox.warning(self, "Config Error", f"Database not found: {db_path}")
            return

        if not macro_name:
            QMessageBox.warning(self, "Config Error", "Macro name is missing in config.ini.")
            return

        progress_bar.setVisible(True)
        progress_bar.setRange(0, 0)
        self.set_status(f"{action_label} started...")

        self.macro_worker = MacroWorker(db_path, macro_name, action_label)
        self.macro_worker.status.connect(self.set_status)
        self.macro_worker.finished.connect(
            lambda success, message: self.on_macro_finished(
                success,
                message,
                progress_bar,
                action_label,
            )
        )
        self.macro_worker.start()

    def start_preview_action(self, action_label: str) -> None:
        query_name = self.archive_preview_queries.get(action_label, "")
        if not query_name:
            QMessageBox.warning(self, "Config Error", "Preview query name is missing in config.ini.")
            return

        if not self.archive_db_path or not Path(self.archive_db_path).exists():
            QMessageBox.warning(
                self, "Config Error", f"Database not found: {self.archive_db_path}"
            )
            return

        self.archive_progress.setVisible(True)
        self.archive_progress.setRange(0, 0)
        self.set_status(f"Loading preview for {action_label}...")
        self._audit("archive_preview_start", f"action={action_label} query={query_name}")

        self.preview_worker = PreviewWorker(
            self.archive_db_path,
            query_name,
            self.archive_preview_limit,
            action_label,
        )
        self.preview_worker.finished.connect(
            lambda success, message, columns, rows: self.on_preview_finished(
                success,
                message,
                columns,
                rows,
                action_label,
            )
        )
        self.preview_worker.start()

    def on_preview_finished(
        self,
        success: bool,
        message: str,
        columns,
        rows,
        action_label: str,
    ) -> None:
        self.archive_progress.setVisible(False)
        self.archive_progress.setRange(0, 100)
        if not success:
            self.set_status(f"Preview failed for {action_label}.")
            QMessageBox.warning(self, "Preview Failed", message)
            self._audit("archive_preview_failed", f"action={action_label} error={message}")
            return

        self.set_status(f"Preview ready for {action_label}.")
        self._audit("archive_preview_success", f"action={action_label} rows={len(rows)}")
        self.show_preview_dialog(action_label, columns, rows)

    def show_preview_dialog(self, action_label: str, columns, rows) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Preview - {action_label}")
        dialog.resize(900, 500)

        layout = QVBoxLayout()

        info = QLabel(
            f"Rows: {len(rows)}  |  Columns: {len(columns)}  |  Limit: {self.archive_preview_limit}"
        )
        info.setStyleSheet("color: #999999; font-size: 10px;")
        layout.addWidget(info)

        if columns:
            table = QTableWidget()
            table.setAlternatingRowColors(True)
            table.setRowCount(len(rows))
            table.setColumnCount(len(columns))
            table.setHorizontalHeaderLabels([str(col) for col in columns])
            for row_idx, row in enumerate(rows):
                for col_idx, value in enumerate(row):
                    item = QTableWidgetItem(str(value)[:80])
                    table.setItem(row_idx, col_idx, item)

            table.resizeColumnsToContents()
            layout.addWidget(table)
        else:
            empty = QLabel("No data available for this preview.")
            empty.setStyleSheet("color: #999999; font-size: 10px;")
            layout.addWidget(empty)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.close)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

        dialog.setLayout(layout)
        dialog.exec()

    def _parse_preview_limit(self, value: str) -> int:
        try:
            limit = int(value)
            return limit if limit > 0 else 50
        except ValueError:
            return 50

    def on_macro_finished(
        self,
        success: bool,
        message: str,
        progress_bar: QProgressBar,
        action_label: str,
    ) -> None:
        progress_bar.setVisible(False)
        progress_bar.setRange(0, 100)
        if success:
            self.set_status(f"{action_label} completed.")
            QMessageBox.information(self, "Success", message)
            self._audit("macro_success", f"action={action_label}")
            if action_label == "Delete All" and self.pending_upload_path:
                path = self.pending_upload_path
                row_count = self.pending_upload_row_count
                self.pending_upload_path = ""
                self.pending_upload_row_count = 0
                self._confirm_upload_and_start(path, row_count)
        else:
            self.set_status(f"{action_label} failed.")
            QMessageBox.warning(self, "Action Failed", message)
            self._audit("macro_failed", f"action={action_label} error={message}")
            if action_label == "Delete All":
                self.pending_upload_path = ""
                self.pending_upload_row_count = 0

    def _audit(self, action: str, details: str) -> None:
        if not self.audit_logger:
            return
        user = getpass.getuser()
        machine = socket.gethostname()
        detail_text = f" {details}" if details else ""
        self.audit_logger.info(f"user={user} machine={machine} action={action}{detail_text}")

    def on_upload_finished(self, success: bool, message: str, errors) -> None:
        self.progress.setVisible(False)
        self.progress.setRange(0, 100)
        if success:
            self.set_status("Completed successfully.")
            QMessageBox.information(self, "Success", message)
            self._audit("upload_success", "")
        else:
            self.set_status("Upload failed.")
            if errors:
                self._show_validation_failure("Upload Failed", message, errors, self.file_input.text())
            else:
                QMessageBox.warning(self, "Upload Failed", message)
            self._audit("upload_failed", f"error={message}")

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
        about_text = """SharePoint Upload Tool

Version: 1.0

Description:
Upload Excel data to SharePoint via an Access bridge with validation, preview, and logging."""
        QMessageBox.information(self, "About", about_text)

    def _show_validation_failure(
        self,
        title: str,
        message: str,
        errors,
        file_path: str,
    ) -> None:
        dialog = QMessageBox(self)
        dialog.setWindowTitle(title)
        dialog.setIcon(QMessageBox.Warning)
        dialog.setText(message)
        advanced_btn = None
        if errors:
            advanced_btn = dialog.addButton("Advanced Info", QMessageBox.ActionRole)
        dialog.addButton("Close", QMessageBox.RejectRole)
        dialog.exec()

        if advanced_btn and dialog.clickedButton() == advanced_btn:
            default_name = "validation_errors.csv"
            if file_path:
                default_name = str(Path(file_path).with_name(
                    f"{Path(file_path).stem}_validation_errors.csv"
                ))
            ValidationErrorsDialog(self, errors, default_name).exec()
