import os
from pathlib import Path
from typing import Dict

from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QPlainTextEdit,
    QPushButton,
    QLineEdit,
    QLabel,
    QFormLayout,
    QScrollArea,
    QWidget,
    QMessageBox,
    QFileDialog,
    QTableWidget,
    QTableWidgetItem,
)

from .config_manager import save_config
from .validator import save_validation_errors_csv


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

        for section in self.config_obj.sections():
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


class ValidationErrorsDialog(QDialog):
    def __init__(self, parent, errors, default_name: str):
        super().__init__(parent)
        self.setWindowTitle("Validation Errors")
        self.resize(800, 500)
        self.errors = errors or []
        self.default_name = default_name

        layout = QVBoxLayout()

        info = QLabel(f"Total errors: {len(self.errors)}")
        info.setStyleSheet("color: #999999; font-size: 10px;")
        layout.addWidget(info)

        table = QTableWidget()
        table.setAlternatingRowColors(True)
        table.setRowCount(len(self.errors))
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Row", "Column", "Reason"])
        for row_idx, error in enumerate(self.errors):
            table.setItem(row_idx, 0, QTableWidgetItem(str(error.get("row_number", ""))))
            table.setItem(row_idx, 1, QTableWidgetItem(str(error.get("column", ""))))
            table.setItem(row_idx, 2, QTableWidgetItem(str(error.get("reason", ""))))
        table.resizeColumnsToContents()
        layout.addWidget(table)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        export_btn = QPushButton("Export CSV")
        export_btn.clicked.connect(self.export_csv)
        btn_layout.addWidget(export_btn)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.close)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def export_csv(self) -> None:
        if not self.errors:
            QMessageBox.information(self, "Export", "No validation errors to export.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Validation Errors",
            self.default_name,
            "CSV Files (*.csv)",
        )
        if not file_path:
            return

        try:
            save_validation_errors_csv(self.errors, Path(file_path))
            QMessageBox.information(self, "Export", f"Saved: {file_path}")
        except Exception as exc:
            QMessageBox.warning(self, "Export Failed", f"Failed to save CSV:\n{exc}")
