import os
import sys
import logging
import logging.handlers
from pathlib import Path
from typing import Optional

from PySide6.QtWidgets import QApplication, QMessageBox

SRC_DIR = Path(__file__).resolve().parent / "src"
if SRC_DIR.exists() and str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

try:
    from src.config_manager import load_config, ConfigError
    from src.splash_screen import create_splash_screen
    from src.ui_main import UploadApp
except ModuleNotFoundError:
    from config_manager import load_config, ConfigError # type: ignore
    from splash_screen import create_splash_screen # type: ignore
    from ui_main import UploadApp # type: ignore


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
def _is_truthy(value: str) -> bool:
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def _setup_audit_logger(config_obj) -> Optional[logging.Logger]:
    try:
        audit_config = config_obj["audit"]
    except KeyError:
        return None

    enabled = _is_truthy(audit_config.get("enabled", "true"))
    if not enabled:
        return None

    log_folder = Path(audit_config.get("log_folder", "logs/audit")).expanduser()
    log_file_name = audit_config.get("log_file", "audit.log")

    try:
        log_folder.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        logger.error(f"Failed to create audit log folder: {exc}")
        return None

    test_path = log_folder / ".write_test"
    try:
        with open(test_path, "w", encoding="utf-8") as file_handle:
            file_handle.write("test")
        test_path.unlink()
    except Exception as exc:
        logger.error(f"No write permission for audit log folder: {exc}")
        return None

    audit_logger = logging.getLogger("audit")
    audit_logger.setLevel(logging.INFO)
    audit_logger.propagate = False

    audit_handler = logging.handlers.RotatingFileHandler(
        log_folder / log_file_name,
        maxBytes=1_000_000,
        backupCount=3,
        encoding="utf-8",
    )
    audit_formatter = logging.Formatter("%(asctime)s %(message)s")
    audit_handler.setFormatter(audit_formatter)

    if not audit_logger.handlers:
        audit_logger.addHandler(audit_handler)

    return audit_logger


def main() -> None:
    app = QApplication(sys.argv)

    splash = create_splash_screen()
    audit_logger = _setup_audit_logger(config)

    try:
        window = UploadApp(config, audit_logger, log_file)

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