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
def _is_truthy(value) -> bool:
    """Interpret a value as boolean truthiness.

    Supports str, bool, int and other scalar values. If dict is passed, checks nested values gracefully.
    """
    if isinstance(value, dict):
        value = value.get("enabled", "false") if "enabled" in value else ""

    try:
        return str(value).strip().lower() in {"1", "true", "yes", "y", "on"}
    except Exception:
        return False


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