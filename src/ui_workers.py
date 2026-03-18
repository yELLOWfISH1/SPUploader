import logging

from PySide6.QtCore import QThread, Signal

from .access_bridge import AccessBridge, AccessBridgeError, AuthenticationError, PermissionError
from .validator import (
    load_excel,
    validate_dataframe,
    collect_validation_errors,
    FileNotFoundError as ValidatorFileNotFoundError,
    InvalidFormatError,
)

logger = logging.getLogger(__name__)


class UploadWorker(QThread):
    status = Signal(str)
    finished = Signal(bool, str, object)

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
                errors = collect_validation_errors(df, self.required_columns)
                self.finished.emit(False, msg, errors)
                return

            with AccessBridge(
                db_path=self.config_obj["eol"]["db_path"],
                linked_table=self.config_obj["eol"]["linked_table"],
                temp_table=self.config_obj["eol"]["temp_table"],
                macro_name=self.config_obj["eol"]["macro_name"],
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
                    self.status.emit("Clearing temp table...")
                    bridge.clear_temp_table()
                except AccessBridgeError as exc:
                    logger.error(f"Temp table clear error: {exc}")
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

            self.finished.emit(True, "Upload completed successfully.", None)
            logger.info("Upload completed successfully.")

        except ValidatorFileNotFoundError as exc:
            self.finished.emit(False, str(exc), None)
        except InvalidFormatError as exc:
            self.finished.emit(False, str(exc), None)
        except Exception as exc:
            logger.exception("Upload failed.")
            self.finished.emit(False, f"Upload failed: {exc}", None)


class ConnectionTestWorker(QThread):
    finished = Signal(bool, str)

    def __init__(self, config_obj):
        super().__init__()
        self.config_obj = config_obj

    def run(self) -> None:
        bridge = AccessBridge(
            db_path=self.config_obj["eol"]["db_path"],
            linked_table=self.config_obj["eol"]["linked_table"],
            temp_table=self.config_obj["eol"]["temp_table"],
            macro_name=self.config_obj["eol"]["macro_name"],
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


class AuthenticationWorker(QThread):
    finished = Signal(bool, str)

    def __init__(self, config_obj, interactive=False):
        super().__init__()
        self.config_obj = config_obj
        self.interactive = interactive

    def run(self) -> None:
        with AccessBridge(
            db_path=self.config_obj["eol"]["db_path"],
            linked_table=self.config_obj["eol"]["linked_table"],
            temp_table=self.config_obj["eol"]["temp_table"],
            macro_name=self.config_obj["eol"]["macro_name"],
        ) as bridge:
            try:
                bridge.ensure_authenticated(timeout_seconds=35, poll_interval=3, interactive=self.interactive)
                try:
                    bridge.refresh_linked_table()
                except Exception as refresh_exc:
                    logger.warning(f"Linked table refresh during auth worker failed: {refresh_exc}")
                self.finished.emit(True, "Authentication successful.")
            except PermissionError as exc:
                self.finished.emit(False, str(exc))
            except AuthenticationError as exc:
                self.finished.emit(False, str(exc))
            except Exception as exc:
                self.finished.emit(False, str(exc))


class MacroWorker(QThread):
    status = Signal(str)
    finished = Signal(bool, str)

    def __init__(self, db_path: str, macro_name: str, action_label: str):
        super().__init__()
        self.db_path = db_path
        self.macro_name = macro_name
        self.action_label = action_label

    def run(self) -> None:
        self.status.emit(f"{self.action_label} in progress...")
        logger.info(f"Starting macro action: {self.action_label}")

        try:
            with AccessBridge(
                db_path=self.db_path,
                linked_table="",
                temp_table="",
                macro_name=self.macro_name,
            ) as bridge:
                bridge.run_macro()
            self.finished.emit(True, f"{self.action_label} completed successfully.")
        except AccessBridgeError as exc:
            logger.error(f"Macro action failed: {exc}")
            self.finished.emit(False, str(exc))
        except Exception as exc:
            logger.exception("Macro action failed.")
            self.finished.emit(False, f"{self.action_label} failed: {exc}")


class PreviewWorker(QThread):
    finished = Signal(bool, str, object, object)

    def __init__(self, db_path: str, query_name: str, limit: int, action_label: str):
        super().__init__()
        self.db_path = db_path
        self.query_name = query_name
        self.limit = limit
        self.action_label = action_label

    def run(self) -> None:
        try:
            with AccessBridge(
                db_path=self.db_path,
                linked_table="",
                temp_table="",
                macro_name="",
            ) as bridge:
                columns, rows = bridge.fetch_preview_rows(self.query_name, self.limit)
            self.finished.emit(True, "Preview loaded.", columns, rows)
        except AccessBridgeError as exc:
            logger.error(f"Preview failed: {exc}")
            self.finished.emit(False, str(exc), [], [])
        except Exception as exc:
            logger.exception("Preview failed.")
            self.finished.emit(False, f"Preview failed: {exc}", [], [])
