import time
import logging
import win32com.client as win32
from typing import Optional
from contextlib import contextmanager

logger = logging.getLogger(__name__)

class AccessBridgeError(Exception):
    """Base exception for Access bridge errors."""
    pass

class AuthenticationError(AccessBridgeError):
    """Raised when authentication fails."""
    pass

class PermissionError(AccessBridgeError):
    """Raised when user lacks permissions."""
    pass

class AccessBridge:
    def __init__(self, db_path: str, linked_table: str, temp_table: str, macro_name: str):
        self.db_path = db_path
        self.linked_table = linked_table
        self.temp_table = temp_table
        self.macro_name = macro_name
        self.access: Optional[object] = None

    def __enter__(self):
        """Context manager entry."""
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit with cleanup."""
        self.close()
        return False

    def open(self) -> None:
        """Open the Access database."""
        logger.info(f"Opening Access DB: {self.db_path}")
        try:
            self.access = win32.Dispatch("Access.Application")
            self.access.Visible = False
            self.access.OpenCurrentDatabase(self.db_path)
        except Exception as e:
            logger.error(f"Failed to open Access database: {e}")
            raise AccessBridgeError(f"Failed to open database: {e}") from e

    def close(self) -> None:
        """Close the Access database and clean up resources."""
        if self.access is not None:
            try:
                logger.info("Closing Access.")
                self.access.Quit()
            except Exception as e:
                logger.warning(f"Error closing Access: {e}")
            finally:
                self.access = None

    def _try_open_linked_table(self) -> None:
        """Attempt to open the linked table to test connection."""
        db = self.access.CurrentDb()
        rs = db.OpenRecordset(self.linked_table)
        rs.Close()

    def refresh_linked_table(self) -> None:
        """Attempt to refresh the linked table metadata/connection."""
        logger.info(f"Refreshing linked table: {self.linked_table}")
        try:
            db = self.access.CurrentDb()
            table_def = db.TableDefs(self.linked_table)
            table_def.RefreshLink()
            logger.info("Linked table refreshed successfully.")
        except Exception as e:
            logger.warning(f"Could not refresh linked table: {e}")

    def ensure_authenticated(self, timeout_seconds: int = 300, poll_interval: int = 5) -> None:
        """
        Silently authenticate by attempting to open the linked table.
        If authentication fails, waits for user to authenticate via background process.
        Never shows the Access window to the user.
        
        Args:
            timeout_seconds: Maximum time to wait for authentication
            poll_interval: Time between retry attempts in seconds
            
        Raises:
            PermissionError: If user lacks permission to access SharePoint list
            AuthenticationError: If authentication cannot be completed
        """
        start_time = time.time()

        while True:
            try:
                logger.info("Checking authentication by opening linked table.")
                self._try_open_linked_table()
                logger.info("Authentication successful.")
                return

            except Exception as e:
                msg = str(e).lower()
                elapsed = time.time() - start_time
                logger.debug(f"Auth check failed (attempt after {elapsed:.1f}s): {e}")

                # Check for permission errors
                if any(term in msg for term in ["permission", "not authorized", "access denied", "401", "403"]):
                    logger.error("User does not have permission to access the SharePoint list.")
                    raise PermissionError(
                        "You do not have permission to access the SharePoint list. "
                        "Please contact your SharePoint administrator."
                    ) from e

                # Check for timeout
                if elapsed > timeout_seconds:
                    logger.error("Authentication timeout exceeded.")
                    raise AuthenticationError(
                        "Authentication did not complete in time. "
                        "Please ensure you're logged in and try again."
                    ) from e

                # Keep Access hidden while waiting for background authentication
                # (Access handles SharePoint auth in background with cached credentials)
                self.access.Visible = False
                logger.debug(f"Waiting for authentication (polling in {poll_interval}s)...")
                time.sleep(poll_interval)

    def test_connection(self) -> None:
        """Test if we can connect to the database and linked table."""
        with self:
            try:
                self.ensure_authenticated(timeout_seconds=10, poll_interval=2)
                logger.info("Test connection successful.")
            except (AuthenticationError, PermissionError) as e:
                logger.error(f"Connection test failed: {e}")
                raise

    def import_excel_to_temp(self, excel_path: str) -> None:
        """
        Import an Excel file into the temporary table.
        
        Args:
            excel_path: Path to the Excel file
            
        Raises:
            AccessBridgeError: If import fails
        """
        logger.info(f"Importing Excel into temp table '{self.temp_table}': {excel_path}")
        try:
            self.access.DoCmd.TransferSpreadsheet(
                TransferType=0,          # acImport
                SpreadsheetType=10,      # acSpreadsheetTypeExcel12
                TableName=self.temp_table,
                FileName=excel_path,
                HasFieldNames=True
            )
            logger.info(f"Successfully imported {excel_path} into {self.temp_table}")
        except Exception as e:
            logger.error(f"Excel import failed: {e}")
            raise AccessBridgeError(f"Failed to import Excel file: {e}") from e

    def run_macro(self) -> None:
        """
        Run the append macro in Access.
        
        Raises:
            AccessBridgeError: If macro execution fails
        """
        logger.info(f"Running macro: {self.macro_name}")
        try:
            self.access.DoCmd.RunMacro(self.macro_name)
            logger.info(f"Macro '{self.macro_name}' completed successfully")
        except Exception as e:
            logger.error(f"Macro execution failed: {e}")
            raise AccessBridgeError(f"Failed to run macro: {e}") from e

    def fetch_preview_rows(self, query_name: str, limit: int = 50):
        """
        Fetch rows from a saved query or table for preview.

        Args:
            query_name: Access query or table name
            limit: Maximum number of rows to return

        Returns:
            Tuple of (columns, rows)
        """
        logger.info(f"Fetching preview rows from '{query_name}' (limit {limit})")
        try:
            db = self.access.CurrentDb()
            rs = db.OpenRecordset(query_name)
            columns = [rs.Fields(i).Name for i in range(rs.Fields.Count)]
            rows = []
            count = 0
            if not rs.EOF:
                rs.MoveFirst()
            while not rs.EOF and count < limit:
                row = []
                for i in range(rs.Fields.Count):
                    value = rs.Fields(i).Value
                    row.append("" if value is None else str(value))
                rows.append(row)
                count += 1
                rs.MoveNext()
            rs.Close()
            return columns, rows
        except Exception as e:
            logger.error(f"Preview query failed: {e}")
            raise AccessBridgeError(f"Failed to fetch preview data: {e}") from e

    def clear_temp_table(self) -> None:
        """Clear the temporary import table (for rollback/cleanup)."""
        logger.info(f"Clearing temp table: {self.temp_table}")
        try:
            db = self.access.CurrentDb()
            db.Execute(f"DELETE FROM [{self.temp_table}]")
            logger.info(f"Temp table '{self.temp_table}' cleared")
        except Exception as e:
            logger.warning(f"Could not clear temp table: {e}")