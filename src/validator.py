import csv
import logging
import pandas as pd
from typing import Tuple, List, Dict
from pathlib import Path

logger = logging.getLogger(__name__)

class ValidationError(Exception):
    """Base exception for validation errors."""
    pass

class FileNotFoundError(ValidationError):
    """Raised when file cannot be found."""
    pass

class InvalidFormatError(ValidationError):
    """Raised when file format is invalid."""
    pass

class MissingColumnsError(ValidationError):
    """Raised when required columns are missing."""
    pass

class DataValidationError(ValidationError):
    """Raised when data fails validation rules."""
    pass

def load_excel(path: str) -> pd.DataFrame:
    """
    Load and parse an Excel file.
    
    Args:
        path: Path to Excel file
        
    Returns:
        DataFrame containing the spreadsheet data
        
    Raises:
        FileNotFoundError: If file doesn't exist
        InvalidFormatError: If file format is invalid
    """
    file_path = Path(path)
    
    if not file_path.exists():
        logger.error(f"File not found: {path}")
        raise FileNotFoundError(f"File not found: {path}")
    
    if file_path.suffix.lower() not in [".xlsx", ".xls"]:
        logger.error(f"Invalid file format: {file_path.suffix}")
        raise InvalidFormatError(f"File must be .xlsx or .xls, got: {file_path.suffix}")
    
    try:
        logger.info(f"Loading Excel file: {path}")
        df = pd.read_excel(path)
        logger.info(f"Successfully loaded {len(df)} rows and {len(df.columns)} columns")
        return df
    except Exception as e:
        logger.error(f"Failed to parse Excel file: {e}")
        raise InvalidFormatError(f"Failed to parse Excel file: {e}") from e

def validate_dataframe(
    df: pd.DataFrame,
    required_columns: List[str],
    check_empty_rows: bool = True,
    check_duplicates: bool = False
) -> Tuple[bool, str]:
    """
    Validate that a DataFrame meets requirements.
    
    Args:
        df: DataFrame to validate
        required_columns: List of column names that must exist
        check_empty_rows: Whether to reject dataframes with null values
        check_duplicates: Whether to warn about duplicate rows
        
    Returns:
        Tuple of (success: bool, message: str)
    """
    # Check for empty dataframe
    if df.empty:
        msg = "Excel file is empty."
        logger.warning(msg)
        return False, msg
    
    # Check required columns exist
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        msg = f"Missing required columns: {', '.join(missing)}"
        logger.warning(msg)
        return False, msg
    
    # Check for null values in required columns
    if check_empty_rows:
        null_cols = [c for c in required_columns if df[c].isnull().any()]
        if null_cols:
            null_count = sum(df[c].isnull().sum() for c in null_cols)
            msg = f"Found {null_count} empty cells in required columns: {', '.join(null_cols)}"
            logger.warning(msg)
            return False, msg
    
    # Warn about duplicates if requested
    if check_duplicates:
        duplicates = df.duplicated(subset=required_columns).sum()
        if duplicates > 0:
            logger.warning(f"Found {duplicates} duplicate rows based on required columns")
    
    logger.info(f"Validation passed for {len(df)} rows with columns: {', '.join(required_columns)}")
    return True, f"Validation passed. Ready to upload {len(df)} rows."


def collect_validation_errors(df: pd.DataFrame, required_columns: List[str]) -> List[Dict[str, str]]:
    errors: List[Dict[str, str]] = []
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        for col in missing:
            errors.append(
                {
                    "row_number": "",
                    "column": col,
                    "reason": "Missing required column",
                }
            )
        return errors

    for col in required_columns:
        null_rows = df[df[col].isnull()].index
        for row_index in null_rows:
            errors.append(
                {
                    "row_number": str(row_index + 2),
                    "column": col,
                    "reason": "Empty required cell",
                }
            )

    return errors


def save_validation_errors_csv(errors: List[Dict[str, str]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", newline="", encoding="utf-8") as file_handle:
        writer = csv.DictWriter(file_handle, fieldnames=["row_number", "column", "reason"])
        writer.writeheader()
        writer.writerows(errors)