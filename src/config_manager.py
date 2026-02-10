import configparser
from pathlib import Path
from typing import Dict, Any

SRC_DIR = Path(__file__).resolve().parent
ROOT_DIR = SRC_DIR.parent
CONFIG_PATHS = [ROOT_DIR / "config.ini", SRC_DIR / "config.ini"]

DEFAULT_CONFIG = {
    "eol": {
        "db_path": "C:\\Path\\To\\Your\\Database.accdb",
        "macro_name": "AppendToSharePoint",
        "delete_macro_name": "DeleteAllEntries",
        "preview_query_incomplete": "qryEolIncompletePreview",
        "preview_row_limit": "50",
        "linked_table": "SP_Devices",
        "temp_table": "TempImportTable",
    },
    "archive": {
        "db_path": "C:\\Path\\To\\Your\\ArchiveDatabase.accdb",
        "macro_new_starters": "NewStartersTransfer",
        "macro_leavers": "LeaversTransfer",
        "macro_mat_leavers": "MatLeaversTransfer",
        "macro_transfers": "TransfersTransfer",
        "preview_query_new_starters": "qryNewStartersPreview",
        "preview_query_leavers": "qryLeaversPreview",
        "preview_query_mat_leavers": "qryMatLeaversPreview",
        "preview_query_transfers": "qryTransfersPreview",
        "preview_row_limit": "50",
    },
    "validation": {
        "required_columns": "DeviceID,Model,EndOfLifeDate",
    },
    "logging": {
        "log_file": "logs/uploader.log",
        "log_level": "INFO",
    },
    "audit": {
        "enabled": "true",
        "log_folder": "logs/audit",
        "log_file": "audit.log",
    },
}

REQUIRED_KEYS = {
    "eol": [
        "db_path",
        "macro_name",
        "delete_macro_name",
        "preview_query_incomplete",
        "preview_row_limit",
        "linked_table",
        "temp_table",
    ],
    "archive": [
        "db_path",
        "macro_new_starters",
        "macro_leavers",
        "macro_mat_leavers",
        "macro_transfers",
        "preview_query_new_starters",
        "preview_query_leavers",
        "preview_query_mat_leavers",
        "preview_query_transfers",
        "preview_row_limit",
    ],
    "validation": ["required_columns"],
    "logging": ["log_file", "log_level"],
    "audit": ["enabled", "log_folder", "log_file"],
}

class ConfigError(Exception):
    """Custom exception for config-related errors."""
    pass

def create_default_config() -> None:
    """Create a default config.ini if it doesn't exist."""
    config_path = CONFIG_PATHS[0]
    if not config_path.exists():
        config = configparser.ConfigParser()
        for section, values in DEFAULT_CONFIG.items():
            config.add_section(section)
            for key, value in values.items():
                config.set(section, key, value)
        save_config(config)
        print(f"Default config created at {config_path}. Please update the paths and settings.")

def validate_config(config: configparser.ConfigParser) -> None:
    """Validate that all required config keys exist."""
    for section, required_keys in REQUIRED_KEYS.items():
        if not config.has_section(section):
            raise ConfigError(f"Missing required section: [{section}]")
        
        for key in required_keys:
            if not config.has_option(section, key):
                raise ConfigError(f"Missing required key '{key}' in section [{section}]")

def load_config() -> configparser.ConfigParser:
    """Load and validate config from config.ini, creating default if needed."""
    config_path = next((path for path in CONFIG_PATHS if path.exists()), None)
    if config_path is None:
        create_default_config()
        raise ConfigError(
            f"Config file created with defaults at {CONFIG_PATHS[0]}. "
            "Please update it with your actual paths and settings."
        )
    
    config = configparser.ConfigParser()
    config.read(config_path)
    
    try:
        validate_config(config)
    except ConfigError as e:
        raise ConfigError(f"Config validation failed: {e}") from e
    
    return config

def save_config(config: configparser.ConfigParser) -> None:
    """Save config to file."""
    config_path = next((path for path in CONFIG_PATHS if path.exists()), CONFIG_PATHS[0])
    with open(config_path, "w") as file_handle:
        config.write(file_handle)