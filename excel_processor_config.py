"""
Configuration settings for the Excel processor component.
"""
from dataclasses import dataclass
from typing import Tuple, Optional, Dict, Any
import os
from pathlib import Path

@dataclass
class ProcessingConfig:
    """Configuration settings for Excel processing."""
    header_row: int = 9
    headers_to_find: Tuple[str, ...] = ("Difference", "Include in CFO Cert Letter", "Explanation")
    sheet_name: str = "SF132 to SF133 Reconciliation"
    output_directory: str = "output"
    backup_directory: str = os.path.join("output", "backups")
    
    # Comment patterns to apply based on different conditions
    comments: Dict[str, str] = None
    
    # Excel repair prevention settings
    clean_external_connections: bool = True  # Clean up external data connections
    preserve_external_data: bool = False     # Whether to preserve external data links
    
    def __post_init__(self):
        """Initialize default values that require computation."""
        if self.comments is None:
            self.comments = {
                "explanation_reasonable": "Explanation Reasonable",
                "explanation_cfo_letter": "Explanation Reasonable; Include in CFO Cert Letter",
                "explanation_required": "Explanation Required"
            }

@dataclass
class FileHandlingConfig:
    """Configuration for file operations."""
    max_retries: int = 3
    retry_delay: float = 2.0
    verify_after_save: bool = True
    create_backups: bool = True
    temp_directory: Optional[str] = None
    
    # File repair prevention settings
    use_com_for_final_save: bool = True  # Use Excel COM for final save to prevent repair issues
    
    def get_temp_dir(self) -> str:
        """Get the temporary directory path."""
        if self.temp_directory:
            Path(self.temp_directory).mkdir(exist_ok=True, parents=True)
            return self.temp_directory
        return None  # Use system default temp directory

@dataclass
class ExcelConfig:
    """Configuration for Excel application interaction."""
    visible: bool = False
    display_alerts: bool = False
    kill_processes_before_start: bool = True
    process_wait_timeout: int = 5
    com_init_timeout: float = 2.0
    
    # Excel repair prevention settings
    disable_links: bool = True          # Disable external links when opening
    ignore_readonly: bool = True        # Ignore readonly recommendation
    disable_auto_recovery: bool = True  # Disable Excel auto-recovery for our files
    
    # COM usage settings
    enable_com: bool = True             # Master switch to enable/disable COM operations
    use_com_for_copy: bool = True       # Use COM for initial file copy
    use_com_for_final_save: bool = True # Use COM for final save
    max_com_retries: int = 2            # Number of times to retry COM operations
    
    # Security settings
    elevate_com_security: bool = False  # Try to elevate COM security context

# Default configurations
DEFAULT_PROCESSING_CONFIG = ProcessingConfig()
DEFAULT_FILE_HANDLING_CONFIG = FileHandlingConfig()
DEFAULT_EXCEL_CONFIG = ExcelConfig()

def load_from_environment() -> Dict[str, Any]:
    """Load configuration values from environment variables."""
    config_values = {}
    
    # Map environment variables to config properties
    env_mappings = {
        "SF132_HEADER_ROW": ("header_row", int),
        "SF132_SHEET_NAME": ("sheet_name", str),
        "SF132_OUTPUT_DIR": ("output_directory", str),
        "SF132_MAX_RETRIES": ("max_retries", int),
        "SF132_VERIFY_SAVE": ("verify_after_save", lambda x: x.lower() == 'true'),
        "SF132_CLEAN_CONNECTIONS": ("clean_external_connections", lambda x: x.lower() == 'true'),
    }
    
    # Extract values from environment
    for env_var, (prop_name, converter) in env_mappings.items():
        if env_var in os.environ:
            config_values[prop_name] = converter(os.environ[env_var])
    
    return config_values
