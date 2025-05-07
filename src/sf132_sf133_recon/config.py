"""
Configuration settings for the SF132 to SF133 Excel processing application.
"""
from dataclasses import dataclass, field
from typing import Dict, Tuple, Optional, Any, List
import os
import json
from pathlib import Path

@dataclass
class AppConfig:
    """Application-wide configuration settings."""
    # Excel processing settings
    enable_com_verification: bool = False  # Set to False to avoid COM verification issues
    max_com_retries: int = 3
    cleanup_temp_files: bool = True
    
    # Output settings
    output_directory: str = "output"
    create_backups: bool = True
    backup_directory: str = "output/backups"
    
    # Logging settings
    log_directory: str = "logs"
    
    # UI settings
    theme: str = "forest-dark"
    window_width: int = 695
    window_height: int = 400

@dataclass
class ProcessingConfig:
    """Configuration settings for Excel processing."""
    header_row: int = 9
    headers_to_find: Tuple[str, ...] = ("Difference", "Include in CFO Cert Letter", "Explanation")
    sheet_name: str = "SF132 to SF133 Reconciliation"
    output_directory: str = "output"
    backup_directory: str = field(default="output/backups")
    
    # Comment patterns to apply based on different conditions
    comments: Dict[str, str] = field(default_factory=lambda: {
        "explanation_reasonable": "Explanation Reasonable",
        "explanation_cfo_letter": "Explanation Reasonable; Include in CFO Cert Letter",
        "explanation_required": "Explanation Required"
    })
    
    # Excel repair prevention settings
    clean_external_connections: bool = True  # Clean up external data connections
    preserve_external_data: bool = False     # Whether to preserve external data links

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
    
    def get_temp_dir(self) -> Optional[str]:
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

def load_app_config(config_file: str = "app_config.json") -> AppConfig:
    """
    Load application configuration from a JSON file, or create default if not exists.
    
    Args:
        config_file: Path to the config file
        
    Returns:
        AppConfig: Application configuration object
    """
    config = AppConfig()
    
    # If config file exists, load it
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r') as f:
                config_dict = json.load(f)
                
            # Update config with loaded values
            for key, value in config_dict.items():
                if hasattr(config, key):
                    setattr(config, key, value)
        except Exception as e:
            print(f"Error loading config: {e}")
    else:
        # Save default config
        save_app_config(config, config_file)
        
    return config

def save_app_config(config: AppConfig, config_file: str = "app_config.json") -> None:
    """
    Save application configuration to a JSON file.
    
    Args:
        config: Configuration object
        config_file: Path to the config file
    """
    try:
        # Create config as dictionary
        config_dict = {key: value for key, value in config.__dict__.items()}
        
        # Save to file
        with open(config_file, 'w') as f:
            json.dump(config_dict, f, indent=4)
    except Exception as e:
        print(f"Error saving config: {e}")

def load_from_environment() -> Dict[str, Any]:
    """
    Load configuration values from environment variables.
    
    Returns:
        Dict[str, Any]: Configuration values from environment
    """
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

# Initialize with defaults, then update from environment
app_config = load_app_config()

# Allow environment variables to override config
env_config = load_from_environment()
for key, value in env_config.items():
    if hasattr(app_config, key):
        setattr(app_config, key, value)