"""
Configuration settings for the SF132 to SF133 Excel processing application.
"""
from dataclasses import dataclass
from typing import Dict, Any
import json
import os
from pathlib import Path

@dataclass
class AppConfig:
    """Application configuration settings"""
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

def load_config(config_file: str = "app_config.json") -> AppConfig:
    """
    Load configuration from a JSON file, or create default if not exists.
    
    Args:
        config_file (str): Path to the config file
        
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
        save_config(config, config_file)
        
    return config

def save_config(config: AppConfig, config_file: str = "app_config.json") -> None:
    """
    Save configuration to a JSON file.
    
    Args:
        config (AppConfig): Configuration object
        config_file (str): Path to the config file
    """
    try:
        # Create config as dictionary
        config_dict = {key: value for key, value in config.__dict__.items()}
        
        # Save to file
        with open(config_file, 'w') as f:
            json.dump(config_dict, f, indent=4)
    except Exception as e:
        print(f"Error saving config: {e}")

# Access the config throughout the application
app_config = load_config()
