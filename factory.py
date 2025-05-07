"""
Factory module for creating configured application components.
"""
from queue import Queue
from typing import Optional

from excel_processor import ExcelProcessor
from excel_processor_config import (
    ProcessingConfig, 
    FileHandlingConfig, 
    ExcelConfig,
    DEFAULT_PROCESSING_CONFIG,
    DEFAULT_FILE_HANDLING_CONFIG, 
    DEFAULT_EXCEL_CONFIG,
    load_from_environment
)
from config import app_config

def create_excel_processor(queue: Queue) -> ExcelProcessor:
    """
    Create and configure an ExcelProcessor instance.
    
    Args:
        queue (Queue): Communication queue
        
    Returns:
        ExcelProcessor: Configured processor instance
    """
    # Load environment-specific overrides
    env_config = load_from_environment()
    
    # Create processor config with app settings and env overrides
    processing_config = ProcessingConfig(
        header_row=DEFAULT_PROCESSING_CONFIG.header_row,
        headers_to_find=DEFAULT_PROCESSING_CONFIG.headers_to_find,
        sheet_name=DEFAULT_PROCESSING_CONFIG.sheet_name,
        output_directory=app_config.output_directory,  # Use app config
    )
    
    # Apply any environment overrides
    for key, value in env_config.items():
        if hasattr(processing_config, key):
            setattr(processing_config, key, value)
    
    # Create file handling config
    file_config = FileHandlingConfig(
        verify_after_save=app_config.enable_com_verification,
        max_retries=app_config.max_com_retries,
        create_backups=app_config.create_backups,
    )
    
    # Create the processor
    return ExcelProcessor(
        queue=queue,
        processing_config=processing_config,
        file_config=file_config,
        excel_config=DEFAULT_EXCEL_CONFIG
    )
