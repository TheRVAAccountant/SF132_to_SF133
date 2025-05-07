"""
Logging configuration and utilities.

This module provides advanced logging setup with performance monitoring
and exception handling capabilities.
"""

import logging
import logging.handlers
import os
import sys
import traceback
from datetime import datetime
from pathlib import Path
import json
import threading
import time
from functools import wraps
from typing import Optional, Dict, Any, Callable, TypeVar, cast

# Type variables for function decorators
T = TypeVar('T')
F = TypeVar('F', bound=Callable[..., Any])

class LogConfig:
    """Configuration settings for logging setup."""
    log_dir: str = "logs"
    main_log: str = "excel_processor.log"
    error_log: str = "error.log"
    performance_log: str = "performance.log"
    max_bytes: int = 5 * 1024 * 1024  # 5MB
    backup_count: int = 5
    file_log_level: int = logging.DEBUG
    console_log_level: int = logging.INFO
    log_format: str = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    detailed_format: str = (
        "%(asctime)s - %(name)s - %(levelname)s - %(pathname)s:%(lineno)d - "
        "%(funcName)s - %(message)s"
    )

def setup_logging(config: Optional[LogConfig] = None) -> logging.Logger:
    """
    Set up main logger with all handlers.
    
    Args:
        config: Optional configuration settings
        
    Returns:
        logging.Logger: Configured logger instance
    """
    # Create logger
    logger = logging.getLogger("ExcelProcessor")
    
    # If logger already has handlers, assume it's configured
    if logger.handlers:
        return logger
        
    logger.setLevel(logging.DEBUG)
    
    # Use default config if none provided
    if config is None:
        config = LogConfig()
    
    # Create log directory
    log_dir = Path(config.log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # Create handlers
    handlers = [
        # Main log handler
        _create_rotating_handler(
            log_dir / config.main_log,
            config.log_format,
            config.file_log_level,
            config.max_bytes,
            config.backup_count
        ),
        # Error log handler with detailed format
        _create_rotating_handler(
            log_dir / config.error_log,
            config.detailed_format,
            logging.ERROR,
            config.max_bytes,
            config.backup_count
        ),
        # Performance log handler
        _create_rotating_handler(
            log_dir / config.performance_log,
            config.detailed_format,
            logging.INFO,
            config.max_bytes,
            config.backup_count
        ),
        # Console handler
        _create_console_handler(config.log_format, config.console_log_level)
    ]
    
    for handler in handlers:
        logger.addHandler(handler)
    
    return logger

def _create_rotating_handler(
    filename: Path,
    format_str: str,
    level: int,
    max_bytes: int,
    backup_count: int
) -> logging.Handler:
    """
    Create a rotating file handler with specified parameters.
    
    Args:
        filename: Log filename
        format_str: Log format string
        level: Logging level
        max_bytes: Maximum file size in bytes
        backup_count: Number of backup files
        
    Returns:
        logging.Handler: Configured rotating file handler
    """
    handler = logging.handlers.RotatingFileHandler(
        filename,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding='utf-8'
    )
    handler.setFormatter(logging.Formatter(format_str))
    handler.setLevel(level)
    return handler

def _create_console_handler(format_str: str, level: int) -> logging.Handler:
    """
    Create console handler for logging.
    
    Args:
        format_str: Log format string
        level: Logging level
        
    Returns:
        logging.Handler: Configured console handler
    """
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(
        logging.Formatter(format_str)
    )
    console_handler.setLevel(level)
    return console_handler

def setup_exception_handling(logger: logging.Logger) -> None:
    """
    Set up global exception handling.
    
    Args:
        logger: Logger to use for exception handling
    """
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            # Handle keyboard interrupt differently
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
            
        logger.critical("Uncaught exception:", exc_info=(exc_type, exc_value, exc_traceback))
        
    sys.excepthook = handle_exception
    
    def handle_thread_exception(args):
        logger.critical("Uncaught thread exception:", 
                       exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
        
    threading.excepthook = handle_thread_exception

class PerformanceMonitor:
    """
    Monitor and log performance metrics for function execution.
    """
    
    def __init__(self, logger: logging.Logger):
        """
        Initialize performance monitor.
        
        Args:
            logger: Logger instance for performance logging
        """
        self.logger = logger
        self._thread_local = threading.local()
    
    def start_operation(self, operation_name: str) -> None:
        """
        Start timing an operation.
        
        Args:
            operation_name: Name of the operation being monitored
        """
        if not hasattr(self._thread_local, 'operations'):
            self._thread_local.operations = {}
        self._thread_local.operations[operation_name] = time.time()
    
    def end_operation(self, operation_name: str, additional_info: Optional[Dict[str, Any]] = None) -> float:
        """
        End timing an operation and log performance metrics.
        
        Args:
            operation_name: Name of the operation being monitored
            additional_info: Additional information to log
            
        Returns:
            float: Duration of the operation in seconds
        """
        if not hasattr(self._thread_local, 'operations'):
            raise RuntimeError(f"Operation {operation_name} was never started")
        
        start_time = self._thread_local.operations.pop(operation_name, None)
        if start_time is None:
            raise RuntimeError(f"Operation {operation_name} was never started")
        
        duration = time.time() - start_time
        
        log_data = {
            'operation': operation_name,
            'duration': duration,
            'thread_id': threading.get_ident()
        }
        
        if additional_info:
            log_data.update(additional_info)
        
        self.logger.info(f"Performance metrics: {json.dumps(log_data)}")
        return duration

def performance_logger(operation_name: Optional[str] = None) -> Callable[[F], F]:
    """
    Decorator for logging function performance.
    
    Args:
        operation_name: Name of the operation (defaults to function name)
        
    Returns:
        Callable: Decorated function
    """
    def decorator(func: F) -> F:
        @wraps(func)
        def wrapper(*args: Any, **kwargs: Any) -> Any:
            # Get logger
            logger = logging.getLogger("ExcelProcessor")
            monitor = PerformanceMonitor(logger)
            
            # Use provided operation name or function name
            op_name = operation_name or func.__name__
            
            try:
                # Start timing
                monitor.start_operation(op_name)
                
                # Execute function
                result = func(*args, **kwargs)
                
                # End timing and log success
                monitor.end_operation(op_name, {
                    'status': 'success',
                    'args_length': len(args),
                    'kwargs_length': len(kwargs)
                })
                
                return result
                
            except Exception as e:
                # Log error with performance metrics
                monitor.end_operation(op_name, {
                    'status': 'error',
                    'error_type': type(e).__name__,
                    'error_message': str(e)
                })
                raise
        
        return cast(F, wrapper)
    return decorator

def log_exception(logger: logging.Logger, exc: Exception, context: Optional[Dict[str, Any]] = None) -> None:
    """
    Log exception with detailed information and context.
    
    Args:
        logger: Logger instance
        exc: Exception to log
        context: Additional context information
    """
    error_info = {
        'timestamp': datetime.now().isoformat(),
        'error_type': type(exc).__name__,
        'error_message': str(exc),
        'traceback': traceback.format_exc(),
        'thread_id': threading.get_ident(),
    }
    
    if context:
        error_info['context'] = context
    
    logger.error(f"Detailed error information: {json.dumps(error_info, indent=2)}")

def get_logger(name: str = "ExcelProcessor") -> logging.Logger:
    """
    Get or create a logger instance with the specified name.
    
    Args:
        name: Logger name
        
    Returns:
        logging.Logger: Logger instance
    """
    return logging.getLogger(name)