import logging
import logging.handlers
import os
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any
import json
from dataclasses import dataclass
import threading
from functools import wraps
import time

@dataclass
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

class LoggerSetup:
    """
    Advanced logging setup with multiple handlers and performance monitoring.
    Provides structured logging with different levels of detail for different purposes.
    """

    def __init__(self, config: Optional[LogConfig] = None):
        """
        Initialize logging setup with configuration.
        
        Args:
            config (Optional[LogConfig]): Logging configuration settings
        """
        self.config = config or LogConfig()
        self.log_dir = Path(self.config.log_dir)
        self._setup_log_directory()
        self._thread_local = threading.local()

    def _setup_log_directory(self) -> None:
        """Create logging directory if it doesn't exist."""
        self.log_dir.mkdir(parents=True, exist_ok=True)

    def _create_rotating_handler(
        self,
        filename: str,
        format_str: str,
        level: int
    ) -> logging.Handler:
        """
        Create a rotating file handler with specified parameters.
        
        Args:
            filename (str): Log filename
            format_str (str): Log format string
            level (int): Logging level
            
        Returns:
            logging.Handler: Configured rotating file handler
        """
        handler = logging.handlers.RotatingFileHandler(
            self.log_dir / filename,
            maxBytes=self.config.max_bytes,
            backupCount=self.config.backup_count,
            encoding='utf-8'
        )
        handler.setFormatter(logging.Formatter(format_str))
        handler.setLevel(level)
        return handler

    def _create_console_handler(self) -> logging.Handler:
        """
        Create console handler for logging.
        
        Returns:
            logging.Handler: Configured console handler
        """
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(
            logging.Formatter(self.config.log_format)
        )
        console_handler.setLevel(self.config.console_log_level)
        return console_handler

    def setup_logging(self) -> logging.Logger:
        """
        Set up main logger with all handlers.
        
        Returns:
            logging.Logger: Configured logger instance
        """
        # Create logger
        logger = logging.getLogger('ExcelProcessor')
        logger.setLevel(logging.DEBUG)

        # Clear any existing handlers
        logger.handlers.clear()

        # Add handlers
        handlers = [
            # Main log handler
            self._create_rotating_handler(
                self.config.main_log,
                self.config.log_format,
                self.config.file_log_level
            ),
            # Error log handler with detailed format
            self._create_rotating_handler(
                self.config.error_log,
                self.config.detailed_format,
                logging.ERROR
            ),
            # Performance log handler
            self._create_rotating_handler(
                self.config.performance_log,
                self.config.detailed_format,
                logging.INFO
            ),
            # Console handler
            self._create_console_handler()
        ]

        for handler in handlers:
            logger.addHandler(handler)

        return logger

class PerformanceMonitor:
    """
    Monitor and log performance metrics for function execution.
    """

    def __init__(self, logger: logging.Logger):
        """
        Initialize performance monitor.
        
        Args:
            logger (logging.Logger): Logger instance for performance logging
        """
        self.logger = logger
        self._thread_local = threading.local()

    def start_operation(self, operation_name: str) -> None:
        """
        Start timing an operation.
        
        Args:
            operation_name (str): Name of the operation being monitored
        """
        if not hasattr(self._thread_local, 'operations'):
            self._thread_local.operations = {}
        self._thread_local.operations[operation_name] = time.time()

    def end_operation(self, operation_name: str, additional_info: Optional[Dict[str, Any]] = None) -> float:
        """
        End timing an operation and log performance metrics.
        
        Args:
            operation_name (str): Name of the operation being monitored
            additional_info (Optional[Dict[str, Any]]): Additional information to log
            
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

def performance_logger(operation_name: Optional[str] = None):
    """
    Decorator for logging function performance.
    
    Args:
        operation_name (Optional[str]): Name of the operation (defaults to function name)
        
    Returns:
        Callable: Decorated function
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            # Get logger
            logger = logging.getLogger('ExcelProcessor')
            monitor = PerformanceMonitor(logger)
            
            # Use provided operation name or function name
            op_name = operation_name or func.__name__
            
            try:
                # Start timing
                monitor.start_operation(op_name)
                
                # Execute function
                result = func(*args, **kwargs)
                
                # End timing and log success
                duration = monitor.end_operation(op_name, {
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
                
        return wrapper
    return decorator

class ErrorHandler:
    """
    Advanced error handling and logging functionality.
    """

    def __init__(self, logger: logging.Logger):
        """
        Initialize error handler.
        
        Args:
            logger (logging.Logger): Logger instance for error logging
        """
        self.logger = logger

    def log_exception(self, exc: Exception, context: Optional[Dict[str, Any]] = None) -> None:
        """
        Log exception with detailed information and context.
        
        Args:
            exc (Exception): Exception to log
            context (Optional[Dict[str, Any]]): Additional context information
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

        self.logger.error(f"Detailed error information: {json.dumps(error_info, indent=2)}")

def get_logger(name: str = 'ExcelProcessor') -> logging.Logger:
    """
    Get or create a logger instance with the specified name.
    
    Args:
        name (str): Logger name
        
    Returns:
        logging.Logger: Configured logger instance
    """
    return logging.getLogger(name)

# Example usage of performance logging decorator
@performance_logger()
def example_function(arg1, arg2):
    """Example function with performance logging."""
    # Function implementation
    pass