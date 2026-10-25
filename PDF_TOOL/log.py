"""
Logging configuration for PDF Converter.
Sets up rotating file logging in the user AppData directory and console logging.
"""

import logging
from logging.handlers import RotatingFileHandler
import os
import sys

from app_paths import get_app_data_dir

def log_system(log_level: int = logging.INFO) -> logging.Logger:
    """
    Configure and initialize application-wide logging.
    Writes logs to a rotating file in user AppData and streams to stderr.
    """
    root_logger = logging.getLogger()
    
    # Avoid duplicate handlers if already configured
    if root_logger.handlers:
        return root_logger

    root_logger.setLevel(log_level)
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    # Rotating File Handler
    try:
        log_dir = os.path.join(get_app_data_dir(), 'logs')
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, 'app.log')

        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=5 * 1024 * 1024,  # 5 MB
            backupCount=3,
            encoding='utf-8'
        )
        file_handler.setLevel(log_level)
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)
        root_logger.info(f"Logging initialized. Log file: {log_file}")
    except Exception as e:
        root_logger.warning(f"Could not initialize file logging handler: {e}")

    return root_logger
