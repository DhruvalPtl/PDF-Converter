"""
Application paths and environment helper module.
Provides secure, platform-independent temporary directory and resource resolution.
"""

import os
import sys
import tempfile
import shutil
import uuid
import logging

logger = logging.getLogger(__name__)

def get_base_dir() -> str:
    """Return the base directory of the running application or PyInstaller bundle."""
    return getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

def get_resource_path(relative_path: str) -> str:
    """
    Safely resolve resource paths for both direct Python execution and PyInstaller bundles.
    """
    base_dir = get_base_dir()
    resolved_path = os.path.normpath(os.path.join(base_dir, relative_path))
    if os.path.exists(resolved_path):
        return resolved_path
    
    # Fallback to checking relative to current working directory
    cwd_path = os.path.normpath(os.path.join(os.getcwd(), relative_path))
    if os.path.exists(cwd_path):
        return cwd_path

    # Fallback to PDF_TOOL relative path
    tool_path = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path))
    if os.path.exists(tool_path):
        return tool_path

    return resolved_path

def get_app_data_dir() -> str:
    """
    Return the platform-appropriate application data directory.
    Windows: %LOCALAPPDATA%/PdfConverter
    Linux/macOS: ~/.local/share/PdfConverter or ~/Library/Application Support/PdfConverter
    """
    if sys.platform == 'win32':
        base = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
    elif sys.platform == 'darwin':
        base = os.path.expanduser('~/Library/Application Support')
    else:
        base = os.environ.get('XDG_DATA_HOME', os.path.expanduser('~/.local/share'))

    app_dir = os.path.join(base, 'PdfConverter')
    os.makedirs(app_dir, exist_ok=True)
    return app_dir

def get_temp_root_dir() -> str:
    """
    Return a secure, user-isolated root directory for temporary files.
    """
    temp_dir = os.path.join(tempfile.gettempdir(), 'PdfConverter')
    os.makedirs(temp_dir, exist_ok=True)
    return temp_dir

def get_temp_dir(subfolder: str = "temp") -> str:
    """
    Return or create a subfolder inside the user's temporary PdfConverter directory.
    """
    target = os.path.join(get_temp_root_dir(), subfolder)
    os.makedirs(target, exist_ok=True)
    return target

def create_session_dir(prefix: str = "job_") -> str:
    """
    Create a unique, isolated directory for an individual conversion task to prevent race conditions.
    """
    session_id = f"{prefix}{uuid.uuid4().hex[:8]}"
    session_path = os.path.join(get_temp_root_dir(), session_id)
    os.makedirs(session_path, exist_ok=True)
    return session_path

def clean_dir(dir_path: str) -> None:
    """
    Safely empty a directory's contents without removing the directory itself.
    """
    if not os.path.exists(dir_path):
        return

    for item in os.listdir(dir_path):
        item_path = os.path.join(dir_path, item)
        try:
            if os.path.isdir(item_path):
                shutil.rmtree(item_path, ignore_errors=True)
            else:
                os.remove(item_path)
        except Exception as e:
            logger.warning(f"Failed to remove temporary item {item_path}: {e}")
