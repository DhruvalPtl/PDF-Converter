"""
History and user configuration persistence manager.
Stores recent conversions, user preferences, and theme settings in user AppData.
"""

import os
import json
import logging
from datetime import datetime
from typing import List, Dict, Any, Optional

from app_paths import get_app_data_dir

logger = logging.getLogger(__name__)

HISTORY_FILE = "history.json"
SETTINGS_FILE = "settings.json"


def _get_history_path() -> str:
    return os.path.join(get_app_data_dir(), HISTORY_FILE)


def _get_settings_path() -> str:
    return os.path.join(get_app_data_dir(), SETTINGS_FILE)


def add_history_entry(tool_name: str, input_files: List[str], output_files: List[str]) -> None:
    """Record a completed conversion entry in history."""
    try:
        entries = get_history_entries(limit=50)
        entry = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "tool": tool_name,
            "inputs": [os.path.basename(f) for f in input_files],
            "outputs": output_files,
            "success": bool(output_files)
        }
        entries.insert(0, entry)
        # Cap at 50 entries
        entries = entries[:50]

        with open(_get_history_path(), "w", encoding="utf-8") as f:
            json.dump(entries, f, indent=2)
    except Exception as e:
        logger.warning(f"Could not save history entry: {e}")


def get_history_entries(limit: int = 20) -> List[Dict[str, Any]]:
    """Retrieve the most recent conversion history entries."""
    path = _get_history_path()
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data[:limit] if isinstance(data, list) else []
    except Exception as e:
        logger.warning(f"Could not load history: {e}")
        return []


def clear_history() -> None:
    """Clear all recorded conversion history."""
    path = _get_history_path()
    if os.path.exists(path):
        try:
            os.remove(path)
        except Exception as e:
            logger.warning(f"Could not clear history: {e}")


def get_setting(key: str, default: Any = None) -> Any:
    """Retrieve a user setting value."""
    path = _get_settings_path()
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            settings = json.load(f)
            return settings.get(key, default)
    except Exception:
        return default


def set_setting(key: str, value: Any) -> None:
    """Save a user setting value."""
    path = _get_settings_path()
    settings = {}
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                settings = json.load(f)
        except Exception:
            settings = {}

    settings[key] = value
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except Exception as e:
        logger.warning(f"Could not save setting '{key}': {e}")
