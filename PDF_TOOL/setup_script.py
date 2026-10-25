"""
Setup script to ensure required application directories and resources exist.
"""

import os
import shutil
import logging
from app_paths import get_app_data_dir, get_temp_root_dir

logger = logging.getLogger(__name__)

def install_resources() -> None:
    """
    Ensure required application directories exist safely in user profile.
    Copies any bundled resources if present.
    """
    try:
        app_data = get_app_data_dir()
        temp_dir = get_temp_root_dir()

        os.makedirs(os.path.join(temp_dir, 'Temp'), exist_ok=True)
        os.makedirs(os.path.join(temp_dir, 'Temp1'), exist_ok=True)
        os.makedirs(os.path.join(app_data, 'logs'), exist_ok=True)

        source_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources')
        if os.path.exists(source_dir):
            for item in os.listdir(source_dir):
                s = os.path.join(source_dir, item)
                d = os.path.join(app_data, item)
                if os.path.isdir(s):
                    shutil.copytree(s, d, dirs_exist_ok=True)
                else:
                    shutil.copy2(s, d)
        logger.info("Application environment initialized successfully.")
    except Exception as e:
        logger.error(f"Error during resource initialization: {e}")
