"""
Unit tests for PyQt6 GUI components, navigation, and icons.
"""

import os
import sys
import pytest

# Ensure PDF_TOOL is in sys.path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "PDF_TOOL"))

from PyQt6.QtWidgets import QApplication
from icons import get_tool_icon
from main import MainWindow
from secondwindow import secondWindow

@pytest.fixture(scope="session")
def qapp():
    """Ensure a single QApplication instance exists for GUI tests."""
    app = QApplication.instance()
    if app is None:
        app = QApplication(["--platform", "offscreen"])
    return app


def test_icons_generation(qapp):
    """Verify dynamic vector icon generation for all tools without disk assets."""
    tools = [
        "Compress PDF", "Merge PDF", "Split PDF", "HTML to PDF",
        "PDF to Excel", "Excel to PDF", "PDF to Word", "Word to PDF",
        "PDF to PowerPoint", "PowerPoint to PDF", "PDF to JPG", "JPG to PDF",
        "app_icon"
    ]
    for tool in tools:
        icon = get_tool_icon(tool, size=64)
        assert icon is not None
        assert not icon.isNull()


def test_main_window_instantiation(qapp):
    """Verify MainWindow initializes without any runtime crashes."""
    window = MainWindow()
    assert window.windowTitle() == "PDF Converter Pro"
    assert window.centralWidget() is not None
    window.close()


def test_second_window_instantiation(qapp):
    """Verify secondWindow initializes for multiple tool types without crashing."""
    for tool_type in ["Merge", "Split", "Compression", "pdf2word"]:
        win = secondWindow(tool_type)
        assert win.value == tool_type
        assert win.drop_area is not None
        assert win.convert_button is not None
        win.close()
