"""
Modern 2026-Era Fluent & Linear-Inspired Theme Manager for PDF Converter Pro.
Provides high-polish Dark Mode and Light Mode stylesheets for PyQt6.
"""

from PyQt6.QtWidgets import QApplication, QWidget
from typing import Optional

LIGHT_THEME_QSS = """
/* Base Window & Backgrounds */
QMainWindow, QWidget#centralWidget, QWidget#dashboardPage, QWidget#studioPage {
    background-color: #F8FAFC;
    color: #0F172A;
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
}

/* Header & Navigation Bar */
QWidget#topNavBar {
    background-color: #FFFFFF;
    border-bottom: 1px solid #E2E8F0;
}

QLabel#brandTitle {
    font-size: 20px;
    font-weight: 800;
    color: #0F172A;
    letter-spacing: -0.5px;
}

QLabel#brandBadge {
    background-color: #EEF2FF;
    color: #4F46E5;
    font-size: 11px;
    font-weight: 700;
    padding: 3px 8px;
    border-radius: 6px;
    border: 1px solid #E0E7FF;
}

/* Category Filter Chips */
QPushButton.categoryChip {
    background-color: #F1F5F9;
    color: #475569;
    font-size: 13px;
    font-weight: 600;
    padding: 6px 14px;
    border-radius: 8px;
    border: 1px solid transparent;
}

QPushButton.categoryChip:hover {
    background-color: #E2E8F0;
    color: #0F172A;
}

QPushButton.categoryChip:checked {
    background-color: #4F46E5;
    color: #FFFFFF;
    border: 1px solid #4338CA;
}

/* Search Bar */
QLineEdit#searchBar {
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 10px;
    padding: 8px 14px;
    font-size: 13px;
    color: #0F172A;
}

QLineEdit#searchBar:focus {
    border: 2px solid #4F46E5;
    background-color: #FFFFFF;
}

/* Tool Cards */
QFrame.toolCard {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
}

QFrame.toolCard:hover {
    border: 1px solid #4F46E5;
    background-color: #FBFBFE;
}

QLabel.cardTitle {
    font-size: 14px;
    font-weight: 700;
    color: #0F172A;
}

QLabel.cardDesc {
    font-size: 11px;
    color: #64748B;
    line-height: 1.3;
}

QLabel.cardArrow {
    font-size: 16px;
    color: #94A3B8;
    font-weight: 600;
}

QFrame.toolCard:hover QLabel.cardArrow {
    color: #4F46E5;
}

/* Studio & Dropzone */
QLabel.stagedHeader {
    font-size: 13px;
    font-weight: 700;
    color: #0F172A;
}

QLabel.studioTitle {
    font-size: 26px;
    font-weight: 800;
    color: #0F172A;
}

QLabel.dropTitle {
    font-size: 16px;
    font-weight: 700;
    color: #0F172A;
}

QLabel.dropSub {
    font-size: 13px;
    color: #64748B;
}

QFrame#dropAreaFrame {
    background-color: #FFFFFF;
    border: 2px dashed #CBD5E1;
    border-radius: 16px;
}

QFrame#dropAreaFrame:hover {
    background-color: #EEF2FF;
    border-color: #4F46E5;
}

QListWidget#stagedList {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 6px;
}

QListWidget#stagedList::item {
    background-color: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    padding: 8px 12px;
    margin-bottom: 4px;
    color: #0F172A;
}

QListWidget#stagedList::item:selected {
    background-color: #EEF2FF;
    border-color: #818CF8;
    color: #312E81;
}

/* Buttons */
QPushButton.primaryBtn {
    background-color: #4F46E5;
    color: #FFFFFF;
    font-size: 14px;
    font-weight: 700;
    padding: 10px 24px;
    border-radius: 10px;
    border: none;
}

QPushButton.primaryBtn:hover {
    background-color: #4338CA;
}

QPushButton.primaryBtn:disabled {
    background-color: #CBD5E1;
    color: #94A3B8;
}

QPushButton.secondaryBtn {
    background-color: #FFFFFF;
    color: #334155;
    font-size: 13px;
    font-weight: 600;
    padding: 8px 16px;
    border-radius: 8px;
    border: 1px solid #CBD5E1;
}

QPushButton.secondaryBtn:hover {
    background-color: #F1F5F9;
    border-color: #94A3B8;
}

QPushButton.dangerBtn {
    background-color: #EF4444;
    color: white;
    font-size: 13px;
    font-weight: 600;
    padding: 8px 16px;
    border-radius: 8px;
    border: none;
}

QPushButton.dangerBtn:hover {
    background-color: #DC2626;
}

QPushButton.successBtn {
    background-color: #10B981;
    color: white;
    font-size: 14px;
    font-weight: 700;
    padding: 10px 24px;
    border-radius: 10px;
    border: none;
}

QPushButton.successBtn:hover {
    background-color: #059669;
}

/* Option Cards & Inputs */
QFrame.optionsCard {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 12px;
}

QLineEdit, QSpinBox, QComboBox {
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 8px;
    padding: 6px 10px;
    font-size: 13px;
    color: #0F172A;
}

QLineEdit:focus, QSpinBox:focus, QComboBox:focus {
    border: 2px solid #4F46E5;
}

/* Progress Bar */
QProgressBar {
    background-color: #E2E8F0;
    border-radius: 6px;
    height: 10px;
    text-align: center;
}

QProgressBar::chunk {
    background-color: #4F46E5;
    border-radius: 6px;
}

/* ScrollBars */
QScrollBar:vertical {
    border: none;
    background: transparent;
    width: 8px;
    margin: 0;
}

QScrollBar::handle:vertical {
    background: #CBD5E1;
    min-height: 30px;
    border-radius: 4px;
}

QScrollBar::handle:vertical:hover {
    background: #94A3B8;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""

DARK_THEME_QSS = """
/* Base Window & Backgrounds */
QMainWindow, QWidget#centralWidget, QWidget#dashboardPage, QWidget#studioPage {
    background-color: #0B0F19;
    color: #F8FAFC;
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
}

/* Header & Navigation Bar */
QWidget#topNavBar {
    background-color: #111827;
    border-bottom: 1px solid #1F2937;
}

QLabel#brandTitle {
    font-size: 20px;
    font-weight: 800;
    color: #F8FAFC;
    letter-spacing: -0.5px;
}

QLabel#brandBadge {
    background-color: #1E1B4B;
    color: #A5B4FC;
    font-size: 11px;
    font-weight: 700;
    padding: 3px 8px;
    border-radius: 6px;
    border: 1px solid #312E81;
}

/* Category Filter Chips */
QPushButton.categoryChip {
    background-color: #1E293B;
    color: #94A3B8;
    font-size: 13px;
    font-weight: 600;
    padding: 6px 14px;
    border-radius: 8px;
    border: 1px solid transparent;
}

QPushButton.categoryChip:hover {
    background-color: #334155;
    color: #F8FAFC;
}

QPushButton.categoryChip:checked {
    background-color: #6366F1;
    color: #FFFFFF;
    border: 1px solid #4F46E5;
}

/* Search Bar */
QLineEdit#searchBar {
    background-color: #1E293B;
    border: 1px solid #334155;
    border-radius: 10px;
    padding: 8px 14px;
    font-size: 13px;
    color: #F8FAFC;
}

QLineEdit#searchBar:focus {
    border: 2px solid #6366F1;
    background-color: #1E293B;
}

/* Tool Cards */
QFrame.toolCard {
    background-color: #111827;
    border: 1px solid #1F2937;
    border-radius: 12px;
}

QFrame.toolCard:hover {
    border: 1px solid #6366F1;
    background-color: #161F30;
}

QLabel.cardTitle {
    font-size: 14px;
    font-weight: 700;
    color: #F8FAFC;
}

QLabel.cardDesc {
    font-size: 11px;
    color: #94A3B8;
    line-height: 1.3;
}

QLabel.cardArrow {
    font-size: 16px;
    color: #475569;
    font-weight: 600;
}

QFrame.toolCard:hover QLabel.cardArrow {
    color: #818CF8;
}

/* Studio & Dropzone */
QLabel.stagedHeader {
    font-size: 13px;
    font-weight: 700;
    color: #0F172A;
}

QLabel.studioTitle {
    font-size: 26px;
    font-weight: 800;
    color: #F8FAFC;
}

QLabel.dropTitle {
    font-size: 16px;
    font-weight: 700;
    color: #F8FAFC;
}

QLabel.dropSub {
    font-size: 13px;
    color: #94A3B8;
}

QFrame#dropAreaFrame {
    background-color: #111827;
    border: 2px dashed #334155;
    border-radius: 16px;
}

QFrame#dropAreaFrame:hover {
    background-color: #1E1B4B;
    border-color: #6366F1;
}

QListWidget#stagedList {
    background-color: #111827;
    border: 1px solid #1F2937;
    border-radius: 12px;
    padding: 6px;
}

QListWidget#stagedList::item {
    background-color: #1E293B;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 8px 12px;
    margin-bottom: 4px;
    color: #F8FAFC;
}

QListWidget#stagedList::item:selected {
    background-color: #2E3856;
    border-color: #6366F1;
    color: #E0E7FF;
}

/* Buttons */
QPushButton.primaryBtn {
    background-color: #6366F1;
    color: #FFFFFF;
    font-size: 14px;
    font-weight: 700;
    padding: 10px 24px;
    border-radius: 10px;
    border: none;
}

QPushButton.primaryBtn:hover {
    background-color: #4F46E5;
}

QPushButton.primaryBtn:disabled {
    background-color: #1E293B;
    color: #475569;
}

QPushButton.secondaryBtn {
    background-color: #1E293B;
    color: #E2E8F0;
    font-size: 13px;
    font-weight: 600;
    padding: 8px 16px;
    border-radius: 8px;
    border: 1px solid #334155;
}

QPushButton.secondaryBtn:hover {
    background-color: #334155;
    border-color: #475569;
}

QPushButton.dangerBtn {
    background-color: #EF4444;
    color: white;
    font-size: 13px;
    font-weight: 600;
    padding: 8px 16px;
    border-radius: 8px;
    border: none;
}

QPushButton.dangerBtn:hover {
    background-color: #DC2626;
}

QPushButton.successBtn {
    background-color: #10B981;
    color: white;
    font-size: 14px;
    font-weight: 700;
    padding: 10px 24px;
    border-radius: 10px;
    border: none;
}

QPushButton.successBtn:hover {
    background-color: #059669;
}

/* Option Cards & Inputs */
QFrame.optionsCard {
    background-color: #111827;
    border: 1px solid #1F2937;
    border-radius: 12px;
    padding: 12px;
}

QLineEdit, QSpinBox, QComboBox {
    background-color: #1E293B;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 6px 10px;
    font-size: 13px;
    color: #F8FAFC;
}

QLineEdit:focus, QSpinBox:focus, QComboBox:focus {
    border: 2px solid #6366F1;
}

/* Progress Bar */
QProgressBar {
    background-color: #1E293B;
    border-radius: 6px;
    height: 10px;
    text-align: center;
}

QProgressBar::chunk {
    background-color: #6366F1;
    border-radius: 6px;
}

/* ScrollBars */
QScrollBar:vertical {
    border: none;
    background: transparent;
    width: 8px;
    margin: 0;
}

QScrollBar::handle:vertical {
    background: #334155;
    min-height: 30px;
    border-radius: 4px;
}

QScrollBar::handle:vertical:hover {
    background: #475569;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""


def apply_theme(theme_name: str = "light", widget: Optional[QWidget] = None) -> None:
    """Apply either 'dark' or 'light' modern stylesheet application-wide."""
    stylesheet = DARK_THEME_QSS if theme_name.lower() == "dark" else LIGHT_THEME_QSS
    app = QApplication.instance()
    if app:
        app.setStyleSheet(stylesheet)
    elif widget:
        widget.setStyleSheet(stylesheet)
