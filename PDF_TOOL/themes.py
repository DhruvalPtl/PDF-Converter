"""
Modern 2026-Era Desktop Theme Manager for PDF Converter.
Provides high-polish Dark Mode and Light Mode stylesheets with Sidebar navigation.
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

/* Sidebar */
QWidget#sidebarWidget {
    background-color: #FFFFFF;
    border-right: 1px solid #E2E8F0;
}

QLabel#sidebarBrand {
    font-size: 17px;
    font-weight: 800;
    color: #0F172A;
    letter-spacing: -0.3px;
}

QLabel#sidebarSectionTitle {
    font-size: 11px;
    font-weight: 700;
    color: #94A3B8;
    letter-spacing: 0.8px;
    margin-top: 14px;
    margin-bottom: 4px;
    padding-left: 4px;
}

QPushButton.sidebarBtn {
    background-color: transparent;
    color: #334155;
    font-size: 13px;
    font-weight: 600;
    text-align: left;
    padding: 10px 14px;
    border-radius: 8px;
    border: 1px solid transparent;
}

QPushButton.sidebarBtn:hover {
    background-color: #F1F5F9;
    color: #0F172A;
}

QPushButton.sidebarBtn:checked {
    background-color: #4F46E5;
    color: #FFFFFF;
    font-weight: 700;
    border: 1px solid #4F46E5;
}

QPushButton#themeToggleBtn {
    background-color: #F8FAFC;
    color: #334155;
    font-size: 13px;
    font-weight: 600;
    padding: 10px 14px;
    border-radius: 8px;
    border: 1px solid #CBD5E1;
    text-align: center;
}

QPushButton#themeToggleBtn:hover {
    background-color: #F1F5F9;
    border-color: #94A3B8;
    color: #0F172A;
}

/* Content Header */
QWidget#contentHeader {
    background-color: #FFFFFF;
    border-bottom: 1px solid #E2E8F0;
}

QLabel#contentHeaderTitle {
    font-size: 22px;
    font-weight: 800;
    color: #0F172A;
    letter-spacing: -0.3px;
}

QLabel#contentHeaderSubtitle {
    font-size: 13px;
    color: #64748B;
    font-weight: 500;
}

/* Search Bar */
QLineEdit#searchBar {
    background-color: #F8FAFC;
    border: 1px solid #CBD5E1;
    border-radius: 8px;
    padding: 8px 14px;
    font-size: 13px;
    color: #0F172A;
}

QLineEdit#searchBar:focus {
    background-color: #FFFFFF;
    border: 2px solid #4F46E5;
}

/* Category Chips */
QPushButton.categoryChip {
    background-color: #F1F5F9;
    color: #334155;
    font-size: 13px;
    font-weight: 600;
    padding: 8px 16px;
    border-radius: 20px;
    border: 1px solid #CBD5E1;
}

QPushButton.categoryChip:hover {
    background-color: #E2E8F0;
    color: #0F172A;
}

QPushButton.categoryChip:checked {
    background-color: #4F46E5;
    color: #FFFFFF;
    font-weight: 700;
    border: 1px solid #4F46E5;
}

/* Tool Cards */
QFrame.toolCard {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
}

QFrame.toolCard:hover {
    border: 1.5px solid #4F46E5;
    background-color: #F8FAFC;
}

QLabel.cardTitle {
    font-size: 16px;
    font-weight: 700;
    color: #0F172A;
}

QLabel.cardDesc {
    font-size: 13px;
    color: #64748B;
    line-height: 1.35;
}

QLabel.cardBadge {
    font-size: 10px;
    font-weight: 700;
    padding: 2px 6px;
    border-radius: 4px;
}

QPushButton#cardActionBtn {
    background-color: #4F46E5;
    color: #FFFFFF;
    font-size: 13px;
    font-weight: 700;
    border-radius: 8px;
    border: none;
    padding: 6px 14px;
}

QPushButton#cardActionBtn:hover, QFrame.toolCard:hover QPushButton#cardActionBtn {
    background-color: #4338CA;
    color: #FFFFFF;
}

/* Studio & Dropzone */
QLabel.stagedHeader {
    font-size: 14px;
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

/* Sidebar */
QWidget#sidebarWidget {
    background-color: #0F172A;
    border-right: 1px solid #1E293B;
}

QLabel#sidebarBrand {
    font-size: 17px;
    font-weight: 800;
    color: #F8FAFC;
    letter-spacing: -0.3px;
}

QLabel#sidebarSectionTitle {
    font-size: 11px;
    font-weight: 700;
    color: #64748B;
    letter-spacing: 0.8px;
    margin-top: 14px;
    margin-bottom: 4px;
    padding-left: 4px;
}

QPushButton.sidebarBtn {
    background-color: transparent;
    color: #94A3B8;
    font-size: 13px;
    font-weight: 600;
    text-align: left;
    padding: 10px 14px;
    border-radius: 8px;
    border: 1px solid transparent;
}

QPushButton.sidebarBtn:hover {
    background-color: #1E293B;
    color: #F8FAFC;
}

QPushButton.sidebarBtn:checked {
    background-color: #6366F1;
    color: #FFFFFF;
    font-weight: 700;
    border: 1px solid #6366F1;
}

QPushButton#themeToggleBtn {
    background-color: #1E293B;
    color: #E2E8F0;
    font-size: 13px;
    font-weight: 600;
    padding: 10px 14px;
    border-radius: 8px;
    border: 1px solid #334155;
    text-align: center;
}

QPushButton#themeToggleBtn:hover {
    background-color: #334155;
    border-color: #475569;
    color: #FFFFFF;
}

/* Content Header */
QWidget#contentHeader {
    background-color: #0F172A;
    border-bottom: 1px solid #1E293B;
}

QLabel#contentHeaderTitle {
    font-size: 22px;
    font-weight: 800;
    color: #F8FAFC;
    letter-spacing: -0.3px;
}

QLabel#contentHeaderSubtitle {
    font-size: 13px;
    color: #94A3B8;
    font-weight: 500;
}

/* Search Bar */
QLineEdit#searchBar {
    background-color: #1E293B;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 8px 14px;
    font-size: 13px;
    color: #F8FAFC;
}

QLineEdit#searchBar:focus {
    background-color: #111827;
    border: 2px solid #6366F1;
}

/* Category Chips */
QPushButton.categoryChip {
    background-color: #1E293B;
    color: #CBD5E1;
    font-size: 13px;
    font-weight: 600;
    padding: 8px 16px;
    border-radius: 20px;
    border: 1px solid #334155;
}

QPushButton.categoryChip:hover {
    background-color: #334155;
    color: #FFFFFF;
}

QPushButton.categoryChip:checked {
    background-color: #6366F1;
    color: #FFFFFF;
    font-weight: 700;
    border: 1px solid #6366F1;
}

/* Tool Cards */
QFrame.toolCard {
    background-color: #111827;
    border: 1px solid #1F2937;
    border-radius: 12px;
}

QFrame.toolCard:hover {
    border: 1.5px solid #6366F1;
    background-color: #161F30;
}

QLabel.cardTitle {
    font-size: 16px;
    font-weight: 700;
    color: #F8FAFC;
}

QLabel.cardDesc {
    font-size: 13px;
    color: #94A3B8;
    line-height: 1.35;
}

QLabel.cardBadge {
    font-size: 10px;
    font-weight: 700;
    padding: 2px 6px;
    border-radius: 4px;
}

QPushButton#cardActionBtn {
    background-color: #6366F1;
    color: #FFFFFF;
    font-size: 13px;
    font-weight: 700;
    border-radius: 8px;
    border: none;
    padding: 6px 14px;
}

QPushButton#cardActionBtn:hover, QFrame.toolCard:hover QPushButton#cardActionBtn {
    background-color: #4F46E5;
    color: #FFFFFF;
}

/* Studio & Dropzone */
QLabel.stagedHeader {
    font-size: 14px;
    font-weight: 700;
    color: #F8FAFC;
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
