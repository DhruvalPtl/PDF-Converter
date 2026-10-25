"""
PDF Converter — Modern Application Dashboard & Studio Launcher.
Features 2026-era Linear/macOS design with a persistent navigation sidebar,
fluid grid layout, live tool search, category filters, instant Dark/Light theme switching,
and seamless single-window stacked navigation.
"""

import sys
import os
import logging
from typing import List, Tuple

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
    QGridLayout, QPushButton, QWidget, QLabel, QMessageBox,
    QScrollArea, QFrame, QDialog, QListWidget, QListWidgetItem,
    QLineEdit, QStackedWidget, QButtonGroup, QSizePolicy
)
from PyQt6.QtGui import QIcon, QFont, QAction
from PyQt6.QtCore import QSize, Qt

from app_paths import get_app_data_dir
from setup_script import install_resources
from log import log_system
from icons import get_tool_icon
from secondwindow import secondWindow
from themes import apply_theme
from history import get_setting, set_setting, get_history_entries, clear_history
from ui_components import ModernToolCard

# Initialize logging and resources
logger = log_system()
install_resources()


# Tool definitions: (Name, Description, Category, ActionKey)
TOOL_DEFINITIONS: List[Tuple[str, str, str, str]] = [
    # Essential PDF tools
    ("Merge PDF", "Combine multiple PDFs in custom sequence", "essential", "Merge"),
    ("Compress PDF", "60-80% size reduction with smart downsampling", "essential", "Compression"),
    ("Split PDF", "Extract page ranges or split into single pages", "essential", "Split"),
    ("Rotate PDF", "Rotate pages clockwise by 90°, 180°, or 270°", "essential", "Rotate"),

    # Convert to PDF
    ("Word to PDF", "Convert Word DOCX documents to standardized PDF", "to_pdf", "word2pdf"),
    ("Excel to PDF", "Convert spreadsheet workbooks to PDF", "to_pdf", "excel2pdf"),
    ("PowerPoint to PDF", "Convert PPTX slide decks to PDF documents", "to_pdf", "pptx2pdf"),
    ("JPG to PDF", "Compile JPG and PNG images into a clean PDF", "to_pdf", "jpg2pdf"),
    ("HTML to PDF", "Render modern web pages to PDF via Chromium", "to_pdf", "Html2pdf"),

    # Convert from PDF
    ("PDF to Word", "Convert PDF to editable DOCX documents", "from_pdf", "pdf2word"),
    ("PDF to Excel", "Extract tables with C++ MuPDF to styled Excel", "from_pdf", "pdf2excel"),
    ("PDF to PowerPoint", "Generate editable slides with textframes from PDF", "from_pdf", "pdf2pptx"),
    ("PDF to JPG", "Render pages as high-resolution PNG images", "from_pdf", "pdf2jpg"),

    # Security & OCR
    ("Protect PDF", "Encrypt documents with military-grade AES-256", "security", "Protect"),
    ("Unlock PDF", "Decrypt and strip password restrictions", "security", "Unlock"),
    ("Watermark PDF", "Stamp custom semi-transparent vector text", "security", "Watermark"),
    ("OCR Scanned PDF", "Make scanned PDFs searchable and selectable", "security", "Ocr"),
]

# Category definitions: (Button Label, Category ID, Display Title, Subtitle Description)
CATEGORY_INFO = [
    ("All Tools (17)", "all", "All Tools", "17 document conversion and management tools"),
    ("⚡ Essential (4)", "essential", "Essential Tools", "4 core tools for merging, splitting, compressing, and rotating"),
    ("🔄 Convert to PDF (5)", "to_pdf", "Convert to PDF", "5 tools to convert Office documents, images, and HTML to PDF"),
    ("🔁 Convert from PDF (4)", "from_pdf", "Convert from PDF", "4 tools to convert PDF pages into Word, Excel, PowerPoint, and images"),
    ("🔒 Security & OCR (4)", "security", "Security & OCR", "4 tools for AES-256 encryption, password removal, watermarking, and OCR"),
]


class HistoryDialog(QDialog):
    """Dialog displaying recently completed conversion tasks with quick open actions."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Recent Conversion History")
        self.resize(680, 440)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        header = QLabel("Recently Completed Conversions:")
        header.setStyleSheet("font-size: 16px; font-weight: 700;")
        layout.addWidget(header)

        self.list_widget = QListWidget()
        entries = get_history_entries(limit=35)
        if not entries:
            self.list_widget.addItem("No conversions recorded yet.")
        else:
            for entry in entries:
                ts = entry.get("timestamp", "")
                tool = entry.get("tool", "")
                inputs = ", ".join(entry.get("inputs", []))
                outputs = entry.get("outputs", [])
                out_str = os.path.basename(outputs[0]) if outputs else "None"
                item_text = f"[{ts}]  {tool.upper()}  •  {inputs}  →  {out_str}"
                item = QListWidgetItem(item_text)
                if outputs:
                    item.setData(Qt.ItemDataRole.UserRole, outputs[0])
                    item.setToolTip(f"Double-click to open: {outputs[0]}")
                self.list_widget.addItem(item)

        self.list_widget.itemDoubleClicked.connect(self.open_output_file)
        layout.addWidget(self.list_widget)

        btn_row = QHBoxLayout()
        clear_btn = QPushButton("Clear History")
        clear_btn.setProperty("class", "secondaryBtn")
        clear_btn.clicked.connect(self.do_clear_history)
        btn_row.addWidget(clear_btn)

        btn_row.addStretch()

        open_btn = QPushButton("Open Selected File")
        open_btn.setProperty("class", "primaryBtn")
        open_btn.clicked.connect(self.open_selected_item)
        btn_row.addWidget(open_btn)

        close_btn = QPushButton("Close")
        close_btn.setProperty("class", "secondaryBtn")
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)

        layout.addLayout(btn_row)

    def open_output_file(self, item: QListWidgetItem):
        file_path = item.data(Qt.ItemDataRole.UserRole)
        if file_path and os.path.exists(file_path):
            if sys.platform == 'win32':
                os.startfile(file_path)
        elif file_path:
            QMessageBox.warning(self, "File Not Found", f"The file no longer exists at:\n{file_path}")

    def open_selected_item(self):
        item = self.list_widget.currentItem()
        if item:
            self.open_output_file(item)

    def do_clear_history(self):
        clear_history()
        self.list_widget.clear()
        self.list_widget.addItem("No conversions recorded yet.")


class MainWindow(QMainWindow):
    """
    Unified Application Window showcasing 17 document tools with a modern sidebar,
    fluid responsive grid, live search, instant theme switching, and seamless stacked studio navigation.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Converter")
        self.setWindowIcon(get_tool_icon("app_icon", size=64))
        self.resize(1220, 840)
        self.setMinimumSize(980, 660)

        # Stored tool cards for dynamic filtering
        self.card_widgets: List[Tuple[ModernToolCard, str, str, str]] = []
        self.active_category = "all"
        self.current_search = ""

        # Load saved theme
        self.current_theme = get_setting("theme", "dark")
        apply_theme(self.current_theme)

        self.setup_ui()

    def setup_ui(self):
        self.central_widget = QWidget()
        self.central_widget.setObjectName("centralWidget")
        self.setCentralWidget(self.central_widget)

        # Horizontal split: Sidebar on Left (240px) + Stacked Workspaces on Right
        self.root_layout = QHBoxLayout(self.central_widget)
        self.root_layout.setContentsMargins(0, 0, 0, 0)
        self.root_layout.setSpacing(0)

        # 1. Left Navigation Sidebar
        self.setup_sidebar()

        # 2. Main Stacked Area
        self.stack = QStackedWidget()
        self.root_layout.addWidget(self.stack, 1)

        # Page 0: Dashboard Page
        self.dashboard_page = QWidget()
        self.dashboard_page.setObjectName("dashboardPage")
        self.setup_dashboard_page()
        self.stack.addWidget(self.dashboard_page)

    def setup_sidebar(self):
        """Construct the fixed-width persistent navigation sidebar."""
        self.sidebar = QWidget()
        self.sidebar.setObjectName("sidebarWidget")
        self.sidebar.setFixedWidth(240)

        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(14, 20, 14, 18)
        sidebar_layout.setSpacing(4)

        # Brand Header (App Icon + Title)
        brand_row = QHBoxLayout()
        brand_row.setSpacing(10)
        brand_row.setContentsMargins(4, 0, 4, 12)

        icon_label = QLabel()
        icon_label.setFixedSize(28, 28)
        icon_label.setPixmap(get_tool_icon("app_icon", size=56).pixmap(28, 28))
        brand_row.addWidget(icon_label)

        brand_title = QLabel("PDF Converter")
        brand_title.setObjectName("sidebarBrand")
        brand_row.addWidget(brand_title)
        brand_row.addStretch()

        sidebar_layout.addLayout(brand_row)

        # Section: TOOLS
        tools_header = QLabel("TOOLS")
        tools_header.setObjectName("sidebarSectionTitle")
        sidebar_layout.addWidget(tools_header)

        self.category_button_group = QButtonGroup(self)
        self.category_button_group.setExclusive(True)

        self.category_buttons = {}
        for idx, (btn_text, cat_id, title, subtitle) in enumerate(CATEGORY_INFO):
            # Escape ampersand so Qt does not render it as mnemonic underline
            btn = QPushButton(btn_text.replace("&", "&&"))
            btn.setCheckable(True)
            btn.setProperty("class", "sidebarBtn")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            if idx == 0:
                btn.setChecked(True)
            self.category_button_group.addButton(btn, idx)
            btn.clicked.connect(lambda _, cid=cat_id, t=title, s=subtitle: self.on_category_selected(cid, t, s))
            sidebar_layout.addWidget(btn)
            self.category_buttons[cat_id] = btn

        sidebar_layout.addSpacing(14)

        # Section: WORKSPACE
        workspace_header = QLabel("WORKSPACE")
        workspace_header.setObjectName("sidebarSectionTitle")
        sidebar_layout.addWidget(workspace_header)

        history_btn = QPushButton("🕒  Recent History")
        history_btn.setProperty("class", "sidebarBtn")
        history_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        history_btn.clicked.connect(self.show_history_dialog)
        sidebar_layout.addWidget(history_btn)

        logs_btn = QPushButton("📁  View Logs")
        logs_btn.setProperty("class", "sidebarBtn")
        logs_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        logs_btn.clicked.connect(self.open_logs_folder)
        sidebar_layout.addWidget(logs_btn)

        about_btn = QPushButton("ℹ️  About")
        about_btn.setProperty("class", "sidebarBtn")
        about_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        about_btn.clicked.connect(self.show_about_dialog)
        sidebar_layout.addWidget(about_btn)

        sidebar_layout.addStretch(1)

        # Footer: Theme Switcher Button
        theme_label = "🌙 Dark Mode" if self.current_theme == "light" else "☀️ Light Mode"
        self.theme_btn = QPushButton(theme_label)
        self.theme_btn.setObjectName("themeToggleBtn")
        self.theme_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.theme_btn.clicked.connect(self.toggle_theme)
        sidebar_layout.addWidget(self.theme_btn)

        self.root_layout.addWidget(self.sidebar)

    def setup_dashboard_page(self):
        """Build the fluid content area with header, search bar, and card grid."""
        dash_layout = QVBoxLayout(self.dashboard_page)
        dash_layout.setContentsMargins(0, 0, 0, 0)
        dash_layout.setSpacing(0)

        # Top Content Header Bar (Title, Subtitle, Live Search Bar)
        content_header = QWidget()
        content_header.setObjectName("contentHeader")
        content_header.setFixedHeight(72)

        header_layout = QHBoxLayout(content_header)
        header_layout.setContentsMargins(28, 0, 28, 0)
        header_layout.setSpacing(16)

        title_col = QVBoxLayout()
        title_col.setSpacing(2)
        title_col.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        self.header_title = QLabel("All Tools")
        self.header_title.setObjectName("contentHeaderTitle")
        title_col.addWidget(self.header_title)

        self.header_subtitle = QLabel("17 document conversion and management tools")
        self.header_subtitle.setObjectName("contentHeaderSubtitle")
        title_col.addWidget(self.header_subtitle)

        header_layout.addLayout(title_col)
        header_layout.addStretch()

        # Search Bar
        self.search_input = QLineEdit()
        self.search_input.setObjectName("searchBar")
        self.search_input.setPlaceholderText("Search tools (e.g. merge, compress, protect)...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setFixedWidth(320)
        self.search_input.textChanged.connect(self.on_search_text_changed)
        header_layout.addWidget(self.search_input)

        dash_layout.addWidget(content_header)

        # Scrollable Cards Grid
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        scroll_area.setStyleSheet("background-color: transparent;")

        grid_container = QWidget()
        grid_container.setStyleSheet("background-color: transparent;")
        self.grid_layout = QGridLayout(grid_container)
        self.grid_layout.setContentsMargins(28, 20, 28, 28)
        self.grid_layout.setSpacing(14)
        self.grid_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        self.build_tool_cards()

        scroll_area.setWidget(grid_container)
        dash_layout.addWidget(scroll_area, 1)

    def build_tool_cards(self):
        """Instantiate ModernToolCards and arrange in 3 fluid responsive columns."""
        self.card_widgets.clear()
        columns = 3

        for index, (tool_name, tool_desc, category, action_key) in enumerate(TOOL_DEFINITIONS):
            card = ModernToolCard(tool_name, tool_desc, category=category)
            card.apply_card_theme(self.current_theme == "dark")
            card.clicked.connect(lambda name=tool_name: self.open_tool(name))
            row = index // columns
            col = index % columns
            self.grid_layout.addWidget(card, row, col)
            self.card_widgets.append((card, tool_name, tool_desc, category))

    def on_category_selected(self, category_id: str, title: str = "All Tools", subtitle: str = ""):
        """Filter tools when a category button is clicked in the sidebar."""
        self.active_category = category_id
        self.header_title.setText(title)
        self.header_subtitle.setText(subtitle)
        self.show_dashboard()
        self.apply_tool_filters()

    def on_search_text_changed(self, text: str):
        """Live search filter callback."""
        self.current_search = text.strip().lower()
        self.apply_tool_filters()

    def apply_tool_filters(self):
        """Show or hide cards based on active category and live search query."""
        visible_cards = []
        for card, name, desc, category in self.card_widgets:
            match_cat = (self.active_category == "all" or category == self.active_category)
            match_search = (
                not self.current_search
                or self.current_search in name.lower()
                or self.current_search in desc.lower()
            )
            is_visible = match_cat and match_search
            card.setVisible(is_visible)
            if is_visible:
                visible_cards.append(card)

        # Re-arrange visible cards into neat 3 columns
        columns = 3
        for idx, card in enumerate(visible_cards):
            r = idx // columns
            c = idx % columns
            self.grid_layout.addWidget(card, r, c)

        if self.current_search:
            self.header_subtitle.setText(f"Showing {len(visible_cards)} matching tools for '{self.current_search}'")
        else:
            for _, cid, title, subtitle in CATEGORY_INFO:
                if cid == self.active_category:
                    self.header_subtitle.setText(subtitle)
                    break

    def open_tool(self, tool_name: str):
        """Open the tool workspace inside the stacked widget."""
        action_map = {name: act for name, _, _, act in TOOL_DEFINITIONS}
        action_type = action_map.get(tool_name, tool_name)

        # Remove previous studio widget if any
        if self.stack.count() > 1:
            prev = self.stack.widget(1)
            self.stack.removeWidget(prev)
            prev.deleteLater()

        # Instantiate fresh studio view and embed in stack
        studio_view = secondWindow(action_type)
        studio_view.closed_to_home.connect(self.show_dashboard)
        self.stack.addWidget(studio_view)
        self.stack.setCurrentIndex(1)

    def show_dashboard(self):
        """Return to dashboard view."""
        self.stack.setCurrentIndex(0)

    def toggle_theme(self):
        """Toggle between Dark and Light mode."""
        self.current_theme = "dark" if self.current_theme == "light" else "light"
        apply_theme(self.current_theme)
        set_setting("theme", self.current_theme)
        self.theme_btn.setText("🌙 Dark Mode" if self.current_theme == "light" else "☀️ Light Mode")
        for card, _, _, _ in self.card_widgets:
            card.apply_card_theme(self.current_theme == "dark")

    def show_history_dialog(self):
        """Open the recent conversions history dialog."""
        dialog = HistoryDialog(self)
        dialog.exec()

    def open_logs_folder(self):
        """Open the application logs folder in file manager."""
        logs_dir = os.path.join(get_app_data_dir(), "logs")
        os.makedirs(logs_dir, exist_ok=True)
        if sys.platform == 'win32':
            os.startfile(logs_dir)
        else:
            QMessageBox.information(self, "Logs Location", f"Log files are stored at:\n{logs_dir}")

    def show_about_dialog(self):
        """Display About application dialog."""
        QMessageBox.about(
            self,
            "About PDF Converter",
            "<h3>PDF Converter</h3>"
            "<p>Version 3.0.0</p>"
            "<p>An enterprise-grade desktop suite for document conversion, "
            "merging, splitting, high-ratio compression, OCR, and encryption.</p>"
            "<p><b>Engine:</b> PyQt6, PyMuPDF (C++ MuPDF), pikepdf (C++ QPDF), openpyxl, python-pptx, Chromium Headless</p>"
        )


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
