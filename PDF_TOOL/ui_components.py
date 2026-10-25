"""
Custom Modern UI Components for PDF Converter.
Includes ModernToolCard, ModernDropArea, and interactive staged file cards.
"""

import os
from PyQt6.QtWidgets import (
    QFrame, QHBoxLayout, QVBoxLayout, QLabel, QPushButton,
    QWidget, QSizePolicy
)
from PyQt6.QtCore import Qt, pyqtSignal, QSize
from PyQt6.QtGui import QIcon, QFont, QPainter, QColor, QPen

from icons import get_tool_icon

CATEGORY_COLORS = {
    "essential": "#EF4444",   # Red
    "to_pdf": "#3B82F6",      # Blue
    "from_pdf": "#10B981",    # Emerald
    "security": "#8B5CF6",    # Purple
}

CATEGORY_PILLS = {
    "essential": ("#DC2626", "#FEE2E2"),   # Red
    "to_pdf": ("#2563EB", "#DBEAFE"),      # Blue
    "from_pdf": ("#059669", "#D1FAE5"),    # Emerald
    "security": ("#7C3AED", "#EDE9FE"),    # Purple
}

CATEGORY_NAMES = {
    "essential": "ESSENTIAL",
    "to_pdf": "TO PDF",
    "from_pdf": "FROM PDF",
    "security": "SECURITY",
}


class ModernToolCard(QFrame):
    """
    Substantial, rich tool card widget representing a single document tool.
    Features large 68x68 vector badge, category accent indicator, bold typography,
    category tag, and interactive 'Open →' action pill.
    """
    clicked = pyqtSignal(str)

    def __init__(self, tool_name: str, description: str, category: str = "essential", parent=None):
        super().__init__(parent)
        self.tool_name = tool_name
        self.description = description
        self.category = category

        self.setProperty("class", "toolCard")
        self.setFixedHeight(122)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(18, 16, 20, 16)
        main_layout.setSpacing(18)

        # 1. Left Color Indicator Strip (5px)
        indicator = QFrame()
        indicator.setFixedWidth(5)
        cat_color = CATEGORY_COLORS.get(category, "#6366F1")
        indicator.setStyleSheet(f"background-color: {cat_color}; border-radius: 2.5px;")
        main_layout.addWidget(indicator)

        # 2. Large Bold Icon Container (68x68)
        icon_label = QLabel()
        icon_label.setFixedSize(68, 68)
        icon_label.setPixmap(get_tool_icon(tool_name, size=136).pixmap(68, 68))
        main_layout.addWidget(icon_label)

        # 3. Content Column (Category Tag + Title + Description)
        text_layout = QVBoxLayout()
        text_layout.setSpacing(4)
        text_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        title_label = QLabel(tool_name)
        title_label.setProperty("class", "cardTitle")
        top_row.addWidget(title_label)

        badge_label = QLabel(CATEGORY_NAMES.get(category, "TOOL"))
        badge_label.setProperty("class", "cardBadge")
        fg, bg = CATEGORY_PILLS.get(category, ("#4F46E5", "#EEF2FF"))
        badge_label.setStyleSheet(f"background-color: {bg}; color: {fg}; font-size: 10px; font-weight: 700; padding: 3px 8px; border-radius: 4px;")
        top_row.addWidget(badge_label)
        top_row.addStretch()

        text_layout.addLayout(top_row)

        desc_label = QLabel(description)
        desc_label.setProperty("class", "cardDesc")
        desc_label.setWordWrap(True)
        text_layout.addWidget(desc_label)

        main_layout.addLayout(text_layout, 1)

        # 4. Action Button ('Open →')
        self.action_btn = QPushButton("Open →")
        self.action_btn.setObjectName("cardActionBtn")
        self.action_btn.setFixedSize(92, 38)
        self.action_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.action_btn.clicked.connect(lambda: self.clicked.emit(self.tool_name))
        self.action_btn.setStyleSheet("""
            QPushButton#cardActionBtn {
                background-color: #4F46E5;
                color: #FFFFFF;
                font-size: 13px;
                font-weight: 700;
                border-radius: 8px;
                border: none;
                padding: 6px 14px;
            }
            QPushButton#cardActionBtn:hover {
                background-color: #4338CA;
            }
        """)
        main_layout.addWidget(self.action_btn, 0, Qt.AlignmentFlag.AlignVCenter)

    def apply_card_theme(self, is_dark: bool = False):
        """Update inline styles for dark or light mode."""
        if is_dark:
            self.action_btn.setStyleSheet("""
                QPushButton#cardActionBtn {
                    background-color: #6366F1;
                    color: #FFFFFF;
                    font-size: 13px;
                    font-weight: 700;
                    border-radius: 8px;
                    border: none;
                    padding: 6px 14px;
                }
                QPushButton#cardActionBtn:hover {
                    background-color: #4F46E5;
                }
            """)
        else:
            self.action_btn.setStyleSheet("""
                QPushButton#cardActionBtn {
                    background-color: #4F46E5;
                    color: #FFFFFF;
                    font-size: 13px;
                    font-weight: 700;
                    border-radius: 8px;
                    border: none;
                    padding: 6px 14px;
                }
                QPushButton#cardActionBtn:hover {
                    background-color: #4338CA;
                }
            """)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit(self.tool_name)
        super().mousePressEvent(event)


class ModernDropArea(QFrame):
    """
    Drag & Drop dropzone with dashed border, upload iconography, and format tags.
    """
    files_dropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setObjectName("dropAreaFrame")
        self.setFixedHeight(170)
        self.is_hovered = False

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(6)

        title = QLabel("Drag & Drop your files here")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setProperty("class", "dropTitle")
        layout.addWidget(title)

        sub = QLabel("or browse from your local storage")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub.setProperty("class", "dropSub")
        layout.addWidget(sub)

        badges_layout = QHBoxLayout()
        badges_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        badges_layout.setSpacing(8)

        formats = ["PDF", "Word .docx", "Excel .xlsx", "PowerPoint .pptx", "Images", "HTML"]
        for fmt in formats:
            lbl = QLabel(fmt)
            lbl.setStyleSheet("""
                background-color: rgba(100, 116, 139, 0.12);
                font-size: 11px;
                font-weight: 600;
                padding: 3px 8px;
                border-radius: 5px;
            """)
            badges_layout.addWidget(lbl)

        layout.addLayout(badges_layout)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            self.is_hovered = True
            self.update()
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.is_hovered = False
        self.update()

    def dropEvent(self, event):
        self.is_hovered = False
        self.update()
        if event.mimeData().hasUrls():
            file_paths = []
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    path = url.toLocalFile()
                    if os.path.isfile(path):
                        file_paths.append(path)
            if file_paths:
                self.files_dropped.emit(file_paths)
                event.acceptProposedAction()
