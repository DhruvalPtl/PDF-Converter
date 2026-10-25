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

CATEGORY_NAMES = {
    "essential": "ESSENTIAL",
    "to_pdf": "TO PDF",
    "from_pdf": "FROM PDF",
    "security": "SECURITY",
}


class ModernToolCard(QFrame):
    """
    Sleek, elevated card widget representing a single document tool.
    Includes category color bar, tinted icon container, typography hierarchy, and hover effects.
    """
    clicked = pyqtSignal(str)

    def __init__(self, tool_name: str, description: str, category: str = "essential", parent=None):
        super().__init__(parent)
        self.tool_name = tool_name
        self.description = description
        self.category = category

        self.setProperty("class", "toolCard")
        self.setFixedHeight(94)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(14, 12, 16, 12)
        main_layout.setSpacing(14)

        # 1. Left Color Indicator Strip
        indicator = QFrame()
        indicator.setFixedWidth(4)
        cat_color = CATEGORY_COLORS.get(category, "#6366F1")
        indicator.setStyleSheet(f"background-color: {cat_color}; border-radius: 2px;")
        main_layout.addWidget(indicator)

        # 2. Icon Container (50x50)
        icon_label = QLabel()
        icon_label.setFixedSize(50, 50)
        icon_label.setPixmap(get_tool_icon(tool_name, size=100).pixmap(50, 50))
        main_layout.addWidget(icon_label)

        # 3. Content Column (Category Tag + Title + Description)
        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)
        text_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        title_label = QLabel(tool_name)
        title_label.setProperty("class", "cardTitle")
        top_row.addWidget(title_label)

        badge_label = QLabel(CATEGORY_NAMES.get(category, "TOOL"))
        badge_label.setProperty("class", "cardBadge")
        top_row.addWidget(badge_label)
        top_row.addStretch()

        text_layout.addLayout(top_row)

        desc_label = QLabel(description)
        desc_label.setProperty("class", "cardDesc")
        desc_label.setWordWrap(True)
        text_layout.addWidget(desc_label)

        main_layout.addLayout(text_layout, 1)

        # 4. Chevron Arrow
        arrow_label = QLabel("→")
        arrow_label.setProperty("class", "cardArrow")
        main_layout.addWidget(arrow_label)

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
