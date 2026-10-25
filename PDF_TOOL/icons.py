"""
Dynamic icon and visual asset generator for PDF Converter.
Generates clean, modern, high-DPI vector-styled QIcons when asset images are not present on disk.
"""

import os
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QColor, QFont, QPen, QBrush, QPainterPath
from PyQt6.QtCore import Qt, QRectF, QPointF
from app_paths import get_resource_path

# Visual color palette for document types
PALETTE = {
    "pdf": QColor("#E53935"),      # Red
    "word": QColor("#1E88E5"),     # Blue
    "excel": QColor("#43A047"),    # Green
    "pptx": QColor("#FB8C00"),     # Orange
    "image": QColor("#8E24AA"),    # Purple
    "html": QColor("#00ACC1"),     # Cyan
    "tool": QColor("#546E7A"),     # Slate
}

def create_badge_icon(badge_text: str, bg_color: QColor, sub_text: str = "", size: int = 128) -> QIcon:
    """
    Renders a crisp, modern rounded-corner badge icon with label and subtitle.
    """
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)

    margin = 8.0
    rect = QRectF(margin, margin, size - 2 * margin, size - 2 * margin)

    # Draw rounded background
    path = QPainterPath()
    path.addRoundedRect(rect, 16.0, 16.0)
    painter.fillPath(path, QBrush(bg_color))

    # Folded corner effect (top-right)
    corner_size = 24.0
    corner_path = QPainterPath()
    corner_path.moveTo(rect.right() - corner_size, rect.top())
    corner_path.lineTo(rect.right(), rect.top() + corner_size)
    corner_path.lineTo(rect.right() - corner_size, rect.top() + corner_size)
    corner_path.closeSubpath()
    painter.fillPath(corner_path, QBrush(QColor(255, 255, 255, 60)))

    # Dynamic font scaling to guarantee crisp rendering without clipping
    badge_len = len(badge_text)
    if badge_len <= 3:
        font_scale = 0.24
    elif badge_len <= 5:
        font_scale = 0.19
    elif badge_len <= 7:
        font_scale = 0.15
    else:
        font_scale = 0.12

    painter.setPen(QColor(255, 255, 255))
    font = QFont("Arial", int(size * font_scale), QFont.Weight.Bold)
    painter.setFont(font)

    if sub_text:
        main_rect = QRectF(rect.left(), rect.top() + size * 0.15, rect.width(), size * 0.35)
        painter.drawText(main_rect, Qt.AlignmentFlag.AlignCenter, badge_text)

        sub_len = len(sub_text)
        sub_scale = 0.14 if sub_len <= 5 else 0.11
        sub_font = QFont("Arial", int(size * sub_scale), QFont.Weight.DemiBold)
        painter.setFont(sub_font)
        painter.setPen(QColor(255, 255, 255, 220))
        sub_rect = QRectF(rect.left(), rect.top() + size * 0.50, rect.width(), size * 0.30)
        painter.drawText(sub_rect, Qt.AlignmentFlag.AlignCenter, sub_text)
    else:
        font = QFont("Arial", int(size * 0.30), QFont.Weight.Bold)
        painter.setFont(font)
        painter.drawText(rect, Qt.AlignmentFlag.AlignCenter, badge_text)

    painter.end()
    return QIcon(pixmap)

def get_tool_icon(tool_name: str, size: int = 128) -> QIcon:
    """
    Returns a QIcon for the specified tool.
    Checks disk first; if not found, dynamically renders an icon.
    """
    # Mapping to potential image filenames
    filename_map = {
        "Compress PDF": "compress.png",
        "Merge PDF": "merge.png",
        "Split PDF": "split.png",
        "HTML to PDF": "htmltopdf.png",
        "PDF to Excel": "pdftoexcel.png",
        "Excel to PDF": "exceltopdf.png",
        "PDF to Word": "pdftoword.png",
        "Word to PDF": "wordtopdf.png",
        "PDF to PowerPoint": "pdftopp.png",
        "PowerPoint to PDF": "pptopdf.png",
        "PDF to JPG": "pdftojpg.png",
        "JPG to PDF": "jpgtopdf.png",
        "app_icon": "1.png"
    }

    if tool_name in filename_map:
        rel_path = os.path.join("images", filename_map[tool_name])
        disk_path = get_resource_path(rel_path)
        if os.path.exists(disk_path):
            return QIcon(disk_path)

    # Dynamic fallback generation
    configs = {
        "Compress PDF": ("ZIP", PALETTE["tool"], "PDF"),
        "Merge PDF": ("MERGE", PALETTE["pdf"], "PDFs"),
        "Split PDF": ("SPLIT", PALETTE["tool"], "PAGES"),
        "HTML to PDF": ("</>", PALETTE["html"], "→ PDF"),
        "PDF to Excel": ("PDF", PALETTE["excel"], "→ XLSX"),
        "Excel to PDF": ("XLSX", PALETTE["excel"], "→ PDF"),
        "PDF to Word": ("PDF", PALETTE["word"], "→ DOCX"),
        "Word to PDF": ("DOCX", PALETTE["word"], "→ PDF"),
        "PDF to PowerPoint": ("PDF", PALETTE["pptx"], "→ PPTX"),
        "PowerPoint to PDF": ("PPTX", PALETTE["pptx"], "→ PDF"),
        "PDF to JPG": ("PDF", PALETTE["image"], "→ JPG"),
        "JPG to PDF": ("JPG", PALETTE["image"], "→ PDF"),
        "Protect PDF": ("LOCK", PALETTE["pdf"], "ENCRYPT"),
        "Unlock PDF": ("OPEN", PALETTE["excel"], "DECRYPT"),
        "Rotate PDF": ("ROTATE", PALETTE["tool"], "90°"),
        "Watermark PDF": ("STAMP", PALETTE["pptx"], "TEXT"),
        "OCR Scanned PDF": ("OCR", PALETTE["image"], "SCAN"),
        "app_icon": ("PDF", PALETTE["pdf"], "")
    }

    badge_text, bg_color, sub_text = configs.get(tool_name, (tool_name[:4].upper(), PALETTE["tool"], ""))
    return create_badge_icon(badge_text, bg_color, sub_text, size=size)
