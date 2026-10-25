"""
Modern Studio Workspace for PDF Converter Pro.
Provides an interactive drag-and-drop workspace, staged file inspection,
drag-to-reorder for merging, options panels for split/rotate/watermark/protect/OCR,
background asynchronous conversion with cancellation, and flexible download options.
"""

import os
import sys
import shutil
import logging
import zipfile
from typing import List, Optional

from PyQt6.QtWidgets import (
    QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QWidget,
    QLabel, QFileDialog, QProgressBar, QMessageBox, QSpacerItem,
    QSizePolicy, QSpinBox, QCheckBox, QLineEdit, QListWidget,
    QListWidgetItem, QFrame, QComboBox, QSlider
)
from PyQt6.QtGui import QIcon, QPainter, QPen, QFont, QColor, QRegularExpressionValidator
from PyQt6.QtCore import Qt, QRegularExpression, pyqtSignal, QSize

from app_paths import get_temp_dir, clean_dir
from pdfconvter import PDFConverter
from icons import get_tool_icon
from worker import ConversionWorker
from history import add_history_entry
from ui_components import ModernDropArea

logger = logging.getLogger(__name__)


def format_file_size(size_bytes: int) -> str:
    """Format file size in human-readable units."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


class secondWindow(QMainWindow):
    """
    Modern Tool Studio Workspace view with staged files, progress monitoring, and output download.
    Can run standalone or embedded.
    """
    closed_to_home = pyqtSignal()

    def __init__(self, value: str, parent=None):
        super().__init__(parent)
        self.value = value
        self.output_files: List[str] = []
        self.worker: Optional[ConversionWorker] = None

        # User-isolated working directories
        self.temp_input_dir = get_temp_dir("Temp")
        self.temp_output_dir = get_temp_dir("Temp1")
        self.obj = PDFConverter(input_dir=self.temp_input_dir, output_dir=self.temp_output_dir)

        self.setWindowTitle(f"PDF Converter Pro — {self.value}")
        self.setWindowIcon(get_tool_icon("app_icon", size=64))
        self.resize(1080, 800)
        self.setMinimumSize(850, 650)

        self.setup_ui()

    def setup_ui(self):
        self.central_widget = QWidget()
        self.central_widget.setObjectName("studioPage")
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(40, 20, 40, 25)
        self.main_layout.setSpacing(14)

        self.setup_header()
        self.setup_title()
        self.setup_file_management()

        # Context-sensitive option panels
        if self.value == "Split":
            self.setup_split_options()
        elif self.value == "Rotate":
            self.setup_rotate_options()
        elif self.value == "Protect":
            self.setup_protect_options()
        elif self.value == "Unlock":
            self.setup_unlock_options()
        elif self.value == "Watermark":
            self.setup_watermark_options()
        elif self.value == "Ocr":
            self.setup_ocr_options()

        self.setup_conversion_controls()

    def setup_header(self):
        """Top navigation bar."""
        top_bar = QHBoxLayout()

        self.back_button = QPushButton("← Back to Dashboard")
        self.back_button.setProperty("class", "secondaryBtn")
        self.back_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.back_button.clicked.connect(self.go_back)
        top_bar.addWidget(self.back_button)

        top_bar.addStretch()

        app_title = QLabel("PDF Converter Pro")
        app_title.setStyleSheet("font-size: 16px; font-weight: 700; color: #64748B;")
        top_bar.addWidget(app_title)

        self.main_layout.addLayout(top_bar)

    def setup_title(self):
        """Header title and subtitle for current action."""
        titles = {
            "Merge": ("Merge PDF Files", "Combine multiple PDF documents in your preferred order (drag items below to reorder)."),
            "Compression": ("Compress PDF", "Intelligently downsample images and optimize streams for 60-80% file size reduction."),
            "Split": ("Split PDF Document", "Extract specific pages, page ranges, or split into single pages."),
            "Rotate": ("Rotate PDF Pages", "Rotate all or specific pages clockwise by 90°, 180°, or 270°."),
            "Protect": ("Protect PDF with Password", "Encrypt PDF with military-grade AES-256 password protection."),
            "Unlock": ("Unlock PDF Document", "Decrypt and remove password restrictions from a PDF file."),
            "Watermark": ("Watermark PDF", "Stamp custom semi-transparent vector text across all pages."),
            "Ocr": ("OCR / Make Searchable", "Extract text from scanned PDFs and generate searchable, selectable documents."),
            "Html2pdf": ("HTML to PDF", "Render modern HTML web documents into PDF via Headless Chromium."),
            "pdf2excel": ("PDF to Excel", "Extract tables with C++ MuPDF engine into styled Excel spreadsheets."),
            "excel2pdf": ("Excel to PDF", "Convert spreadsheet workbooks into standardized PDF documents."),
            "pdf2word": ("PDF to Word", "Convert PDF layouts into editable Word (.docx) documents."),
            "word2pdf": ("Word to PDF", "Convert Microsoft Word (.docx) files into clean PDF documents."),
            "pdf2pptx": ("PDF to PowerPoint", "Generate editable PowerPoint presentations with textframes from PDF pages."),
            "pptx2pdf": ("PowerPoint to PDF", "Convert PowerPoint slide decks into PDF documents."),
            "pdf2jpg": ("PDF to JPG", "Render PDF pages into high-resolution images."),
            "jpg2pdf": ("JPG to PDF", "Assemble JPG and PNG images into a single PDF document.")
        }
        title_text, subtitle_text = titles.get(self.value, (self.value, "Process your documents easily."))

        title_container = QVBoxLayout()
        title_label = QLabel(title_text)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setProperty("class", "studioTitle")

        sub_label = QLabel(subtitle_text)
        sub_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub_label.setStyleSheet("font-size: 13px; color: #64748B; margin-top: 2px;")

        title_container.addWidget(title_label)
        title_container.addWidget(sub_label)
        self.main_layout.addLayout(title_container)

    def setup_file_management(self):
        """Drag-and-drop zone, browse button, and staged files list with reordering."""
        self.drop_area = ModernDropArea()
        self.drop_area.files_dropped.connect(self.add_files)
        self.main_layout.addWidget(self.drop_area)

        # Browse files row
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        self.browse_btn = QPushButton("Browse Files")
        self.browse_btn.setProperty("class", "primaryBtn")
        self.browse_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.browse_btn.clicked.connect(self.select_input_file)
        btn_row.addWidget(self.browse_btn)

        self.clear_btn = QPushButton("Clear All")
        self.clear_btn.setProperty("class", "secondaryBtn")
        self.clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_btn.clicked.connect(self.clear_staged_files)
        btn_row.addWidget(self.clear_btn)

        btn_row.addStretch()
        self.main_layout.addLayout(btn_row)

        # Staged files list with header
        list_header = QHBoxLayout()
        hint_text = " (Drag items to change merge sequence)" if self.value == "Merge" else ""
        self.files_count_label = QLabel(f"Staged Files (0){hint_text}:")
        self.files_count_label.setProperty("class", "stagedHeader")
        list_header.addWidget(self.files_count_label)
        list_header.addStretch()

        self.remove_item_btn = QPushButton("Remove Selected")
        self.remove_item_btn.setStyleSheet("font-size: 12px; color: #EF4444; border: none; background: transparent; font-weight: 600;")
        self.remove_item_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.remove_item_btn.clicked.connect(self.remove_selected_file)
        list_header.addWidget(self.remove_item_btn)
        self.main_layout.addLayout(list_header)

        # File list with Drag-to-Reorder enabled!
        self.file_list_widget = QListWidget()
        self.file_list_widget.setObjectName("stagedList")
        self.file_list_widget.setFixedHeight(120)
        self.file_list_widget.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        self.file_list_widget.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.main_layout.addWidget(self.file_list_widget)

    # ----------------------------------------------------------------------
    # Option Panels
    # ----------------------------------------------------------------------
    def setup_split_options(self):
        card = QFrame()
        card.setProperty("class", "optionsCard")
        layout = QHBoxLayout(card)
        layout.setContentsMargins(15, 8, 15, 8)

        lbl_range = QLabel("Page Range:")
        lbl_range.setStyleSheet("font-weight: 600;")
        layout.addWidget(lbl_range)

        self.spin_box_from = QSpinBox()
        self.spin_box_from.setRange(0, 9999)
        self.spin_box_from.setPrefix("From: ")
        layout.addWidget(self.spin_box_from)

        self.spin_box_to = QSpinBox()
        self.spin_box_to.setRange(0, 9999)
        self.spin_box_to.setPrefix("To: ")
        layout.addWidget(self.spin_box_to)

        layout.addSpacing(15)

        self.checkbox_single_page = QCheckBox("Split all single pages")
        self.checkbox_single_page.setStyleSheet("font-weight: 600;")
        layout.addWidget(self.checkbox_single_page)

        layout.addSpacing(15)

        lbl_extract = QLabel("Pages:")
        layout.addWidget(lbl_extract)

        self.input_pages_to_extract = QLineEdit()
        self.input_pages_to_extract.setPlaceholderText("e.g. 1, 3, 5")
        self.input_pages_to_extract.setFixedWidth(110)
        validator = QRegularExpressionValidator(QRegularExpression(r"[0-9,\s]+"))
        self.input_pages_to_extract.setValidator(validator)
        layout.addWidget(self.input_pages_to_extract)

        self.main_layout.addWidget(card)

    def setup_rotate_options(self):
        card = QFrame()
        card.setProperty("class", "optionsCard")
        layout = QHBoxLayout(card)
        layout.setContentsMargins(20, 8, 20, 8)

        lbl = QLabel("Rotation Angle:")
        lbl.setStyleSheet("font-weight: 600;")
        layout.addWidget(lbl)

        self.rotate_combo = QComboBox()
        self.rotate_combo.addItems(["90° Clockwise", "180° Flip", "270° (Counter-Clockwise)"])
        layout.addWidget(self.rotate_combo)
        layout.addStretch()
        self.main_layout.addWidget(card)

    def setup_protect_options(self):
        card = QFrame()
        card.setProperty("class", "optionsCard")
        layout = QHBoxLayout(card)
        layout.setContentsMargins(20, 8, 20, 8)

        lbl = QLabel("Set Password:")
        lbl.setStyleSheet("font-weight: 600;")
        layout.addWidget(lbl)

        self.protect_pwd_input = QLineEdit()
        self.protect_pwd_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.protect_pwd_input.setPlaceholderText("Enter password")
        self.protect_pwd_input.setFixedWidth(180)
        layout.addWidget(self.protect_pwd_input)

        lbl2 = QLabel("Confirm:")
        lbl2.setStyleSheet("font-weight: 600;")
        layout.addWidget(lbl2)

        self.protect_pwd_confirm = QLineEdit()
        self.protect_pwd_confirm.setEchoMode(QLineEdit.EchoMode.Password)
        self.protect_pwd_confirm.setPlaceholderText("Confirm password")
        self.protect_pwd_confirm.setFixedWidth(180)
        layout.addWidget(self.protect_pwd_confirm)

        layout.addStretch()
        self.main_layout.addWidget(card)

    def setup_unlock_options(self):
        card = QFrame()
        card.setProperty("class", "optionsCard")
        layout = QHBoxLayout(card)
        layout.setContentsMargins(20, 8, 20, 8)

        lbl = QLabel("PDF Password:")
        lbl.setStyleSheet("font-weight: 600;")
        layout.addWidget(lbl)

        self.unlock_pwd_input = QLineEdit()
        self.unlock_pwd_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.unlock_pwd_input.setPlaceholderText("Enter current password to unlock")
        self.unlock_pwd_input.setFixedWidth(240)
        layout.addWidget(self.unlock_pwd_input)

        layout.addStretch()
        self.main_layout.addWidget(card)

    def setup_watermark_options(self):
        card = QFrame()
        card.setProperty("class", "optionsCard")
        layout = QHBoxLayout(card)
        layout.setContentsMargins(20, 8, 20, 8)

        lbl_text = QLabel("Watermark Text:")
        lbl_text.setStyleSheet("font-weight: 600;")
        layout.addWidget(lbl_text)

        self.watermark_input = QLineEdit("CONFIDENTIAL")
        self.watermark_input.setFixedWidth(160)
        layout.addWidget(self.watermark_input)

        layout.addSpacing(15)

        lbl_angle = QLabel("Angle:")
        lbl_angle.setStyleSheet("font-weight: 600;")
        layout.addWidget(lbl_angle)

        self.watermark_angle_combo = QComboBox()
        self.watermark_angle_combo.addItems(["45° Diagonal", "0° Horizontal", "90° Vertical", "-45° Diagonal"])
        layout.addWidget(self.watermark_angle_combo)

        layout.addSpacing(15)

        lbl_opacity = QLabel("Opacity:")
        lbl_opacity.setStyleSheet("font-weight: 600;")
        layout.addWidget(lbl_opacity)

        self.opacity_slider = QSlider(Qt.Orientation.Horizontal)
        self.opacity_slider.setRange(10, 90)
        self.opacity_slider.setValue(30)
        self.opacity_slider.setFixedWidth(100)
        layout.addWidget(self.opacity_slider)

        layout.addStretch()
        self.main_layout.addWidget(card)

    def setup_ocr_options(self):
        card = QFrame()
        card.setProperty("class", "optionsCard")
        layout = QHBoxLayout(card)
        layout.setContentsMargins(20, 8, 20, 8)

        lbl = QLabel("OCR Mode:")
        lbl.setStyleSheet("font-weight: 600;")
        layout.addWidget(lbl)

        info_lbl = QLabel("High-Accuracy C++ MuPDF OCR Engine with Searchable Text Layer")
        info_lbl.setStyleSheet("color: #64748B; font-size: 13px;")
        layout.addWidget(info_lbl)
        layout.addStretch()
        self.main_layout.addWidget(card)

    def setup_conversion_controls(self):
        """Convert button, cancel button, progress bar, and download buttons."""
        controls_layout = QVBoxLayout()

        # Status label
        self.status_label = QLabel("Ready to process")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("font-size: 13px; color: #64748B; font-weight: 500;")
        controls_layout.addWidget(self.status_label)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(8)
        self.progress_bar.setTextVisible(False)
        controls_layout.addWidget(self.progress_bar)

        # Action Buttons Row
        action_row = QHBoxLayout()
        action_row.addStretch()

        self.convert_button = QPushButton(f"⚡ Execute {self.value}")
        self.convert_button.setProperty("class", "primaryBtn")
        self.convert_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.convert_button.clicked.connect(self.convert_file)
        action_row.addWidget(self.convert_button)

        # Cancel button
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setProperty("class", "secondaryBtn")
        self.cancel_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self.cancel_conversion)
        action_row.addWidget(self.cancel_button)

        self.download_button = QPushButton("Save Output")
        self.download_button.setProperty("class", "successBtn")
        self.download_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.download_button.setEnabled(False)
        self.download_button.clicked.connect(self.download_file)
        action_row.addWidget(self.download_button)

        action_row.addStretch()
        controls_layout.addLayout(action_row)

        self.main_layout.addLayout(controls_layout)

    # ----------------------------------------------------------------------
    # Event Handlers & Business Logic
    # ----------------------------------------------------------------------
    def get_ordered_staged_files(self) -> List[str]:
        """Read ordered file paths directly from the reorderable QListWidget."""
        paths = []
        for i in range(self.file_list_widget.count()):
            item = self.file_list_widget.item(i)
            file_path = item.data(Qt.ItemDataRole.UserRole)
            if file_path and os.path.exists(file_path):
                paths.append(file_path)
        return paths

    def go_back(self):
        """Navigate back to the main launcher screen."""
        self.closed_to_home.emit()
        self.close()

    def add_files(self, file_paths: List[str]):
        """Add files to staged list and update UI."""
        existing = set(self.get_ordered_staged_files())
        added_count = 0
        for path in file_paths:
            if path not in existing and os.path.isfile(path):
                size_str = format_file_size(os.path.getsize(path))
                item = QListWidgetItem(f"📄  {os.path.basename(path)}  •  {size_str}")
                item.setData(Qt.ItemDataRole.UserRole, path)
                item.setToolTip(path)
                self.file_list_widget.addItem(item)
                existing.add(path)
                added_count += 1

        total = self.file_list_widget.count()
        hint = " (Drag items to change merge sequence)" if self.value == "Merge" else ""
        self.files_count_label.setText(f"Staged Files ({total}){hint}:")
        if added_count > 0:
            self.status_label.setText(f"Added {added_count} file(s).")

    def remove_selected_file(self):
        """Remove currently selected item from the staged list."""
        current_row = self.file_list_widget.currentRow()
        if current_row >= 0:
            self.file_list_widget.takeItem(current_row)
            total = self.file_list_widget.count()
            hint = " (Drag items to change merge sequence)" if self.value == "Merge" else ""
            self.files_count_label.setText(f"Staged Files ({total}){hint}:")
            self.status_label.setText("Item removed.")

    def clear_staged_files(self):
        """Clear all staged input files."""
        self.file_list_widget.clear()
        hint = " (Drag items to change merge sequence)" if self.value == "Merge" else ""
        self.files_count_label.setText(f"Staged Files (0){hint}:")
        self.output_files.clear()
        self.download_button.setEnabled(False)
        self.status_label.setText("Cleared all files.")
        self.progress_bar.setValue(0)

    def select_input_file(self):
        """Open system file picker."""
        file_filters = {
            "Merge": "PDF Files (*.pdf)",
            "Compression": "PDF Files (*.pdf)",
            "Split": "PDF Files (*.pdf)",
            "Rotate": "PDF Files (*.pdf)",
            "Protect": "PDF Files (*.pdf)",
            "Unlock": "PDF Files (*.pdf)",
            "Watermark": "PDF Files (*.pdf)",
            "Ocr": "PDF Files (*.pdf)",
            "Html2pdf": "HTML Files (*.html *.htm)",
            "pdf2excel": "PDF Files (*.pdf)",
            "excel2pdf": "Excel Files (*.xlsx *.xls)",
            "pdf2word": "PDF Files (*.pdf)",
            "word2pdf": "Word Documents (*.docx *.doc)",
            "pdf2pptx": "PDF Files (*.pdf)",
            "pptx2pdf": "PowerPoint Files (*.pptx *.ppt)",
            "pdf2jpg": "PDF Files (*.pdf)",
            "jpg2pdf": "Image Files (*.jpg *.jpeg *.png)"
        }
        filter_str = file_filters.get(self.value, "Supported Files (*.*);;All Files (*.*)")
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Select input files", "", filter_str)
        if file_paths:
            self.add_files(file_paths)

    def convert_file(self):
        """Start asynchronous conversion worker with user options."""
        staged = self.get_ordered_staged_files()
        if not staged:
            QMessageBox.warning(self, "No Files Selected", "Please select or drop files to convert first.")
            return

        clean_dir(self.temp_output_dir)

        # Parameters for various tools
        from_page = 0
        to_page = 0
        is_single = False
        extract_pages = []
        password = ""
        rotation_angle = 90
        watermark_text = ""
        watermark_opacity = 0.3
        watermark_angle = 45

        if self.value == "Split":
            from_page = self.spin_box_from.value()
            to_page = self.spin_box_to.value()
            is_single = self.checkbox_single_page.isChecked()
            raw_extract = self.input_pages_to_extract.text().strip()
            if raw_extract:
                for part in raw_extract.split(","):
                    part = part.strip()
                    if part.isdigit():
                        extract_pages.append(int(part))

            if not is_single and not extract_pages and (from_page == 0 or to_page == 0):
                QMessageBox.warning(
                    self, "Split Options Needed",
                    "Please specify a page range, select 'Split all single pages', or enter pages to extract."
                )
                return

        elif self.value == "Rotate":
            angle_map = {0: 90, 1: 180, 2: 270}
            rotation_angle = angle_map.get(self.rotate_combo.currentIndex(), 90)

        elif self.value == "Protect":
            pwd1 = self.protect_pwd_input.text()
            pwd2 = self.protect_pwd_confirm.text()
            if not pwd1:
                QMessageBox.warning(self, "Password Required", "Please enter a password to protect the document.")
                return
            if pwd1 != pwd2:
                QMessageBox.warning(self, "Password Mismatch", "Passwords do not match. Please verify and retry.")
                return
            password = pwd1

        elif self.value == "Unlock":
            password = self.unlock_pwd_input.text()
            if not password:
                QMessageBox.warning(self, "Password Required", "Please enter the password to unlock the document.")
                return

        elif self.value == "Watermark":
            watermark_text = self.watermark_input.text().strip()
            if not watermark_text:
                watermark_text = "CONFIDENTIAL"
            watermark_opacity = self.opacity_slider.value() / 100.0
            angle_map = {0: 45, 1: 0, 2: 90, 3: -45}
            watermark_angle = angle_map.get(self.watermark_angle_combo.currentIndex(), 45)

        # UI state during conversion
        self.convert_button.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.browse_btn.setEnabled(False)
        self.download_button.setEnabled(False)
        self.progress_bar.setValue(0)
        self.status_label.setText("Processing...")

        # Launch worker
        self.worker = ConversionWorker(
            action_type=self.value,
            converter=self.obj,
            input_files=staged,
            output_dir=self.temp_output_dir,
            from_page=from_page,
            to_page=to_page,
            is_single_page=is_single,
            pages_to_extract=extract_pages,
            password=password,
            rotation_angle=rotation_angle,
            watermark_text=watermark_text,
            watermark_opacity=watermark_opacity,
            watermark_angle=watermark_angle,
            parent=self
        )
        self.worker.progress.connect(self.on_worker_progress)
        self.worker.finished_task.connect(self.on_worker_finished)
        self.worker.start()

    def cancel_conversion(self):
        """User requested cancellation."""
        if self.worker and self.worker.isRunning():
            self.status_label.setText("Cancelling...")
            self.cancel_button.setEnabled(False)
            self.worker.cancel()

    def on_worker_progress(self, percent: int, message: str):
        self.progress_bar.setValue(percent)
        self.status_label.setText(message)

    def on_worker_finished(self, success: bool, message: str, output_files: List[str]):
        self.convert_button.setEnabled(True)
        self.cancel_button.setEnabled(False)
        self.browse_btn.setEnabled(True)

        if success and output_files:
            self.output_files = output_files
            self.download_button.setEnabled(True)
            self.status_label.setText(f"Completed! {len(output_files)} file(s) generated.")
            add_history_entry(self.value, self.get_ordered_staged_files(), output_files)
            QMessageBox.information(self, "Success", f"{message}\n\nClick 'Save Output' to save your file(s).")
        else:
            self.status_label.setText("Conversion stopped.")
            if "cancelled" in message.lower():
                QMessageBox.information(self, "Cancelled", message)
            else:
                QMessageBox.critical(self, "Error", message)

    def download_file(self):
        """Save generated output files locally or as a ZIP archive."""
        if not self.output_files:
            QMessageBox.information(self, "No Output", "No generated files found to save.")
            return

        if len(self.output_files) == 1:
            src_file = self.output_files[0]
            default_name = os.path.basename(src_file)
            ext = os.path.splitext(default_name)[1].lower()
            save_path, _ = QFileDialog.getSaveFileName(
                self, "Save File", default_name, f"Converted File (*{ext});;All Files (*.*)"
            )
            if save_path:
                try:
                    shutil.copy2(src_file, save_path)
                    QMessageBox.information(self, "Saved", f"File saved successfully to:\n{save_path}")
                    self.status_label.setText("File saved successfully.")
                except Exception as e:
                    QMessageBox.critical(self, "Save Error", f"Could not save file: {e}")
        else:
            default_zip = f"{self.value}_output.zip"
            save_path, _ = QFileDialog.getSaveFileName(
                self, "Save All Files as ZIP", default_zip, "ZIP Archives (*.zip)"
            )
            if save_path:
                if not save_path.lower().endswith(".zip"):
                    save_path += ".zip"
                try:
                    with zipfile.ZipFile(save_path, "w", zipfile.ZIP_DEFLATED) as zip_out:
                        for f in self.output_files:
                            if os.path.exists(f):
                                zip_out.write(f, arcname=os.path.basename(f))
                    QMessageBox.information(self, "Saved", f"All files packaged and saved to:\n{save_path}")
                    self.status_label.setText("Archive saved successfully.")
                except Exception as e:
                    QMessageBox.critical(self, "Save Error", f"Could not save ZIP archive: {e}")
