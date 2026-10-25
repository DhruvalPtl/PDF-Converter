"""
Background Worker thread for non-blocking, cancellable file conversions in PDF Converter.
Subclasses QThread and emits real-time progress and completion signals.
"""

import logging
from typing import List, Optional
from PyQt6.QtCore import QThread, pyqtSignal
from pdfconvter import PDFConverter

logger = logging.getLogger(__name__)


class ConversionWorker(QThread):
    """
    Executes heavy document conversion tasks on a background thread
    to keep the PyQt UI completely responsive. Supports cooperative cancellation.
    """
    # Signals
    started_task = pyqtSignal()
    progress = pyqtSignal(int, str)  # (percentage: 0-100, status_message)
    finished_task = pyqtSignal(bool, str, list)  # (success, message, output_files)

    def __init__(
        self,
        action_type: str,
        converter: PDFConverter,
        input_files: List[str],
        output_dir: str,
        from_page: int = 0,
        to_page: int = 0,
        is_single_page: bool = False,
        pages_to_extract: Optional[List[int]] = None,
        password: str = "",
        rotation_angle: int = 90,
        watermark_text: str = "",
        watermark_opacity: float = 0.3,
        watermark_angle: int = 45,
        parent=None
    ):
        super().__init__(parent)
        self.action_type = action_type
        self.converter = converter
        self.input_files = input_files
        self.output_dir = output_dir
        self.from_page = from_page
        self.to_page = to_page
        self.is_single_page = is_single_page
        self.pages_to_extract = pages_to_extract or []
        self.password = password
        self.rotation_angle = rotation_angle
        self.watermark_text = watermark_text
        self.watermark_opacity = watermark_opacity
        self.watermark_angle = watermark_angle
        self._is_cancelled = False

    def cancel(self):
        """Request cooperative cancellation."""
        self._is_cancelled = True

    def is_cancelled(self) -> bool:
        return self._is_cancelled

    def run(self):
        self.started_task.emit()
        self.progress.emit(10, "Starting processing...")

        def on_progress(percent: int, text: str):
            self.progress.emit(percent, text)

        try:
            output_files: List[str] = []
            success_msg = "Operation completed successfully!"

            if self.action_type == "Merge":
                output_files = self.converter.pdfmerger(
                    self.input_files, self.output_dir, progress_callback=on_progress, is_cancelled=self.is_cancelled
                )
                success_msg = f"Merged {len(self.input_files)} PDF files successfully!"

            elif self.action_type == "Compression":
                output_files = self.converter.pdfcompress(
                    self.input_files, self.output_dir, progress_callback=on_progress, is_cancelled=self.is_cancelled
                )
                success_msg = "Files compressed successfully!"

            elif self.action_type == "Split":
                self.progress.emit(25, "Splitting pages...")
                if self.is_single_page:
                    output_files.extend(
                        self.converter.single_pages(self.input_files, self.output_dir, is_cancelled=self.is_cancelled)
                    )
                if self.pages_to_extract:
                    output_files.extend(
                        self.converter.extract_pages(self.pages_to_extract, self.input_files, self.output_dir, is_cancelled=self.is_cancelled)
                    )
                if self.from_page > 0 and self.to_page >= self.from_page:
                    output_files.extend(
                        self.converter.split_range(self.from_page, self.to_page, self.input_files, self.output_dir, is_cancelled=self.is_cancelled)
                    )
                success_msg = f"Generated {len(output_files)} split PDF files!"

            elif self.action_type == "Rotate":
                output_files = self.converter.pdfrotate(
                    self.input_files, rotation_angle=self.rotation_angle, output_dir=self.output_dir,
                    progress_callback=on_progress, is_cancelled=self.is_cancelled
                )
                success_msg = f"Rotated pages by {self.rotation_angle}° successfully!"

            elif self.action_type == "Protect":
                output_files = self.converter.pdfprotect(
                    self.input_files, user_password=self.password, output_dir=self.output_dir,
                    progress_callback=on_progress, is_cancelled=self.is_cancelled
                )
                success_msg = "Password protection applied successfully!"

            elif self.action_type == "Unlock":
                output_files = self.converter.pdfunlock(
                    self.input_files, password=self.password, output_dir=self.output_dir,
                    progress_callback=on_progress, is_cancelled=self.is_cancelled
                )
                success_msg = "Password removed from PDF successfully!"

            elif self.action_type == "Watermark":
                output_files = self.converter.pdfwatermark(
                    self.input_files, watermark_text=self.watermark_text or "CONFIDENTIAL",
                    opacity=self.watermark_opacity, angle=self.watermark_angle, output_dir=self.output_dir,
                    progress_callback=on_progress, is_cancelled=self.is_cancelled
                )
                success_msg = "Watermark stamped on document successfully!"

            elif self.action_type == "Ocr":
                output_files = self.converter.pdfocr(
                    self.input_files, output_dir=self.output_dir,
                    progress_callback=on_progress, is_cancelled=self.is_cancelled
                )
                success_msg = "OCR processed and searchable PDF generated!"

            elif self.action_type == "Html2pdf":
                self.progress.emit(30, "Converting HTML to PDF...")
                output_files = self.converter.html2pdf(self.input_files, self.output_dir, is_cancelled=self.is_cancelled)
                success_msg = "HTML files converted to PDF!"

            elif self.action_type == "pdf2word":
                output_files = self.converter.pdf2word(
                    self.input_files, self.output_dir, progress_callback=on_progress, is_cancelled=self.is_cancelled
                )
                success_msg = "PDF converted to Word document (.docx)!"

            elif self.action_type == "word2pdf":
                self.progress.emit(40, "Converting Word to PDF...")
                output_files = self.converter.word2pdf(self.input_files, self.output_dir, is_cancelled=self.is_cancelled)
                success_msg = "Word documents converted to PDF!"

            elif self.action_type == "excel2pdf":
                self.progress.emit(40, "Converting Excel to PDF...")
                output_files = self.converter.excel2pdf(self.input_files, self.output_dir, is_cancelled=self.is_cancelled)
                success_msg = "Excel spreadsheets converted to PDF!"

            elif self.action_type == "pdf2excel":
                self.progress.emit(35, "Extracting tables from PDF...")
                output_files = self.converter.pdf2excel(self.input_files, self.output_dir, is_cancelled=self.is_cancelled)
                success_msg = "Tables extracted to Excel (.xlsx)!"

            elif self.action_type == "pdf2pptx":
                self.progress.emit(30, "Generating PowerPoint slides...")
                output_files = self.converter.pdf_to_pptx(self.input_files, self.output_dir, is_cancelled=self.is_cancelled)
                success_msg = "PDF converted to PowerPoint presentation (.pptx)!"

            elif self.action_type == "pptx2pdf":
                self.progress.emit(40, "Converting PowerPoint to PDF...")
                output_files = self.converter.pptx2pdf(self.input_files, self.output_dir, is_cancelled=self.is_cancelled)
                success_msg = "PowerPoint presentation converted to PDF!"

            elif self.action_type == "pdf2jpg":
                self.progress.emit(35, "Rendering PDF pages to images...")
                output_files = self.converter.pdf2images(self.input_files, self.output_dir, is_cancelled=self.is_cancelled)
                success_msg = f"Rendered {len(output_files)} page images!"

            elif self.action_type == "jpg2pdf":
                self.progress.emit(40, "Assembling images into PDF...")
                output_files = self.converter.image2pdf(self.input_files, self.output_dir, is_cancelled=self.is_cancelled)
                success_msg = "Images combined into PDF!"

            else:
                raise ValueError(f"Unknown conversion action: {self.action_type}")

            if self._is_cancelled:
                self.progress.emit(0, "Cancelled by user.")
                self.finished_task.emit(False, "Operation was cancelled.", [])
            elif not output_files:
                self.progress.emit(100, "Done (No output files generated).")
                self.finished_task.emit(False, "No output files were produced. Please check your input files.", [])
            else:
                self.progress.emit(100, "Done!")
                self.finished_task.emit(True, success_msg, output_files)

        except InterruptedError:
            self.progress.emit(0, "Cancelled by user.")
            self.finished_task.emit(False, "Operation was cancelled by user.", [])
        except Exception as e:
            logger.error(f"Worker task '{self.action_type}' failed: {e}", exc_info=True)
            self.progress.emit(0, "Error occurred")
            self.finished_task.emit(False, f"Conversion failed: {str(e)}", [])
