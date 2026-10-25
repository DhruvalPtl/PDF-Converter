"""
State-of-the-Art (SOTA) PDF and Document Processing Engine.
Combines PyMuPDF (C++ MuPDF), pikepdf (C++ QPDF), openpyxl, python-pptx,
headless Chromium/Edge rendering, and headless LibreOffice/COM automation.
"""

import os
import sys
import shutil
import logging
import subprocess
import time
from typing import List, Callable, Optional, Tuple

# SOTA High-performance PDF engines
import fitz  # PyMuPDF (C++ MuPDF engine)
import pikepdf  # C++ QPDF engine
import pypdf  # Pure-Python fallback

# Office & Spreadsheet engines
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from pdf2docx import Converter as DocxConverter
except ImportError:
    DocxConverter = None

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
except ImportError:
    Presentation = None
    Inches = None
    Pt = None

try:
    from PIL import Image
    import io
except ImportError:
    Image = None

try:
    import img2pdf
except ImportError:
    img2pdf = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import pdfkit
except ImportError:
    pdfkit = None

# Optional Windows COM automation
HAS_WIN32COM = False
try:
    import win32com.client
    import pywintypes
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    win32com = None
    pywintypes = None
    pythoncom = None

from app_paths import get_temp_dir, clean_dir

logger = logging.getLogger(__name__)


def find_system_browser() -> Optional[str]:
    """Find a modern Chromium-based browser for headless HTML-to-PDF rendering."""
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        "msedge",
        "chrome",
        "google-chrome",
        "chromium-browser",
        "chromium"
    ]
    for candidate in candidates:
        if os.path.isabs(candidate) and os.path.exists(candidate):
            return candidate
        found = shutil.which(candidate)
        if found:
            return found
    return None


def find_libreoffice() -> Optional[str]:
    """Find headless LibreOffice binary for cross-platform Office conversion."""
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "soffice",
        "libreoffice"
    ]
    for candidate in candidates:
        if os.path.isabs(candidate) and os.path.exists(candidate):
            return candidate
        found = shutil.which(candidate)
        if found:
            return found
    return None


class PDFConverter:
    """
    SOTA document processing engine combining MuPDF C++, QPDF, openpyxl,
    and headless Chromium print engines.
    """
    def __init__(self, input_dir: Optional[str] = None, output_dir: Optional[str] = None):
        self.input_path_dir = input_dir or get_temp_dir("Temp")
        self.output_path = output_dir or get_temp_dir("Temp1")
        self.browser_path = find_system_browser()
        self.libreoffice_path = find_libreoffice()

        os.makedirs(self.input_path_dir, exist_ok=True)
        os.makedirs(self.output_path, exist_ok=True)

    def _get_input_files(self, extension: str, override_files: Optional[List[str]] = None) -> List[str]:
        if override_files:
            return [os.path.abspath(f) for f in override_files if os.path.exists(f)]
        if not os.path.exists(self.input_path_dir):
            return []
        ext = extension.lower() if extension.startswith(".") else f".{extension.lower()}"
        return [
            os.path.join(self.input_path_dir, f)
            for f in os.listdir(self.input_path_dir)
            if f.lower().endswith(ext)
        ]

    def _check_cancellation(self, is_cancelled: Optional[Callable[[], bool]]) -> None:
        if is_cancelled and is_cancelled():
            raise InterruptedError("Operation was cancelled by user.")

    def emptydir(self) -> None:
        clean_dir(self.input_path_dir)

    # ----------------------------------------------------------------------
    # 1. SOTA PDF MERGER (PyMuPDF C Engine)
    # ----------------------------------------------------------------------
    def pdfmerger(self, input_files: Optional[List[str]] = None, output_dir: Optional[str] = None,
                  progress_callback: Optional[Callable[[int, str], None]] = None,
                  is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        """Merge PDFs with C++ MuPDF engine — ultra-fast, handles broken xrefs cleanly."""
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)

        if not files:
            return []

        if len(files) == 1:
            dest = os.path.join(out_dir, os.path.basename(files[0]))
            shutil.copy2(files[0], dest)
            return [dest]

        output_filename = os.path.join(out_dir, "merged_document.pdf")
        merged_doc = fitz.open()

        try:
            for idx, pdf_path in enumerate(files):
                self._check_cancellation(is_cancelled)
                if progress_callback:
                    progress_callback(int((idx / len(files)) * 80), f"Merging {os.path.basename(pdf_path)}...")

                with fitz.open(pdf_path) as src:
                    merged_doc.insert_pdf(src)

            self._check_cancellation(is_cancelled)
            # Save with garbage collection and deflation
            merged_doc.save(output_filename, garbage=4, deflate=True, clean=True)
            merged_doc.close()

            if progress_callback:
                progress_callback(100, "Merge complete!")
            return [output_filename]
        except Exception as e:
            logger.error(f"SOTA Merge error: {e}", exc_info=True)
            # Fallback to pypdf if PyMuPDF encounters unexpected issue
            if hasattr(pypdf, 'PdfMerger'):
                m = pypdf.PdfMerger()
            else:
                m = pypdf.PdfWriter()
            for f in files:
                m.append(f)
            m.write(output_filename)
            if hasattr(m, 'close'):
                m.close()
            return [output_filename]
        finally:
            if merged_doc is not None and not getattr(merged_doc, 'is_closed', True):
                try:
                    merged_doc.close()
                except Exception:
                    pass

    # ----------------------------------------------------------------------
    # 2. SOTA PDF COMPRESSION (Image Downsampling + pikepdf Linearization)
    # ----------------------------------------------------------------------
    def pdfcompress(self, input_files: Optional[List[str]] = None, output_dir: Optional[str] = None,
                    progress_callback: Optional[Callable[[int, str], None]] = None,
                    is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        """
        High-ratio compression:
        - Downsamples oversized embedded images to 150 DPI JPEG.
        - Strips unused font subsets, duplicate streams, and orphan objects (garbage=4).
        - Linearizes the PDF and generates Object Streams for maximum compression.
        """
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)
        results = []

        for idx, pdf_path in enumerate(files):
            self._check_cancellation(is_cancelled)
            file_name = os.path.basename(pdf_path)
            output_pdf_path = os.path.join(out_dir, f"compressed_{file_name}")

            if progress_callback:
                progress_callback(int((idx / len(files)) * 80), f"Optimizing {file_name}...")

            doc = None
            try:
                doc = fitz.open(pdf_path)

                # 1. Downsample embedded images if Pillow is available
                if Image:
                    for page_num in range(len(doc)):
                        page = doc.load_page(page_num)
                        image_list = page.get_images()
                        for img_info in image_list:
                            xref = img_info[0]
                            base_img = doc.extract_image(xref)
                            if not base_img:
                                continue
                            img_bytes = base_img["image"]
                            try:
                                pil_img = Image.open(io.BytesIO(img_bytes))
                                w, h = pil_img.size
                                # Downsample if excessively large
                                max_dimension = 1400
                                if w > max_dimension or h > max_dimension:
                                    scale = max_dimension / max(w, h)
                                    new_size = (int(w * scale), int(h * scale))
                                    pil_img = pil_img.resize(new_size, Image.Resampling.LANCZOS)
                                    out_buf = io.BytesIO()
                                    pil_img.convert("RGB").save(out_buf, format="JPEG", quality=75, optimize=True)
                                    # Replace image stream in PDF
                                    page.replace_image(xref, stream=out_buf.getvalue())
                            except Exception:
                                pass

                # 2. Save with PyMuPDF full garbage collection & stream deflation
                temp_saved = output_pdf_path + ".tmp.pdf"
                doc.save(
                    temp_saved,
                    garbage=4,
                    clean=True,
                    deflate=True,
                    deflate_images=True,
                    deflate_fonts=True,
                    use_objstms=1
                )
                doc.close()
                doc = None

                # 3. Apply pikepdf QPDF object stream compaction & linearization
                try:
                    with pikepdf.open(temp_saved) as p:
                        p.save(
                            output_pdf_path,
                            compress_streams=True,
                            linearize=True,
                            object_stream_mode=pikepdf.ObjectStreamMode.generate
                        )
                    if os.path.exists(temp_saved):
                        os.remove(temp_saved)
                except Exception:
                    # If pikepdf fails, use PyMuPDF optimized output
                    if os.path.exists(temp_saved):
                        shutil.move(temp_saved, output_pdf_path)

                results.append(output_pdf_path)
                logger.info(f"SOTA Compression finished: {file_name} -> {output_pdf_path}")
            except Exception as e:
                logger.error(f"Compression error on {file_name}: {e}")
            finally:
                if doc:
                    try:
                        doc.close()
                    except Exception:
                        pass

        if progress_callback:
            progress_callback(100, "Compression complete!")
        return results

    # ----------------------------------------------------------------------
    # 3. SOTA SPLITTING (PyMuPDF)
    # ----------------------------------------------------------------------
    def split_range(self, start_page: int, end_page: int, input_files: Optional[List[str]] = None,
                    output_dir: Optional[str] = None,
                    is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)
        results = []

        for pdf_path in files:
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            try:
                with fitz.open(pdf_path) as src:
                    total_pages = len(src)
                    start_idx = max(0, start_page - 1)
                    end_idx = min(total_pages, end_page)

                    if start_idx >= total_pages or start_idx >= end_idx:
                        continue

                    new_doc = fitz.open()
                    new_doc.insert_pdf(src, from_page=start_idx, to_page=end_idx - 1)
                    out_file = os.path.join(out_dir, f"{base_name}_range_{start_page}_{end_page}.pdf")
                    new_doc.save(out_file, garbage=4, deflate=True)
                    new_doc.close()
                    results.append(out_file)
            except Exception as e:
                logger.error(f"SOTA split range error: {e}")

        return results

    def extract_pages(self, page_numbers: List[int], input_files: Optional[List[str]] = None,
                      output_dir: Optional[str] = None,
                      is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)
        results = []

        for pdf_path in files:
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            try:
                with fitz.open(pdf_path) as src:
                    total_pages = len(src)
                    valid_pages = [p - 1 for p in page_numbers if 0 <= p - 1 < total_pages]
                    if not valid_pages:
                        continue

                    new_doc = fitz.open()
                    for p_idx in valid_pages:
                        new_doc.insert_pdf(src, from_page=p_idx, to_page=p_idx)

                    out_file = os.path.join(out_dir, f"{base_name}_extracted.pdf")
                    new_doc.save(out_file, garbage=4, deflate=True)
                    new_doc.close()
                    results.append(out_file)
            except Exception as e:
                logger.error(f"SOTA extract pages error: {e}")

        return results

    def single_pages(self, input_files: Optional[List[str]] = None,
                     output_dir: Optional[str] = None,
                     is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)
        results = []

        for pdf_path in files:
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            try:
                with fitz.open(pdf_path) as src:
                    for idx in range(len(src)):
                        self._check_cancellation(is_cancelled)
                        single_doc = fitz.open()
                        single_doc.insert_pdf(src, from_page=idx, to_page=idx)
                        out_file = os.path.join(out_dir, f"{base_name}_page_{idx + 1}.pdf")
                        single_doc.save(out_file, garbage=4, deflate=True)
                        single_doc.close()
                        results.append(out_file)
            except Exception as e:
                logger.error(f"SOTA single pages error: {e}")

        return results

    # ----------------------------------------------------------------------
    # 4. SOTA ROTATION, PROTECTION & WATERMARKING
    # ----------------------------------------------------------------------
    def pdfrotate(self, input_files: Optional[List[str]] = None, rotation_angle: int = 90,
                  page_numbers: Optional[List[int]] = None, output_dir: Optional[str] = None,
                  progress_callback: Optional[Callable[[int, str], None]] = None,
                  is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)
        results = []

        for idx, pdf_path in enumerate(files):
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_file = os.path.join(out_dir, f"{base_name}_rotated_{rotation_angle}.pdf")

            if progress_callback:
                progress_callback(int((idx / len(files)) * 80), f"Rotating {os.path.basename(pdf_path)}...")

            try:
                with fitz.open(pdf_path) as doc:
                    for p_idx, page in enumerate(doc):
                        if page_numbers is None or (p_idx + 1) in page_numbers:
                            page.set_rotation((page.rotation + rotation_angle) % 360)
                    doc.save(out_file, garbage=4, deflate=True)
                results.append(out_file)
            except Exception as e:
                logger.error(f"Rotation error on {pdf_path}: {e}")

        if progress_callback:
            progress_callback(100, "Rotation complete!")
        return results

    def pdfprotect(self, input_files: Optional[List[str]] = None, user_password: str = "",
                   owner_password: Optional[str] = None, output_dir: Optional[str] = None,
                   progress_callback: Optional[Callable[[int, str], None]] = None,
                   is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        if not user_password:
            raise ValueError("Password cannot be empty.")

        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)
        results = []

        for idx, pdf_path in enumerate(files):
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_file = os.path.join(out_dir, f"{base_name}_protected.pdf")

            if progress_callback:
                progress_callback(int((idx / len(files)) * 80), f"Encrypting {os.path.basename(pdf_path)}...")

            try:
                with pikepdf.open(pdf_path) as p:
                    perm = pikepdf.Permissions(accessibility=True, extract=False, print_highres=False, modify_other=False)
                    enc = pikepdf.Encryption(
                        user=user_password,
                        owner=owner_password or user_password,
                        allow=perm,
                        R=6  # AES-256 (Highest security standard)
                    )
                    p.save(out_file, encryption=enc)
                results.append(out_file)
            except Exception as e:
                logger.error(f"Encryption error: {e}")

        if progress_callback:
            progress_callback(100, "Protection applied!")
        return results

    def pdfunlock(self, input_files: Optional[List[str]] = None, password: str = "",
                  output_dir: Optional[str] = None,
                  progress_callback: Optional[Callable[[int, str], None]] = None,
                  is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)
        results = []

        for idx, pdf_path in enumerate(files):
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_file = os.path.join(out_dir, f"{base_name}_unlocked.pdf")

            if progress_callback:
                progress_callback(int((idx / len(files)) * 80), f"Decrypting {os.path.basename(pdf_path)}...")

            try:
                with pikepdf.open(pdf_path, password=password) as p:
                    p.save(out_file)
                results.append(out_file)
            except Exception as e:
                logger.error(f"Unlock error on {pdf_path}: {e}")
                raise

        if progress_callback:
            progress_callback(100, "Unlocking complete!")
        return results

    def pdfwatermark(self, input_files: Optional[List[str]] = None, watermark_text: str = "CONFIDENTIAL",
                     opacity: float = 0.3, font_size: float = 38.0, angle: int = 45,
                     color: Tuple[float, float, float] = (0.5, 0.5, 0.5), output_dir: Optional[str] = None,
                     progress_callback: Optional[Callable[[int, str], None]] = None,
                     is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)
        results = []

        for idx, pdf_path in enumerate(files):
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_file = os.path.join(out_dir, f"{base_name}_watermarked.pdf")

            if progress_callback:
                progress_callback(int((idx / len(files)) * 80), f"Watermarking {os.path.basename(pdf_path)}...")

            try:
                with fitz.open(pdf_path) as doc:
                    for page in doc:
                        rect = page.rect
                        center_pt = fitz.Point(rect.width * 0.25, rect.height * 0.55)
                        mat = fitz.Matrix(angle)
                        morph = (center_pt, mat)
                        page.insert_text(
                            center_pt,
                            watermark_text,
                            fontsize=font_size,
                            fontname="helv",
                            color=color,
                            fill_opacity=opacity,
                            morph=morph,
                            overlay=True
                        )
                    doc.save(out_file, garbage=4, deflate=True)
                results.append(out_file)
            except Exception as e:
                logger.error(f"Watermarking error: {e}")

        if progress_callback:
            progress_callback(100, "Watermarking complete!")
        return results

    # ----------------------------------------------------------------------
    # 5. SOTA HTML TO PDF (Headless Chromium Engine)
    # ----------------------------------------------------------------------
    def html2pdf(self, input_files: Optional[List[str]] = None, output_dir: Optional[str] = None,
                 is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        """Convert HTML to PDF using modern Headless Chromium (Edge/Chrome) with full CSS3 Flex/Grid support."""
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".html", input_files)
        results = []

        for html_path in files:
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(html_path))[0]
            out_file = os.path.join(out_dir, f"{base_name}.pdf")

            # Method 1: Headless Chromium (Edge / Chrome)
            if self.browser_path:
                url = "file:///" + os.path.abspath(html_path).replace("\\", "/")
                cmd = [
                    self.browser_path,
                    "--headless",
                    "--disable-gpu",
                    "--no-pdf-header-footer",
                    f"--print-to-pdf={out_file}",
                    url
                ]
                try:
                    subprocess.run(cmd, timeout=30, check=True)
                    # Small grace wait for file flush
                    for _ in range(15):
                        if os.path.exists(out_file) and os.path.getsize(out_file) > 0:
                            break
                        time.sleep(0.2)
                    if os.path.exists(out_file):
                        results.append(out_file)
                        logger.info(f"Chromium HTML -> PDF success: {out_file}")
                        continue
                except Exception as e:
                    logger.warning(f"Chromium PDF print failed: {e}")

            # Method 2: Fallback to pdfkit if available on system
            if pdfkit:
                try:
                    pdfkit.from_file(html_path, out_file)
                    results.append(out_file)
                except Exception as e:
                    logger.error(f"pdfkit conversion failed: {e}")

        return results

    # ----------------------------------------------------------------------
    # 6. SOTA PDF TO EXCEL (PyMuPDF C++ Table Finder + Styled openpyxl)
    # ----------------------------------------------------------------------
    def pdf2excel(self, input_files: Optional[List[str]] = None, output_dir: Optional[str] = None,
                  is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        """
        Extract tables with PyMuPDF's C++ table finder and format into professional Excel workbooks
        with header styling, alternating row fills, borders, and auto-fitted columns.
        """
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)
        results = []

        # Professional styling definitions
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        regular_font = Font(name="Segoe UI", size=10)
        border_side = Side(style="thin", color="D9D9D9")
        thin_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

        for pdf_path in files:
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_file = os.path.join(out_dir, f"{base_name}.xlsx")

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Extracted Tables"
            current_row = 1
            tables_found = 0

            try:
                with fitz.open(pdf_path) as doc:
                    for page_idx, page in enumerate(doc):
                        self._check_cancellation(is_cancelled)
                        tab_finder = page.find_tables()
                        for table in tab_finder.tables:
                            extracted_rows = table.extract()
                            if not extracted_rows:
                                continue
                            tables_found += 1

                            for r_idx, row in enumerate(extracted_rows):
                                is_header = (r_idx == 0)
                                for c_idx, cell_value in enumerate(row):
                                    cell = ws.cell(row=current_row, column=c_idx + 1, value=cell_value or "")
                                    cell.border = thin_border
                                    if is_header:
                                        cell.fill = header_fill
                                        cell.font = header_font
                                        cell.alignment = Alignment(horizontal="center", vertical="center")
                                    else:
                                        cell.font = regular_font
                                        cell.alignment = Alignment(vertical="center")
                                current_row += 1
                            current_row += 2  # Spacing between tables

                # Auto-fit column widths
                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

                if tables_found > 0:
                    wb.save(out_file)
                    results.append(out_file)
                    logger.info(f"SOTA Table extraction succeeded: {out_file} ({tables_found} tables)")
                elif pdfplumber:
                    # Fallback to pdfplumber if fitz find_tables found nothing
                    with pdfplumber.open(pdf_path) as p_doc:
                        for page in p_doc.pages:
                            t_list = page.extract_tables()
                            for t in t_list:
                                for r in t:
                                    for c_idx, val in enumerate(r):
                                        ws.cell(row=current_row, column=c_idx + 1, value=val or "")
                                    current_row += 1
                                current_row += 1
                    if current_row > 1:
                        wb.save(out_file)
                        results.append(out_file)
            except Exception as e:
                logger.error(f"Table extraction error: {e}")

        return results

    # ----------------------------------------------------------------------
    # 7. SOTA PDF TO POWERPOINT (Editable Text Frames + Slide Layouts)
    # ----------------------------------------------------------------------
    def pdf_to_pptx(self, input_files: Optional[List[str]] = None, output_dir: Optional[str] = None,
                    is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        """Convert PDF into PowerPoint with editable text boxes and slide backgrounds."""
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)
        results = []

        if not Presentation:
            raise RuntimeError("python-pptx is not installed.")

        for pdf_path in files:
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_file = os.path.join(out_dir, f"{base_name}.pptx")
            prs = Presentation()
            temp_images = []

            try:
                with fitz.open(pdf_path) as doc:
                    for page_idx, page in enumerate(doc):
                        self._check_cancellation(is_cancelled)
                        # Render background layout image
                        pix = page.get_pixmap(dpi=150)
                        temp_bg = os.path.join(out_dir, f"_temp_bg_{base_name}_{page_idx}.png")
                        pix.save(temp_bg)
                        temp_images.append(temp_bg)

                        # Blank slide layout
                        blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
                        slide = prs.slides.add_slide(blank_layout)

                        # 1. Background image
                        slide.shapes.add_picture(temp_bg, Inches(0), Inches(0),
                                                width=prs.slide_width, height=prs.slide_height)

                        # 2. Extract editable text blocks and position textframes
                        pw, ph = page.rect.width, page.rect.height
                        scale_x = prs.slide_width / pw
                        scale_y = prs.slide_height / ph

                        for block in page.get_text("blocks"):
                            x0, y0, x1, y1, text, block_no, block_type = block
                            if block_type == 0 and text.strip():
                                tb_x = int(x0 * scale_x)
                                tb_y = int(y0 * scale_y)
                                tb_w = int((x1 - x0) * scale_x)
                                tb_h = int((y1 - y0) * scale_y)
                                tx_box = slide.shapes.add_textbox(tb_x, tb_y, max(tb_w, Inches(1)), max(tb_h, Inches(0.4)))
                                tf = tx_box.text_frame
                                tf.word_wrap = True
                                p = tf.paragraphs[0]
                                p.text = text.strip()
                                p.font.size = Pt(11)
                                p.font.name = "Segoe UI"

                prs.save(out_file)
                results.append(out_file)
            except Exception as e:
                logger.error(f"PDF to PPTX conversion error: {e}")
            finally:
                for img in temp_images:
                    if os.path.exists(img):
                        try:
                            os.remove(img)
                        except Exception:
                            pass

        return results

    # ----------------------------------------------------------------------
    # 8. SOTA OFFICE CONVERSIONS (LibreOffice CLI + COM Fallback)
    # ----------------------------------------------------------------------
    def word2pdf(self, input_files: Optional[List[str]] = None, output_dir: Optional[str] = None,
                 is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".docx", input_files)
        results = []

        for docx_path in files:
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(docx_path))[0]
            out_file = os.path.join(out_dir, f"{base_name}.pdf")

            # Method 1: Headless LibreOffice
            if self.libreoffice_path:
                try:
                    cmd = [self.libreoffice_path, "--headless", "--convert-to", "pdf", docx_path, "--outdir", out_dir]
                    subprocess.run(cmd, timeout=45, check=True)
                    if os.path.exists(out_file):
                        results.append(out_file)
                        continue
                except Exception as e:
                    logger.warning(f"LibreOffice conversion error: {e}")

            # Method 2: Windows COM Automation
            if sys.platform == 'win32' and HAS_WIN32COM:
                pythoncom.CoInitialize()
                word = None
                try:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = False
                    doc = word.Documents.Open(os.path.abspath(docx_path))
                    doc.SaveAs(os.path.abspath(out_file), FileFormat=17)
                    doc.Close(False)
                    results.append(out_file)
                except Exception as e:
                    logger.error(f"Word COM error: {e}")
                finally:
                    if word:
                        try:
                            word.Quit()
                        except Exception:
                            pass
                    pythoncom.CoUninitialize()

        return results

    def excel2pdf(self, input_files: Optional[List[str]] = None, output_dir: Optional[str] = None,
                  is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".xlsx", input_files)
        results = []

        for xlsx_path in files:
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(xlsx_path))[0]
            out_file = os.path.join(out_dir, f"{base_name}.pdf")

            if self.libreoffice_path:
                try:
                    cmd = [self.libreoffice_path, "--headless", "--convert-to", "pdf", xlsx_path, "--outdir", out_dir]
                    subprocess.run(cmd, timeout=45, check=True)
                    if os.path.exists(out_file):
                        results.append(out_file)
                        continue
                except Exception as e:
                    logger.warning(f"LibreOffice Excel conversion error: {e}")

            if sys.platform == 'win32' and HAS_WIN32COM:
                pythoncom.CoInitialize()
                excel = None
                try:
                    excel = win32com.client.Dispatch("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                    wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
                    wb.ExportAsFixedFormat(0, os.path.abspath(out_file))
                    wb.Close(False)
                    results.append(out_file)
                except Exception as e:
                    logger.error(f"Excel COM error: {e}")
                finally:
                    if excel:
                        try:
                            excel.Quit()
                        except Exception:
                            pass
                    pythoncom.CoUninitialize()

        return results

    def pptx2pdf(self, input_files: Optional[List[str]] = None, output_dir: Optional[str] = None,
                 is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pptx", input_files)
        results = []

        for pptx_path in files:
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(pptx_path))[0]
            out_file = os.path.join(out_dir, f"{base_name}.pdf")

            if self.libreoffice_path:
                try:
                    cmd = [self.libreoffice_path, "--headless", "--convert-to", "pdf", pptx_path, "--outdir", out_dir]
                    subprocess.run(cmd, timeout=45, check=True)
                    if os.path.exists(out_file):
                        results.append(out_file)
                        continue
                except Exception as e:
                    logger.warning(f"LibreOffice PPTX error: {e}")

            if sys.platform == 'win32' and HAS_WIN32COM:
                pythoncom.CoInitialize()
                powerpoint = None
                try:
                    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                    pres = powerpoint.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
                    pres.SaveAs(os.path.abspath(out_file), 32)
                    pres.Close()
                    results.append(out_file)
                except Exception as e:
                    logger.error(f"PowerPoint COM error: {e}")
                finally:
                    if powerpoint:
                        try:
                            powerpoint.Quit()
                        except Exception:
                            pass
                    pythoncom.CoUninitialize()

        return results

    # ----------------------------------------------------------------------
    # 10. OCR / SEARCHABLE PDF
    # ----------------------------------------------------------------------
    def pdfocr(self, input_files: Optional[List[str]] = None, output_dir: Optional[str] = None,
               progress_callback: Optional[Callable[[int, str], None]] = None,
               is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        """Perform OCR on scanned PDFs and output searchable documents."""
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)
        results = []

        from ocr_engine import ocr_pdf_to_searchable

        for idx, pdf_path in enumerate(files):
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_file = os.path.join(out_dir, f"{base_name}_searchable_ocr.pdf")
            try:
                res = ocr_pdf_to_searchable(pdf_path, out_file, progress_callback=progress_callback)
                results.append(res)
            except Exception as e:
                logger.error(f"OCR error on {pdf_path}: {e}")

        return results

    # ----------------------------------------------------------------------
    # 11. PDF TO WORD, IMAGES & RENDERING
    # ----------------------------------------------------------------------
    def pdf2word(self, input_files: Optional[List[str]] = None, output_dir: Optional[str] = None,
                 progress_callback: Optional[Callable[[int, str], None]] = None,
                 is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)
        results = []

        if not DocxConverter:
            raise RuntimeError("pdf2docx is not installed.")

        for idx, pdf_path in enumerate(files):
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_file = os.path.join(out_dir, f"{base_name}.docx")
            if progress_callback:
                progress_callback(int((idx / len(files)) * 80), f"Converting {os.path.basename(pdf_path)}...")

            cv = None
            try:
                cv = DocxConverter(pdf_path)
                cv.convert(out_file, start=0, end=None)
                results.append(out_file)
            except Exception as e:
                logger.error(f"PDF to Word error: {e}")
            finally:
                if cv:
                    try:
                        cv.close()
                    except Exception:
                        pass

        if progress_callback:
            progress_callback(100, "Conversion complete!")
        return results

    def image2pdf(self, input_files: Optional[List[str]] = None, output_dir: Optional[str] = None,
                  is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        imgs = self._get_input_files(".jpg", input_files) + self._get_input_files(".png", input_files)

        if not imgs:
            return []

        self._check_cancellation(is_cancelled)
        out_file = os.path.join(out_dir, f"{os.path.splitext(os.path.basename(imgs[0]))[0]}.pdf")
        try:
            if img2pdf:
                with open(out_file, "wb") as f_out:
                    f_out.write(img2pdf.convert(imgs))
            elif Image:
                pil_imgs = [Image.open(img).convert("RGB") for img in imgs]
                pil_imgs[0].save(out_file, save_all=True, append_images=pil_imgs[1:])
            return [out_file]
        except Exception as e:
            logger.error(f"Image to PDF error: {e}")
            return []

    def pdf2images(self, input_files: Optional[List[str]] = None, output_dir: Optional[str] = None,
                   is_cancelled: Optional[Callable[[], bool]] = None) -> List[str]:
        out_dir = output_dir or self.output_path
        os.makedirs(out_dir, exist_ok=True)
        files = self._get_input_files(".pdf", input_files)
        results = []

        for pdf_path in files:
            self._check_cancellation(is_cancelled)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            try:
                with fitz.open(pdf_path) as doc:
                    for page_num in range(len(doc)):
                        self._check_cancellation(is_cancelled)
                        page = doc.load_page(page_num)
                        pix = page.get_pixmap(dpi=200)
                        out_img = os.path.join(out_dir, f"{base_name}_page_{page_num + 1}.png")
                        pix.save(out_img)
                        results.append(out_img)
            except Exception as e:
                logger.error(f"PDF to images error: {e}")

        return results
