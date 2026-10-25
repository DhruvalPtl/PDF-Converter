"""
Optical Character Recognition (OCR) Engine for Scanned Documents.
Provides automatic detection of scanned image-only PDFs, C++ MuPDF text recognition,
and generation of searchable PDFs with invisible text layers.
"""

import os
import sys
import logging
import urllib.request
from typing import Optional, List, Tuple
import fitz

from app_paths import get_app_data_dir

logger = logging.getLogger(__name__)

TESSDATA_URL = "https://github.com/tesseract-ocr/tessdata_fast/raw/main/eng.traineddata"


def get_tessdata_dir() -> str:
    """Return the local tessdata directory inside user AppData."""
    tess_dir = os.path.join(get_app_data_dir(), "tessdata")
    os.makedirs(tess_dir, exist_ok=True)
    return tess_dir


def ensure_tessdata() -> Optional[str]:
    """Ensure eng.traineddata is available for OCR, downloading if necessary."""
    tess_dir = get_tessdata_dir()
    lang_file = os.path.join(tess_dir, "eng.traineddata")

    if os.path.exists(lang_file) and os.path.getsize(lang_file) > 1000000:
        return tess_dir

    try:
        logger.info("Downloading OCR language model (eng.traineddata)...")
        urllib.request.urlretrieve(TESSDATA_URL, lang_file)
        if os.path.exists(lang_file):
            logger.info("OCR model ready.")
            return tess_dir
    except Exception as e:
        logger.warning(f"Could not download OCR language model: {e}")

    return None


def is_scanned_pdf(doc_or_path, sample_pages: int = 3) -> bool:
    """Detect if a PDF document contains mostly images without extractable text."""
    close_when_done = False
    if isinstance(doc_or_path, str):
        if not os.path.exists(doc_or_path):
            return False
        doc = fitz.open(doc_or_path)
        close_when_done = True
    else:
        doc = doc_or_path

    try:
        pages_to_check = min(len(doc), sample_pages)
        if pages_to_check == 0:
            return False

        total_text_len = 0
        total_images = 0

        for i in range(pages_to_check):
            page = doc[i]
            total_text_len += len(page.get_text().strip())
            total_images += len(page.get_images())

        # If pages have images but virtually zero text characters, it's scanned
        return total_images > 0 and total_text_len < 20
    finally:
        if close_when_done and doc:
            doc.close()


def ocr_pdf_to_searchable(input_path: str, output_path: str, progress_callback=None) -> str:
    """
    Perform OCR on a scanned PDF and save as a searchable PDF with selectable text.
    """
    tess_dir = ensure_tessdata()
    if not tess_dir:
        raise RuntimeError("Tesseract language data is not available.")

    doc = fitz.open(input_path)
    total_pages = len(doc)
    logger.info(f"Performing OCR on {input_path} ({total_pages} pages)...")

    # In PyMuPDF, to create a searchable PDF with OCR:
    # Use doc.convert_to_pdf() with OCR
    try:
        pdf_bytes = doc.convert_to_pdf()
        out_doc = fitz.open("pdf", pdf_bytes)

        for idx, page in enumerate(doc):
            if progress_callback:
                progress_callback(int((idx / total_pages) * 80), f"OCR Page {idx + 1} of {total_pages}...")
            # Extract OCR textpage
            tp = page.get_textpage_ocr(tessdata=tess_dir)
            # Insert invisible text onto corresponding output page
            out_page = out_doc[idx]
            text = tp.extractText()
            if text:
                out_page.insert_text((50, 50), text, render_mode=3)  # render_mode 3 = invisible text

        out_doc.save(output_path, garbage=4, deflate=True)
        out_doc.close()
        doc.close()

        if progress_callback:
            progress_callback(100, "OCR complete!")
        return output_path
    except Exception as e:
        logger.error(f"OCR error: {e}", exc_info=True)
        doc.close()
        raise
