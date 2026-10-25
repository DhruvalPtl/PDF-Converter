"""
Unit tests for SOTA capabilities:
- Modern HTML to PDF
- PyMuPDF C++ table extraction to styled openpyxl
- PPTX with editable textframes
"""

import os
import sys
import tempfile
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "PDF_TOOL"))

import fitz
import openpyxl
from pptx import Presentation
from pdfconvter import PDFConverter, find_system_browser


def test_sota_html2pdf(tmp_path):
    """Test modern HTML to PDF conversion via system browser engine."""
    browser = find_system_browser()
    if not browser:
        pytest.skip("No system Chromium browser available for test.")

    html_path = tmp_path / "page.html"
    with open(html_path, "w", encoding="utf-8") as f:
        f.write("""<!DOCTYPE html>
        <html>
        <head><style>body { font-family: sans-serif; color: #1565C0; }</style></head>
        <body>
          <h1>Modern SOTA PDF Generation</h1>
          <p>Testing flexbox, modern typography and clean rendering.</p>
        </body>
        </html>""")

    out_dir = str(tmp_path / "html_out")
    converter = PDFConverter(output_dir=out_dir)
    results = converter.html2pdf(input_files=[str(html_path)], output_dir=out_dir)

    assert len(results) == 1
    assert os.path.exists(results[0])
    assert os.path.getsize(results[0]) > 0

    with fitz.open(results[0]) as doc:
        assert len(doc) == 1
        text = doc[0].get_text()
        assert "Modern SOTA PDF Generation" in text


def test_sota_pdf2excel(tmp_path):
    """Test extracting tables with PyMuPDF into styled openpyxl workbook."""
    pdf_path = tmp_path / "table_doc.pdf"
    doc = fitz.open()
    page = doc.new_page()

    # Draw a table with gridlines and text
    page.draw_rect(fitz.Rect(50, 50, 300, 150))
    page.draw_line(fitz.Point(50, 90), fitz.Point(300, 90))
    page.draw_line(fitz.Point(175, 50), fitz.Point(175, 150))
    page.insert_text(fitz.Point(60, 75), "Header A")
    page.insert_text(fitz.Point(185, 75), "Header B")
    page.insert_text(fitz.Point(60, 120), "Val 1")
    page.insert_text(fitz.Point(185, 120), "Val 2")
    doc.save(str(pdf_path))
    doc.close()

    out_dir = str(tmp_path / "excel_out")
    converter = PDFConverter(output_dir=out_dir)
    results = converter.pdf2excel(input_files=[str(pdf_path)], output_dir=out_dir)

    assert len(results) == 1
    assert os.path.exists(results[0])

    wb = openpyxl.load_workbook(results[0])
    assert "Extracted Tables" in wb.sheetnames
    ws = wb["Extracted Tables"]
    # Check that workbook has cells
    assert ws.max_row >= 1


def test_sota_pdf_to_pptx(tmp_path):
    """Test PDF to editable PPTX conversion."""
    pdf_path = tmp_path / "slides.pdf"
    doc = fitz.open()
    page = doc.new_page(width=720, height=405)
    page.insert_text(fitz.Point(100, 100), "Quarterly Report", fontsize=24)
    page.insert_text(fitz.Point(100, 180), "Financial summary and performance KPIs.", fontsize=14)
    doc.save(str(pdf_path))
    doc.close()

    out_dir = str(tmp_path / "pptx_out")
    converter = PDFConverter(output_dir=out_dir)
    results = converter.pdf_to_pptx(input_files=[str(pdf_path)], output_dir=out_dir)

    assert len(results) == 1
    assert os.path.exists(results[0])

    prs = Presentation(results[0])
    assert len(prs.slides) == 1
    slide = prs.slides[0]

    # Verify slide contains text frames with editable text
    extracted_texts = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            extracted_texts.append(shape.text_frame.text)

    all_text = " ".join(extracted_texts)
    assert "Quarterly Report" in all_text
