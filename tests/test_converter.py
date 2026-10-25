"""
Unit and integration tests for PDF Converter core modules.
"""

import os
import sys
import tempfile
import pytest

# Ensure PDF_TOOL is in sys.path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "PDF_TOOL"))

import pypdf
from app_paths import get_resource_path, get_temp_dir, create_session_dir, clean_dir
from pdfconvter import PDFConverter


@pytest.fixture
def sample_pdf_files(tmp_path):
    """Generate two minimal valid PDF files for testing."""
    pdf1_path = tmp_path / "test1.pdf"
    pdf2_path = tmp_path / "test2.pdf"

    for path, text in [(pdf1_path, "Page 1 Content"), (pdf2_path, "Page 2 Content")]:
        writer = pypdf.PdfWriter()
        writer.add_blank_page(width=200, height=200)
        writer.add_blank_page(width=200, height=200)
        with open(path, "wb") as f:
            writer.write(f)

    return [str(pdf1_path), str(pdf2_path)]


def test_app_paths():
    """Verify safe path generation and directory cleanup."""
    temp_dir = get_temp_dir("test_sub")
    assert os.path.exists(temp_dir)

    session_1 = create_session_dir()
    session_2 = create_session_dir()
    assert session_1 != session_2
    assert os.path.exists(session_1)

    # Test clean_dir
    dummy_file = os.path.join(temp_dir, "sample.txt")
    with open(dummy_file, "w") as f:
        f.write("hello")
    assert os.path.exists(dummy_file)
    clean_dir(temp_dir)
    assert not os.path.exists(dummy_file)


def test_pdf_merger(sample_pdf_files, tmp_path):
    """Test merging multiple PDFs."""
    out_dir = str(tmp_path / "merged_out")
    converter = PDFConverter(output_dir=out_dir)

    results = converter.pdfmerger(input_files=sample_pdf_files, output_dir=out_dir)
    assert len(results) == 1
    assert os.path.exists(results[0])

    reader = pypdf.PdfReader(results[0])
    # Each sample PDF had 2 pages -> 2 + 2 = 4 pages
    assert len(reader.pages) == 4


def test_pdf_compression(sample_pdf_files, tmp_path):
    """Test PDF compression."""
    out_dir = str(tmp_path / "compress_out")
    converter = PDFConverter(output_dir=out_dir)

    results = converter.pdfcompress(input_files=[sample_pdf_files[0]], output_dir=out_dir)
    assert len(results) == 1
    assert os.path.exists(results[0])


def test_pdf_split_range(sample_pdf_files, tmp_path):
    """Test extracting a range of pages from a PDF."""
    out_dir = str(tmp_path / "split_out")
    converter = PDFConverter(output_dir=out_dir)

    results = converter.split_range(start_page=1, end_page=1, input_files=[sample_pdf_files[0]], output_dir=out_dir)
    assert len(results) == 1
    assert os.path.exists(results[0])

    reader = pypdf.PdfReader(results[0])
    assert len(reader.pages) == 1


def test_pdf_extract_pages(sample_pdf_files, tmp_path):
    """Test extracting specific page numbers."""
    out_dir = str(tmp_path / "extract_out")
    converter = PDFConverter(output_dir=out_dir)

    results = converter.extract_pages(page_numbers=[2], input_files=[sample_pdf_files[0]], output_dir=out_dir)
    assert len(results) == 1
    assert os.path.exists(results[0])

    reader = pypdf.PdfReader(results[0])
    assert len(reader.pages) == 1


def test_pdf_single_pages(sample_pdf_files, tmp_path):
    """Test splitting into single pages."""
    out_dir = str(tmp_path / "single_out")
    converter = PDFConverter(output_dir=out_dir)

    results = converter.single_pages(input_files=[sample_pdf_files[0]], output_dir=out_dir)
    assert len(results) == 2
    for r in results:
        assert os.path.exists(r)
        reader = pypdf.PdfReader(r)
        assert len(reader.pages) == 1


def test_pdf2images_and_image2pdf(sample_pdf_files, tmp_path):
    """Test PDF to PNG rendering and PNG to PDF conversion."""
    out_dir = str(tmp_path / "img_out")
    converter = PDFConverter(output_dir=out_dir)

    # Convert PDF to images
    img_results = converter.pdf2images(input_files=[sample_pdf_files[0]], output_dir=out_dir)
    assert len(img_results) == 2
    for img in img_results:
        assert os.path.exists(img)

    # Convert images back to PDF
    pdf_out = str(tmp_path / "pdf_from_img")
    pdf_results = converter.image2pdf(input_files=img_results, output_dir=pdf_out)
    assert len(pdf_results) == 1
    assert os.path.exists(pdf_results[0])
