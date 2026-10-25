"""
Unit tests for Advanced PDF operations: Protection, Unlocking, Rotation, Watermarking, CLI, and History.
"""

import os
import sys
import tempfile
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "PDF_TOOL"))

import pypdf
from pdfconvter import PDFConverter
from history import add_history_entry, get_history_entries, clear_history, get_setting, set_setting
from themes import apply_theme
import cli


@pytest.fixture
def blank_pdf(tmp_path):
    """Create a minimal 2-page PDF file."""
    path = tmp_path / "doc.pdf"
    writer = pypdf.PdfWriter()
    writer.add_blank_page(width=200, height=200)
    writer.add_blank_page(width=200, height=200)
    with open(path, "wb") as f:
        writer.write(f)
    return str(path)


def test_pdf_protect_and_unlock(blank_pdf, tmp_path):
    """Test encrypting a PDF with a password and unlocking it."""
    out_dir = str(tmp_path / "sec_out")
    converter = PDFConverter(output_dir=out_dir)

    # 1. Protect
    prot_results = converter.pdfprotect(input_files=[blank_pdf], user_password="mysecretpassword", output_dir=out_dir)
    assert len(prot_results) == 1
    prot_path = prot_results[0]
    assert os.path.exists(prot_path)

    # Verify encrypted
    reader = pypdf.PdfReader(prot_path)
    assert reader.is_encrypted

    # 2. Unlock
    unlock_dir = str(tmp_path / "unlock_out")
    unlock_results = converter.pdfunlock(input_files=[prot_path], password="mysecretpassword", output_dir=unlock_dir)
    assert len(unlock_results) == 1
    assert os.path.exists(unlock_results[0])

    unlocked_reader = pypdf.PdfReader(unlock_results[0])
    assert not unlocked_reader.is_encrypted
    assert len(unlocked_reader.pages) == 2


def test_pdf_rotate(blank_pdf, tmp_path):
    """Test rotating PDF pages by 90 degrees."""
    out_dir = str(tmp_path / "rot_out")
    converter = PDFConverter(output_dir=out_dir)

    results = converter.pdfrotate(input_files=[blank_pdf], rotation_angle=90, output_dir=out_dir)
    assert len(results) == 1
    assert os.path.exists(results[0])

    reader = pypdf.PdfReader(results[0])
    assert reader.pages[0].rotation == 90


def test_pdf_watermark(blank_pdf, tmp_path):
    """Test stamping vector text watermark on PDF."""
    out_dir = str(tmp_path / "wm_out")
    converter = PDFConverter(output_dir=out_dir)

    results = converter.pdfwatermark(
        input_files=[blank_pdf], watermark_text="CONFIDENTIAL",
        opacity=0.4, angle=45, output_dir=out_dir
    )
    assert len(results) == 1
    assert os.path.exists(results[0])

    reader = pypdf.PdfReader(results[0])
    assert len(reader.pages) == 2


def test_history_and_settings(tmp_path):
    """Test history recording, retrieval, clearing, and user settings."""
    clear_history()
    add_history_entry("Merge", ["f1.pdf", "f2.pdf"], ["merged.pdf"])

    entries = get_history_entries(limit=10)
    assert len(entries) >= 1
    assert entries[0]["tool"] == "Merge"
    assert "merged.pdf" in entries[0]["outputs"][0]

    clear_history()
    assert len(get_history_entries(limit=10)) == 0

    # Settings test
    set_setting("theme", "dark")
    assert get_setting("theme") == "dark"
    set_setting("theme", "light")
    assert get_setting("theme") == "light"


def test_themes():
    """Verify applying light and dark themes does not throw errors."""
    apply_theme("light")
    apply_theme("dark")


def test_cli_execution(blank_pdf, tmp_path):
    """Test CLI commands execution."""
    out_rotate = str(tmp_path / "cli_rotated.pdf")
    exit_code = cli.main(["rotate", blank_pdf, "--angle", "180", "-o", out_rotate])
    assert exit_code == 0
    assert os.path.exists(out_rotate)

    out_wm = str(tmp_path / "cli_wm.pdf")
    exit_code_wm = cli.main(["watermark", blank_pdf, "--text", "SAMPLE", "-o", out_wm])
    assert exit_code_wm == 0
    assert os.path.exists(out_wm)
