"""
Command Line Interface (CLI) for PDF Converter Pro.
Allows headless, automated execution of PDF conversions, merging, splitting,
watermarking, and encryption directly from the terminal.
"""

import argparse
import sys
import os
import logging
from typing import List, Optional

from pdfconvter import PDFConverter

# Configure CLI logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger("pdf_cli")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pdfconv",
        description="PDF Converter Pro — Command Line Interface"
    )
    subparsers = parser.add_subparsers(dest="command", required=True, help="Available sub-commands")

    # 1. Merge
    p_merge = subparsers.add_parser("merge", help="Merge multiple PDF files into one")
    p_merge.add_argument("files", nargs="+", help="Input PDF files in merge order")
    p_merge.add_argument("-o", "--output", required=True, help="Output merged PDF file or directory")

    # 2. Compress
    p_compress = subparsers.add_parser("compress", help="Compress PDF content streams")
    p_compress.add_argument("files", nargs="+", help="Input PDF files")
    p_compress.add_argument("-o", "--output", required=True, help="Output directory or file")

    # 3. Split
    p_split = subparsers.add_parser("split", help="Split PDF into pages or extract page ranges")
    p_split.add_argument("file", help="Input PDF file to split")
    p_split.add_argument("-o", "--output", required=True, help="Output directory")
    split_group = p_split.add_mutually_exclusive_group(required=True)
    split_group.add_argument("--single", action="store_true", help="Split into single-page documents")
    split_group.add_argument("--range", nargs=2, type=int, metavar=("START", "END"), help="Extract page range (1-indexed)")
    split_group.add_argument("--pages", nargs="+", type=int, help="Extract specific page numbers (e.g. 1 3 5)")

    # 4. Rotate
    p_rotate = subparsers.add_parser("rotate", help="Rotate PDF pages clockwise")
    p_rotate.add_argument("files", nargs="+", help="Input PDF files")
    p_rotate.add_argument("--angle", type=int, choices=[90, 180, 270], default=90, help="Rotation angle in degrees")
    p_rotate.add_argument("-o", "--output", required=True, help="Output directory or file")

    # 5. Protect
    p_protect = subparsers.add_parser("protect", help="Encrypt PDF with a password")
    p_protect.add_argument("files", nargs="+", help="Input PDF files")
    p_protect.add_argument("-p", "--password", required=True, help="User password for encryption")
    p_protect.add_argument("-o", "--output", required=True, help="Output directory or file")

    # 6. Unlock
    p_unlock = subparsers.add_parser("unlock", help="Decrypt password-protected PDF")
    p_unlock.add_argument("files", nargs="+", help="Input PDF files")
    p_unlock.add_argument("-p", "--password", required=True, help="Current password to decrypt")
    p_unlock.add_argument("-o", "--output", required=True, help="Output directory or file")

    # 7. Watermark
    p_wm = subparsers.add_parser("watermark", help="Stamp vector text watermark on PDF")
    p_wm.add_argument("files", nargs="+", help="Input PDF files")
    p_wm.add_argument("-t", "--text", default="CONFIDENTIAL", help="Watermark text")
    p_wm.add_argument("--opacity", type=float, default=0.3, help="Watermark opacity (0.1 to 1.0)")
    p_wm.add_argument("--angle", type=int, default=45, help="Watermark angle (e.g. 45, 0, 90)")
    p_wm.add_argument("-o", "--output", required=True, help="Output directory or file")

    # 8. PDF to Images
    p_p2i = subparsers.add_parser("pdf2images", help="Render PDF pages as PNG images")
    p_p2i.add_argument("files", nargs="+", help="Input PDF files")
    p_p2i.add_argument("-o", "--output", required=True, help="Output directory")

    # 9. Images to PDF
    p_i2p = subparsers.add_parser("image2pdf", help="Compile images into a single PDF")
    p_i2p.add_argument("files", nargs="+", help="Input JPG/PNG image files")
    p_i2p.add_argument("-o", "--output", required=True, help="Output PDF file")

    # 10. PDF to Word
    p_p2w = subparsers.add_parser("pdf2word", help="Convert PDF to Word (.docx)")
    p_p2w.add_argument("files", nargs="+", help="Input PDF files")
    p_p2w.add_argument("-o", "--output", required=True, help="Output directory")

    # 11. PDF to Excel
    p_p2e = subparsers.add_parser("pdf2excel", help="Extract PDF tables into Excel (.xlsx)")
    p_p2e.add_argument("files", nargs="+", help="Input PDF files")
    p_p2e.add_argument("-o", "--output", required=True, help="Output directory")

    return parser


def main(args: Optional[List[str]] = None) -> int:
    parser = build_parser()
    parsed = parser.parse_args(args)
    converter = PDFConverter()

    out = os.path.abspath(parsed.output)
    out_dir = out if os.path.isdir(out) or not out.endswith(".pdf") and not out.endswith(".docx") and not out.endswith(".xlsx") else os.path.dirname(out)
    os.makedirs(out_dir, exist_ok=True)

    try:
        results = []
        if parsed.command == "merge":
            results = converter.pdfmerger(input_files=parsed.files, output_dir=out_dir)
            if out.lower().endswith(".pdf") and results and results[0] != out:
                import shutil
                shutil.move(results[0], out)
                results = [out]

        elif parsed.command == "compress":
            results = converter.pdfcompress(input_files=parsed.files, output_dir=out_dir)

        elif parsed.command == "split":
            if parsed.single:
                results = converter.single_pages(input_files=[parsed.file], output_dir=out_dir)
            elif parsed.range:
                results = converter.split_range(parsed.range[0], parsed.range[1], input_files=[parsed.file], output_dir=out_dir)
            elif parsed.pages:
                results = converter.extract_pages(parsed.pages, input_files=[parsed.file], output_dir=out_dir)

        elif parsed.command == "rotate":
            results = converter.pdfrotate(input_files=parsed.files, rotation_angle=parsed.angle, output_dir=out_dir)

        elif parsed.command == "protect":
            results = converter.pdfprotect(input_files=parsed.files, user_password=parsed.password, output_dir=out_dir)

        elif parsed.command == "unlock":
            results = converter.pdfunlock(input_files=parsed.files, password=parsed.password, output_dir=out_dir)

        elif parsed.command == "watermark":
            results = converter.pdfwatermark(
                input_files=parsed.files, watermark_text=parsed.text,
                opacity=parsed.opacity, angle=parsed.angle, output_dir=out_dir
            )

        elif parsed.command == "pdf2images":
            results = converter.pdf2images(input_files=parsed.files, output_dir=out_dir)

        elif parsed.command == "image2pdf":
            results = converter.image2pdf(input_files=parsed.files, output_dir=out_dir)

        elif parsed.command == "pdf2word":
            results = converter.pdf2word(input_files=parsed.files, output_dir=out_dir)

        elif parsed.command == "pdf2excel":
            results = converter.pdf2excel(input_files=parsed.files, output_dir=out_dir)

        # If a specific target file path was supplied and single result produced, move to requested destination
        if (out.lower().endswith((".pdf", ".docx", ".xlsx", ".pptx"))) and len(results) == 1 and results[0] != out:
            import shutil
            shutil.move(results[0], out)
            results = [out]

        if results:
            print(f"SUCCESS: Generated {len(results)} file(s):")
            for r in results:
                print(f"  -> {r}")
            return 0
        else:
            print("WARNING: No output files were produced.")
            return 1

    except Exception as e:
        logger.error(f"Error executing command '{parsed.command}': {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
