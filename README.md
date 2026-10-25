# PDF Converter

**PDF Converter** is an all-in-one desktop and command-line document management suite built with Python and PyQt6. It provides 17 essential tools to convert, merge, split, compress, watermark, protect, unlock, rotate, and OCR documents with a modern responsive UI, dark/light theme switching, and non-blocking background concurrency.

---

## ✨ Comprehensive Features

### 📄 17 Integrated Tools:
1. **Merge PDF**: Combine multiple PDFs in custom sequence with visual drag-to-reorder.
2. **Compress PDF**: Optimize and deflate PDF content streams with smart downsampling.
3. **Split PDF**: Split into single pages, extract page ranges, or cherry-pick specific pages.
4. **Rotate PDF**: Rotate all or specific pages clockwise by 90°, 180°, or 270°.
5. **Protect PDF**: Encrypt documents with military-grade AES-256 passwords.
6. **Unlock PDF**: Decrypt and strip password restrictions from protected PDFs.
7. **Watermark PDF**: Stamp custom diagonal semi-transparent vector text across all pages.
8. **OCR Scanned PDF**: Make scanned image-only PDFs searchable and selectable.
9. **HTML to PDF**: Convert web documents into clean PDF files via Chromium engine.
10. **PDF to Word**: Convert PDF layouts into editable `.docx` Word documents.
11. **Word to PDF**: Convert Word `.docx` documents into standardized PDFs.
12. **PDF to Excel**: Extract tabular data from PDFs into styled `.xlsx` workbooks.
13. **Excel to PDF**: Render spreadsheets into standardized PDF documents.
14. **PDF to PowerPoint**: Generate presentation slide decks (`.pptx`) from PDF pages.
15. **PowerPoint to PDF**: Convert slide presentations into PDFs.
16. **PDF to JPG**: Render high-resolution PNG/JPG images of each page.
17. **JPG to PDF**: Compile multiple photos/images into a single PDF document.

### 🎨 GUI & UX Innovations:
- **Dark Mode & Light Mode**: Toggle anytime between a sleek Dark Theme and a crisp Light Theme.
- **Drag-to-Reorder**: Reorder files directly by dragging them in the staging list (crucial for PDF merging).
- **Cancellable Background Concurrency**: Heavy conversions run in background `QThread` workers with live progress reporting and a one-click "Cancel" button.
- **Recent Conversion History**: History dialog with double-click "Open File" and "Show in Folder" shortcuts.
- **Dynamic Vector Assets**: High-DPI vector icons generated programmatically—zero missing image asset bugs.
- **Secure File Storage**: User-isolated temporary and application data directories in `%LOCALAPPDATA%`, eliminating multi-user collisions and permission failures.

### 💻 Dual Interface: GUI & CLI
Run the full graphical dashboard or automate workflows via the built-in terminal CLI:
```bash
# Merge PDFs
python PDF_TOOL/cli.py merge doc1.pdf doc2.pdf -o merged.pdf

# Watermark
python PDF_TOOL/cli.py watermark input.pdf --text "CONFIDENTIAL" --opacity 0.3 -o watermarked.pdf

# Password Protect
python PDF_TOOL/cli.py protect input.pdf --password "secret123" -o protected.pdf

# Unlock
python PDF_TOOL/cli.py unlock protected.pdf --password "secret123" -o unlocked.pdf

# Rotate
python PDF_TOOL/cli.py rotate input.pdf --angle 90 -o rotated.pdf

# Split
python PDF_TOOL/cli.py split input.pdf --range 1 5 -o output_dir/
```

---

## 🚀 Getting Started

### 1. Installation
```bash
git clone https://github.com/DhruvalPtl/PDF-Converter.git
cd PDF-Converter

# Set up virtual environment
uv venv .venv
uv pip install -r PDF_TOOL/requirements.txt
```

### 2. Launch the Application
```bash
python PDF_TOOL/main.py
```

### 3. Run the Test Suite
```bash
pytest tests/ -v
```

---

## 🏗️ Continuous Integration
Automated testing via GitHub Actions ([`.github/workflows/ci.yml`](file:///.github/workflows/ci.yml)) runs on every push and pull request across Windows and Ubuntu on Python 3.11 and 3.12.

---

## 📄 License
This project is licensed under the MIT License.
