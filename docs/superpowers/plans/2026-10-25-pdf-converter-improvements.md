# PDF-Converter Comprehensive Improvement Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Transform PDF-Converter into a robust, secure, high-performance, and visually modern desktop application with non-blocking concurrency, cross-platform path handling, clean dependencies, and zero runtime crashes.

**Architecture:**
- Decouple GUI from conversion logic: introduce `AppPaths` for platform-safe temporary and persistent storage, and `ConversionWorker` running on `QThread` for non-blocking UI.
- Modularize and modernize `PDFConverter` using `pypdf` (with fallback to `PyPDF2`), safe COM automation guards, unified multi-file handling, and robust exception handling.
- Modernize PyQt6 GUI with responsive layouts (removing hardcoded 500px margins), a dynamic SVG/QIcon system for missing assets, full drag-and-drop file list feedback, and single-window stack navigation.

**Tech Stack:** Python 3.12, PyQt6, pypdf, pdf2docx, python-docx, python-pptx, openpyxl, pdfplumber, img2pdf, PyMuPDF (fitz), pywin32, pytest.

---

### Task 1: Repository Hygiene & Git Clean-up
- Clean `.gitignore` to ignore `.venv`, `__pycache__`, `*.pyc`, `dist`, `build`, and logs.
- Untrack committed `__pycache__` files from git.
- Prune `requirements.txt` to only the real essential dependencies.

### Task 2: Centralized Safe Paths & Logging (`app_paths.py` & `log.py`)
- Implement `app_paths.py` with:
  - `get_resource_path()` supporting dev and PyInstaller (`sys._MEIPASS`).
  - `get_temp_dir()` and session-isolated working directories in user's OS temp folder.
  - Safe directory cleanup.
- Refactor `log.py` to write to AppData/temp with rotating file handler and console handler.
- Unit tests for path resolution and directory creation.

### Task 3: Core Conversion Engine Overhaul (`pdfconvter.py`)
- Modernize `PDFConverter` to accept explicit input/output paths in addition to directory scanning.
- Upgrade to `pypdf` with fallback.
- Fix all unhandled exceptions and errors (`pywintypes` import, `UnboundLocalError` in merger and excel, fitz double-closing).
- Implement proper multi-file naming in `pdf2images()` and `pdf_to_pptx()`.
- Improve compression and table extraction.
- Unit tests for core conversions.

### Task 4: Background Concurrency (`worker.py`)
- Implement `ConversionWorker` subclassing `QThread` (or `QRunnable`).
- Emit `progress(int, str)`, `finished(bool, str, list)`, and `error(str)` signals.
- Ensure the Qt main thread never blocks during document conversions.

### Task 5: UI Modernization & Bug Fixes (`main.py` & `secondwindow.py`)
- Dynamic icon generator (`icons.py`) generating crisp SVG/QPixmap icons so the app has attractive icons without requiring external files.
- Fix `DragDropArea`: fix missing separator bug (`os.path.join`), add visual file list with remove actions and file size.
- Replace hardcoded margins (`margin: 4px 0px 2px 500px`) with responsive layout boxes.
- Fix ZIP export file path truncation.
- Implement "Back to Home" navigation and wire all top-bar buttons.
- Verification and end-to-end launch test.
