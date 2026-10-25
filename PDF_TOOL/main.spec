# -*- mode: python ; coding: utf-8 -*-
import os

datas = []
if os.path.exists("resources"):
    datas.append(("resources", "resources"))
if os.path.exists("images"):
    datas.append(("images", "images"))
if os.path.exists("wkhtmltopdf"):
    datas.append(("wkhtmltopdf", "wkhtmltopdf"))

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'pypdf',
        'pdf2docx',
        'pdfplumber',
        'fitz',
        'img2pdf',
        'openpyxl',
        'pptx',
        'pikepdf',
        'win32com.client',
        'pythoncom'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'scipy', 'tkinter', 'sqlalchemy'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='PDF Converter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="images/1.png" if os.path.exists("images/1.png") else None
)
