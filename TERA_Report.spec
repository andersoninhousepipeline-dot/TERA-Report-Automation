# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for TERA Report Generator
# Build with:  pyinstaller TERA_Report.spec --clean

block_cipher = None

a = Analysis(
    ['tera_report_generator.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Bundled fonts (full TTF files)
        ('fonts',         'fonts'),
        # App icon
        ('tera_icon.png', '.'),
        ('tera_icon.ico', '.'),
    ],
    hiddenimports=[
        # Excel reading
        'xlrd', 'xlrd.biffh', 'xlrd.book', 'xlrd.compdoc',
        'openpyxl', 'openpyxl.cell', 'openpyxl.worksheet',
        # PDF generation
        'reportlab', 'reportlab.pdfgen', 'reportlab.pdfgen.canvas',
        'reportlab.lib', 'reportlab.lib.colors', 'reportlab.lib.utils',
        'reportlab.lib.styles', 'reportlab.lib.enums',
        'reportlab.platypus', 'reportlab.pdfbase', 'reportlab.pdfbase.ttfonts',
        'reportlab.pdfbase.pdfmetrics',
        # PDF inspection / comparison
        'pdfplumber', 'pdfminer', 'pdfminer.high_level',
        'pdfminer.layout', 'pdfminer.converter',
        'pypdfium2',
        # Data / image
        'pandas', 'pandas.io.formats.style',
        'PIL', 'PIL.Image', 'PIL.ImageDraw', 'PIL.ImageFont',
        # Qt
        'PyQt6', 'PyQt6.QtWidgets', 'PyQt6.QtCore', 'PyQt6.QtGui',
        'PyQt6.QtPrintSupport',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'scipy', 'numpy.distutils'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TERA Report',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # no black console window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='tera_icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='TERA Report',
)
