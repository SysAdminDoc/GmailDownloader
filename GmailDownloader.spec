from PyInstaller.building.build_main import Analysis, EXE, PYZ


analysis = Analysis(
    ['gmaildownloader.py'],
    pathex=['.'],
    binaries=[],
    datas=[('icon.png', '.')],
    hiddenimports=['anthropic'],
    hooksconfig={},
    runtime_hooks=[],
    # PDFium/Tesseract are optional source-mode integrations; the packaged app
    # uses its ImageMagick fallback instead of pulling their native stacks into
    # the base executable.
    excludes=['pypdfium2', 'pypdfium2_raw', 'pytesseract'],
    noarchive=False,
)
pyz = PYZ(analysis.pure)
exe = EXE(
    pyz,
    analysis.scripts,
    analysis.binaries,
    analysis.datas,
    [],
    name='GmailDownloader',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='icon.ico',
)
