# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['pwa_extractor.py'],
    pathex=[],
    binaries=[],
    datas=[('App_Logo.ico', '.'), ('README.md', '.')],
    hiddenimports=['app', 'backend', 'PySide6.QtPdf', 'PySide6.QtPdfWidgets'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='pwa_extractor',
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
    icon=['App_Logo.ico'],
)
