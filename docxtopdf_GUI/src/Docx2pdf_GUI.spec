# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Docx2pdf_GUI.py'],
    pathex=[],
    binaries=[('C:/Python312/python312.dll', '.')],
    datas=[('C:/Python312/Lib/site-packages/tkinterdnd2/tkdnd', 'tkinterdnd2/tkdnd')],
    hiddenimports=[],
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
    name='Docx2pdf_GUI',
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
    icon=['Docx2PDF_logo.ico'],
)
