# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['word2md\\gui.py'],
    pathex=[],
    binaries=[],
    datas=[('word2md', 'word2md')],
    hiddenimports=['word2md', 'word2md.converter', 'word2md.gui', 'word2md.scanner', 'fitz', 'pptx', 'openpyxl'],
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
    name='OmnMarkdown',
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
)
