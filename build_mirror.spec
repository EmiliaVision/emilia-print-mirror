# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Emilia Print Mirror
Build with: pyinstaller build_mirror.spec
"""

block_cipher = None

a = Analysis(
    ['src/mirror_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets/icon.svg', 'assets'),
    ],
    hiddenimports=[
        'win32print',
        'win32api',
        'win32con',
        'PyQt6.QtWidgets',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
        'PyQt6.QtSvg',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='EmiliaPrintMirror',
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
    icon=None,  # No icon file (uses embedded SVG)
    uac_admin=True,
)
