# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Spooler Queue Copy Mirror App
Build with: pyinstaller build_mirror.spec
"""

block_cipher = None

a = Analysis(
    ['src/mirror_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'win32print',
        'win32api',
        'win32con',
        'PyQt6.QtWidgets',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
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
    name='SpoolerQueueCopy',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI application
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon here: icon='icon.ico'
    uac_admin=True,  # Request administrator privileges
)
