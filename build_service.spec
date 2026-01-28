# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Emilia Print Mirror Service
Build with: uv run pyinstaller build_service.spec

Key discovery: win32timezone is required for win32print.EnumJobs to work.
Without it, EnumJobs silently returns 0 jobs instead of the actual queue.
"""

block_cipher = None

a = Analysis(
    ['src/mirror_service.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        # Core pywin32 modules
        'win32print',
        'win32api',
        'win32con',
        'pywintypes',
        'pythoncom',
        # CRITICAL: win32timezone is required for EnumJobs to work correctly
        'win32timezone',
        # Windows service modules (for service mode)
        'win32event',
        'win32service',
        'win32serviceutil',
        'servicemanager',
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
    name='EmiliaMirrorService',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
    uac_admin=True,
)
